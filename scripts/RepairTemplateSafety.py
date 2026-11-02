"""Bounded I/O and publication for the offline alignment utility; no Office."""
from dataclasses import dataclass, fields
import hashlib
import io
import json
import math
import os
from pathlib import Path
import re
import struct
import tempfile
import time
from zipfile import ZipFile
from lxml import etree as E

NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}


def require(valid, message):
    if not valid:
        raise ValueError(message)


@dataclass(frozen=True)
class Limits:
    max_file_bytes: int = 512 * 1024 * 1024
    max_members: int = 10000
    max_directory_bytes: int = 64 * 1024 * 1024
    max_member_bytes: int = 64 * 1024 * 1024
    max_archive_bytes: int = 512 * 1024 * 1024
    max_processed_bytes: int = 4 * 1024 * 1024 * 1024
    max_xml_bytes: int = 40 * 1024 * 1024
    max_xml_nodes: int = 1000000
    max_xml_depth: int = 128
    max_work: int = 25000000
    max_seconds: float = 120

    def __post_init__(self):
        for field in fields(self):
            value = getattr(self, field.name)
            if field.name == 'max_seconds':
                require(isinstance(value, (int, float)) and not isinstance(value, bool)
                        and math.isfinite(value) and value > 0, 'Invalid elapsed-time budget.')
            else:
                require(type(value) is int and value > 0, 'Invalid resource budget.')


class Budget:
    def __init__(self, limits):
        self.limits = limits
        self.started = time.monotonic()
        self.processed = 0
        self.work = 0

    def check(self, work=0, processed=0):
        self.work += work
        self.processed += processed
        require(self.work <= self.limits.max_work, 'Presentation work budget exceeded.')
        require(self.processed <= self.limits.max_processed_bytes, 'Processed-byte budget exceeded.')
        require(time.monotonic() - self.started <= self.limits.max_seconds, 'Repair time budget exceeded.')


def fingerprint(path, budget):
    result = hashlib.sha256()
    with path.open('rb') as source:
        while True:
            block = source.read(65536)
            if not block:
                break
            budget.check(processed=len(block))
            result.update(block)
    return result.hexdigest()


def preflight_zip(path, budget):
    """Bound central-directory allocation before constructing ZipFile."""
    limits = budget.limits
    size = path.stat().st_size
    require(size <= limits.max_file_bytes, 'Presentation file exceeds byte budget.')
    with path.open('rb') as source:
        source.seek(max(0, size - 65557))
        tail = source.read(65557)
        end = tail.rfind(b'PK\x05\x06')
        record = None
        while end >= 0:
            if end + 22 <= len(tail):
                candidate = struct.unpack_from('<4s4H2LH', tail, end)
                if end + 22 + candidate[-1] == len(tail):
                    record = candidate
                    break
            end = tail.rfind(b'PK\x05\x06', 0, end)
        require(record is not None, 'Invalid ZIP end record.')
        _, disk, directory_disk, on_disk, count, directory_size, offset, _ = record
        require(disk == directory_disk == 0 and on_disk == count, 'Split ZIP packages are unsupported.')
        require(count <= limits.max_members, 'ZIP member count exceeds budget.')
        require(directory_size <= limits.max_directory_bytes and offset + directory_size <= size,
                'ZIP central directory exceeds bounds.')
        source.seek(offset)
        stop = offset + directory_size
        scanned = 0
        while source.tell() < stop:
            budget.check(work=1)
            header = source.read(46)
            require(len(header) == 46 and header[:4] == b'PK\x01\x02', 'Invalid ZIP directory record.')
            scanned += 1
            require(scanned <= limits.max_members, 'Actual ZIP entry count exceeds budget.')
            name, extra, comment = struct.unpack_from('<3H', header, 28)
            source.seek(name + extra + comment, 1)
            require(source.tell() <= stop, 'Truncated ZIP directory.')
        require(scanned == count and source.tell() == stop, 'ZIP directory count differs.')


def inspect_archive(archive, budget):
    members = archive.infolist()
    require(len(members) <= budget.limits.max_members, 'ZIP member count exceeds budget.')
    seen = set()
    total = 0
    for item in members:
        budget.check(work=1)
        name = item.orig_filename
        require(name == item.filename and not any(ord(c) < 32 for c in name), 'Ambiguous raw ZIP member name.')
        segments = name.rstrip('/').split('/')
        require(name and not name.startswith('/') and '\\' not in name and ':' not in name
                and all(segment not in ('', '.', '..') for segment in segments), 'Unsafe package member name.')
        require(name.casefold() not in seen, 'Duplicate or ambiguous package member.')
        seen.add(name.casefold())
        require(not item.flag_bits & 1, 'Encrypted package member is unsupported.')
        require(0 <= item.file_size <= budget.limits.max_member_bytes, 'ZIP member exceeds byte budget.')
        total += item.file_size
        require(total <= budget.limits.max_archive_bytes, 'Expanded package exceeds byte budget.')
    return members


def read_member(archive, item, budget):
    parts, size = [], 0
    with archive.open(item) as source:
        while True:
            block = source.read(min(65536, budget.limits.max_member_bytes - size + 1))
            if not block:
                break
            size += len(block)
            require(size <= budget.limits.max_member_bytes, 'Actual ZIP member exceeds byte budget.')
            budget.check(processed=len(block))
            parts.append(block)
    require(size == item.file_size, 'ZIP member length differs.')
    return b''.join(parts)


def parse_xml(data, budget):
    require(len(data) <= budget.limits.max_xml_bytes, 'Slide XML exceeds byte budget.')
    budget.check()
    # Count while building, rather than discovering a node/depth overrun after
    # materializing the entire tree. The bounded tree is retained for editing.
    context = E.iterparse(io.BytesIO(data), events=('start', 'end', 'comment', 'pi'), resolve_entities=False,
                          no_network=True, load_dtd=False, huge_tree=False)
    depth = count = 0
    for event, element in context:
        if event == 'start':
            depth += 1
            count += 1
            require(count <= budget.limits.max_xml_nodes and depth <= budget.limits.max_xml_depth,
                    'Slide XML structure exceeds budget.')
        elif event == 'end':
            depth -= 1
        else:
            count += 1
            require(count <= budget.limits.max_xml_nodes, 'Slide XML structure exceeds budget.')
        budget.check(work=1)
    root = context.root
    require(root is not None and not root.getroottree().docinfo.doctype, 'Presentation DTDs are unsupported.')
    return root


def validate_repair(source, staged, budget):
    preflight_zip(source, budget)
    preflight_zip(staged, budget)
    with ZipFile(source) as before, ZipFile(staged) as after:
        members = inspect_archive(before, budget)
        inspect_archive(after, budget)
        require(before.namelist() == after.namelist(), 'Package members changed.')
        for item in members:
            old_data = read_member(before, item, budget)
            new_data = read_member(after, after.getinfo(item.filename), budget)
            if re.fullmatch(r'ppt/slides/slide\d+\.xml', item.filename):
                old, new = parse_xml(old_data, budget), parse_xml(new_data, budget)
                require(old.xpath('//a:t/text()', namespaces=NS) == new.xpath('//a:t/text()', namespaces=NS), 'Slide text changed.')
                for tag in ('timing', 'transition'):
                    x, y = old.find('p:' + tag, NS), new.find('p:' + tag, NS)
                    require((E.tostring(x) if x is not None else None) == (E.tostring(y) if y is not None else None), 'Native effects changed.')
                require(all(p.get('anchor') == 'ctr' for p in new.xpath('//p:txBody/a:bodyPr | //a:txBody/a:bodyPr', namespaces=NS)), 'Vertical alignment verification failed.')
                require(all(p.get('algn') == 'ctr' for p in new.xpath('//p:txBody/a:p/a:pPr | //a:txBody/a:p/a:pPr', namespaces=NS)), 'Text alignment verification failed.')
            else:
                require(old_data == new_data, 'An untouched package member changed.')


def publish_new(staged, target):
    """Publish a complete same-directory file without replacing a destination."""
    if os.name == 'nt':
        os.rename(staged, target)  # Windows fails rather than replacing an existing target.
    else:
        os.link(staged, target)
        staged.unlink()


def repair(source, output, transform, limits=None):
    source, output = Path(source).resolve(), Path(output).resolve()
    sidecar = output.with_suffix('.validation.json')
    require(source != output and source != sidecar and not output.exists() and not sidecar.exists(),
            'Choose new deck and validation paths; existing files are never overwritten.')
    budget = Budget(limits or Limits())
    preflight_zip(source, budget)
    original_hash = fingerprint(source, budget)
    output.parent.mkdir(parents=True, exist_ok=True)
    staged = staged_report = None
    published = False
    slides = fields_count = expanded = 0
    try:
        with tempfile.NamedTemporaryFile(prefix='.ppi-repair-', suffix='.pptx', dir=output.parent, delete=False) as temp:
            staged = Path(temp.name)
        with ZipFile(source) as zin, ZipFile(staged, 'w') as zout:
            for info in inspect_archive(zin, budget):
                data = read_member(zin, info, budget)
                if re.fullmatch(r'ppt/slides/slide\d+\.xml', info.filename):
                    data, count, fit = transform(data, budget)
                    slides += 1; fields_count += count; expanded += fit
                zout.writestr(info, data)
        validate_repair(source, staged, budget)
        require(fingerprint(source, budget) == original_hash, 'Input changed during repair; nothing published.')
        report = {'slides': slides, 'textFramesCentered': fields_count, 'singleLabelsFitted': expanded,
                  'inputUnchanged': True, 'textAndOtherPartsPreserved': True, 'powerPointRenderingVerified': False}
        with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', prefix='.ppi-repair-', suffix='.json', dir=output.parent, delete=False) as temp:
            staged_report = Path(temp.name)
            json.dump(report, temp, indent=2)
        budget.check()
        publish_new(staged, output)
        published = True
        publish_new(staged_report, sidecar)
    except BaseException:
        if published:
            output.unlink(missing_ok=True)
        raise
    finally:
        for temp in (staged, staged_report):
            if temp is not None:
                temp.unlink(missing_ok=True)
    print(json.dumps(report))
    return report
