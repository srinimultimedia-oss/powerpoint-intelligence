"""Conservative artifact-only cleanup; defaults to planning, requires --apply.

Historical source/evidence is content-addressed in a verified ZIP before its
expanded copy is removed. Executables are never archived or restored.
"""
import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import stat
import zipfile


def digest(path):
    h = hashlib.sha256()
    with path.open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def linked(path):
    return bool(getattr(path.lstat(), 'st_file_attributes', 0) & 0x400) or path.is_symlink()


def confined(path, root):
    absolute = path.resolve(strict=False)
    if not absolute.is_relative_to(root) or absolute == root:
        raise ValueError('Cleanup target outside artifacts: ' + str(path))
    current = path
    while current != root:
        if linked(current):
            raise ValueError('Reparse cleanup target refused: ' + str(path))
        current = current.parent
    return absolute


def files_under(directory):
    for base, dirs, names in os.walk(directory, followlinks=False):
        dirs[:] = [d for d in dirs if not linked(Path(base, d))]
        for name in names:
            file = Path(base, name)
            if not linked(file):
                yield file


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--apply', action='store_true')
    parser.add_argument('--keep', type=int, default=2)
    parser.add_argument('--audit-name', default='maintenance-20261103')
    args = parser.parse_args()
    if args.keep < 2 or not re.fullmatch(r'maintenance-[A-Za-z0-9-]+', args.audit_name):
        raise ValueError('At least two recent runs and an owned audit name are required.')
    repo = Path(__file__).resolve().parent.parent
    root = (repo / 'artifacts').resolve()
    audit = root / args.audit_name
    audit.mkdir(exist_ok=True)
    if linked(root) or linked(audit):
        raise ValueError('Artifact/audit reparse root refused.')
    selected = []
    retained = []

    def old_group(parent, pattern, relative_child=None, keep=None):
        groups = sorted((p for p in parent.glob(pattern) if p.is_dir() and not linked(p)), key=lambda p: p.stat().st_mtime, reverse=True)
        limit = args.keep if keep is None else keep
        retained.extend(str(p.relative_to(root)) for p in groups[:limit])
        for group in groups[limit:]:
            target = group / relative_child if relative_child else group
            if target.is_dir():
                selected.append((target, 'snapshot' if relative_child else 'generated'))

    old_group(root, 'validation-*', 'source')
    old_group(root, 'easySetup-build-*')
    old_group(root / 'piidx', 'suite-*')
    old_group(root / 'demo-contracts', 'piidx-ui-*')
    old_group(root / 'source-fingerprint-tests', '*')
    old_group(root / 'piidx-recovery-tests', '*')
    binary = {'.exe', '.dll', '.pdb', '.obj', '.lib', '.nupkg', '.zip', '.db', '.sqlite', '.sqlite3', '.wal', '.shm'}
    evidence = {'.json', '.log', '.txt', '.csv', '.xml', '.md', '.png', '.jpg', '.svg', '.pptx', '.potx', '.config', '.manifest', '.vsto'}
    secret = {'.pfx', '.p12', '.snk', '.key', '.pem'}
    generated_parts = {'artifacts', 'node_modules', 'bin', 'obj', '__pycache__'}
    installed = root / 'installed-demo-receipt.json'
    active_parent = None
    if installed.exists():
        item = json.loads(installed.read_text(encoding='utf-8-sig'))
        active_parent = Path(item.get('Manifest', '').split('|')[0]).parent.resolve()
    plan = []
    for target, category in selected:
        confined(target, root)
        if active_parent and (active_parent == target or active_parent.is_relative_to(target)):
            retained.append(str(target.relative_to(root)) + ': active installation')
            continue
        for file in files_under(target):
            # Target was confined above; files_under never follows reparse entries.
            # Recheck the full path immediately before archival/deletion below.
            if file.suffix.lower() in secret or file.name.startswith('.env') or file.name == 'previous-setup.exe':
                retained.append(str(file.relative_to(root)) + ': signing/rollback material')
                continue
            relative = file.relative_to(target)
            source_input = category == 'snapshot' and not any(p in generated_parts for p in relative.parts)
            if source_input and file.suffix.lower() in binary:
                retained.append(str(file.relative_to(root)) + ': historical binary source input')
                continue
            preserve = file.suffix.lower() not in binary and (source_input or file.suffix.lower() in evidence or file.suffix.lower() in {'.cs', '.ps1', '.py'})
            info = file.stat()
            plan.append({'path': str(file.relative_to(root)), 'bytes': info.st_size, 'mtime_ns': info.st_mtime_ns, 'sha256': digest(file), 'action': 'archive' if preserve else 'discard-generated'})
            if len(plan) % 2000 == 0:
                print('Inventoried ' + str(len(plan)) + ' candidate files', flush=True)
    document = {'root': str(root), 'retained': retained, 'protected': ['stable runners/receipts', 'refs', 'merge backups', 'scanner evidence', 'current installer outside artifacts', 'native acceptance folders outside selected groups'], 'files': plan}
    plan_file = audit / 'cleanup-plan.json'
    if plan_file.exists():
        raise ValueError('Audit plan already exists; choose a fresh --audit-name.')
    plan_file.write_text(json.dumps(document, indent=2), encoding='utf-8')
    print(json.dumps({'planned_files': len(plan), 'bytes': sum(p['bytes'] for p in plan), 'plan': str(plan_file), 'apply': args.apply}), flush=True)
    if not args.apply:
        return
    archive = audit / 'historical-source-evidence.zip'
    objects = set()
    with zipfile.ZipFile(archive, 'x', compression=zipfile.ZIP_DEFLATED, compresslevel=6) as package:
        for index, item in enumerate(plan):
            if item['action'] != 'archive' or item['sha256'] in objects:
                continue
            file = root / item['path']
            confined(file, root)
            if digest(file) != item['sha256']:
                raise ValueError('Artifact changed before archival; cleanup aborted.')
            package.write(file, 'objects/' + item['sha256'])
            objects.add(item['sha256'])
            if index and index % 2000 == 0:
                print('Archiving item ' + str(index), flush=True)
        package.writestr('manifest.json', json.dumps(document, separators=(',', ':')))
    with zipfile.ZipFile(archive) as package:
        for entry in package.infolist():
            if not entry.filename.startswith('objects/'):
                continue
            h = hashlib.sha256()
            with package.open(entry) as stream:
                for block in iter(lambda: stream.read(1024 * 1024), b''):
                    h.update(block)
            if h.hexdigest() != entry.filename.split('/')[-1]:
                raise ValueError('Archive verification failed; nothing deleted.')
    removed = []; skipped = []
    for item in plan:
        file = root / item['path']
        confined(file, root)
        if digest(file) != item['sha256']:
            skipped.append({'path': item['path'], 'reason': 'changed after plan'})
            continue
        try:
            file.unlink()
            removed.append(item)
        except OSError as error:
            skipped.append({'path': item['path'], 'reason': type(error).__name__})
    # Empty directory removal is nonrecursive and restricted to planned subtrees.
    for target, _ in selected:
        for base, dirs, _ in os.walk(target, topdown=False, followlinks=False):
            folder = Path(base)
            if folder.exists() and not linked(folder):
                confined(folder, root)
                try:
                    folder.rmdir()
                except OSError:
                    pass
    freed = sum(p['bytes'] for p in removed)
    result = {'removed_files': len(removed), 'removed_bytes': freed, 'archive_bytes': archive.stat().st_size, 'net_reclaimed_bytes': freed - archive.stat().st_size - plan_file.stat().st_size, 'archived_unique_objects': len(objects), 'archive_sha256': digest(archive), 'skipped': skipped, 'cleanup_scope': str(root)}
    (audit / 'cleanup-result.json').write_text(json.dumps(result, indent=2), encoding='utf-8')
    print(json.dumps(result), flush=True)


if __name__ == '__main__':
    main()
