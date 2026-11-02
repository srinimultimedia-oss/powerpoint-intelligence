"""Inert, bounded fixtures for the offline repair security boundary; no Office."""
import contextlib
import hashlib
import importlib.util
import io
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest import mock
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / 'scripts'))
SPEC = importlib.util.spec_from_file_location('repair_alignment', ROOT / 'scripts/Repair-TemplateAlignment.py')
REPAIR = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = REPAIR
SPEC.loader.exec_module(REPAIR)
SLIDE = b'''<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:bodyPr/><a:p><a:r><a:t>Original text</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:transition/><p:timing/></p:sld>'''


class RepairSecurityTests(unittest.TestCase):
    def setUp(self):
        parent = ROOT / 'artifacts/security-remediation-tests'
        parent.mkdir(parents=True, exist_ok=True)
        self.temp = tempfile.TemporaryDirectory(prefix='ppi-repair-security-', dir=parent)
        self.addCleanup(self.clean_temp)
        self.root = Path(self.temp.name)
        self.source = self.root / 'source.pptx'
        self.output = self.root / 'repaired.pptx'
        self.make_source()

    def clean_temp(self):
        target = Path(self.temp.name).resolve()
        parent = (ROOT / 'artifacts/security-remediation-tests').resolve()
        if not target.is_relative_to(parent) or not target.name.startswith('ppi-repair-security-'):
            raise ValueError('Unexpected test cleanup target')
        self.temp.cleanup()

    def make_source(self, slide=SLIDE, extra=()):
        with ZipFile(self.source, 'w') as archive:
            archive.writestr('ppt/slides/slide1.xml', slide)
            archive.writestr('docProps/unchanged.bin', b'owned fixture\x00')
            for name, content in extra:
                entry = ZipInfo(name)
                entry.filename = name  # Preserve the hostile raw name on Windows too.
                archive.writestr(entry, content)

    def run_repair(self, **kwargs):
        with contextlib.redirect_stdout(io.StringIO()):
            return REPAIR.repair(self.source, self.output, **kwargs)

    def require_clean(self):
        self.assertFalse(self.output.exists())
        self.assertFalse(self.output.with_suffix('.validation.json').exists())
        self.assertFalse(list(self.root.glob('.ppi-repair-*')))

    def test_preserves_text_source_untouched_parts_and_native_effects(self):
        original = self.source.read_bytes()
        self.run_repair()
        self.assertEqual(original, self.source.read_bytes())
        with ZipFile(self.source) as before, ZipFile(self.output) as after:
            self.assertEqual(before.namelist(), after.namelist())
            self.assertEqual(before.read('docProps/unchanged.bin'), after.read('docProps/unchanged.bin'))
            old = REPAIR.E.fromstring(before.read('ppt/slides/slide1.xml'))
            new = REPAIR.E.fromstring(after.read('ppt/slides/slide1.xml'))
            self.assertEqual(old.xpath('//a:t/text()', namespaces=REPAIR.NS), new.xpath('//a:t/text()', namespaces=REPAIR.NS))
            for tag in ('timing', 'transition'):
                self.assertEqual(REPAIR.E.tostring(old.find('p:' + tag, REPAIR.NS)), REPAIR.E.tostring(new.find('p:' + tag, REPAIR.NS)))
            self.assertEqual(new.xpath('//a:pPr/@algn', namespaces=REPAIR.NS), ['ctr'])
            self.assertEqual(new.xpath('//a:bodyPr/@anchor', namespaces=REPAIR.NS), ['ctr'])
        self.assertFalse(json.loads(self.output.with_suffix('.validation.json').read_text())['powerPointRenderingVerified'])

    def test_existing_deck_not_overwritten(self):
        self.output.write_bytes(b'sentinel')
        with self.assertRaises(ValueError): self.run_repair()
        self.assertEqual(self.output.read_bytes(), b'sentinel')

    def test_existing_sidecar_not_overwritten(self):
        sidecar = self.output.with_suffix('.validation.json')
        sidecar.write_text('sentinel')
        with self.assertRaises(ValueError): self.run_repair()
        self.assertEqual(sidecar.read_text(), 'sentinel')
        self.assertFalse(self.output.exists())

    def test_input_output_identity_refused(self):
        original = self.source.read_bytes()
        with self.assertRaises(ValueError): REPAIR.repair(self.source, self.source)
        self.assertEqual(original, self.source.read_bytes())

    def test_duplicate_members_rejected(self):
        with ZipFile(self.source, 'a') as archive:
            with self.assertWarns(UserWarning): archive.writestr('ppt/slides/slide1.xml', SLIDE)
        with self.assertRaises(ValueError): self.run_repair()
        self.require_clean()

    def test_member_count_budget(self):
        with self.assertRaises(ValueError): self.run_repair(limits=REPAIR.Limits(max_members=1))
        self.require_clean()

    def test_member_bytes_budget(self):
        with self.assertRaises(ValueError): self.run_repair(limits=REPAIR.Limits(max_member_bytes=16))
        self.require_clean()

    def test_total_inflated_budget(self):
        with self.assertRaises(ValueError): self.run_repair(limits=REPAIR.Limits(max_archive_bytes=32))
        self.require_clean()

    def test_xml_bytes_budget(self):
        with self.assertRaises(ValueError): self.run_repair(limits=REPAIR.Limits(max_xml_bytes=32))
        self.require_clean()

    def test_xml_node_budget(self):
        with self.assertRaises(ValueError): self.run_repair(limits=REPAIR.Limits(max_xml_nodes=3))
        self.require_clean()

    def test_xml_depth_budget(self):
        with self.assertRaises(ValueError): self.run_repair(limits=REPAIR.Limits(max_xml_depth=2))
        self.require_clean()

    def test_actual_processed_bytes_budget(self):
        with self.assertRaises(ValueError): self.run_repair(limits=REPAIR.Limits(max_processed_bytes=16))
        self.require_clean()

    def test_small_compressed_bomb_rejected_before_inflating(self):
        with ZipFile(self.source, 'w', compression=ZIP_DEFLATED) as archive:
            archive.writestr('ppt/slides/slide1.xml', SLIDE)
            archive.writestr('media/repeated.bin', b'A' * 8192)
        self.assertLess(self.source.stat().st_size, 2048)
        with self.assertRaises(ValueError): self.run_repair(limits=REPAIR.Limits(max_member_bytes=2048))
        self.require_clean()

    def test_source_mutation_prevents_publication(self):
        original = REPAIR.SAFETY.validate_repair
        def mutate_after_validation(source, staged, budget):
            original(source, staged, budget)
            self.source.write_bytes(b'changed owned fixture')
        with mock.patch.object(REPAIR.SAFETY, 'validate_repair', side_effect=mutate_after_validation):
            with self.assertRaises(ValueError): self.run_repair()
        self.require_clean()

    def test_box_label_fit_geometry_matches_existing_alignment_contract(self):
        slide = REPAIR.E.fromstring(SLIDE)
        tree = slide.find('p:cSld/p:spTree', REPAIR.NS)
        label = tree[0]
        p, a = '{' + REPAIR.NS['p'] + '}', REPAIR.A
        for shape, x, y, width, height in ((label, 20, 20, 10, 10), (REPAIR.E.SubElement(tree, p + 'sp'), 0, 0, 100, 100)):
            props = REPAIR.E.SubElement(shape, p + 'spPr')
            transform = REPAIR.E.SubElement(props, a + 'xfrm')
            REPAIR.E.SubElement(transform, a + 'off', x=str(x), y=str(y))
            REPAIR.E.SubElement(transform, a + 'ext', cx=str(width), cy=str(height))
            if shape is not label:
                REPAIR.E.SubElement(props, a + 'solidFill')
                REPAIR.E.SubElement(props, a + 'prstGeom', prst='rect')
        self.make_source(REPAIR.E.tostring(slide))
        self.run_repair()
        with ZipFile(self.output) as archive:
            repaired = REPAIR.E.fromstring(archive.read('ppt/slides/slide1.xml'))
        self.assertEqual(REPAIR.bounds(repaired.find('p:cSld/p:spTree/p:sp', REPAIR.NS)), (6, 12, 88, 76))

    def test_work_budget(self):
        with self.assertRaises(ValueError): self.run_repair(limits=REPAIR.Limits(max_work=1))
        self.require_clean()

    def test_elapsed_budget(self):
        budget = REPAIR.Budget(REPAIR.Limits(max_seconds=1))
        budget.started -= 2
        with self.assertRaises(ValueError): budget.check()

    def test_dtd_rejected_without_fetching(self):
        self.make_source(b'<!DOCTYPE sld [<!ENTITY x SYSTEM "file:///never-read">]>' + SLIDE)
        with self.assertRaises(ValueError): self.run_repair()
        self.require_clean()

    def test_unsafe_member_name_rejected_without_extraction(self):
        for name in ('../outside', '/absolute', 'C:/drive', 'ppt\\outside', 'ppt/./x'):
            with self.subTest(name=name):
                self.make_source(extra=[(name, b'owned')])
                with self.assertRaises(ValueError): self.run_repair()
                self.require_clean()

    def test_validation_failure_never_publishes(self):
        with mock.patch.object(REPAIR.SAFETY, 'validate_repair', side_effect=ValueError('injected validation failure')):
            with self.assertRaises(ValueError): self.run_repair()
        self.require_clean()

    def test_second_publication_failure_rolls_back_own_deck(self):
        original = REPAIR.SAFETY.publish_new
        def fail_sidecar(staged, target):
            if Path(target).suffix == '.json': raise OSError('injected publication failure')
            return original(staged, target)
        with mock.patch.object(REPAIR.SAFETY, 'publish_new', side_effect=fail_sidecar):
            with self.assertRaises(OSError): self.run_repair()
        self.require_clean()

    def test_preservation_validation_remains_enabled_under_optimized_python(self):
        code = "import sys;sys.path.insert(0,r'%s');from RepairTemplateSafety import require;require(False,'sentinel')" % (ROOT / 'scripts')
        result = subprocess.run([sys.executable, '-B', '-O', '-c', code], capture_output=True, text=True)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn('ValueError: sentinel', result.stderr)


if __name__ == '__main__': unittest.main()
