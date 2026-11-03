import hashlib
import importlib.util
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile

REPO = Path(__file__).resolve().parents[2]
SCRIPT = REPO / 'scripts/Maintain-Artifacts.py'
spec = importlib.util.spec_from_file_location('maintenance', SCRIPT)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


class MaintenanceTests(unittest.TestCase):
    def test_archive_then_remove_only_selected_artifacts(self):
        parent = REPO / 'artifacts/maintenance-policy-tests'
        parent.mkdir(exist_ok=True)
        with tempfile.TemporaryDirectory(dir=parent) as temp:
            root = Path(temp)
            (root / 'scripts').mkdir()
            shutil.copyfile(SCRIPT, root / 'scripts/Maintain-Artifacts.py')
            artifacts = root / 'artifacts'
            for index in range(3):
                run = artifacts / ('validation-' + str(index))
                (run / 'source/apps').mkdir(parents=True)
                (run / 'source/apps/code.cs').write_text('historical source')
                (run / 'results.json').write_text('evidence')
                os.utime(run, (1000 + index, 1000 + index))
            old = artifacts / 'validation-0/source'
            (old / 'artifacts').mkdir()
            (old / 'artifacts/old.exe').write_bytes(b'inert generated output')
            (old / 'apps/input.dll').write_bytes(b'historical binary fixture')
            (old / 'apps/key.pfx').write_bytes(b'inert secret fixture')
            (artifacts / 'refs').mkdir()
            (artifacts / 'refs/dependency.dll').write_bytes(b'required')
            (artifacts / 'scene-contracts-x64').mkdir()
            stable = artifacts / 'scene-contracts-x64/SceneContractTests.exe'
            stable.write_bytes(b'stable runner fixture')
            outside = root / 'outside.txt'
            outside.write_text('must remain')
            result = subprocess.run([sys.executable, str(root / 'scripts/Maintain-Artifacts.py'), '--apply'], capture_output=True, text=True)
            self.assertEqual(result.returncode, 0, result.stderr)
            audit = artifacts / 'maintenance-20261103'
            plan = json.loads((audit / 'cleanup-plan.json').read_text())
            with zipfile.ZipFile(audit / 'historical-source-evidence.zip') as archive:
                for item in plan['files']:
                    if item['action'] == 'archive':
                        data = archive.read('objects/' + item['sha256'])
                        self.assertEqual(hashlib.sha256(data).hexdigest(), item['sha256'])
            self.assertFalse((old / 'apps/code.cs').exists())
            self.assertFalse((old / 'artifacts/old.exe').exists())
            self.assertTrue((old / 'apps/input.dll').exists())
            self.assertTrue((old / 'apps/key.pfx').exists())
            self.assertTrue((artifacts / 'validation-0/results.json').exists())
            self.assertTrue((artifacts / 'validation-1/source/apps/code.cs').exists())
            self.assertTrue((artifacts / 'validation-2/source/apps/code.cs').exists())
            self.assertEqual(stable.read_bytes(), b'stable runner fixture')
            self.assertEqual(outside.read_text(), 'must remain')
            self.assertTrue((artifacts / 'refs/dependency.dll').exists())

    def test_boundary_refuses_root_and_sibling(self):
        root = (REPO / 'artifacts').resolve()
        with self.assertRaises(ValueError):
            module.confined(root, root)
        with self.assertRaises(ValueError):
            module.confined(root.parent / (root.name + '-sibling'), root)


if __name__ == '__main__':
    unittest.main()
