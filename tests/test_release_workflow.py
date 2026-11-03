"""在仓库外执行Actions临时Python脚本，验证发布标签检查。"""
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest


class ReleaseWorkflowTests(unittest.TestCase):
    def run_validation(self, tag):
        root = Path(__file__).resolve().parents[1]
        workflow = (root / '.github/workflows/release.yml').read_text(encoding='utf-8')
        block = workflow.split('        shell: python\n        run: |\n', 1)[1].split('      - ', 1)[0]
        script = '\n'.join(line[10:] for line in block.splitlines())
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'action_step.py'
            path.write_text(script, encoding='utf-8')
            env = dict(os.environ, GITHUB_WORKSPACE=str(root), GITHUB_REF_NAME=tag)
            env.pop('PYTHONPATH', None)
            return subprocess.run([sys.executable, '-I', str(path)], cwd=root,
                                  env=env, capture_output=True, text=True)

    def test_valid_tag_from_temporary_script(self):
        from bill_splitter.version import VERSION
        result = self.run_validation('v' + VERSION)
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_mismatched_tag_is_rejected(self):
        result = self.run_validation('v999.0.0')
        self.assertNotEqual(result.returncode, 0)
        self.assertIn('Tag must match VERSION', result.stderr)
