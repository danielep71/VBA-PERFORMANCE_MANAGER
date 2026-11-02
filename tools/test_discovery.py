"""Every tools/test_*.py module runs under unittest discovery (#81).

CI runs ``python -m unittest discover -s tools -p "test_*.py"``. These tests
prove that discovery reaches every test module in tools/, so a new test file
runs without being wired into tools/vba_lint.py or a workflow, and that a
module with no test cases cannot pass unnoticed.
"""
from __future__ import annotations

import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

TOOLS = Path(__file__).resolve().parent


def _discovered_modules(suite: unittest.TestSuite) -> set[str]:
    modules: set[str] = set()
    for item in suite:
        if isinstance(item, unittest.TestSuite):
            modules |= _discovered_modules(item)
        else:
            modules.add(type(item).__module__)
    return modules


class DiscoveryTests(unittest.TestCase):
    def test_every_test_module_contributes_tests(self) -> None:
        expected = {path.stem for path in TOOLS.glob("test_*.py")}
        suite = unittest.defaultTestLoader.discover(str(TOOLS), pattern="test_*.py")
        discovered = _discovered_modules(suite)
        self.assertEqual(sorted(expected - discovered), [],
                         "test modules without discoverable unittest cases")
        failed_imports = sorted(m for m in discovered if m == "unittest.loader")
        self.assertEqual(failed_imports, [], "a test module failed to import")

    def test_an_unwired_new_test_file_is_discovered_and_run(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            folder = Path(td)
            (folder / "test_unwired.py").write_text(
                "import unittest\n\n"
                "class Unwired(unittest.TestCase):\n"
                "    def test_runs(self):\n"
                "        self.fail('unwired test ran')\n",
                encoding="utf-8",
            )
            proc = subprocess.run(
                [sys.executable, "-m", "unittest", "discover", "-s", td, "-p", "test_*.py"],
                capture_output=True, text=True, encoding="utf-8", errors="replace",
            )
        self.assertNotEqual(proc.returncode, 0)
        self.assertIn("unwired test ran", proc.stderr)
        self.assertIn("Ran 1 test", proc.stderr)


if __name__ == "__main__":
    unittest.main(verbosity=2)
