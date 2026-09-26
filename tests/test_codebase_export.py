"""End-to-end checks for the portable codebase export command."""

import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "tools" / "export_codebase.py"


class CodebaseExportTests(unittest.TestCase):
    def test_export_includes_application_source_and_excludes_auxiliary_trees(self):
        """Application source stays exportable while docs, tests and tooling stay out."""
        with tempfile.TemporaryDirectory() as directory:
            project = Path(directory) / "project"
            project.mkdir()
            (project / "app.py").write_text("print('uygulama')\n", encoding="utf-8")
            (project / "requirements.txt").write_text("Flask==3.0.3\n", encoding="utf-8")
            (project / "templates").mkdir()
            (project / "templates" / "panel.html").write_text("<h1>Pano</h1>\n", encoding="utf-8")
            (project / "static").mkdir()
            (project / "static" / "stil.css").write_text("body {}\n", encoding="utf-8")
            (project / "docs").mkdir()
            (project / "docs" / "KILAVUZ.md").write_text(
                "# Kılavuz\n\n```python\nprint('örnek')\n```\n", encoding="utf-8"
            )
            (project / "tests").mkdir()
            (project / "tests" / "test_app.py").write_text("assert True\n", encoding="utf-8")
            (project / "artifact_work").mkdir()
            (project / "artifact_work" / "yardimci.py").write_text("legacy = True\n", encoding="utf-8")
            (project / "data").mkdir()
            (project / "data" / "kartlar.xlsx").write_bytes(b"private-data")
            (project / ".venv").mkdir()
            (project / ".venv" / "gizli.py").write_text("secret = 1\n", encoding="utf-8")
            (project / "outputs").mkdir()
            (project / "outputs" / "sonuc.txt").write_text("test output\n", encoding="utf-8")
            (project / ".investigation").mkdir()
            (project / ".investigation" / "eski_kod.py").write_text("legacy = True\n", encoding="utf-8")

            output = project / "codebase.md"
            completed = subprocess.run(
                [sys.executable, str(SCRIPT), "--root", str(project), "--output", str(output)],
                capture_output=True,
                encoding="utf-8",
                errors="replace",
                check=False,
            )

            self.assertEqual(completed.returncode, 0, completed.stderr)

            report = output.read_text(encoding="utf-8")
            self.assertIn("# PDGM Codebase Export", report)
            self.assertIn("`app.py`", report)
            self.assertIn("`requirements.txt`", report)
            self.assertIn("`templates/panel.html`", report)
            self.assertIn("`static/stil.css`", report)
            self.assertIn("print('uygulama')", report)
            self.assertNotIn("kartlar.xlsx", report)
            self.assertNotIn("gizli.py", report)
            self.assertNotIn("sonuc.txt", report)
            self.assertNotIn("eski_kod.py", report)
            self.assertNotIn("KILAVUZ.md", report)
            self.assertNotIn("test_app.py", report)
            self.assertNotIn("yardimci.py", report)
            self.assertNotIn("codebase.md`", report)
            self.assertIn("4 dosya", completed.stdout)


if __name__ == "__main__":
    unittest.main()
