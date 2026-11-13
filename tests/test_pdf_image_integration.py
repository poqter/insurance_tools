"""Pure source-level regression checks; no production secrets or customer files."""
import ast
from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[1]
TREE = ast.parse((ROOT / "app.py").read_text(encoding="utf-8"))


def assignment(name):
    return next(n.value for n in TREE.body if isinstance(n, ast.Assign)
                and any(isinstance(t, ast.Name) and t.id == name for t in n.targets))


class PdfToolIntegrationTests(unittest.TestCase):
    def test_entry_and_icon_registered(self):
        apps = assignment("APP_DEFINITIONS")
        self.assertIn("pdf_image_tools", [ast.literal_eval(k) for k in apps.keys])
        icons = ast.literal_eval(assignment("HOME_ICONS"))
        self.assertTrue(icons["pdf_image_tools"].startswith("<svg"))
        self.assertEqual(len(apps.keys), 13)

    def test_all_existing_roles_receive_only_new_tool(self):
        permissions = ast.literal_eval(assignment("USER_PERMISSIONS"))
        self.assertEqual(set(permissions), {"Admin", "Manager1", "Basic", "Crew", "Dream"})
        self.assertTrue(all(p["pdf_image_tools"] for p in permissions.values()))
        self.assertFalse(permissions["Basic"]["commission_calculator"])
        self.assertTrue(permissions["Admin"]["commission_calculator"])

    def test_local_assets_and_no_remote_document_submission(self):
        root = ROOT / "modules/pdf_tools"
        for name in ["index.html", "app.js", "style.css", "vendor/pdf-lib.min.js",
                     "vendor/pdf.min.mjs", "vendor/pdf.worker.min.mjs", "vendor/jszip.min.js"]:
            self.assertTrue((root / name).is_file(), name)
        script = (root / "app.js").read_text()
        self.assertNotIn("streamlit:setComponentValue", script)
        self.assertNotIn("localStorage", script)
        self.assertNotIn("sessionStorage", script)
        self.assertIn('"streamlit:setFrameHeight"', script)


if __name__ == "__main__":
    unittest.main()
