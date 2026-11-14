import os
import sys
import unittest
import tempfile
from unittest.mock import patch, MagicMock

import openpyxl
from openpyxl.styles import Font

# Add Scripts folder to path
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS_DIR = os.path.join(REPO_ROOT, "Scripts")
if SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, SCRIPTS_DIR)

from update_inventory import update_excel, _clear_red_font
from daily import decide_git_push_action, _confirm_missing_slug_push, step_git_push


class TestInventoryUpdate(unittest.TestCase):
    def setUp(self):
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.excel_path = os.path.join(self.tmp_dir.name, "Test_Tubex.xlsx")

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Inventory"

        # Row 1: Title/Date
        ws.cell(row=1, column=1, value="Date:")
        ws.cell(row=1, column=10, value="=TODAY()")

        # Row 2: Headers
        headers = [
            "Item ID", "Category", "Item Name", "UOM",
            "Opening", "Received", "Issued", "Balance",
            "WIP", "Tube Equiv", "Status / Remarks"
        ]
        for c, h in enumerate(headers, 1):
            ws.cell(row=2, column=c, value=h)

        # Row 3: Restored item (was red, now back in ERP)
        ws.cell(row=3, column=1, value=1001)
        ws.cell(row=3, column=2, value="Cap")
        ws.cell(row=3, column=3, value="CAP 13MM")
        ws.cell(row=3, column=4, value="PCS")
        ws.cell(row=3, column=5, value=0.0)
        ws.cell(row=3, column=6, value=0.0)
        ws.cell(row=3, column=7, value=0.0)
        ws.cell(row=3, column=11, value="Not active in ERP")
        for col in range(1, 12):
            ws.cell(row=3, column=col).font = Font(name="Segoe UI", size=9, color="FFFF0000")

        # Row 4: Missing active item (slug, will be marked inactive)
        ws.cell(row=4, column=1, value=1002)
        ws.cell(row=4, column=2, value="Slug")
        ws.cell(row=4, column=3, value="SLUG 30X4.8")
        ws.cell(row=4, column=4, value="kg")
        ws.cell(row=4, column=5, value=100.0)
        ws.cell(row=4, column=6, value=50.0)
        ws.cell(row=4, column=7, value=25.0)
        # Give row 4 some legacy red styling to ensure it gets cleaned
        for col in range(1, 12):
            ws.cell(row=4, column=col).font = Font(name="Segoe UI", size=9, color="FFFF0000")

        # Row 5: Configured inactive item (ID 406 from custom_aliases.json)
        ws.cell(row=5, column=1, value=406)
        ws.cell(row=5, column=2, value="Lacquer")
        ws.cell(row=5, column=3, value="OLD LACQUER")
        ws.cell(row=5, column=4, value="kg")
        ws.cell(row=5, column=5, value=20.0)
        ws.cell(row=5, column=6, value=0.0)
        ws.cell(row=5, column=7, value=0.0)

        # Row 6: Regular active slug present in ERP
        ws.cell(row=6, column=1, value=1003)
        ws.cell(row=6, column=2, value="Slug")
        ws.cell(row=6, column=3, value="SLUG 25X4.2")
        ws.cell(row=6, column=4, value="kg")
        ws.cell(row=6, column=5, value=200.0)
        ws.cell(row=6, column=6, value=100.0)
        ws.cell(row=6, column=7, value=50.0)

        wb.save(self.excel_path)

    def tearDown(self):
        self.tmp_dir.cleanup()

    def test_clear_red_font_helper(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        for c in range(1, 12):
            ws.cell(row=1, column=c).font = Font(name="Arial", size=10, bold=True, color="FFFF0000")

        _clear_red_font(ws, 1)

        for c in range(1, 12):
            font = ws.cell(row=1, column=c).font
            self.assertEqual(font.color.rgb, "FF1A1A2E")
            self.assertEqual(font.name, "Arial")
            self.assertEqual(font.size, 10)
            self.assertTrue(font.bold)

    def test_update_excel_zeroes_and_marks_inactive_without_red(self):
        xls_items = {
            1001: {'name': 'CAP 13MM', 'opening': 500, 'inward': 200, 'out': 100, 'unit': 'PCS', 'balance': 600, 'category': 'Cap'},
            1003: {'name': 'SLUG 25X4.2', 'opening': 200, 'inward': 100, 'out': 50, 'unit': 'kg', 'balance': 250, 'category': 'Slug'},
        }

        updated, not_in_excel, missing_critical, missing_items, new_critical_slugs = update_excel(
            self.excel_path, xls_items, "01-11-2026 to 13-11-2026"
        )

        wb = openpyxl.load_workbook(self.excel_path)
        ws = wb['Inventory']

        # 1. Restored Item 1001 (Row 3): Col K cleared, values updated, red fonts cleared
        self.assertEqual(ws.cell(row=3, column=5).value, 500)
        self.assertEqual(ws.cell(row=3, column=6).value, 200)
        self.assertEqual(ws.cell(row=3, column=7).value, 100)
        self.assertIn(ws.cell(row=3, column=11).value, ("", None))
        for col in range(1, 12):
            font = ws.cell(row=3, column=col).font
            color_rgb = getattr(font.color, 'rgb', None) if (font and font.color) else None
            self.assertNotIn(color_rgb, ('FFFF0000', 'FF0000'))
            self.assertEqual(color_rgb, "FF1A1A2E")

        # 2. Missing Active Item 1002 (Row 4): Col K marked "Not active in ERP", E/F/G zeroed, NO RED FONT
        self.assertEqual(ws.cell(row=4, column=5).value, 0.0)
        self.assertEqual(ws.cell(row=4, column=6).value, 0.0)
        self.assertEqual(ws.cell(row=4, column=7).value, 0.0)
        self.assertEqual(ws.cell(row=4, column=11).value, "Not active in ERP")
        self.assertEqual(ws.cell(row=4, column=11).font.color.rgb, "FF1A1A2E")
        self.assertEqual(ws.cell(row=4, column=11).font.name, "Segoe UI")
        for col in range(1, 12):
            font = ws.cell(row=4, column=col).font
            color_rgb = getattr(font.color, 'rgb', None) if (font and font.color) else None
            self.assertNotIn(color_rgb, ('FFFF0000', 'FF0000'), f"Col {col} in row 4 must not be red!")

        # 3. Configured Inactive Item 406 (Row 5): Col K marked "Inactive (Not in use)", E/F/G zeroed, grey italic
        self.assertEqual(ws.cell(row=5, column=5).value, 0.0)
        self.assertEqual(ws.cell(row=5, column=6).value, 0.0)
        self.assertEqual(ws.cell(row=5, column=7).value, 0.0)
        self.assertEqual(ws.cell(row=5, column=11).value, "Inactive (Not in use)")
        for col in range(1, 12):
            font = ws.cell(row=5, column=col).font
            self.assertEqual(font.color.rgb, "FF888888")
            self.assertTrue(font.italic)

        # 4. Check return metrics
        self.assertEqual(len(missing_critical), 1)  # 1002 is Slug
        self.assertEqual(missing_critical[0][0], 1002)


class TestDailyGitPushGating(unittest.TestCase):
    def test_decide_git_push_action_hard_blockers(self):
        # Pipeline failure
        action, msg = decide_git_push_action(
            success=False, critical_errors=[], new_slug_warns=[], inv_warns=[]
        )
        self.assertEqual(action, "gated")
        self.assertIn("Core pipeline experienced failure", msg)

        # Cross-check discrepancy
        action, msg = decide_git_push_action(
            success=True, critical_errors=["Machine discrepancy"], new_slug_warns=[], inv_warns=[]
        )
        self.assertEqual(action, "gated")
        self.assertIn("Core cross-check found 1 discrepancy", msg)

        # New slug in ERP missing from Excel
        action, msg = decide_git_push_action(
            success=True, critical_errors=[], new_slug_warns=["New Slug ID 9999"], inv_warns=[]
        )
        self.assertEqual(action, "gated")
        self.assertIn("Found 1 new SLUG inventory item(s)", msg)

    def test_decide_git_push_action_skip_git_flag(self):
        # When skip_git is True and no hard blockers exist
        action, msg = decide_git_push_action(
            success=True, critical_errors=[], new_slug_warns=[], inv_warns=["Missing slug"], skip_git=True
        )
        self.assertEqual(action, "skip_flag")
        self.assertIsNone(msg)

        # But if there's a hard blocker, gated takes priority
        action, msg = decide_git_push_action(
            success=False, critical_errors=[], new_slug_warns=[], inv_warns=[], skip_git=True
        )
        self.assertEqual(action, "gated")

    def test_decide_git_push_action_missing_active_inventory(self):
        # inv_warns present -> confirm_missing_active
        action, msg = decide_git_push_action(
            success=True, critical_errors=[], new_slug_warns=[], inv_warns=["Item ID 3503 missing"], skip_git=False
        )
        self.assertEqual(action, "confirm_missing_active")
        self.assertIsNone(msg)

    def test_decide_git_push_action_clean_push(self):
        # Clean run
        action, msg = decide_git_push_action(
            success=True, critical_errors=[], new_slug_warns=[], inv_warns=[], skip_git=False
        )
        self.assertEqual(action, "push")
        self.assertIsNone(msg)

    def test_confirm_missing_slug_push_auto_yes(self):
        res = _confirm_missing_slug_push(["Item ID 3503 [Slug] missing"], auto_yes=True)
        self.assertTrue(res)

    def test_confirm_missing_slug_push_non_interactive(self):
        with patch("sys.stdin") as mock_stdin:
            mock_stdin.isatty.return_value = False
            res = _confirm_missing_slug_push(["Item ID 3503 [Slug] missing"], auto_yes=False)
            self.assertFalse(res)

    def test_confirm_missing_slug_push_interactive_yes(self):
        with patch("sys.stdin") as mock_stdin:
            mock_stdin.isatty.return_value = True
            with patch("builtins.input", return_value="y"):
                res = _confirm_missing_slug_push(["Item ID 3503 [Slug] missing"], auto_yes=False)
                self.assertTrue(res)

            with patch("builtins.input", return_value="YES"):
                res = _confirm_missing_slug_push(["Item ID 3503 [Slug] missing"], auto_yes=False)
                self.assertTrue(res)

    def test_confirm_missing_slug_push_interactive_no(self):
        with patch("sys.stdin") as mock_stdin:
            mock_stdin.isatty.return_value = True
            with patch("builtins.input", return_value="n"):
                res = _confirm_missing_slug_push(["Item ID 3503 [Slug] missing"], auto_yes=False)
                self.assertFalse(res)

            with patch("builtins.input", return_value=""):
                res = _confirm_missing_slug_push(["Item ID 3503 [Slug] missing"], auto_yes=False)
                self.assertFalse(res)

            with patch("builtins.input", side_effect=EOFError):
                res = _confirm_missing_slug_push(["Item ID 3503 [Slug] missing"], auto_yes=False)
                self.assertFalse(res)

    def test_step_git_push_skip(self):
        res = step_git_push(skip=True, show_header=False)
        self.assertFalse(res)


if __name__ == '__main__':
    unittest.main()
