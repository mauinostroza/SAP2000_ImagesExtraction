import unittest

from sap_ui_automation import _normalize_menu_text


class SapUiAutomationTests(unittest.TestCase):
    def test_normalize_menu_text_removes_ampersands_and_shortcuts(self):
        self.assertEqual(_normalize_menu_text("&Set 3D View...\tCtrl+3"), "set 3d view...")
        self.assertEqual(_normalize_menu_text("  Show Load Assigns  "), "show load assigns")


if __name__ == "__main__":
    unittest.main()
