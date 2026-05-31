import unittest

from sap_ui_automation import GA_ROOT, GA_ROOTOWNER, SAP2000UIController, _normalize_menu_text


class SapUiAutomationTests(unittest.TestCase):
    def test_normalize_menu_text_removes_ampersands_and_shortcuts(self):
        self.assertEqual(_normalize_menu_text("&Set 3D View...\tCtrl+3"), "set 3d view...")
        self.assertEqual(_normalize_menu_text("  Show Load Assigns  "), "show load assigns")

    def test_window_context_accepts_main_window_and_owned_dialog(self):
        controller = SAP2000UIController.__new__(SAP2000UIController)
        controller.hwnd_principal = 100
        controller.sap_pid = 500

        controller._get_window_pid = lambda hwnd: 500 if hwnd in {100, 101} else 999
        controller._get_ancestor = lambda hwnd, flag: {
            (100, GA_ROOT): 100,
            (100, GA_ROOTOWNER): 100,
            (101, GA_ROOT): 100,
            (101, GA_ROOTOWNER): 100,
            (102, GA_ROOT): 102,
            (102, GA_ROOTOWNER): 102,
        }[(hwnd, flag)]

        self.assertTrue(controller._window_is_sap_context(100))
        self.assertTrue(controller._window_is_sap_context(101))
        self.assertFalse(controller._window_is_sap_context(102))


if __name__ == "__main__":
    unittest.main()
