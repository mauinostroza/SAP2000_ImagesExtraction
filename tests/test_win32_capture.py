import os
import unittest

from win32_capture import _score_sap2000_window


class Win32CaptureWindowSelectionTests(unittest.TestCase):
    def test_rejects_own_capture_window(self):
        current_pid = os.getpid()
        self.assertIsNone(
            _score_sap2000_window("SAP2000 Capture", "TkTopLevel", current_pid, current_pid)
        )

    def test_prefers_real_sap2000_title(self):
        current_pid = os.getpid()
        real_score = _score_sap2000_window(
            "SAP2000 v23 - Modelo1.sdb",
            "Afx:00400000:8:10011:0:0",
            current_pid + 1,
            current_pid,
        )
        other_score = _score_sap2000_window(
            "CSI.SAP2000",
            "WindowsForms10.Window.8.app.0.141b42a_r7_ad1",
            current_pid + 2,
            current_pid,
        )
        self.assertIsNotNone(real_score)
        self.assertIsNotNone(other_score)
        self.assertGreaterEqual(real_score, other_score)

    def test_rejects_non_sap_window(self):
        current_pid = os.getpid()
        self.assertIsNone(
            _score_sap2000_window(
                "Bloc de notas",
                "Notepad",
                current_pid + 1,
                current_pid,
            )
        )


if __name__ == "__main__":
    unittest.main()
