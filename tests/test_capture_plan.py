import unittest

from capture_plan import _dict_to_config, serialize_plan
from view_controller import DisplayType, ViewConfig, ViewType


class CapturePlanTests(unittest.TestCase):
    def test_dict_to_config_reads_camera_and_flags(self):
        cfg = _dict_to_config(
            {
                "filename": "vista_01",
                "view_type": "ISO_3D",
                "display_type": "GEOMETRY_ONLY",
                "azimuth": "210",
                "elevation": "35",
                "is_extruded": "true",
                "ui_automation_required": "si",
            },
            1,
        )

        self.assertEqual(cfg.filename, "vista_01")
        self.assertEqual(cfg.view_type, ViewType.ISO_3D)
        self.assertEqual(cfg.display_type, DisplayType.GEOMETRY_ONLY)
        self.assertEqual(cfg.azimuth, 210.0)
        self.assertEqual(cfg.elevation, 35.0)
        self.assertTrue(cfg.is_extruded)
        self.assertTrue(cfg.ui_automation_required)

    def test_serialize_plan_keeps_camera_and_flags(self):
        data = serialize_plan(
            [
                ViewConfig(
                    filename="vista_02",
                    view_type=ViewType.ISO_3D,
                    display_type=DisplayType.LOAD_CASE,
                    case_name="DEAD",
                    azimuth=225,
                    elevation=30,
                    is_extruded=True,
                    ui_automation_required=True,
                )
            ]
        )

        self.assertEqual(
            data[0],
            {
                "filename": "vista_02",
                "view_type": "ISO_3D",
                "display_type": "LOAD_CASE",
                "case_name": "DEAD",
                "mode_number": 1,
                "window_number": 0,
                "render_delay": 0.0,
                "azimuth": 225,
                "elevation": 30,
                "is_extruded": True,
                "ui_automation_required": True,
                "description": "",
            },
        )


if __name__ == "__main__":
    unittest.main()
