import unittest
from unittest.mock import patch

from view_controller import DisplayType, ViewConfig, ViewController, ViewType


class _FakeView:
    def __init__(self, set_view_results, unzoom_result=0, refresh_result=0):
        self._set_view_results = list(set_view_results)
        self._unzoom_result = unzoom_result
        self._refresh_result = refresh_result
        self.set_view_calls = []
        self.unzoom_calls = []
        self.refresh_calls = []

    def SetView(self, window_number, view_type):
        self.set_view_calls.append((window_number, view_type))
        if self._set_view_results:
            result = self._set_view_results.pop(0)
            if isinstance(result, Exception):
                raise result
            return result
        return 1

    def UnzoomAll(self, window_number):
        self.unzoom_calls.append(window_number)
        if isinstance(self._unzoom_result, Exception):
            raise self._unzoom_result
        return self._unzoom_result

    def RefreshView(self, window_number, hold):
        self.refresh_calls.append((window_number, hold))
        if isinstance(self._refresh_result, Exception):
            raise self._refresh_result
        return self._refresh_result

    def SetActiveDisplayCase(self, case_name):
        return 0


class _FakeModel:
    def __init__(self, view):
        self.View = view


class _FakeUI:
    def __init__(self, enabled=True, set_vista_result=True, set_extrusion_result=False):
        self.enabled = enabled
        self.set_vista_result = set_vista_result
        self.set_extrusion_result = set_extrusion_result
        self.set_vista_calls = []
        self.set_extrusion_calls = []

    def set_vista(self, view_type, azimuth=None, elevation=None):
        self.set_vista_calls.append((view_type, azimuth, elevation))
        return self.set_vista_result

    def set_extrusion(self, enabled):
        self.set_extrusion_calls.append(enabled)
        return self.set_extrusion_result


class ViewControllerTests(unittest.TestCase):
    @patch("view_controller.time.sleep", return_value=None)
    def test_apply_retries_set_view_with_window_one(self, _sleep):
        fake_view = _FakeView([1, 0])
        controller = ViewController(_FakeModel(fake_view), base_render_delay=0.0)
        cfg = ViewConfig(filename="geom", view_type=ViewType.ISO_3D, window_number=0)

        controller.apply(cfg)

        self.assertEqual(fake_view.set_view_calls, [(0, 0), (1, 0)])
        self.assertEqual(fake_view.unzoom_calls, [1])
        self.assertEqual(fake_view.refresh_calls, [(1, True)])

    @patch("view_controller.time.sleep", return_value=None)
    def test_apply_continues_when_set_view_fails(self, _sleep):
        fake_view = _FakeView([1, 1])
        controller = ViewController(_FakeModel(fake_view), base_render_delay=0.0)
        cfg = ViewConfig(
            filename="geom",
            view_type=ViewType.PLAN_XY,
            display_type=DisplayType.GEOMETRY_ONLY,
            window_number=0,
        )

        controller.apply(cfg)

        self.assertEqual(fake_view.set_view_calls, [(0, 1), (1, 1)])
        self.assertEqual(fake_view.unzoom_calls, [0])
        self.assertEqual(fake_view.refresh_calls, [(0, True)])

    @patch("view_controller.time.sleep", return_value=None)
    def test_apply_continues_when_set_view_raises_exception(self, _sleep):
        fake_view = _FakeView([RuntimeError("SetView"), RuntimeError("SetView")])
        controller = ViewController(_FakeModel(fake_view), base_render_delay=0.0)
        cfg = ViewConfig(filename="geom", view_type=ViewType.ISO_3D, window_number=0)

        controller.apply(cfg)

        self.assertEqual(fake_view.set_view_calls, [(0, 0), (1, 0)])
        self.assertEqual(fake_view.unzoom_calls, [0])
        self.assertEqual(fake_view.refresh_calls, [(0, True)])

    @patch("view_controller.time.sleep", return_value=None)
    def test_apply_continues_when_zoom_and_refresh_raise(self, _sleep):
        fake_view = _FakeView(
            [0],
            unzoom_result=RuntimeError("UnzoomAll"),
            refresh_result=RuntimeError("RefreshView"),
        )
        controller = ViewController(_FakeModel(fake_view), base_render_delay=0.0)
        cfg = ViewConfig(filename="geom", view_type=ViewType.ISO_3D, window_number=0)

        controller.apply(cfg)

        self.assertEqual(fake_view.set_view_calls, [(0, 0)])
        self.assertEqual(fake_view.unzoom_calls, [0])
        self.assertEqual(fake_view.refresh_calls, [(0, True)])

    @patch("view_controller.time.sleep", return_value=None)
    def test_apply_forwards_custom_angles_to_ui_fallback(self, _sleep):
        fake_view = _FakeView([RuntimeError("SetView"), RuntimeError("SetView")])
        fake_ui = _FakeUI(enabled=True, set_vista_result=True)
        controller = ViewController(_FakeModel(fake_view), base_render_delay=0.0, ui_controller=fake_ui)
        cfg = ViewConfig(
            filename="geom",
            view_type=ViewType.ISO_3D,
            azimuth=210,
            elevation=25,
        )

        controller.apply(cfg)

        self.assertEqual(len(fake_ui.set_vista_calls), 1)
        self.assertEqual(fake_ui.set_vista_calls[0][1:], (210, 25))

    @patch("view_controller.time.sleep", return_value=None)
    def test_apply_requires_ui_when_flagged(self, _sleep):
        fake_view = _FakeView([RuntimeError("SetView"), RuntimeError("SetView")])
        controller = ViewController(_FakeModel(fake_view), base_render_delay=0.0)
        cfg = ViewConfig(
            filename="geom",
            view_type=ViewType.ISO_3D,
            ui_automation_required=True,
        )

        with self.assertRaisesRegex(RuntimeError, "requiere automatizacion UI armada"):
            controller.apply(cfg)

    @patch("view_controller.time.sleep", return_value=None)
    def test_apply_tracks_extrusion_via_ui_when_requested(self, _sleep):
        fake_view = _FakeView([0])
        fake_ui = _FakeUI(enabled=True, set_vista_result=False, set_extrusion_result=False)
        controller = ViewController(_FakeModel(fake_view), base_render_delay=0.0, ui_controller=fake_ui)
        cfg = ViewConfig(
            filename="geom",
            view_type=ViewType.ISO_3D,
            is_extruded=True,
        )

        controller.apply(cfg)

        self.assertEqual(fake_ui.set_extrusion_calls, [True])


if __name__ == "__main__":
    unittest.main()
