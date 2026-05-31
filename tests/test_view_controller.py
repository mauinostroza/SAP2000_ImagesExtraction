import unittest
from unittest.mock import patch

from view_controller import DisplayType, ViewConfig, ViewController, ViewType


class _FakeView:
    def __init__(self, set_view_results):
        self._set_view_results = list(set_view_results)
        self.set_view_calls = []
        self.unzoom_calls = []
        self.refresh_calls = []

    def SetView(self, window_number, view_type):
        self.set_view_calls.append((window_number, view_type))
        if self._set_view_results:
            return self._set_view_results.pop(0)
        return 1

    def UnzoomAll(self, window_number):
        self.unzoom_calls.append(window_number)
        return 0

    def RefreshView(self, window_number, hold):
        self.refresh_calls.append((window_number, hold))
        return 0

    def SetActiveDisplayCase(self, case_name):
        return 0


class _FakeModel:
    def __init__(self, view):
        self.View = view


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


if __name__ == "__main__":
    unittest.main()
