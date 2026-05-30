import unittest

from sap_bridge import SapBridge


class SapBridgeNameCoercionTests(unittest.TestCase):
    def test_coerce_names_splits_delimited_string_results(self):
        self.assertEqual(
            SapBridge._coerce_names("DEAD\nLIVE\nWIND"),
            ["DEAD", "LIVE", "WIND"],
        )
        self.assertEqual(
            SapBridge._coerce_names("DEAD; LIVE; WIND"),
            ["DEAD", "LIVE", "WIND"],
        )

    def test_coerce_names_splits_delimited_items_inside_api_tuple(self):
        self.assertEqual(
            SapBridge._coerce_names((0, 3, "DEAD\tLIVE\tWIND")),
            ["DEAD", "LIVE", "WIND"],
        )

    def test_coerce_names_preserves_names_with_spaces(self):
        self.assertEqual(
            SapBridge._coerce_names("LOAD PATTERN WITH SPACES"),
            ["LOAD PATTERN WITH SPACES"],
        )


if __name__ == "__main__":
    unittest.main()
