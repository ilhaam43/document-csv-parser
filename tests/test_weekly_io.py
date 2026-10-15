import tempfile
import unittest
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from weekly.constants import MASTER_COLUMNS
from weekly.io import iter_master_csv, write_master_csv


class WeeklyIoTests(unittest.TestCase):
    def test_generated_master_csv_round_trips_typed_values(self) -> None:
        record = {column: None for column in MASTER_COLUMNS}
        record.update(
            {
                "No.": 1,
                "quote_num": "Q-1",
                "fab_upload": datetime(2026, 9, 24, 10, 30),
                "Count FAB": 1,
                "otc": Decimal("1250.50"),
                "mrc": Decimal("0"),
                "OTC+MRC": Decimal("1250.50"),
                "OTC+MRC (inMio)": Decimal("0.0012505"),
            }
        )
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "master.csv"
            write_master_csv(path, [record])
            loaded = list(iter_master_csv(path))

        self.assertEqual(len(loaded), 1)
        self.assertEqual(loaded[0]["No."], 1)
        self.assertEqual(loaded[0]["fab_upload"], datetime(2026, 9, 24, 10, 30))
        self.assertEqual(loaded[0]["otc"], Decimal("1250.50"))
        self.assertIsNone(loaded[0]["cash_dt"])


if __name__ == "__main__":
    unittest.main()
