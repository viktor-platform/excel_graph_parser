import unittest
from pathlib import Path
from unittest.mock import patch

from viktor import File
from viktor.external.spreadsheet import SpreadsheetCalculation, SpreadsheetResult

from excel_graph_parser import ExcelChartParser

SPREADSHEET_PATH = Path(__file__).parent / "spreadsheet.xlsx"


class TestExcelChartParser(unittest.TestCase):

    def test_get_plotly_figure(self):
        spreadsheet = SpreadsheetCalculation.from_path(SPREADSHEET_PATH, inputs=[])
        parser = ExcelChartParser(spreadsheet)

        with patch("viktor.external.spreadsheet.SpreadsheetCalculation.evaluate") as mock_evaluate:
            mock_evaluate.return_value = SpreadsheetResult(values={}, file=File.from_path(SPREADSHEET_PATH))

            with self.subTest("lineChart with categories"):
                fig = parser.get_plotly_figure("lineChart")
                self.assertListEqual(fig._data[0]['x'], [10, 20, 30])
                self.assertListEqual(fig._data[0]['y'], [100, 200, 300])

            with self.subTest("lineChart without categories"):
                fig = parser.get_plotly_figure("lineChart-no-cat")
                self.assertListEqual(fig._data[0]['x'], [1, 2, 3])  # fall back to index
                self.assertListEqual(fig._data[0]['y'], [100, 200, 300])

            with self.subTest("scatterChart with categories"):
                fig = parser.get_plotly_figure("scatterChart")
                self.assertListEqual(fig._data[0]['x'], [10, 20, 30])
                self.assertListEqual(fig._data[0]['y'], [100, 200, 300])

            with self.subTest("scatterChart without categories"):
                fig = parser.get_plotly_figure("scatterChart-no-cat")
                self.assertListEqual(fig._data[0]['x'], [1, 2, 3])  # fall back to index
                self.assertListEqual(fig._data[0]['y'], [100, 200, 300])
