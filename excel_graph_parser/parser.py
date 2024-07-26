from io import BytesIO

from openpyxl import load_workbook, Workbook
from openpyxl.cell import Cell
from viktor import File
from viktor.external.spreadsheet import SpreadsheetCalculation

import plotly.graph_objects as go

ALLOWED_FIGURE_TYPES = ["lineChart", "scatterChart", "barChart", "pieChart"]


class SpreadsheetParser:
    def __init__(self, spreadsheet_calculation: SpreadsheetCalculation = None):
        file = spreadsheet_calculation._file
        if isinstance(file, File):
            with file.open_binary() as r:
                self.workbook = load_workbook(filename=r, data_only=True)
        elif isinstance(file, BytesIO):
            self.workbook = load_workbook(filename=file, data_only=True)
        else:
            raise NotImplementedError

        self._spreadsheet_calculation = spreadsheet_calculation

        # Gather charts by looping through sheets
        self._charts_map = {}
        untitled_index = 1
        for sheet_name in self.workbook.sheetnames:
            sheet = self.workbook[sheet_name]
            for chart in sheet._charts:  # could be empty
                if chart.title:
                    chart_title = ''.join([title_element.t for title_element in chart.title.tx.rich.p[0].r])
                else:
                    chart_title = f"Untitled {untitled_index}"
                    untitled_index += 1

                self._charts_map[chart_title] = chart

    def get_plotly_figure_by_title(self, title: str) -> go.Figure:
        """Gets plotly figure by title"""
        if title not in self._charts_map:
            raise ValueError(f"No chart found with title: {title}")

        spreadsheet = self._spreadsheet_calculation
        result = spreadsheet.evaluate(include_filled_file=False)
        wb = load_workbook(BytesIO(result.file_content), data_only=True)
        try:
            chart_data = self._parse_chart_data(title, wb)
        finally:
            wb.close()

        return self._create_plotly_figure(chart_data)

    def _parse_chart_data(self, chart_title: str, wb: Workbook) -> dict:
        """Extracts chart data from the provided workbook"""
        chart = self._charts_map[chart_title]

        # Get the general chart elements
        chart_type = chart.tagname
        if chart_type not in ALLOWED_FIGURE_TYPES:
            raise TypeError(
                f"Chart '{chart_title}' (type {chart_type}) cannot be parsed. Allowed types are: {ALLOWED_FIGURE_TYPES}"
            )

        x_axis_title, y_axis_title = None, None
        if chart_type != "pieChart":
            if chart.x_axis.title:
                x_axis_title = chart.x_axis.title.tx.rich.p[0].r[-1].t
            if chart.y_axis.title:
                y_axis_title = chart.y_axis.title.tx.rich.p[0].r[-1].t

        # Get series data
        series = []
        input_cat_range = ""
        input_cat_format = None
        for i, series in enumerate(chart.series):
            if chart_type == "scatterChart":
                if series.xVal:
                    if series.xVal.strRef:
                        input_cat_range = series.xVal.strRef.f
                    elif series.xVal.numRef:
                        input_cat_range = series.xVal.numRef.f
                        input_cat_format = series.xVal.numRef.numCache.formatCode

                input_val_range = series.yVal.numRef.f
                input_val_format = series.yVal.numRef.numCache.formatCode

            else:
                if series.cat:
                    # if no category data in the sequence, use the one that was set for the previous sequence
                    if series.cat.strRef:
                        input_cat_range = series.cat.strRef.f
                    elif series.cat.numRef:
                        input_cat_range = series.cat.numRef.f
                        input_cat_format = series.cat.numRef.numCache.formatCode

                input_val_range = series.val.numRef.f
                input_val_format = series.val.numRef.numCache.formatCode

            input_cat_format = None if input_cat_format == "General" else input_cat_format
            input_val_format = None if input_val_format == "General" else input_val_format

            # category_axis_data
            chart_sheet_name = (
                input_cat_range
                .replace("(", "")
                .replace(")", "")
                .replace("'", "")
                .split(sep="!")[0]
            )
            chart_cat_range = (
                input_cat_range
                .replace('(', '')
                .replace(')', '')
                .replace("'", "")
                .replace(f"{chart_sheet_name}!", "")
                .replace('$', "")
            )

            cat_range_start = chart_cat_range.split(",")[0]
            cat_range_end = chart_cat_range.split(",")[-1] if "," in chart_cat_range else chart_cat_range
            cat_data = []
            for element in wb[chart_sheet_name][f"{cat_range_start}:{cat_range_end}"]:
                cat_data = [e.value for e in element if type(e) == Cell]

            # value_axis_data
            chart_sheet_name = (
                input_val_range
                .replace('(', '')
                .replace(')', '')
                .replace("'", "")
                .split(sep="!")[0]
            )
            chart_val_range = (
                input_val_range
                .replace('(', '')
                .replace(')', '')
                .replace("'", "")
                .replace(f"{chart_sheet_name}!", "")
            )
            val_range_start = chart_val_range.split(",")[0]
            val_range_end = chart_val_range.split(",")[-1] if "," in chart_val_range else chart_val_range
            val_data = []
            for element in wb[chart_sheet_name][f"{val_range_start}:{val_range_end}"]:
                val_data = [e.value for e in element if type(e) == Cell]

            series_name = series.tx.v if series.tx else None
            ser = {
                "category_axis_data": cat_data,
                "value_axis_data": val_data,
                "category_value_format": input_cat_format,
                "values_value_format": input_val_format,
                "series_name": series_name if series_name else None
            }
            series.append(ser)

        chart_data = {
            "chart_title": chart_title,
            "chart_type": chart_type,
            "x_axis_title": x_axis_title,
            "y_axis_title": y_axis_title,
            "series": series,
        }

        return chart_data

    @staticmethod
    def _create_plotly_figure(chart_data: dict) -> go.Figure:
        """Creates plotly figure based on the extracted chart data"""
        fig = go.Figure()
        if chart_data["chart_type"] == "lineChart":
            for ser in chart_data["series"]:
                fig.add_trace(go.Scatter(
                    x=ser["category_axis_data"],
                    y=ser["value_axis_data"],
                    mode='lines',
                    name=ser["series_name"]
                ))
            fig.update_layout(
                title_text=chart_data["chart_title"],
                xaxis_title=chart_data["x_axis_title"],
                yaxis_title=chart_data["y_axis_title"],
                yaxis_tickformat=chart_data["series"][0]["values_value_format"],
                xaxis_tickformat=chart_data["series"][0]["category_value_format"],
            )
        if chart_data["chart_type"] == "barChart":
            for ser in chart_data["series"]:
                fig.add_trace(go.Bar(x=ser["category_axis_data"], y=ser["value_axis_data"]))
            fig.update_layout(
                title_text=chart_data["chart_title"],
                xaxis_title=chart_data["x_axis_title"],
                yaxis_title=chart_data["y_axis_title"],
                yaxis_tickformat=chart_data["series"][0]["values_value_format"],
                xaxis_tickformat=chart_data["series"][0]["category_value_format"],
            )
        if chart_data["chart_type"] == "pieChart":
            for ser in chart_data["series"]:
                fig.add_trace(go.Pie(labels=ser["category_axis_data"], values=ser["value_axis_data"]))
            fig.update_layout(
                title_text=chart_data["chart_title"],
            )
        if chart_data["chart_type"] == "scatterChart":
            for ser in chart_data["series"]:
                fig.add_trace(go.Scatter(x=ser["category_axis_data"], y=ser["value_axis_data"]))
            fig.update_layout(
                title_text=chart_data["chart_title"],
                xaxis_title=chart_data["x_axis_title"],
                yaxis_title=chart_data["y_axis_title"],
                yaxis_tickformat=chart_data["series"][0]["values_value_format"],
                xaxis_tickformat=chart_data["series"][0]["category_value_format"],
            )

        return fig
