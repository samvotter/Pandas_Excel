__author__ = "Sam Van Otterloo"

'''
This module is designed to facilitate reading and writing Pandas Dataframes 
to an Excel Workbook environment
'''

from pandas import ExcelWriter, read_html, DataFrame
from requests import get
from re import sub
from typing import Dict


class ConditionalFormat:

    def __init__(
            self,
            cell_format: Dict,
            condition_type: str = 'cell',
            criteria: str = '',
            value=0,
            minimum=0,
            maximum=0,
    ):
        self.format = cell_format
        self.condition_type = condition_type
        self.criteria = criteria
        self.value = value
        self.minimum = minimum
        self.maximum = maximum


# if the DataTable has formatting options it should be entered as a Dict where
# {column: [ConditionalFormat,]}
class DataTable:

    def __init__(
            self,
            title: str,
            df: DataFrame,
            sheet: str,
            index: bool = True,
            formatting: Dict = None
    ):
        self.title = title
        self.df = df
        self.sheet = sheet
        self.index = index

        self.start_row: int = 0
        self.start_col: int = 0

        height = self.df.shape[0]
        width = self.df.shape[1]

        self.end_row: int = self.start_row + height
        self.end_col: int = self.start_col + width

        self.formatting = formatting

    def set_start(
            self,
            start_row: int = 0,
            start_col: int = 0
    ):

        self.start_row: int = start_row
        self.start_col: int = start_col

        height = self.df.shape[0]
        width = self.df.shape[1]

        self.end_row: int = start_row + height
        self.end_col: int = start_col + width


class Chart:

    def __init__(
            self,
            source: DataTable,
            chart_type: str,
            name: str,
            nw_corner: str,
            category_col: str,
            values_col: str
    ):
        self.source = source
        self.chart_type = chart_type
        self.name = name
        self.nw_corner = nw_corner
        self.category_col = category_col
        self.values_col = values_col


class ExcelManager:

    def __init__(
            self,
            outfile_name: str
    ):
        self.outfile = outfile_name
        self.writer = ExcelWriter(
            self.outfile,
            engine='xlsxwriter'
        )
        # Exalphabet
        self.exalpha = {
            num: chr(65+num)
            for num in range(26)
        }

    def write_data_table(
            self,
            dt: DataTable,
            set_start: bool = False,
            start_row: int = 0,
            start_col: int = 0
    ) -> None:
        if set_start:
            dt.set_start(
                start_row,
                start_col
            )
        sheet = dt.sheet
        if sheet not in self.writer.sheets:
            self.writer.sheets[sheet] = self.writer.book.add_worksheet(
                sheet
            )
        self.writer.sheets[sheet].write_string(
            dt.start_row,
            dt.start_col,
            dt.title
        )
        dt.set_start(
            dt.start_row + 1,
            dt.start_col
        )
        dt.df.to_excel(
            self.writer,
            sheet,
            startrow=dt.start_row,
            startcol=dt.start_col,
            index=dt.index
        )
        if dt.formatting:
            for column in dt.formatting:
                for form in dt.formatting[column]:
                    format_row_start = dt.start_row+2
                    format_row_end = dt.end_row+1
                    col = self.exalpha[dt.start_col + dt.df.columns.get_loc(column)+1]
                    cell_format = self.writer.book.add_format(form.format)
                    if form.value:
                        self.writer.sheets[sheet].conditional_format(
                            f"{col}{format_row_start}:{col}{format_row_end}",
                            {
                                'type': form.condition_type,
                                'criteria': form.criteria,
                                'value': form.value,
                                'format': cell_format
                            }
                        )
                    else:
                        self.writer.sheets[sheet].conditional_format(
                            f"{col}{format_row_start}:{col}{format_row_end}",
                            {
                                'type': form.condition_type,
                                'criteria': form.criteria,
                                'minimum': form.minimum,
                                'maximum': form.maximum,
                                'format': cell_format
                            }
                        )

    # charts can only be written AFTER the chart.source has been written to the workbook FIRST
    def write_chart(
            self,
            chart: Chart
    ) -> None:
        sheet = chart.source.sheet
        if chart.category_col == 'index':
            c_loc = self.exalpha[chart.source.start_col]
        else:
            c_loc = self.exalpha[chart.source.start_col + chart.source.df.columns.get_loc(chart.category_col)+1]
        v_loc = self.exalpha[chart.source.start_col + chart.source.df.columns.get_loc(chart.values_col)+1]

        if sheet not in self.writer.sheets:
            self.writer.book.add_worksheet(
                sheet
            )
        chart_to_add = self.writer.book.add_chart({'type': chart.chart_type})
        chart_to_add.add_series(
            {
                'name': chart.name,
                'categories': f"={sheet}!${c_loc}${chart.source.start_row+2}:${c_loc}${chart.source.end_row+1}",
                'values': f"={sheet}!${v_loc}${chart.source.start_row+2}:${v_loc}${chart.source.end_row+1}"
            }
        )
        self.writer.sheets[sheet].insert_chart(chart.nw_corner, chart_to_add)

    def save_close(
            self
    ) -> None:
        print("Saving Workbook . . .")
        self.writer.save()
        print("Closing Workbook . . .")
        self.writer.close()
        print("Success!")


# retrieve a Dataframe from a url
def get_dataframe_from_url(
        url: str,
        username: str = None,
        password: str = None
) -> DataFrame:
    if username and password:
        prefix = r"https://"
        suffix = sub(prefix, '', url)
        prefix += f"{username}:{password}@"
        url = prefix + suffix
    processed_url = get(
        url,
        verify=False
    )
    return read_html(processed_url.content)[0]
