from multiprocessing import Pool, cpu_count
import regex as re
from typing import List, Optional, Tuple, Union, Callable, Any
import pandas as pd
import xlwings as xw
from xlwings.main import Book, Range
import traceback
import logging
from argparse import ArgumentParser, ArgumentTypeError
from functools import partial
import timeit

logging.basicConfig(
        format='%(asctime)s %(levelname)-8s %(message)s',
        level=logging.INFO,
        datefmt='%Y-%m-%d %H:%M:%S')

# E.g. "[C:\Users\[username]\Desktop\Book1.xlsx]Sheet1!A1:B5"
path_pattern = re.compile(r"^\[(.+?)\](.+?)!(.+)$")
my_date_handler = lambda year, month, day, **kwargs: "%04i-%02i-%02i" % (year, month, day)

def get_sheet_names(book: Book) -> List[str]:
    return [sheet.name for sheet in book.sheets]

def normalize_data(shape: Tuple[int, int], input) -> List[List[str]]:
    """
    Constructing pd.DataFrame requires a List of Lists. Data must be normalized accordingly.
    """
    row, column = shape
    if row == 1 and column == 1:
        return [[input]]
    elif row == 1 and column > 1:
        return [input]
    elif row > 1 and column == 1:
        return list(map(lambda x: [x], input))
    elif row > 1 and column > 1:
        return input
    else:
        raise ValueError(f"Shape invalid: [{shape}]")

def extract_cell_values(excel_range: Range) -> pd.DataFrame:
    data = excel_range.options(dates=my_date_handler, numbers=str, empty="").value
    return pd.DataFrame(normalize_data(excel_range.shape, data))

def extract_cell_formulas(excel_range: Range) -> pd.DataFrame:
    data = excel_range.options(dates=str, numbers=str, empty="").formula
    return pd.DataFrame(normalize_data(excel_range.shape, data))

def get_cell_value(cell) -> str:
    """
    Try getting a cell value. If it is in a merged area, get the value of the merged area.
    """
    if cell.value is not None:
        return str(cell.value)
    elif cell.MergeArea.Cells(1, 1).Value is not None:
        return str(cell.MergeArea.Cells(1, 1).Value)
    return ""

def extract_cell_property_by_enumerating(excel_range: Range, extract_func: Callable[[Any], str]):
    """
    The slowest method, but also the most flexible (any type of properties can be extracted).
    """
    data = [[extract_func(excel_range.api.Cells(row_idx, col_idx)) for col_idx in range(1, excel_range.api.Columns.Count + 1)]
            for row_idx in range(1, excel_range.api.Rows.Count + 1)]
    return pd.DataFrame(data)

def initialize():
    global cached_excels
    cached_excels = {}

def extracting_data(request: str, extract_func: Callable[[Range], pd.DataFrame]) -> Tuple[bool, Union[pd.DataFrame, str]]:
    result: Tuple[bool, Union[pd.DataFrame, str]] = (False, "")
    try:
        if request == "EXIT":
            for (_, value) in cached_excels.items():
                (temp_app, temp_book) = value
                temp_book.close()
                temp_app.kill()
            result = (True, "")
        else:
            matches = re.search(path_pattern, request)
            if matches:
                (excel_path, excel_sheet, excel_range) = (matches.group(1), matches.group(2), matches.group(3))
                if excel_path not in cached_excels:
                    temp_app = xw.App(visible=False)
                    temp_app.calculation = "manual"
                    temp_app.screen_updating = False
                    try:
                        temp_book = temp_app.books.open(excel_path)
                        cached_excels[excel_path] = (temp_app, temp_book)
                    except:
                        temp_app.kill() # if file is problematic, kill the temp instance.
                        raise
                if excel_sheet not in get_sheet_names(cached_excels[excel_path][1]):
                    raise ValueError(f"Unknown sheet: {excel_sheet}")
                cells: Range = cached_excels[excel_path][1].sheets[excel_sheet].range(excel_range)
                result = (True, extract_func(cells))
            else:
                raise ValueError(f"Invalid request format: [{request}]")
    except Exception:
        result = (False, traceback.format_exc())
    return result


def check_input_num_core(value: str) -> str:
    value = value.strip().lower()
    specified_cpu_count = int(value)
    max_cpu_count = cpu_count()
    if (specified_cpu_count < 0):
        raise ArgumentTypeError(f"Number of cores invalid. [{specified_cpu_count}]")
    if (specified_cpu_count == 0):
        specified_cpu_count = max_cpu_count
    if specified_cpu_count > max_cpu_count:
        raise ArgumentTypeError(f"Max CPUs allowed: [{max_cpu_count}]. (Demanded: [{specified_cpu_count}])")
    return value

def check_input_cell_mode(value: str) -> str:
    value = value.strip().lower()
    if value not in {"value_fast", "value_aggressive", "formula"}:
        raise ArgumentTypeError(f"Cell value mode invalid. [{value}]")
    return value

if __name__ == "__main__":
    pool: Optional[Pool] = None
    # From Console.
    specified_cpu_count = 1
    try:
        # Process input arguments.
        parser = ArgumentParser()
        parser.add_argument("-num_core", dest="num_core", default="0", help="Number of cores for parallelism.",
                            metavar="a positive number", type=check_input_num_core)
        parser.add_argument("-cell_mode", dest="cell_mode", default="fast", help="The mode to extract value in cell.",
                            metavar="value_fast|value_aggressive|formula", type=check_input_cell_mode)
        inputs = parser.parse_args()

        specified_cpu_count = int(inputs.num_core)
        cell_value_mode = inputs.cell_mode

        extract_func = {
            "value_fast": extract_cell_values,
            "value_aggressive": partial(extract_cell_property_by_enumerating, extract_func=get_cell_value),
            "formula": extract_cell_formulas
        }[cell_value_mode]

        with open(r"requests.txt", "r") as input_file:
            requests = input_file.read().splitlines()

        # Start the pool of Workers (Process).
        pool = Pool(specified_cpu_count, initialize, ())

        data_cell_values_all_at_once = list(map(lambda request: (request, extract_func), requests))

        logging.info("Start processing Excel files.")

        results = pool.starmap(extracting_data, data_cell_values_all_at_once)

        logging.info("Finish processing Excel files.")

        for result in results:
            bln, df = result
            if bln:
                logging.info(f"Collected Dataframe: {df.shape}.")
                if df.shape[0] > 1 and df.shape[1] > 0:
                    logging.info(f"First row: [{'|'.join([df.iloc[1, i] for i in range(df.shape[1])])}]")
            else:
                logging.warning(df)

        logging.info("Done")
    except Exception as ex:     # Fatal exception.
        logging.error(traceback.format_exc())
    finally:
        if pool:
            # Clean up!
            pool.starmap(extracting_data, [("EXIT", None) for i in range(0, specified_cpu_count)])
            pool.close()
