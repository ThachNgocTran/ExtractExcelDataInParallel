# Extract Excel Data In Parallel
A small script to extract Excel Data in parallel into Dataframes, using xlwings and Python's multiprocessing. Why need Parallelism? Because accessing low-level cells is slow, e.g. a few ms for one cell. Low-level cells allow to leverage many kinds of functionalities supported by VBA.

My article: [Extract Excel Data In Parallel](https://medium.com/@thachngoctran/extract-excel-data-in-parallel-181838c4ed37)

## Prerequisites

+ Anaconda3-2019.10-Windows-x86_64.exe
+ Python v3.7.4 (64 bit)
+ Additional package: `xlwings`, `regex` (see `requirements.txt`)
+ Microsoft Office (Excel)

## Getting started

1. Put Excel files of interest into `data` folder.
2. Update the extraction request in `requests.txt`.
3. In a Virtual Environment, run as followed:

```python
python.exe main_program -num_core 4 -cell_mode value_aggressive
```

4. After extracting ranges in Excel, they are converted into Pandas Dataframe. What to do with those Dataframes is up to you. E.g. export to CSV files.

## Notes

1. Switch `-cell_mode`: `value_aggressive` to get cells individually (even merged cells), `value_fast` to get cells all at once (very fast, lightning fast!), `formula` to get cells' formula.
2. Switch `-num_core`: if `0`, get the maximum CPU cores available.
3. The program can be extended, for example, writing another function to extract Background Color in Conditional Formatting, e.g. `rang.api.Cells(1, 1).DisplayFormat.Interior.Color`. Expect that going through cells individually is a must because such property is not available natively in `xlwings`.
4. `xlwings` is a wrapper around Microsoft's Library to control Excel instance. It is as powerful as VBA. When wanting something which is not available through `xlwings`, search for doing it in VBA instead. Then in `xlwings`, call `xlwing_range.api.` to access the desired property, like above.

## Sample Logging Result

```bash
2019-12-03 20:17:35 INFO     Start processing Excel files.
2019-12-03 20:21:50 INFO     Finish processing Excel files.
2019-12-03 20:21:50 INFO     Collected Dataframe: (5001, 2).
2019-12-03 20:21:50 INFO     First row: [Central America and the Caribbean Antigua and Barbuda |Central America and the Caribbean Antigua and Barbuda ]
2019-12-03 20:21:50 INFO     Collected Dataframe: (5001, 2).
2019-12-03 20:21:50 INFO     First row: [Baby Food|Online]
2019-12-03 20:21:50 INFO     Collected Dataframe: (5001, 2).
2019-12-03 20:21:50 INFO     First row: [M|12/20/2013]
2019-12-03 20:21:50 INFO     Collected Dataframe: (5001, 2).
2019-12-03 20:21:50 INFO     First row: [957081544|01/11/2014]
2019-12-03 20:21:50 INFO     Done
```

## Sample File Source

[1] [Downloads 18 - Sample CSV Files / Data Sets for Testing (till 1.5 Million Records) - Sales](http://eforexcel.com/wp/downloads-18-sample-csv-files-data-sets-for-testing-sales/)
