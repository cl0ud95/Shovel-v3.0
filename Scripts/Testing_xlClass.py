from Shovel_Classes import ExcelTable
import pandas as pd
import numpy as np

def test_df():
    testing_tbl = ExcelTable("shovel", "Sheet2", "testing")
    df = testing_tbl.dataframe(first_col_as_index=True)
    print(df)

def test_headerlist():
    testing_tbl = ExcelTable("shovel", "Sheet2", "testing")
    list_header = testing_tbl.headerlist()
    print(list_header)

def test_write_to_table():
    data = [(6, "n2nanchor_2", "AnchorForce2D"), (7, "n2nanchor_3", "AnchorForce2D")]
    df = pd.DataFrame(data)
    testing_tbl = ExcelTable("shovel", "Sheet2", "testing")
    result = testing_tbl.write_df_to_table(df, wipe_table=True)
    print(result)

def test_search_delete():
    testing_tbl = ExcelTable("shovel", "Sheet2", "testing")
    testing_tbl.search_and_delete({"Element":["Plate_10", "Plate_11", "Plate_99", "Plate_3"]})

def test_write_dict_to_table():
    data = {(8, "Plate_11"):[[1,4,46,788,33,5667], [5,2,66,11,98555,123]]}
    testing_tbl = ExcelTable("Shovel.xlsm", "Sheet2", "testing")
    testing_tbl.write_dict_to_table(data)

def test_df_to_dict():
    testing_tbl = ExcelTable("shovel", "Sheet2", "testing")
    df = testing_tbl.dataframe(first_col_as_index=True)
    dictionary = testing_tbl.df_to_dict("No.")
    print(dictionary)
def max_table_col():
    testing_tbl = ExcelTable("shovel", "Sheet2", "testing")
    maxnum = int(max(testing_tbl.table.ListColumns("No.").DataBodyRange())[0])
    print(maxnum)
# test_headerlist()
# test_write_to_table()
# test_search_delete()
test_write_dict_to_table()
# test_df_to_dict()
# max_table_col()
