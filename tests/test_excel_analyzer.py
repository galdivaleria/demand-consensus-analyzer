import pytest
import pandas as pd
import numpy as np
from io import BytesIO

# the project modules live in src/ so adjust path if necessary
import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src")))
import excel_analyzer


def create_excel(df: pd.DataFrame) -> BytesIO:
    """Helper to write a DataFrame to an in-memory Excel file and return a BytesIO."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    buffer.seek(0)
    return buffer


def test_load_excel_file_success():
    df = pd.DataFrame({"Key": ["A", "B"], "Value": [1, 2]})
    buf = create_excel(df)
    sheets, workbook = excel_analyzer.load_excel_file(buf)
    assert sheets == ["Sheet1"]
    # the workbook should allow reading back the same sheet
    reloaded = pd.read_excel(workbook, sheet_name="Sheet1")
    pd.testing.assert_frame_equal(df, reloaded)


def test_load_excel_file_invalid():
    # passing a non-excel, non-csv file should not crash; we may fall back to
    # treating it as CSV. The main requirement is that the function returns
    # something reasonable and does not raise.
    buf = BytesIO(b"this is not an excel file")
    sheets, workbook = excel_analyzer.load_excel_file(buf)
    assert sheets is not None
    # workbook may be a dataframe or an ExcelFile; ensure it has one of the
    # expected interfaces. We don't care about content, just that nothing blew
    # up.
    if isinstance(workbook, pd.DataFrame):
        assert hasattr(workbook, 'columns')
    else:
        assert hasattr(workbook, 'sheet_names')


def test_compare_excel_files_missing_key_column(capfd):
    df1 = pd.DataFrame({"NotKey": [1, 2], "Value": [100, 200]})
    df2 = pd.DataFrame({"NotKey": [1, 2], "Value": [110, 210]})
    result = excel_analyzer.compare_excel_files(df1, df2, threshold=5)
    # function should return None when Key column missing; we no longer rely on
    # streamlit output because the stub is a no-op in tests.
    assert result is None



def test_compare_excel_files_basic_threshold():
    df1 = pd.DataFrame({"Key": ["x", "y"], "Metric": [100, 50]})
    df2 = pd.DataFrame({"Key": ["x", "y"], "Metric": [110, 40]})
    # For threshold 5, row x is an outlier (10% diff), y is not (20% diff? Actually y diff is 20% > threshold)
    result = excel_analyzer.compare_excel_files(df1, df2, threshold=5)
    assert isinstance(result, pd.DataFrame)
    # There should be two rows
    assert len(result) == 2
    # Verify computed columns
    row_x = result[result['Key'] == 'x'].iloc[0]
    assert row_x['Metric_File1'] == 100.0
    assert row_x['Metric_File2'] == 110.0
    assert pytest.approx(row_x['Metric_Diff%'], rel=1e-3) == 10.0
    # numpy.bool_ truthiness is acceptable
    assert bool(row_x['Metric_IsOutlier']) is True
    row_y = result[result['Key'] == 'y'].iloc[0]
    assert pytest.approx(row_y['Metric_Diff%'], rel=1e-3) == 20.0
    assert bool(row_y['Metric_IsOutlier']) is True


def test_compare_excel_files_non_numeric():
    df1 = pd.DataFrame({"Key": ["1"], "A": ["foo"], "B": [100]})
    df2 = pd.DataFrame({"Key": ["1"], "A": ["bar"], "B": [110]})
    result = excel_analyzer.compare_excel_files(df1, df2, threshold=5)
    # Should ignore non-numeric column A and only compare B
    assert 'A_File1' not in result.columns
    assert 'B_File1' in result.columns


def test_compare_zero_division():
    df1 = pd.DataFrame({"Key": ["k1"], "M": [0]})
    df2 = pd.DataFrame({"Key": ["k1"], "M": [5]})
    result = excel_analyzer.compare_excel_files(df1, df2, threshold=1)
    assert result.loc[0, 'M_Diff%'] == 100
    assert bool(result.loc[0, 'M_IsOutlier']) is True


def test_ignore_formula_injection():
    # formulas or strings beginning with '=' should not be treated as numeric values
    df1 = pd.DataFrame({"Key": ["a"], "Val": ["=HYPERLINK(\"http://malicious\",\"click\")"]})
    df2 = pd.DataFrame({"Key": ["a"], "Val": [123]})
    result = excel_analyzer.compare_excel_files(df1, df2, threshold=1)
    # the non-numeric column should be ignored entirely
    assert 'Val_File1' not in result.columns


def test_get_plant_summary():
    # basic call without material column should behave as before
    df = pd.DataFrame({
        "Key": ["a", "b", "c"],
        "Plant": ["P1", "P1", "P2"],
        "X_IsOutlier": [True, False, True],
        "Y_IsOutlier": [False, False, True]
    })
    summary = excel_analyzer.get_plant_summary(df, plant_column="Plant")
    # there should be two plants
    assert set(summary['Plant']) == {"P1", "P2"}
    # check records counts
    assert summary[summary['Plant'] == 'P1']['Records'].iloc[0] == 2
    assert summary[summary['Plant'] == 'P2']['Records'].iloc[0] == 1
    # Materials column shouldn't exist when not requested
    assert 'Materials' not in summary.columns


def test_get_plant_summary_no_column():
    df = pd.DataFrame({"Key": [1], "A": [10]})
    assert excel_analyzer.get_plant_summary(df, plant_column="NonExistent") is None
    # specifying a material column when plant is missing should also return None
    assert excel_analyzer.get_plant_summary(df, plant_column="NonExistent", material_column="A") is None


def test_load_excel_file_csv():
    df = pd.DataFrame({"Key": ["A"], "Val": [1]})
    buf = BytesIO()
    df.to_csv(buf, index=False)
    buf.seek(0)
    sheets, workbook = excel_analyzer.load_excel_file(buf)
    assert sheets == ["Sheet1"]
    # reading the returned object should give identical frame
    assert isinstance(workbook, pd.DataFrame)
    pd.testing.assert_frame_equal(df, workbook)


def test_plant_summary_without_outliers():
    df = pd.DataFrame({
        "Key": [1,2],
        "Plant": ["X","Y"],
        "A": [10,20]
    })
    summary = excel_analyzer.get_plant_summary(df, plant_column="Plant")
    # even though no "IsOutlier" columns exist, it should produce records
    assert summary['Records'].tolist() == [1,1]
    assert summary['Outliers'].tolist() == [0,0]
    # still no Materials column unless explicitly requested
    assert 'Materials' not in summary.columns


def test_highlight_outliers():
    df = pd.DataFrame({
        "Key":["a","b"],
        "M_IsOutlier":[True, False]
    })
    styled = excel_analyzer.highlight_outliers(df)
    # style should be a Styler object and should not error when rendering
    assert hasattr(styled, 'to_html')




def test_compare_excel_files_with_plant_and_material():
    df1 = pd.DataFrame({
        "Key": ["x"],
        "Plant": ["P1"],
        "Material": ["M1"],
        "Metric": [100]
    })
    df2 = pd.DataFrame({
        "Key": ["x"],
        "Plant": ["P1"],
        "Material": ["M1"],
        "Metric": [110]
    })
    result = excel_analyzer.compare_excel_files(
        df1, df2, threshold=5, plant_column="Plant", material_column="Material"
    )
    assert 'Plant' in result.columns
    assert 'Material' in result.columns
    assert result.loc[0, 'Plant'] == 'P1'
    assert result.loc[0, 'Material'] == 'M1'


def test_get_plant_summary_with_materials():
    df = pd.DataFrame({
        "Key": ["a", "b", "c"],
        "Plant": ["P1", "P1", "P2"],
        "Material": ["X", "Y", "X"],
        "X_IsOutlier": [True, False, True],
        "Y_IsOutlier": [False, False, True]
    })
    summary = excel_analyzer.get_plant_summary(
        df, plant_column="Plant", material_column="Material"
    )
    assert 'Materials' in summary.columns
    assert summary[summary['Plant'] == 'P1']['Materials'].iloc[0] == 2
    assert summary[summary['Plant'] == 'P2']['Materials'].iloc[0] == 1


def test_threshold_edge_cases():
    df1 = pd.DataFrame({"Key":["a"],"M":[0]})
    df2 = pd.DataFrame({"Key":["a"],"M":[0]})
    # threshold zero should still work
    result = excel_analyzer.compare_excel_files(df1, df2, threshold=0)
    assert result.loc[0, 'M_Diff%'] == 0
    assert bool(result.loc[0, 'M_IsOutlier']) is False

    # negative threshold (nonsense) should treat everything as outlier
    result = excel_analyzer.compare_excel_files(df1, df2, threshold=-1)
    assert bool(result.loc[0, 'M_IsOutlier']) is True


def test_map_threshold_option():
    # verify both legacy and newly added thresholds
    assert excel_analyzer.map_threshold_option(">1%") == 1
    assert excel_analyzer.map_threshold_option(">2%") == 2
    assert excel_analyzer.map_threshold_option(">5%") == 5
    assert excel_analyzer.map_threshold_option(">15%") == 15
    assert excel_analyzer.map_threshold_option(">30%") == 30
    assert excel_analyzer.map_threshold_option(">60%") == 60
    assert excel_analyzer.map_threshold_option(">90%") == 90
    assert excel_analyzer.map_threshold_option("Custom", custom=42) == 42
    with pytest.raises(ValueError):
        excel_analyzer.map_threshold_option("Unknown")
    with pytest.raises(ValueError):
        excel_analyzer.map_threshold_option("Custom")


def test_filter_outliers_by_metric():
    df = pd.DataFrame({
        "Key":[1,2,3],
        "A_IsOutlier":[True, False, True],
        "B_IsOutlier":[False, True, False]
    })
    mask_any = excel_analyzer.filter_outliers_by_metric(df)
    assert mask_any.tolist() == [True, True, True]
    mask_a = excel_analyzer.filter_outliers_by_metric(df, metric_name="A")
    assert mask_a.tolist() == [True, False, True]
    mask_b = excel_analyzer.filter_outliers_by_metric(df, metric_name="B")
    assert mask_b.tolist() == [False, True, False]
    mask_unknown = excel_analyzer.filter_outliers_by_metric(df, metric_name="C")
    assert mask_unknown.tolist() == [False, False, False]



def test_threshold_validation():
    """Passing a non-numeric threshold should raise a TypeError or be handled gracefully."""
    df1 = pd.DataFrame({"Key":["a"],"M":[1]})
    df2 = pd.DataFrame({"Key":["a"],"M":[2]})
    with pytest.raises((TypeError, ValueError)):
        # the implementation expects a number; a string should not silently succeed
        excel_analyzer.compare_excel_files(df1, df2, threshold="foo")


def test_compare_excel_files_empty():
    """Comparing two empty dataframes should not crash and should return None or empty frame."""
    df1 = pd.DataFrame(columns=["Key", "A"])
    df2 = pd.DataFrame(columns=["Key", "A"])
    result = excel_analyzer.compare_excel_files(df1, df2, threshold=5)
    assert result is None or result.empty


def test_csv_injection_ignored():
    """Strings beginning with '=' in CSV input should not be treated as formulas/numbers."""
    # create a dataframe as if read from a CSV file containing a CSV injection
    df1 = pd.DataFrame({"Key": ["1"], "Val": ["=HYPERLINK(\"http://malicious\",\"click\")"]})
    df2 = pd.DataFrame({"Key": ["1"], "Val": [100]})
    result = excel_analyzer.compare_excel_files(df1, df2, threshold=1)
    # the injected value should be ignored (non-numeric)
    assert 'Val_File1' not in result.columns

