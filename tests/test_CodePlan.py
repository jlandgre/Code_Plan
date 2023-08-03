#Version 8/3/23
"""
These commands are helpful for avoiding typing
python -m pytest test_CodePlan.py -v -s

#2345678901234567890123456789012345678901234567890123456789012345678901234567890
"""
import sys, os
import pandas as pd
import numpy as np
import pytest

#Append the roll_scripts subdirectory to sys.path
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
sys.path.append(parent_dir)

#Import project-specific class(es)
from CodePlan import VBAToCodePlan

#Toggle to activate various print statements within tests (if IsPrint:)
IsPrint = True

@pytest.fixture
def file1():
    return "code1.vb"

@pytest.fixture
def file2():
    return "code2.vb"

@pytest.fixture
def cls1(file1):
    return VBAToCodePlan(file1)

@pytest.fixture
def cls2(file2):
    return VBAToCodePlan(file2)

@pytest.fixture
def sep():
    #line break separator for df_plan values
    return ",\n"

def test_create_df_plan_args_col(cls1, sep):
    """
    Create the arg column in df_plan by parsing the args_temp column
    JDL 8/3/23
    """
    cls1.read_vba_code_file()
    cls1.init_df_code()
    cls1.combine_split_lines()
    cls1.set_filters()
    cls1.parse_start_lines()
    cls1.create_df_plan_args_col()

    assert cls1.df_plan.shape == (3, 4)

    lst_expected = ["cls|Variant|ByRef" + sep + "arg1|Variant|ByVal" + sep +
                    "arg2|Variant|ByRef|Optional",
                    "cls|Variant|ByRef" + sep + "arg1|Variant|ByRef",
                    "cls|Variant|ByRef" + sep + "i|Integer|ByRef" + sep + 
                    "j|Integer|ByRef" + sep + "arg2|Variant|ByRef"]
    for idx, s in enumerate(cls1.df_plan["arguments"]):
        assert s == lst_expected[idx]

    output_file = "df_plan.xlsx"
    cls1.df_plan.to_excel(output_file, index=False)
    

def test_parse_arglist(cls1, sep):
    """
    Parse a VBA arglist string (read from df_plan args_temp row/column
    JDL 8/3/23
    """
    s1 = "arg1 As Integer"
    s1_expected = "arg1|Integer|ByRef"
    assert cls1.parse_arglist(s1) == s1_expected

    s2 = "arg2"
    s2_expected = "arg2|Variant|ByRef"
    assert cls1.parse_arglist(s2) == s2_expected

    s3 = "ByRef arg3"
    s3_expected = "arg3|Variant|ByRef"
    assert cls1.parse_arglist(s3) == s3_expected

    s4 = "Optional ByVal arg4"
    s4_expected = "arg4|Variant|ByVal|Optional"
    assert cls1.parse_arglist(s4) == s4_expected

    s5 = "Optional ByVal arg5 As Integer"
    s5_expected = "arg5|Integer|ByVal|Optional"
    assert cls1.parse_arglist(s5) == s5_expected

    #Parse compound arglist
    s = "arg1 As Integer, arg2, ByRef arg3, Optional ByVal arg4, " \
        + "Optional ByVal arg5 As Integer"
    s_expected = s1_expected + sep + s2_expected + sep + \
        s3_expected + sep + s4_expected + sep + s5_expected    
    assert cls1.parse_arglist(s) == s_expected

def test_parse_start_lines(cls1):
    """
    Parse the sub and function start lines
    """
    cls1.read_vba_code_file()
    cls1.init_df_code()
    cls1.combine_split_lines()
    cls1.set_filters()
    cls1.parse_start_lines()

    assert cls1.df_plan.shape == (3,4)  #3 routines, 4 columns

    assert cls1.df_plan.loc[0,"routine_name"] == "ExampleProcedure"
    assert cls1.df_plan.loc[0,"type"] == "Boolean"
    assert cls1.df_plan.loc[0,"args_temp"] == "cls, ByVal arg1, Optional arg2"
    assert cls1.df_plan.loc[0,"line_num_start"] == 11

    assert cls1.df_plan.loc[1,"routine_name"] == "Method1"
    assert cls1.df_plan.loc[1,"type"] == "Boolean"
    assert cls1.df_plan.loc[1,"args_temp"] == "cls, arg1"
    assert cls1.df_plan.loc[1,"line_num_start"] == 31

    assert cls1.df_plan.loc[2,"routine_name"] == "Method2"
    assert cls1.df_plan.loc[2,"type"] == "Variant"
    assert cls1.df_plan.loc[2,"args_temp"] == "cls, i As Integer, j As Integer, arg2"
    assert cls1.df_plan.loc[2,"line_num_start"] == 45

def test_parse_startline(cls1):
    """
    Parse an individual line that defines a function or sub
    """
    s1 = "Function ExampleProcedure(cls, ByVal arg1, Optional arg2) As Boolean"
    name, args, type, is_fn, is_sub = cls1.parse_startline(s1)
    tup_results = "ExampleProcedure", "cls, ByVal arg1, Optional arg2",\
        "Boolean", True, False
    check_parse_startline_results(tup_results, name, args, type, is_fn, is_sub)

    s2 = "Function ExampleProcedure()"
    name, args, type, is_fn, is_sub = cls1.parse_startline(s2)
    tup_results = "ExampleProcedure", "", "Variant", True, False
    check_parse_startline_results(tup_results, name, args, type, is_fn, is_sub)

    s3 = "Sub ExampleProcedure(arg1, arg2)"
    name, args, type, is_fn, is_sub = cls1.parse_startline(s3)
    tup_results = "ExampleProcedure", "arg1, arg2", "Sub", False, True
    check_parse_startline_results(tup_results, name, args, type, is_fn, is_sub)

def check_parse_startline_results(tup_results, name, args, type, is_fn, is_sub):
    assert name == tup_results[0]
    assert args == tup_results[1]
    assert type == tup_results[2]
    assert is_fn == tup_results[3]
    assert is_sub == tup_results[4]

def test_set_filters(cls1):
    cls1.read_vba_code_file()
    cls1.init_df_code()
    cls1.combine_split_lines()
    cls1.set_filters()

    assert list(cls1.df_code[cls1.fil_starts].index) == [10, 30, 44]
    assert list(cls1.df_code[cls1.fil_ends].index) == [23, 38, 55]

def test_combine_split_lines(cls2):
    """
    Combine split VBA code lines into a single line
    """
    cls2.read_vba_code_file()
    cls2.init_df_code()
    cls2.combine_split_lines()
    assert cls2.df_code.index.size == 20

    s = "If Not .Method1(cls, arg1) Then GoTo ErrorExit"
    assert cls2.df_code.loc[12, 'stripped_code'] == s

    s = "errs.RecordErr errs, \"ExampleProcedure\", ExampleProcedure"
    assert cls2.df_code.loc[18, 'stripped_code'] == s

def test_init_df_code(cls1):
    cls1.read_vba_code_file()
    cls1.init_df_code()
    assert cls1.df_code.shape == (57,2)

def test_read_vba_code_file(cls1):
    cls1.read_vba_code_file()
    assert cls1.vba_code[0:15] == "Option Explicit"

def test_init_cls(cls1):
    assert cls1.file_name == "code1.vb"

