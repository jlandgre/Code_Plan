#Version 8/1/23
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

def test_parse_args(cls1):
    """
    Parse the args_temp df column
    """
    cls1.read_vba_code_file()
    cls1.init_df_code()
    cls2.combine_split_lines()
    cls1.set_filters()
    cls1.parse_start_lines()

    cls1.parse_args()
    assert cls1.df_plan.shape == (3,5)  #3 routines, 5 columns with "args" added
    s = ""
    assert cls1.loc[0, "args"] == s


def test_parse_start_lines(cls1):
    """
    Parse the sub and function start lines
    """
    cls1.read_vba_code_file()
    cls1.init_df_code()
    cls2.combine_split_lines()
    cls1.set_filters()
    cls1.parse_start_lines()
    #print('\n\n', cls1.df_plan, '\n\n')

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
    assert cls1.df_plan.loc[2,"args_temp"] == "cls, i, j, arg2"
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
    cls2.combine_split_lines()
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

