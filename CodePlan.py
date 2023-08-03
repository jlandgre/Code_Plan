import pandas as pd
import re
#2345678901234567890123456789012345678901234567890123456789012345678901234567890

class VBAToCodePlan:
    def __init__(self, file_name):
        self.file_name = file_name

        #VBA code as a string and DataFrame
        self.vba_code = ""
        self.df_code = None

        #VBA routine first lines Regular Expression patterns
        self.line1_pattern = r"(Function|Sub)\s+(\w+)\((.*?)\)(?:\s+As\s+(\w+))?"

        #Code plan
        self.df_plan = None

    def parse_start_lines(self):
        """
        Parse the sub and function start lines
        """
        lst_names = []
        lst_args = []
        lst_types = []
        lst_line_nos = []

        # Iterate through lines that define functions and subs
        for line_no, line_code in zip(self.df_code.loc[self.fil_starts].index,
                                      self.df_code.loc[self.fil_starts, 'stripped_code']):
            name, args, type, is_fn, is_sub = self.parse_startline(line_code)
            lst_names.append(name)
            lst_args.append(args)
            lst_types.append(type)
            lst_line_nos.append(line_no + 1)
        
        # Initialize Code Plan DF and Add lists as columns
        self.df_plan = pd.DataFrame()
        self.df_plan['routine_name'] = lst_names
        self.df_plan['type'] = lst_types
        self.df_plan['args_temp'] = lst_args
        self.df_plan['line_num_start'] = lst_line_nos

    def parse_startline(self, line_code):
        """
        Parse an individual line that defines a function or sub
        """

        # Check for function and sub matches
        is_fn, is_sub = False, False
        fn_match = re.match(self.line1_pattern, line_code)
        #print("\n\n", fn_match.group(1), fn_match.group(2), fn_match.group(3), fn_match.group(4))

        if fn_match:
            if fn_match.group(1) == "Function":
                is_fn = True
            elif fn_match.group(1) == "Sub":
                is_sub = True

            name = fn_match.group(2)
            args = fn_match.group(3)
            type = fn_match.group(4)

        # Set type if not specified in Function line or if Sub
        if type is None:
            if is_fn: type = "Variant"
            if is_sub: type = "Sub"
        
        #Return results tuple
        return name, args, type, is_fn, is_sub

    def set_filters(self):
        """
        Set filters for function starts and ends
        """
        self.fil_starts = self.df_code['stripped_code'].str.startswith('Function')
        self.fil_starts = self.fil_starts | \
            self.df_code['stripped_code'].str.startswith('Sub')
        
        self.fil_ends = self.df_code['stripped_code'].str.startswith('End Function')
        self.fil_ends = self.fil_ends | \
            self.df_code['stripped_code'].str.startswith('End Sub')

    def combine_split_lines(self):
        """
        Combine split VBA code lines into a single line
        """
        
        self.lst_rows_deleted = []
        n_rows = len(self.df_code)
        idx = 0
        while idx < n_rows - 1:

            # Get the range index from the row number
            index_idx = self.df_code.iloc[idx].name
            s = self.df_code.loc[index_idx, "stripped_code"]
            if s.endswith("_"):
                index_idx_next = self.df_code.iloc[idx + 1].name
                next_row_string = self.df_code.loc[index_idx_next, "stripped_code"]
                self.df_code.loc[index_idx, "stripped_code"] = s[:-1] + next_row_string

                # Drop the next row whose string got combined
                self.df_code.drop(index_idx_next, inplace=True)
                self.lst_rows_deleted.append(index_idx_next)
                n_rows -= 1
            else:
                idx += 1
        
        #Reset to range index with consecutive integers
        self.df_code.reset_index(drop=True, inplace=True)
        
    def init_df_code(self):
        """
        Initialize the DataFrame with the VBA code
        """

        # Column with stripped leading/trailing spaces
        stripped_lines = [line.strip() for line in self.vba_code.split('\n')]

        self.df_code = pd.DataFrame({'orig_code': self.vba_code.split('\n'), 
                                     'stripped_code': stripped_lines})
    """
    Read the VBA code from the specified file into a string
    """
    def read_vba_code_file(self):
            with open(self.file_name, 'r') as file:
                self.vba_code = file.read()


    def CreateCodePlanProcedure(self):
        self.read_vba_code_file()
        self.init_df_code()
        self.combine_split_lines()
        self.set_filters()
        self.parse_start_lines()
        