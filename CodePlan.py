#Version 8/28/23
import pandas as pd
import re
#2345678901234567890123456789012345678901234567890123456789012345678901234567890

class VBAToCodePlan:
    def __init__(self, file_name):
        self.file_name = file_name

        # VBA code as a string and DataFrame
        self.vba_code = ""
        self.df_code = None
        self.df_lookup = None # docstring lookup table row subset of df_code

        # df_code filters
        self.fil_starts = None
        self.fil_ends = None
        self.fil_docstrings = None

        # VBA Sub/Function first line and Dim statement Regular Expression patterns
        self.line1_pattern = r"(Function|Sub)\s+(\w+)\((.*?)\)(?:\s+As\s+(\w+))?"
        self.dim_pattern = r"(\w+)\s(?:As\s)(New\s)?(\w+)(?:,\s*|$)"

        # Code plan
        self.df_plan = None

    def CreateCodePlanProcedure(self):
        """
        Procedure to parse VBA code into a code plan
        combine_split_lines resets df_code range index but subsequent code assumes a
        consistent index in the df_code DataFrame.  

        JDL 8/3/2023
        """
        self.read_vba_code_file()
        self.init_df_code()
        self.combine_split_lines()
        self.set_filters()
        self.drop_public_private()
        self.parse_start_lines()
        self.create_df_plan_args_col()
        self.create_docstrings_df()
        self.add_plan_docstring_col()
        self.create_internal_vars_df()
        self.add_plan_internal_vars_col()


    def read_vba_code_file(self):
        """
        Read the VBA code from the specified file into a string

        JDL 8/1/2023
        """
        with open(self.file_name, "r") as file:
            self.vba_code = file.read()

    def init_df_code(self):
        """
        Initialize the DataFrame with the VBA code

        JDL 8/1/2023
        """

        # Column with stripped leading/trailing spaces
        stripped_lines = [line.strip() for line in self.vba_code.split("\n")]

        self.df_code = pd.DataFrame({"orig_code": self.vba_code.split("\n"), 
                                     "stripped_code": stripped_lines})

    def combine_split_lines(self):
        """
        Combine split VBA code lines into a single line

        JDL 8/1/2023
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

    def set_filters(self):
        """
        Set filters for function starts and ends

        JDL 8/1/2023    Modified 8/28/23 switch to using regex for starts and ends
        """
        #Set filters for first and last lines in functions and subs
        col = "stripped_code"
        re_starts = r"^(Public|Private)?\s*(Sub|Function)\s+"
        self.fil_starts = self.df_code[col].str.match(re_starts)
        re_ends = r"^(Public|Private)?\s*End (Sub|Function)\s+"
        self.fil_ends = self.df_code[col].str.match(re_ends)

        # Comment fence boundaries and rows with VBA dim statements
        self.fil_bounds = self.df_code[col].str.startswith("\'------")
        self.fil_dims = self.df_code[col].str.startswith("Dim ")

    def drop_public_private(self):
        """
        Strip leading "Public " or "Private " from function/sub names

        JDL 8/28/2023
        """
        col = "stripped_code"
        re_strip = r"^(Public|Private)\s+"
        subset = self.df_code.loc[self.fil_starts, col]
        self.df_code.loc[self.fil_starts, col] = subset.str.replace(re_strip, "", regex=True)

    def parse_start_lines(self):
        """
        Parse the sub and function start lines

        JDL 8/1/2023
        """
        lst_names = []
        lst_args = []
        lst_types = []
        lst_idx = []

        # Iterate through lines that define functions and subs
        fil = self.fil_starts
        for idx, line_code in zip(self.df_code.loc[fil].index,
                                      self.df_code.loc[fil, "stripped_code"]):
            name, args, type, is_fn, is_sub = self.parse_startline(line_code)
            lst_names.append(name)
            lst_args.append(args)
            lst_types.append(type)
            lst_idx.append(idx)
        
        # Initialize Code Plan DF and Add lists as columns
        self.df_plan = pd.DataFrame()
        self.df_plan["routine_name"] = lst_names
        self.df_plan["type"] = lst_types
        self.df_plan["args_temp"] = lst_args
        self.df_plan["idx_start"] = lst_idx

    def parse_startline(self, line_code):
        """
        Parse an individual line that defines a function or sub
        (helper function for parse_start_lines)

        JDL 8/1/2023
        """

        # Check for function and sub matches
        is_fn, is_sub = False, False
        fn_match = re.match(self.line1_pattern, line_code)

        if fn_match:
            if fn_match.group(1) == "Function":
                is_fn = True
            elif fn_match.group(1) == "Sub":
                is_sub = True

            name, args, type = fn_match.group(2), fn_match.group(3), fn_match.group(4)

        # Set type if not specified in Function line or if Sub
        if type is None:
            if is_fn: type = "Variant"
            if is_sub: type = "Sub"
        
        #Return results tuple
        return name, args, type, is_fn, is_sub

    def create_df_plan_args_col(self):
        """
        Create the df_plan arguments col by parsing and then dropping args_temp col

        JDL 8/3/23
        """
        self.df_plan["arguments"] = self.df_plan["args_temp"].apply(self.parse_arglist)
        self.df_plan.drop(columns=["args_temp"], inplace=True)
        
    def parse_arglist(self, arglist):
        """
        Parse a VBA arglist string (read from df_plan args_temp row/column
        (helper function for create_df_plan_args_col)

        JDL 8/3/23
        """
        def ParseArg(arg):
            has_type = False
            lst = arg.split(" As ")
            if len(lst) > 1:
                has_type = True
                lst = lst[0].split(" ") + [lst[1]]
            else:
                lst = arg.split(" ")
            return lst, has_type

        lst_arg_code_plan = []
    
        for arg in arglist.split(", "):
            parsed_arg, has_type = ParseArg(arg)
            
            #Set defaults
            IsOptional, arg_by = False, "ByRef"

            #if arg is Optional
            if parsed_arg[0] == "Optional":
                IsOptional = True
                parsed_arg = parsed_arg[1:]

            #if arg has overt ByRef or ByVal specification
            if parsed_arg[0] in ["ByRef", "ByVal"]: 
                arg_by = parsed_arg[0]
                parsed_arg = parsed_arg[1:]
            
            #Set sub-variables and create parsed list is "code plan" format
            arg_name = parsed_arg[0]
            arg_type = "Variant"
            if has_type: arg_type = parsed_arg[-1]
            arg_code_plan = "|".join([arg_name, arg_type, arg_by])
            if IsOptional: arg_code_plan = arg_code_plan + "|Optional"
            
            lst_arg_code_plan.append(arg_code_plan)
        return ",\n".join(lst_arg_code_plan)

    def create_docstrings_df(self):
        """
        Locate docstring rows between function/sub start rows and previous
        function/sub end row
        
        JDL 8/3/23
        """
        # Combined filter for starts, ends and comment fence boundaries
        fil = self.fil_starts | self.fil_ends | self.fil_bounds

        # Modify a copy/subset of df_code to use as a lookup table
        df_docstr = self.df_code.copy()
        ser_idx_shift = df_docstr.loc[fil].index.to_series().shift(1, fill_value=0)
        df_docstr.loc[fil, "prev_end_idx"] = ser_idx_shift

        # Parse docstrings for functions
        for idx in df_docstr[self.fil_starts].index:
            idx_prev = int(df_docstr.loc[idx, "prev_end_idx"])

            #idx_prev +1 for first row after end except file begin or if boundary row
            if (idx_prev > 0) or (idx_prev in df_docstr[self.fil_bounds].index):
                idx_prev +=1

            # Combine lines in block between previous end and function starts
            docstring = "\n".join(self.create_lst_lines(idx, idx_prev))
            df_docstr.loc[idx, "docstring"] = docstring

        # Set Class attribute and subset cols
        self.df_lookup =  df_docstr[self.fil_starts]
        self.df_lookup = self.df_lookup[ ["prev_end_idx", "docstring"]]

    def create_lst_lines(self, idx, idx_prev):
        """
        Helper function to create list of cleaned up lines in a docstring block
        (omits lines after encountering VBA blank line or blank comment line so
        won't include Created/Modified Date row etc.)

        JDL 8/4/23
        """
        docstring_lines = []
        for i in range(idx_prev, idx):
            line = self.df_code.loc[i, "stripped_code"]

            # Strip single quote (VBA comment character) and leading whitespace
            if line.startswith("'") and not line.strip() == "'":
                docstring_lines.append(line[1:].strip())
            
            #Stop if the line is blank or only a single quote + whitespace
            elif line.strip() == "'" or line.strip() == "":
                break
        return docstring_lines

    def add_plan_docstring_col(self):
        """
        Merge docstring column from df_lookup into df_plan

        JDL 8/3/23
        """
        self.df_plan = self.df_plan.merge(self.df_lookup[["docstring"]], 
                                          left_on="idx_start", right_index=True)

    def create_internal_vars_df(self):
        """
        Extract/parse VBA variables from Dim statement rows into a lookup table
        indexed by function/sub first line index
        
        JDL 8/3/23
        """
        df_lookup = self.df_code.copy()
        
        #Populate rows with index of their function/sub's first line
        df_lookup.loc[self.fil_starts, "idx_start"] = \
                df_lookup[self.fil_starts].index.values
        
        df_lookup["idx_start"].ffill(inplace=True)

        #Initialize column for internal variable strings
        df_lookup["vars_internal"] = [[] for _ in range(len(df_lookup))]
        
        # Parse VBA Dim statements - omit df_lookup rows prior to first Function or Sub (NaN idx_start)
        #fil_internal_rows = ~df_lookup["idx_start"].isnull() | self.fil_dims

        #for idx in df_lookup[fil_internal_rows].index:
        for idx in df_lookup[self.fil_dims].index:
            idx_start = int(df_lookup.loc[idx, "idx_start"])
            lst_parsed_dims = self.parse_dim_statement(df_lookup.loc[idx, "stripped_code"])
                    
            #Append Dim statement list onto previous from other Dims in this function/sub
            df_lookup.at[idx_start, "vars_internal"] = \
                df_lookup.loc[idx_start, "vars_internal"] + lst_parsed_dims
        
        #Subset to just start rows and convert the temp lists to strings
        df_lookup = df_lookup[self.fil_starts]
        df_lookup["vars_internal"] = \
            df_lookup["vars_internal"].apply(lambda lst: ",\n".join(map(str, lst)))
        
        self.df_lookup = df_lookup[["vars_internal"]]

    def parse_dim_statement(self, s):
        """
        Helper function to parse VBA Dim statement into a list of strings
        for each variable declared in the Dim statement

        JDL 8/4/23
        """
        
        #Pattern match on string without "Dim " prefix aka [4:] slice
        dim_match = re.findall(self.dim_pattern, s[4:])
        
        #Build list of parsed strings for each variable
        lst_parsed_vars = []
        for match in dim_match:
            
            #If match[1] populated, New descriptor is present
            lst_parsed = [match[0], match[2]]
            if len(match[1]) > 0: lst_parsed.append("New")
            s_parsed = "|".join(lst_parsed)
            
            #Append variable's parsed string to the list
            lst_parsed_vars.append(s_parsed)
        return lst_parsed_vars

    def add_plan_internal_vars_col(self):
        """
        Merge internal_vars column from df_lookup into df_plan

        JDL 8/4/23
        """
        self.df_plan = self.df_plan.merge(self.df_lookup[["vars_internal"]], 
                                          left_on="idx_start", right_index=True)


