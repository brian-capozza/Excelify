from __future__ import annotations
import xlsxwriter
import pandas as pd
import re
import collections

# Setting up imports from github in VS code:
# 1. ctrl + shift + p
# 2. Python: Create Environment
# 3. Select Interpreter
# 4. pip install git+https://github.com/PCXIT/PCXModule.git#subdirectory=PCX

# dictionary of allowed functions in the totals column
AGG_CONVERSION = {'AVERAGE': 1,
                  'COUNT': 2,
                  'COUNTA': 3,
                  'MAX': 4,
                  'MIN': 5,
                  'PRODUCT': 6,
                  'STDEV.S': 7,
                  'STDEV.P': 8,
                  'SUM': 9,
                  'VAR.S': 10,
                  'VAR.P': 11,
                  'MEDIAN': 12,
                  'MODE.SNGL': 13,
                  'LARGE': 14,
                  'SMALL': 15,
                  'PERCENTILE.INC': 16,
                  'QUARTILE.INC': 17,
                  'PERCENTILE.EXC': 18,
                  'QUARTILE.EXC': 19}

def makehash():
    return collections.defaultdict(makehash)

class Helper:

    def clean_string(input_string: str):
        """
        Information
        -
        Removes spaces and special characters from a string

        Parameters:
        -
        input_string (str): the string to be cleaned

        Returns:
        -
        cleaned_string (str): the cleaned string after cleaning operations
        """
        
        # Define regex pattern to match special characters and spaces
        pattern = r'[^a-zA-Z0-9]'
        
        # Use re.sub to replace special characters and spaces with an empty string
        cleaned_string = re.sub(pattern, '', input_string)
        
        return cleaned_string
    
    def _format_string(input_str: str =''):
        """
        Information
        -
        Private function used to replace backtick surrounded strings with the appropraite formatting
        to access the column data (`column_name` -> row['column_name'])

        Parameters
        -
        input_str (str): the string to format

        Returns
        -
        str: the formatted string

        """
        def _replace_match(match):
            column_name = match.group(1)
            return f"row['{column_name}']"

        pattern = re.compile(r'`([^`]+)`')
        result_str = pattern.sub(_replace_match, input_str)
        return result_str


class Table:

    """
    Information
    -
    Initializes an Excel Table object that can be manipulated and displayed in a workbook

    Parameters:
    -
    name (str): the name of the table to add to your sheet
    title (str): the title of the table
    data (Dataframe): the data that the table will contain
    start_col (int): optional parameter to change what column the table starts on: default is 0
    title_row (int): optional parameter to change what row the title goes on: default is 1
    header_row (int): optional parameter to change what row the header goes on: default is 1
    data_row (int): optional parameter to change what row the data goes on: default is 2
    total_row (str): optional parameter to change whether the total goes above or below
        (top or bottom) the table: default is bottom
    freeze_panes (bool): optional parameter to specify whether the header should be frozen or not

    Attibutes
    -
    formats (dict): Dictionary containing formatting settings for columns.
    conditions (dict): Dictionary containing conditional formatting settings for columns.
    agg_functions (dict): Dictionary containing aggregation functions for columns.
    operation_functions (dict): Dictionary containing operation functions for columns.
    links (dict): Dictionary containing link mappings for columns.
    data_validation (dict): Dictionary containing data validation settings for columns.
    data_validation_condition (dict): Dictionary containing data validation condition settings for columns.
    data_validation_extras (dict): Dictionary containing extra data validation settings for columns.
    """

    def __init__(self,
                 name: str, 
                 data: pd.DataFrame,
                 title: str =None,
                 start_col: int =0, 
                 title_row: int =1, 
                 header_row: int =1, 
                 data_row: int =2, 
                 total_row: str ='bottom', 
                 freeze_panes: int | tuple =None,
                 title_align: str = 'center',
                 header_comments: dict[str: str] = makehash()):
        self.helper = Helper
        self.name = self.helper.clean_string(name)
        self.title = title
        self.data = data
        self.start_col = start_col
        self.title_row = title_row
        self.header_row = header_row
        self.data_row = data_row
        self.start_row = data_row + 1
        self.end_data_row = len(data) + (data_row - 1)
        self.header = data.columns.tolist()
        self.end_col = len(self.header) - 1 + start_col
        self.total_row = total_row
        self.freeze_panes = freeze_panes
        self.hidden_columns = None
        self.title_align = title_align
        self.formats = makehash()
        self.conditions = makehash()
        self.total_formats = makehash()
        self.total_conditions = makehash()
        self.agg_functions = makehash()
        self.operation_functions = makehash()
        self.links = makehash()
        self.data_validation = makehash()
        self.data_validation_condition = makehash()
        self.data_validation_extras = makehash()
        self.header_comments = header_comments
    
    def _convert(self, base_string: str ='', table_name: str ='', aggregation_functions: str | list =''):
        """
        Information
        -
        Convert a base string containing column names enclosed in backticks to a formula 
        using Excel's AGGREGATE function for each column.

        Parameters:
        -
        base_string (str): The base string containing column names enclosed in backticks (`).
        table_name (str): The name of the Excel table containing the columns.
        aggregation_functions (list or str): A list of aggregation functions to apply to each column,
            or a single aggregation function as a string. Supported aggregation functions include 'AVERAGE',
            'COUNT', 'COUNTA', 'MAX', 'MIN', 'PRODUCT', 'STDEV.S', 'STDEV.P', 'SUM', 'VAR.S', 'VAR.P'.

        Raises:
        -
        ValueError: If the number of non-empty strings between backticks does not match the length of the aggregation functions list.
        TypeError: If the aggregation_functions parameter is neither a list nor a string.
        KeyError: If an unsupported aggregation function is provided.

        Returns:
        -
        str: The converted formula string.
        """

        # Split the base string using backticks (`) as the delimiter
        parts = base_string.split('`')

        if isinstance(aggregation_functions, list):
            # Check if the number of non-empty strings between backticks is equal to the length of the aggregation functions list
            if len([part for part in parts if part]) - 1 != len(aggregation_functions):
                raise ValueError("Number of non-empty strings between backticks must be equal to the length of the aggregation functions list.")
        elif not isinstance(aggregation_functions, str):
            raise TypeError('aggregation function must be of type list or str')

        # Construct the final formula using the AGGREGATE function
        result = "=("
        for i, part in enumerate(parts):
            if i % 2 == 0:  # Even indices contain regular text
                result += part
            else:
                # Odd indices contain the column name between backticks
                try:
                    if isinstance(aggregation_functions, list):
                        result += f"_xlfn.AGGREGATE({AGG_CONVERSION[aggregation_functions[i // 2].upper()]}, 5, {table_name}[{part}])"
                    elif isinstance(aggregation_functions, str):
                        result += f"_xlfn.AGGREGATE({AGG_CONVERSION[aggregation_functions.upper()]}, 5, {table_name}[{part}])"
                except KeyError as key_error:
                    raise KeyError(f'no {key_error} aggregation function')

        result += ')'

        return result


    # Use _column_ in condition for current value.
    def add_column_formats(self, columns: str | list ='', format: str | list ='', condition: str | list =None):
        """
        Information
        -
        Add formatting to specified columns in the Excel worksheet.

        Parameters
        -
        columns (str or list of str): The name of the column(s) to which formatting will be applied.
        format (str or list of str): The formatting to apply to the specified column(s).
        condition (str or list of str, optional): The condition(s) for applying the formatting. 
            If specified, it must be parallel to the 'format' parameter.

        Raises
        -
        Exception: If 'format' and 'condition' are not parallel in terms of their types and lengths.
        Exception: If 'columns' parameter is neither a column name nor a list of column names.
        KeyError: If the sheet name specified in 'columns' is not found.

        Notes
        -
        - If 'columns' is a list of column names, 'format' and 'condition' should also be lists of equal lengths.
        - If 'condition' is not specified, formatting will be applied to all cells in the specified column(s).
        - If 'condition' is specified, formatting will be applied based on the condition(s) provided.
        """
        try:
            if (isinstance(format, list) and isinstance(condition, list) and len(format) != len(condition)) or (isinstance(format, list) and isinstance(condition, str)) or (isinstance(format, str) and isinstance(condition, list)):
                raise Exception('format and condition must be parallel')
            if isinstance(columns, list):
                for col in columns:
                    self.formats[col] = format
                    self.conditions[col] = condition
            elif isinstance(columns, str):
                self.formats[columns] = format
                self.conditions[columns] = condition
            else:
                raise Exception('columns parameter must be a column name or list of column names')
        except KeyError as key_error:
            raise KeyError(f'No sheet named {key_error}')
        except:
            print('error')


    # Use _column_ in condition for current value.
    def add_total_formats(self, columns: str | list ='', format: str | list ='', condition: str | list =None):
        """
        Information
        -
        Add formatting to specified total columns in the Excel worksheet.

        Parameters
        -
        columns (str or list of str): The name of the column(s) to which formatting will be applied.
        format (str or list of str): The formatting to apply to the specified column(s).
        condition (str or list of str, optional): The condition(s) for applying the formatting. 
            If specified, it must be parallel to the 'format' parameter.

        Raises
        -
        Exception: If 'format' and 'condition' are not parallel in terms of their types and lengths.
        Exception: If 'columns' parameter is neither a column name nor a list of column names.
        KeyError: If the sheet name specified in 'columns' is not found.

        Notes
        -
        - If 'columns' is a list of column names, 'format' and 'condition' should also be lists of equal lengths.
        - If 'condition' is not specified, formatting will be applied to all cells in the specified column(s).
        - If 'condition' is specified, formatting will be applied based on the condition(s) provided.
        """
        try:
            if (isinstance(format, list) and isinstance(condition, list) and len(format) != len(condition)) or (isinstance(format, list) and isinstance(condition, str)) or (isinstance(format, str) and isinstance(condition, list)):
                raise Exception('format and condition must be parallel')
            if isinstance(columns, list):
                for col in columns:
                    self.total_formats[col] = format
                    self.total_conditions[col] = condition
            elif isinstance(columns, str):
                self.total_formats[columns] = format
                self.total_conditions[columns] = condition
            else:
                raise Exception('columns parameter must be a column name or list of column names')
        except KeyError as key_error:
            raise KeyError(f'No sheet named {key_error}')
        except:
            print('error')

            
    def add_total_operation(self, columns: str | list, agg_function: str | list =None, operation_function: str =None):
        """
        Information
        -
        Add total operation to specified columns in the Excel worksheet.

        Parameters
        -
        columns (str or list of str): The name of the column(s) to which total operation will be applied.
        agg_function (str or list of str, optional): The aggregation function(s) to be used for totaling the specified column(s).
        operation_function (str, optional): The operation function to be applied to the totaled column(s).

        Raises
        -
        ValueError: If 'operation_function' is specified without 'agg_function'.
        ValueError: If neither 'agg_function' nor 'operation_function' is provided.
        ValueError: If 'agg_function' is a list and 'operation_function' is not None.
        Exception: If 'columns' parameter is neither a column name nor a list of column names.
        KeyError: If the sheet name specified in 'columns' is not found.

        Notes
        -
        - If 'columns' is a list of column names, 'agg_function' should be a string or a list of strings corresponding to each column.
        - If 'agg_function' is specified without 'operation_function', the default operation is used (e.g., sum).
        - If 'operation_function' is specified, 'agg_function' must also be specified.
        """
        if agg_function is None and operation_function is not None:
            raise ValueError("Parameter operation_function requires use of agg_function")
        if agg_function is None and operation_function is None:
            raise ValueError("Parameter agg_function or operation_function is required")
        if agg_function is not None and operation_function is None and isinstance(agg_function, list):
            raise ValueError("Parameter agg_function must be string without operation_function")
        try:
            if isinstance(columns, list):
                for col in columns:
                    self.agg_functions[col] = agg_function
                    self.operation_functions[col] = operation_function
            elif isinstance(columns, str):
                self.agg_functions[columns] = agg_function
                self.operation_functions[columns] = operation_function
            else:
                raise Exception('columns parameter must be a column name or list of column names')
        except KeyError as key_error:
            raise KeyError(f'No sheet named {key_error}')
        

    def add_links(self, column: str, link_mapping: dict):
        """
        Information
        -
        Add hyperlink mappings to the specified column in the Excel Table.

        Parameters
        -
        column (str): The name of the column to which hyperlink mappings will be added.
        link_mapping (dict): A dictionary containing the mapping of cell values to hyperlink URLs.

        Notes
        -
        - The 'link_mapping' dictionary should have cell values as keys and corresponding hyperlink URLs as values.
        - Hyperlinks will be applied to cells in the specified 'column' based on the mappings provided in 'link_mapping'.
        """
        self.links[column] = makehash()
        for value in link_mapping.keys():
            self.links[column][value] = link_mapping[value]

    
    def add_data_validation(self, column, dropdown_values, condition =None, type: str ='list', criteria =None, minimum =None, maximum =None, input_title =None, input_message =None, error_title =None, error_message =None, error_type ='information'):
        """
        Information
        -
        Add data validation to the specified column in the Excel worksheet.

        Parameters
        -
        column (str): The name of the column to which data validation will be applied.
        dropdown_values (list): A list of dropdown values for the data validation.
        condition (callable, optional): A condition function to apply additional validation rules.
        type (str, optional): The type of data validation to apply. Default is 'list'.
        criteria (str, optional): The criteria to apply for the data validation.
        minimum (int or float, optional): The minimum value for numeric data validation.
        maximum (int or float, optional): The maximum value for numeric data validation.
        input_title (str, optional): The title for the input message.
        input_message (str, optional): The input message displayed when the cell is selected.
        error_title (str, optional): The title for the error message.
        error_message (str, optional): The error message displayed when invalid data is entered.
        error_type (str, optional): The type of error message. Default is 'information'.

        Notes
        -
        - Data validation can be applied to restrict input based on a list of dropdown values, numerical criteria, or custom conditions.
        - Additional validation rules can be applied using the 'condition' parameter, which should be a callable accepting the cell value as input and returning True or False.
        """
        self.data_validation[column] = dropdown_values
        self.data_validation_condition[column] = condition
        self.data_validation_extras[column] = {
            'validate': type,
            'value': dropdown_values,
            'criteria': criteria,
            'minimum': minimum,
            'maximum': maximum,
            'input_title': input_title,
            'input_message': input_message,
            'error_title': error_title,
            'error_message': error_message,
            'error_type': error_type
        }

    
    def hide_columns(self, columns: str | list =''):
        if self.hidden_columns is None:
            if isinstance(columns, str):
                self.hidden_columns = [columns]
            elif isinstance(columns, list):
                self.hidden_columns = columns
        elif self.hidden_columns is not None:
            if isinstance(columns, str):
                self.hidden_columns.append(columns)
            elif isinstance(columns, list):
                self.hidden_columns += columns



    
    def _get_table_dimensions(self):
        """
        Information
        -
        Retrieve the dimensions of the data table in the Excel worksheet.

        Returns
        -
        tuple: A tuple containing the following dimensions:
            - data_row (int): The starting row index of the data table.
            - start_col (str): The starting column index of the data table (e.g., 'A', 'B', etc.).
            - end_data_row (int): The ending row index of the data table.
            - end_col (str): The ending column index of the data table (e.g., 'A', 'B', etc.).
        """
        return (
            self.data_row,
            self.start_col,
            self.end_data_row,
            self.end_col,
        )

    
    def set_worksheet(self, worksheet: xlsxwriter.Workbook.worksheet_class):
        """
        Information
        -
        Set the worksheet for the Excel workbook.

        Parameters
        -
        worksheet (worksheet): The worksheet object to be set.

        Notes
        -
        - This method sets the worksheet object for the Excel workbook, allowing subsequent operations to be performed on it.
        """
        self.worksheet = worksheet




class Chart:

    """
    Represents a chart object with various properties and series data.

    Parameters
    -
    name (str): The name of the chart.
    title (str): The title of the chart.
    title_size (int): The font size of the chart title.
    type (str): The type of the chart (e.g., 'line', 'bar', etc.).
    x_axis_name (str): The name of the x-axis.
    y_axis_name (str): The name of the primary y-axis.
    y2_axis_name (str): The name of the secondary y-axis.
    location (tuple): The location of the chart within a worksheet (row, column).
    size (tuple): The size of the chart within a worksheet (width, height).

    Attributes
    -
    series (list): A list containing series data for the chart.
    combined (list): A list containing combined series data for the chart.
    y2_axis_sw (bool): A boolean indicating whether the secondary y-axis is enabled.
    """

    def __init__(self,
                 name: str ='', 
                 title: str ='',
                 title_size: int =16,
                 type: str ='',
                 subtype: bool | str =None,
                 x_axis_name: str ='',
                 y_axis_name: str ='',
                 y2_axis_name: str ='',
                 location: tuple =(0, 0),
                 size: tuple =(480, 288),
                 x_axis_text_rotation: int = 0):
        
        self.name = name
        self.title = title
        self.title_size = title_size
        self.type = type
        self.subtype = subtype
        self.x_axis_name = x_axis_name
        self.y_axis_name = y_axis_name
        self.y2_axis_name = y2_axis_name
        self.location = location
        self.size = size
        self.x_axis_text_rotation = x_axis_text_rotation

        self.series = []
        self.combined = []

        self.y2_axis_sw = False
    

    def add_value(self, table: Table, category: str, column: str, y2_axis: bool =False, marker: str ='none', color: str =None):
        """
        Information
        -
        Add a value to the chart based on the specified table, category, and column.

        Parameters
        -
        table (Table): The Table object containing the data.
        category (str): The category column used for the x-axis.
        column (str): The column whose values are plotted on the chart.
        y2_axis (bool, optional): Indicates whether the series should be plotted on the secondary y-axis. Default is False.
        marker (str, optional): The marker type for the series. Default is 'none'.
        color (str, optional): The color for markers, fill, and line. Default is None.

        Raises
        -
        ValueError: If attempting to add a series to the secondary y-axis without enabling it first.

        """
        if self.y2_axis_sw is False and y2_axis is True:
            self.y2_axis_sw = True

        name_letter = xlsxwriter.utility.xl_col_to_name(table.header.index(column) + table.start_col)
        category_letter = xlsxwriter.utility.xl_col_to_name(table.header.index(category) + table.start_col)
        marker_params = {'type': marker}
        if color is not None:
            marker_params['fill'] = {'color': color}
            marker_params['border'] = {'color': color}
        
        fill_params = {'color': color} if color is not None else makehash()
        line_params = {'color': color} if color is not None else makehash()
        self.series.append(
                {
                    'name':         f"='{table.worksheet}'!${name_letter}${table.header_row + 1}",
                    'categories':   f"='{table.worksheet}'!${category_letter}${table.start_row}:${category_letter}${table.data_row}",
                    'values':       f"='{table.worksheet}'!${name_letter}${table.start_row}:${name_letter}${table.data_row}",
                    'y2_axis':      y2_axis,
                    'marker':       marker_params,
                    'fill':         fill_params,
                    'line':         line_params
                }
            )

    def combine(self, chart: Chart):
        """
        Information
        -
        Combine another chart with this chart.

        Parameters
        -
        chart (Chart): The chart object to combine with this chart.

        """
        self.combined.append(chart)


class CollapsibleTable:
    
    def __init__(self,
                 header: list,
                 data: dict[str: pd.DataFrame],
                 title: str =None,
                 data_properties: dict[str: dict['hidden': bool]] =None,
                 start_col: int =0,
                 title_row: int =1, 
                 header_row: int =1, 
                 data_row: int =2, 
                 total_row: str ='bottom', 
                 freeze_panes: int | tuple =None,
                 title_align: str = 'center',
                 header_comments: dict[str: str] = makehash()):
        self.helper = Helper
        self.title = title
        self.data = data
        self.start_col = start_col
        self.title_row = title_row
        self.header_row = header_row
        self.data_row = data_row
        self.start_row = data_row + 1
        self.header = header
        self.end_col = len(self.header) - 1 + start_col
        self.freeze_panes = freeze_panes
        self.hidden_columns = None
        self.title_align = title_align
        self.total_row = total_row
        self.formats = makehash()
        self.conditions = makehash()
        self.total_formats = makehash()
        self.total_conditions = makehash()
        self.operation_functions = makehash()
        self.links = makehash()
        self.data_validation = makehash()
        self.data_validation_condition = makehash()
        self.data_validation_extras = makehash()
        self.header_comments = header_comments
        self.levels = list(data.keys())
        self.data_properties = data_properties


    # Use _column_ in condition for current value.
    def add_total_formats(self, columns: str | list ='', format: str | list ='', condition: str | list =None):
        """
        Information
        -
        Add formatting to specified total columns in the Excel worksheet.

        Parameters
        -
        columns (str or list of str): The name of the column(s) to which formatting will be applied.
        format (str or list of str): The formatting to apply to the specified column(s).
        condition (str or list of str, optional): The condition(s) for applying the formatting. 
            If specified, it must be parallel to the 'format' parameter.

        Raises
        -
        Exception: If 'format' and 'condition' are not parallel in terms of their types and lengths.
        Exception: If 'columns' parameter is neither a column name nor a list of column names.
        KeyError: If the sheet name specified in 'columns' is not found.

        Notes
        -
        - If 'columns' is a list of column names, 'format' and 'condition' should also be lists of equal lengths.
        - If 'condition' is not specified, formatting will be applied to all cells in the specified column(s).
        - If 'condition' is specified, formatting will be applied based on the condition(s) provided.
        """
        try:
            if (isinstance(format, list) and isinstance(condition, list) and len(format) != len(condition)) or (isinstance(format, list) and isinstance(condition, str)) or (isinstance(format, str) and isinstance(condition, list)):
                raise Exception('format and condition must be parallel')
            if isinstance(columns, list):
                for col in columns:
                    self.total_formats[col] = format
                    self.total_conditions[col] = condition
            elif isinstance(columns, str):
                self.total_formats[columns] = format
                self.total_conditions[columns] = condition
            else:
                raise Exception('columns parameter must be a column name or list of column names')
        except KeyError as key_error:
            raise KeyError(f'No sheet named {key_error}')
        except:
            print('error')


    def add_total(self, line: pd.DataFrame):
        """
        Information
        -
        Add total operation to specified columns in the Excel worksheet.

        Parameters
        -
        columns (str or list of str): The name of the column(s) to which total operation will be applied.
        agg_function (str or list of str, optional): The aggregation function(s) to be used for totaling the specified column(s).
        operation_function (str, optional): The operation function to be applied to the totaled column(s).

        Raises
        -
        ValueError: If 'operation_function' is specified without 'agg_function'.
        ValueError: If neither 'agg_function' nor 'operation_function' is provided.
        ValueError: If 'agg_function' is a list and 'operation_function' is not None.
        Exception: If 'columns' parameter is neither a column name nor a list of column names.
        KeyError: If the sheet name specified in 'columns' is not found.

        Notes
        -
        - If 'columns' is a list of column names, 'agg_function' should be a string or a list of strings corresponding to each column.
        - If 'agg_function' is specified without 'operation_function', the default operation is used (e.g., sum).
        - If 'operation_function' is specified, 'agg_function' must also be specified.
        """
        print(len(line))
        print(line)
        if len(line) > 1:
            raise Exception('Total must be of length 1')
        else:
            self.total_row_values = line


    # Use _column_ in condition for current value.
    def add_column_formats(self, levels: str | list, columns: str | list ='', format: str | list ='', condition: str | list =None):
        """
        Information
        -
        Add formatting to specified columns in the Excel worksheet.

        Parameters
        -
        columns (str or list of str): The name of the column(s) to which formatting will be applied.
        format (str or list of str): The formatting to apply to the specified column(s).
        condition (str or list of str, optional): The condition(s) for applying the formatting. 
            If specified, it must be parallel to the 'format' parameter.

        Raises
        -
        Exception: If 'format' and 'condition' are not parallel in terms of their types and lengths.
        Exception: If 'columns' parameter is neither a column name nor a list of column names.
        KeyError: If the sheet name specified in 'columns' is not found.

        Notes
        -
        - If 'columns' is a list of column names, 'format' and 'condition' should also be lists of equal lengths.
        - If 'condition' is not specified, formatting will be applied to all cells in the specified column(s).
        - If 'condition' is specified, formatting will be applied based on the condition(s) provided.
        """
        try:
            if (isinstance(format, list) and isinstance(condition, list) and len(format) != len(condition)) or (isinstance(format, list) and isinstance(condition, str)) or (isinstance(format, str) and isinstance(condition, list)):
                raise Exception('format and condition must be parallel')
            if isinstance(columns, list):
                if isinstance(levels, str):
                    for col in columns:
                        self.formats[levels][col] = format
                        self.conditions[levels][col] = condition
                elif isinstance(levels, list):
                    for level in levels:
                        for col in columns:
                            self.formats[level][col] = format
                            self.conditions[level][col] = condition
            elif isinstance(columns, str):
                if isinstance(levels, str):
                    self.formats[levels][columns] = format
                    self.conditions[levels][columns] = condition
                elif isinstance(levels, list):
                    for level in levels:
                        self.formats[level][col] = format
                        self.conditions[level][col] = condition
            else:
                raise Exception('columns parameter must be a column name or list of column names')
        except KeyError as key_error:
            raise KeyError(f'No sheet named {key_error}')
        except:
            print('error')
        

    def add_links(self, level: str, column: str, link_mapping: dict):
        """
        Information
        -
        Add hyperlink mappings to the specified column in the Excel Table.

        Parameters
        -
        column (str): The name of the column to which hyperlink mappings will be added.
        link_mapping (dict): A dictionary containing the mapping of cell values to hyperlink URLs.

        Notes
        -
        - The 'link_mapping' dictionary should have cell values as keys and corresponding hyperlink URLs as values.
        - Hyperlinks will be applied to cells in the specified 'column' based on the mappings provided in 'link_mapping'.
        """
        self.links[level][column] = makehash()
        for value in link_mapping.keys():
            self.links[level][column][value] = link_mapping[value]

    
    def add_data_validation(self, column, level, dropdown_values, condition =None, type: str ='list', criteria =None, minimum =None, maximum =None, input_title =None, input_message =None, error_title =None, error_message =None, error_type ='information'):
        """
        Information
        -
        Add data validation to the specified column in the Excel worksheet.

        Parameters
        -
        column (str): The name of the column to which data validation will be applied.
        dropdown_values (list): A list of dropdown values for the data validation.
        condition (callable, optional): A condition function to apply additional validation rules.
        type (str, optional): The type of data validation to apply. Default is 'list'.
        criteria (str, optional): The criteria to apply for the data validation.
        minimum (int or float, optional): The minimum value for numeric data validation.
        maximum (int or float, optional): The maximum value for numeric data validation.
        input_title (str, optional): The title for the input message.
        input_message (str, optional): The input message displayed when the cell is selected.
        error_title (str, optional): The title for the error message.
        error_message (str, optional): The error message displayed when invalid data is entered.
        error_type (str, optional): The type of error message. Default is 'information'.

        Notes
        -
        - Data validation can be applied to restrict input based on a list of dropdown values, numerical criteria, or custom conditions.
        - Additional validation rules can be applied using the 'condition' parameter, which should be a callable accepting the cell value as input and returning True or False.
        """
        self.data_validation[level][column] = dropdown_values
        self.data_validation_condition[level][column] = condition
        self.data_validation_extras[level][column] = {
            'validate': type,
            'value': dropdown_values,
            'criteria': criteria,
            'minimum': minimum,
            'maximum': maximum,
            'input_title': input_title,
            'input_message': input_message,
            'error_title': error_title,
            'error_message': error_message,
            'error_type': error_type
        }


    
    def hide_columns(self, columns: str | list =''):
        if self.hidden_columns is None:
            if isinstance(columns, str):
                self.hidden_columns = [columns]
            elif isinstance(columns, list):
                self.hidden_columns = columns
        elif self.hidden_columns is not None:
            if isinstance(columns, str):
                self.hidden_columns.append(columns)
            elif isinstance(columns, list):
                self.hidden_columns += columns

    
    def set_worksheet(self, worksheet: xlsxwriter.Workbook.worksheet_class):
        """
        Information
        -
        Set the worksheet for the Excel workbook.

        Parameters
        -
        worksheet (worksheet): The worksheet object to be set.

        Notes
        -
        - This method sets the worksheet object for the Excel workbook, allowing subsequent operations to be performed on it.
        """
        self.worksheet = worksheet



class ExcelWriter:
    """
    Information
    -
    Represents an Excel writer object used to create and manipulate Excel workbooks.

    Parameters:
    -
    excel_path (str): The file path where the Excel workbook will be saved.

    Attributes:
    -
    workbook: An xlsxwriter Workbook object representing the Excel workbook being created.
    worksheet: A dictionary to store worksheet objects.
    excel_path (str): The file path where the Excel workbook will be saved.
    """
    
    def __init__(self, excel_path):
        self.workbook = xlsxwriter.Workbook(excel_path, {"nan_inf_to_errors": True})
        self.worksheet = makehash()
        self.excel_path = excel_path

        self.build_excel_workbook()

    
    def build_excel_workbook(self):
        """
        Information
        -
        Builds the excel workbook xlsxwriter class
        """
        
        self.workbook = xlsxwriter.Workbook(self.excel_path, {"nan_inf_to_errors": True})


    def get_worksheet(self, sheet_name):
        """
        Information
        -
        Retrieves a worksheet object from the Excel workbook.

        Parameters:
        -
        sheet_name (str): The name of the worksheet to retrieve.

        Returns:
        -
        Worksheet: A worksheet object representing the specified worksheet in the Excel workbook.
        """
        return self.worksheet[sheet_name]['worksheet']


    def add_worksheet(self, sheet_name):
        """
        Information
        -
        Adds a new worksheet to the Excel workbook.

        Parameters:
        -
        sheet_name (str): The name of the new worksheet to add.
        """
        self.worksheet[sheet_name] = makehash()
        self.worksheet[sheet_name]['worksheet'] = self.workbook.add_worksheet(sheet_name)
        self.worksheet[sheet_name]['tables'] = makehash()
        self.worksheet[sheet_name]['charts'] = makehash()


    def hide_worksheet(self, worksheet: str | list):
        """
        Information
        -
        Hides the specified worksheet(s) in the Excel workbook.

        Parameters:
        -
        worksheet (str or list): The name of the worksheet to hide, or a list of worksheet names to hide.

        Notes:
        -
        - If the input is a string, it hides the specified worksheet.
        - If the input is a list, it hides each worksheet in the list.

        """
        if isinstance(worksheet, str):
            self.worksheet[worksheet]['worksheet'].hide()
        if isinstance(worksheet, list):
            for sheet in worksheet:
                self.worksheet[sheet]['worksheet'].hide()


    def _get_excel_format(self, format):
        """
        Information
        -
        helper function used to access excel formats

        Parameters
        -
        format: the name of the format that you want to access
            Bold
            BoldWrapped
            MergeCenter
            TwoDecimalNum
            Link
            TwoDecimalNumBold
            PercentBold
            Percent
            OrangePercent
            RedPercent
            Blank
            RedBackground
            GreenBackground
            YellowBackground
            DateFormat

        Returns
        -
        the chosen excel format
        """
        #excel cell formatting
        workbook = self.workbook
        if format == "Bold": return workbook.add_format({'bold': True})
        if format == "BoldWrapped": return workbook.add_format({'text_wrap': True, 'bold': True})
        if format == "MergeCenter": return workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center'})
        if format == "MergeLeft": return workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'left'})
        if format == "MergeRight": return workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'right'})
        if format == "BoldNum": return workbook.add_format({'bold': True, 'num_format': '#,##0'})
        if format == "Num": return workbook.add_format({'num_format': '#,##0'})
        if format == "TwoDecimalNum": return workbook.add_format({'num_format': '#,##0.00'})
        if format == "Link": return workbook.add_format({'underline': True, 'font_color': 'blue'})
        if format == "TwoDecimalNumBold": return workbook.add_format({'bold': True, 'num_format': '#,##0.00'})
        if format == "PercentBold": return workbook.add_format({'bold': True, 'num_format': '0%'})
        if format == "Percent": return workbook.add_format({'num_format': '0%'})
        if format == "OrangePercent": return workbook.add_format({'font_color': 'orange', 'num_format': '0%'})
        if format == "RedPercent": return workbook.add_format({'font_color': 'red', 'num_format': '0%'})
        if format == "Blank": return workbook.add_format()
        if format == "RedBackground": return workbook.add_format({'bg_color' : '#F25757', 'num_format': '#,##0.00'})
        if format == "GreenBackground": return workbook.add_format({'bg_color' : '#57f257', 'num_format': '#,##0.00'})
        if format == "YellowBackground": return workbook.add_format({'bg_color' : '#f2f257', 'num_format': '#,##0.00'})
        if format == "DateFormat": return workbook.add_format({'num_format': 'mm/dd/yyyy'})
        if format == "AccountingNoSign": return workbook.add_format({'num_format': 43})
        if format == "AccountingNoSignBold": return workbook.add_format({'num_format': 43, 'bold': True})
        if format == "AccountingNoSignGreen": return workbook.add_format({'num_format': 43, 'bg_color' : '#57f257'})
        if format == "AccountingNoSignRed": return workbook.add_format({'num_format': 43, 'bg_color' : '#F25757'})
        if format == "Accounting": return workbook.add_format({'num_format': 44})
        if format == "AccountingBold": return workbook.add_format({'num_format': 44, 'bold': True})
        if format == "AccountingGreen": return workbook.add_format({'num_format': 44, 'bg_color' : '#57f257'})
        if format == "AccountingRed": return workbook.add_format({'num_format': 44, 'bg_color' : '#F25757'})
        if format == "Datetime": return workbook.add_format({'num_format': 'm/d/yyyy h:mm:ss AM/PM'})
        if format == "DatetimeBold": return workbook.add_format({'num_format': 'm/d/yyyy h:mm:ss AM/PM', 'bold': True})
        if format == "Date": return workbook.add_format({'num_format': 'm/d/yyyy'})
        if format == "DateBold": return workbook.add_format({'num_format': 'm/d/yyyy', 'bold': True})
        if format == "Time": return workbook.add_format({'num_format': 'h:mm:ss AM/PM'})
        if format == "TimeBold": return workbook.add_format({'num_format': 'h:mm:ss AM/PM', 'bold': True})

    
    def add_table(self, table: Table, sheet_name: str):
        """
        Information
        -
        Adds a table to the specified worksheet in the Excel workbook.

        Parameters:
        -
        table (Table): The Table object representing the table to be added.
        sheet_name (str): The name of the worksheet where the table will be added.

        Notes:
        -
        - The table object should contain the necessary information and data to construct the table.
        - If the table data is empty, it adds a row with 'No Results' to display in the table.
        - It handles different cases for the total_row parameter ('top' or 'bottom') to position the total row accordingly.
        - It formats the table header and merges the title cell if provided.
        - It adds the table to the worksheet and fills in the data from the table object.
        - It applies formatting, links, and data validation (dropdowns) as specified in the table object.
        - It calculates and writes aggregate functions for specified columns if provided.
        """

        def _write_total_formula(formatting: bool):
            if formatting:
                if col in table.agg_functions.keys() and col not in table.operation_functions.keys() and isinstance(table.agg_functions[col], str):
                    worksheet.write_formula(total_row, table.start_col + table.header.index(col), 
                                                    f'=_xlfn.AGGREGATE({str(AGG_CONVERSION.get(table.agg_functions[col]))}, 5, {table.name}[{col}])', self._get_excel_format(table.total_formats[col]))
                elif col in table.agg_functions.keys() and col in table.agg_functions.keys() and isinstance(table.agg_functions[col], str):
                    if table.operation_functions[col] is None:
                        worksheet.write_formula(total_row, table.start_col + table.header.index(col), 
                                                    table._convert(_column_, table.name, table.agg_functions[col]), self._get_excel_format(table.total_formats[col]))
                    else:
                        worksheet.write_formula(total_row, table.start_col + table.header.index(col), 
                                                    table._convert(table.operation_functions[col], table.name, table.agg_functions[col]), self._get_excel_format(table.total_formats[col]))
                elif col in table.agg_functions.keys() and isinstance(table.agg_functions[col], list):
                    worksheet.write_formula(total_row, table.start_col + table.header.index(col), 
                                                    table._convert(table.operation_functions[col], table.name, table.agg_functions[col]), self._get_excel_format(table.total_formats[col]))
                else:
                    worksheet.write(total_row, table.start_col + table.header.index(col), '', self._get_excel_format(table.total_formats[col]))
            else:
                if col in table.agg_functions.keys() and col not in table.operation_functions.keys() and isinstance(table.agg_functions[col], str):
                    worksheet.write_formula(total_row, table.start_col + table.header.index(col), 
                                                    f'=_xlfn.AGGREGATE({str(AGG_CONVERSION.get(table.agg_functions[col]))}, 5, {table.name}[{col}])')
                elif col in table.agg_functions.keys() and col in table.agg_functions.keys() and isinstance(table.agg_functions[col], str):
                    if table.operation_functions[col] is None:
                        worksheet.write_formula(total_row, table.start_col + table.header.index(col), 
                                                    table._convert(_column_, table.name, table.agg_functions[col]))
                    else:
                        worksheet.write_formula(total_row, table.start_col + table.header.index(col), 
                                                    table._convert(table.operation_functions[col], table.name, table.agg_functions[col]))
                elif col in table.agg_functions.keys() and isinstance(table.agg_functions[col], list):
                    worksheet.write_formula(total_row, table.start_col + table.header.index(col), 
                                                    table._convert(table.operation_functions[col], table.name, table.agg_functions[col]))
                else:
                    worksheet.write(total_row, table.start_col + table.header.index(col), '')

        table.set_worksheet(sheet_name)

        excel_header = []
        
        for col in table.data.columns:
            excel_header.append({'header': col, 
                        'header_format': self._get_excel_format('BoldWrapped'), 
                        'format': self._get_excel_format('TwoDecimalNum')})
            
        
        no_results_data = makehash()
        if table.data.empty:
            table.end_data_row += 1
            for index in range(len(table.data.columns)):
                if index == 0:
                    no_results_data[table.data.columns[index]] = 'No Results'
                else:
                    no_results_data[table.data.columns[index]] = ''

            table.data = pd.DataFrame([no_results_data])

        

        worksheet = self.worksheet[sheet_name]['worksheet']

        table_val = table.total_row.lower()
        use_header = table.header_row
        use_title = table.title
        if table.header_row is None:
            table.header_row = table.data_row
        if table.title_row is None:
            table.title_row = table.data_row

        if table_val == 'top':
            total_row = table.header_row
            table.total_row = table.header_row
            table.header_row += 1
            table.data_row += 1
            table.end_data_row += 1
        elif table_val == 'bottom':
            total_row = len(table.data) + table.data_row
        else:
            raise ValueError('total_row arguments must be of value "top" or "bottom"')
        


        if table.header != None:
            if table.freeze_panes != None:
                if isinstance(table.freeze_panes, int):
                    worksheet.freeze_panes(table.header_row + 1, table.freeze_panes)
                elif isinstance(table.freeze_panes, tuple):
                    worksheet.freeze_panes(table.freeze_panes[0], table.freeze_panes[1])
            if use_title is not None:
                if table.title_align.lower() == 'center':
                    worksheet.merge_range(f'{xlsxwriter.utility.xl_col_to_name(table.start_col)}{table.title_row}:{xlsxwriter.utility.xl_col_to_name(table.start_col + len(table.header) - 1)}{table.title_row}', table.title, self._get_excel_format("MergeCenter"))
                elif table.title_align.lower() == 'right':
                    worksheet.merge_range(f'{xlsxwriter.utility.xl_col_to_name(table.start_col)}{table.title_row}:{xlsxwriter.utility.xl_col_to_name(table.start_col + len(table.header) - 1)}{table.title_row}', table.title, self._get_excel_format("MergeRight"))
                elif table.title_align.lower() == 'left':
                    worksheet.merge_range(f'{xlsxwriter.utility.xl_col_to_name(table.start_col)}{table.title_row}:{xlsxwriter.utility.xl_col_to_name(table.start_col + len(table.header) - 1)}{table.title_row}', table.title, self._get_excel_format("MergeLeft"))
            if use_header is not None:
                for a in range(table.start_col, len(table.header) + table.start_col):
                    worksheet.write(table.header_row, a, table.header[a - table.start_col], self._get_excel_format("BoldWrapped"))
                    if table.header_comments.get(str(table.header[a - table.start_col])) is not None:
                        worksheet.write_comment(table.header_row, a, table.header_comments[table.header[a - table.start_col]])


        if use_header is not None:
            worksheet.add_table(table.header_row, 
                                table.start_col,
                                table.end_data_row,
                                table.end_col,
                                {'columns': excel_header,
                                'name': table.name})
        else:
            worksheet.add_table(table.header_row, 
                                table.start_col,
                                table.end_data_row,
                                table.end_col,
                                {'header_row': False,
                                'name': table.name})


        for index, row in table.data.iterrows():
            for col in table.data.columns:
                _column_ = row[col]
                # Write value
                if col in table.formats.keys() and table.conditions[col] is None:   
                    worksheet.write(table.data_row, table.start_col + table.header.index(col), row[col], self._get_excel_format(table.formats[col]))
                elif col in table.formats.keys() and table.conditions[col] is not None:
                    if isinstance(table.conditions[col], str):
                        if eval(table.helper._format_string(table.conditions[col])):
                            worksheet.write(table.data_row, table.start_col + table.header.index(col), row[col], self._get_excel_format(table.formats[col]))
                        else:
                            worksheet.write(table.data_row, table.start_col + table.header.index(col), row[col])
                    elif isinstance(table.conditions[col], list):
                        written = False
                        for index in range(len(table.conditions[col])):
                            if eval(table.helper._format_string(table.conditions[col][index])):
                                worksheet.write(table.data_row, table.start_col + table.header.index(col), row[col], self._get_excel_format(table.formats[col][index]))
                                written = True
                        if not written:
                            worksheet.write(table.data_row, table.start_col + table.header.index(col), row[col])
                else:
                    worksheet.write(table.data_row, table.start_col + table.header.index(col), row[col])

                # Write links
                if col in table.links.keys():
                    if row[col] in table.links[col].keys():
                        worksheet.write_url(table.data_row, table.start_col + table.header.index(col), table.links[col][row[col]], string=str(row[col]))
                
                # Write data validation (dropdown)
                if col in table.data_validation.keys() and table.data_validation_condition[col] is None:
                    data_validation = makehash()
                    for parameter in table.data_validation_extras[col].keys():
                        if table.data_validation_extras[col][parameter] != None:
                            data_validation[parameter] = table.data_validation_extras[col][parameter]
                    worksheet.data_validation(table.data_row, table.start_col + table.header.index(col), table.data_row, table.start_col + table.header.index(col), data_validation)
                elif col in table.data_validation.keys() and table.data_validation_condition[col] is not None:
                    if isinstance(table.data_validation_condition[col], str):
                        if eval(table.helper._format_string(table.data_validation_condition[col])):
                            data_validation = makehash()
                            for parameter in table.data_validation_extras[col].keys():
                                if table.data_validation_extras[col][parameter] != None:
                                    data_validation[parameter] = table.data_validation_extras[col][parameter]
                            worksheet.data_validation(table.data_row, table.start_col + table.header.index(col), table.data_row, table.start_col + table.header.index(col), data_validation)
                    if isinstance(table.data_validation_condition[col], list):
                        written = False
                        for index in range(len(table.data_validation_condition[col])):
                            if eval(table.helper._format_string(table.data_validation_condition[col][index])):
                                data_validation = makehash()
                                for parameter in table.data_validation_extras[col].keys():
                                    if table.data_validation_extras[col][parameter] != None:
                                        data_validation[parameter] = table.data_validation_extras[col][parameter]
                                worksheet.data_validation(table.data_row, table.start_col + table.header.index(col), table.data_row, table.start_col + table.header.index(col), data_validation)
                                written = True
                        if not written:
                            worksheet.write(table.data_row, table.start_col + table.header.index(col), row[col])
            table.data_row += 1

        # Write Total Row
        if table_val is not None:
            for col in table.data.columns:
                if use_header is None:
                    _column_ = '`Column' + str(table.header.index(col) + 1) + '`'
                else:
                    _column_ = '`' + col + '`'
                if col in table.total_formats.keys() and table.total_conditions[col] is None:
                    _write_total_formula(formatting=True)
                elif col in table.total_formats.keys() and table.total_conditions[col] is not None:
                    if isinstance(table.total_conditions[col], str):
                        if eval(table.helper._format_string(table.total_conditions[col])):
                            _write_total_formula(formatting=True)
                        else:
                            _write_total_formula(formatting=False)
                    elif isinstance(table.total_conditions[col], list):
                        written = False
                        for index in range(len(table.total_conditions[col])):
                            if eval(table.helper._format_string(table.total_conditions[col][index])):
                                _write_total_formula(formatting=True)
                                written = True
                        if not written:
                            _write_total_formula(formatting=False)
                else:
                    _write_total_formula(formatting=False)

        if table.hidden_columns is not None:
            for column in table.hidden_columns:
                col_val = table.header.index(column)
                worksheet.set_column(col_val, col_val, None, None, {'hidden': True})


    
    def add_chart(self, chart: Chart, sheet_name: str):
        """
        Information
        -
        Adds a chart to the specified worksheet in the Excel workbook.

        Parameters:
        -
        chart (Chart): The Chart object representing the chart to be added.
        sheet_name (str): The name of the worksheet where the chart will be added.

        Notes:
        -
        - The chart object should contain the necessary properties and series data to construct the chart.
        - If combining multiple charts, the combined attribute of the chart object should be populated.
        - It creates a new xlsxwriter chart object and adds series data from the chart object.
        - It sets the chart title, axis names, and secondary y-axis if specified in the chart object.
        - If combining multiple charts, it merges them into a single chart.
        - It inserts the finalized chart into the specified location on the worksheet.
        """
        worksheet = self.worksheet[sheet_name]['worksheet']

        chart_list = []

        chart.combined.append(chart)

        for charts in chart.combined:
            c = self.workbook.add_chart({'type': charts.type})
            for series in charts.series:
                c.add_series(series)

            c.set_title({'name': charts.title, 'name_font': {'size': charts.title_size}})
            c.set_x_axis({'name': charts.x_axis_name, 'num_font': {'rotation': charts.x_axis_text_rotation}})
            c.set_y_axis({'name': charts.y_axis_name})
            c.set_size({'width': charts.size[0], 'height': charts.size[1]})
            if charts.y2_axis_sw:
                c.set_y2_axis({'name': charts.y2_axis_name})

            chart_list.append(c)

        main_chart = chart_list[-1]
        chart_list.pop()
        for combine_chart in chart_list:
            main_chart.combine(combine_chart)
                
        worksheet.insert_chart(chart.location[0], chart.location[1], main_chart)



    def add_collapsible_table(self, table: CollapsibleTable, sheet_name: str):

        def _write_total_formula(formatting: bool):
            try:
                if formatting:
                    worksheet.write(total_row, table.start_col + table.header.index(col), table.total_row_values[col].iloc[0], self._get_excel_format(table.total_formats[col]))
                else:
                    worksheet.write(total_row, table.start_col + table.header.index(col), table.total_row_values[col].iloc[0])
            except:
                pass

        table.set_worksheet(sheet_name)
        worksheet = self.worksheet[sheet_name]['worksheet']


        table_val = table.total_row.lower()
        use_header = table.header_row
        use_title = table.title
        if table.header_row is None:
            table.header_row = table.data_row
        if table.title_row is None:
            table.title_row = table.data_row

        if table_val == 'top':
            total_row = table.header_row
            table.total_row = table.header_row
            table.header_row += 1
            table.data_row += 1
        elif table_val == 'bottom':
            total_row = len(table.data) + table.data_row
        else:
            raise ValueError('total_row arguments must be of value "top" or "bottom"')

        if table.header != None:
            if table.freeze_panes != None:
                if isinstance(table.freeze_panes, int):
                    worksheet.freeze_panes(table.header_row + 1, table.freeze_panes)
                elif isinstance(table.freeze_panes, tuple):
                    worksheet.freeze_panes(table.freeze_panes[0], table.freeze_panes[1])
            if use_title is not None:
                if table.title_align.lower() == 'center':
                    worksheet.merge_range(f'{xlsxwriter.utility.xl_col_to_name(table.start_col)}{table.title_row}:{xlsxwriter.utility.xl_col_to_name(table.start_col + len(table.header) - 1)}{table.title_row}', table.title, self._get_excel_format("MergeCenter"))
                elif table.title_align.lower() == 'right':
                    worksheet.merge_range(f'{xlsxwriter.utility.xl_col_to_name(table.start_col)}{table.title_row}:{xlsxwriter.utility.xl_col_to_name(table.start_col + len(table.header) - 1)}{table.title_row}', table.title, self._get_excel_format("MergeRight"))
                elif table.title_align.lower() == 'left':
                    worksheet.merge_range(f'{xlsxwriter.utility.xl_col_to_name(table.start_col)}{table.title_row}:{xlsxwriter.utility.xl_col_to_name(table.start_col + len(table.header) - 1)}{table.title_row}', table.title, self._get_excel_format("MergeLeft"))
            if use_header is not None:
                for a in range(table.start_col, len(table.header) + table.start_col):
                    worksheet.write(table.header_row, a, table.header[a - table.start_col], self._get_excel_format("BoldWrapped"))
                    if table.header_comments.get(str(table.header[a - table.start_col])) is not None:
                        worksheet.write_comment(table.header_row, a, table.header_comments[table.header[a - table.start_col]])

                worksheet.autofilter(table.header_row, table.start_col, table.header_row, table.end_col)


        def write_line(line_info, line_header, level, level_name):
            for col in line_header:
                _column_ = line_info[line_header.index(col)]
                # Write value
                if col in table.formats[level_name].keys() and table.conditions[level_name][col] is None:   
                    worksheet.write(table.data_row, table.start_col + table.header.index(col), line_info[line_header.index(col)], self._get_excel_format(table.formats[level_name][col]))
                elif col in table.formats[level_name].keys() and table.conditions[level_name][col] is not None:
                    if isinstance(table.conditions[level_name][col], str):
                        if eval(table.helper._format_string(table.conditions[level_name][col])):
                            worksheet.write(table.data_row, table.start_col + table.header.index(col), line_info[line_header.index(col)], self._get_excel_format(table.formats[level_name][col]))
                        else:
                            worksheet.write(table.data_row, table.start_col + table.header.index(col), line_info[line_header.index(col)])
                    elif isinstance(table.conditions[level_name][col], list):
                        written = False
                        for index in range(len(table.conditions[level_name][col])):
                            if eval(table.helper._format_string(table.conditions[level_name][col][index])):
                                worksheet.write(table.data_row, table.start_col + table.header.index(col), line_info[line_header.index(col)], self._get_excel_format(table.formats[level_name][col][index]))
                                written = True
                        if not written:
                            worksheet.write(table.data_row, table.start_col + table.header.index(col), line_info[line_header.index(col)])
                else:
                    worksheet.write(table.data_row, table.start_col + table.header.index(col), line_info[line_header.index(col)])

                # Write links
                if col in table.links.keys():
                    if line_info[line_header.index(col)] in table.links[level_name][col].keys():
                        worksheet.write_url(table.data_row, table.start_col + table.header.index(col), table.links[level_name][col][line_info[line_header.index(col)]], string=str(line_info[line_header.index(col)]))
                
                # Write data validation (dropdown)
                if col in table.data_validation[level_name].keys() and table.data_validation_condition[level_name][col] is None:
                    data_validation = makehash()
                    for parameter in table.data_validation_extras[level_name][col].keys():
                        if table.data_validation_extras[level_name][col][parameter] != None:
                            data_validation[parameter] = table.data_validation_extras[level_name][col][parameter]
                    worksheet.data_validation(table.data_row, table.start_col + table.header.index(col), table.data_row, table.start_col + table.header.index(col), data_validation)
                elif col in table.data_validation.keys() and table.data_validation_condition[level_name][col] is not None:
                    if isinstance(table.data_validation_condition[level_name][col], str):
                        if eval(table.helper._format_string(table.data_validation_condition[level_name][col])):
                            data_validation = makehash()
                            for parameter in table.data_validation_extras[level_name][col].keys():
                                if table.data_validation_extras[level_name][col][parameter] != None:
                                    data_validation[parameter] = table.data_validation_extras[level_name][col][parameter]
                            worksheet.data_validation(table.data_row, table.start_col + table.header.index(col), table.data_row, table.start_col + table.header.index(col), data_validation)
                    if isinstance(table.data_validation_condition[level_name][col], list):
                        written = False
                        for index in range(len(table.data_validation_condition[level_name][col])):
                            if eval(table.helper._format_string(table.data_validation_condition[level_name][col][index])):
                                data_validation = makehash()
                                for parameter in table.data_validation_extras[level_name][col].keys():
                                    if table.data_validation_extras[level_name][col][parameter] != None:
                                        data_validation[parameter] = table.data_validation_extras[level_name][col][parameter]
                                worksheet.data_validation(table.data_row, table.start_col + table.header.index(col), table.data_row, table.start_col + table.header.index(col), data_validation)
                                written = True
                        if not written:
                            worksheet.write(table.data_row, table.start_col + table.header.index(col), line_info[line_header.index(col)])
                if level != 0:
                    if table.data_properties is not None:
                        if table.data_properties.get(level_name) is not None:
                            if table.data_properties[level_name].get('hidden') is not None:
                                worksheet.set_row(table.data_row, None, None, {'level': level, 'hidden': table.data_properties[level_name]['hidden']})
                            else:
                                worksheet.set_row(table.data_row, None, None, {'level': level, 'hidden': True})
                        else:
                            worksheet.set_row(table.data_row, None, None, {'level': level, 'hidden': True})
                    else:
                        worksheet.set_row(table.data_row, None, None, {'level': level, 'hidden': True})
                else:
                    worksheet.set_row(table.data_row, None, None, {'level': level, 'hidden': False})
            table.data_row += 1

        def _recursive_data_looper(data_dict: dict, position: int, data: pd.DataFrame= None, filter: dict[str: str]= None):
            if filter is None:
                filter = makehash()
            if data is None:
                data = data_dict[list(data_dict.keys())[position]]

            if len(list(data_dict.keys())) == position:
                data = data_dict[list(data_dict.keys())[position - 1]]
                return
            else:
                while len(filter.keys()) != position:
                    filter.pop(list(data_dict.keys())[len(filter.keys()) - 1])
                if len(filter.keys()) > 0:
                    query_params = []
                    for key, value in filter.items():
                        query_params.append(f"`{key}` == '{value}'")
                    
                    data = data_dict[list(data_dict.keys())[position]].copy().query(' and '.join(query_params))
                else:
                    data = data_dict[list(data_dict.keys())[position]]

                for line in data.itertuples(index=False):
                    if position == 0:
                        filter = makehash()
                    if len(list(data_dict.keys())) != position + 1:
                        filter[list(data_dict.keys())[position]] = line[data.columns.get_loc(str(list(data_dict.keys())[position]))]
                    _recursive_data_looper(data_dict=data_dict, position=position+1, data=data, filter=filter)

                    write_line(line_info=line, line_header=list(data_dict[list(data_dict.keys())[position]].columns), level=position, level_name=str(list(data_dict.keys())[position]))

        _recursive_data_looper(table.data, 0)


        # Write Total Row
        if table_val is not None:
            for col in table.header:
                if use_header is None:
                    _column_ = '`Column' + str(table.header.index(col) + 1) + '`'
                else:
                    _column_ = '`' + col + '`'
                if col in table.total_formats.keys() and table.total_conditions[col] is None:
                    _write_total_formula(formatting=True)
                elif col in table.total_formats.keys() and table.total_conditions[col] is not None:
                    if isinstance(table.total_conditions[col], str):
                        if eval(table.helper._format_string(table.total_conditions[col])):
                            _write_total_formula(formatting=True)
                        else:
                            _write_total_formula(formatting=False)
                    elif isinstance(table.total_conditions[col], list):
                        written = False
                        for index in range(len(table.total_conditions[col])):
                            if eval(table.helper._format_string(table.total_conditions[col][index])):
                                _write_total_formula(formatting=True)
                                written = True
                        if not written:
                            _write_total_formula(formatting=False)
                else:
                    _write_total_formula(formatting=False)

        if table.hidden_columns is not None:
            for column in table.hidden_columns:
                col_val = table.header.index(column)
                worksheet.set_column(col_val, col_val, None, None, {'hidden': True})

        

        

        


    
    def extend_column_width(self, sheet_name: str, column_widths: list):
        """
        Information
        -
        Extend the width of columns in the specified Excel worksheet.

        Parameters
        -
        sheet_name (str): The name of the worksheet in which to extend column widths.
        column_widths (list): A list containing the width of each column.

        Notes
        -
        - The 'column_widths' list should contain the width for each column in the order they appear in the worksheet.
        - The width of each column will be set according to the corresponding value in the 'column_widths' list.
        """
        for col_num, width in enumerate(column_widths):
            self.worksheet[sheet_name]['worksheet'].set_column(col_num, col_num, width)

    
    def close_workbook(self):
        """
        Information
        -
        Closes the Excel workbook, saving any changes made.

        Notes:
        -
        - It is important to call this method to ensure that any changes made to the workbook are saved.
        """
        self.workbook.close()





def main():
    args = {'-i': '--dist'}
    excel_path = 'C:\\Users\\brian.capozza\\OneDrive - PCX Aerostructures, LLC\\Documents\\Python312\\github_repos\\PCXModule\\PCX\\XLSX\\test.xlsx'

    level_one_data = {'Computer Name': ['PCXNW-JCAPOZZA', 'PCXNW-BCAPOZZA', 'PCXNW-JCRUZ'], 
                      'CVE Count': [665, 113, 465]}
    level_two_data = {'Computer Name': ['PCXNW-JCAPOZZA', 'PCXNW-JCAPOZZA', 'PCXNW-JCAPOZZA', 'PCXNW-BCAPOZZA', 'PCXNW-BCAPOZZA', 'PCXNW-BCAPOZZA', 'PCXNW-BCAPOZZA', 'PCXNW-JCRUZ', 'PCXNW-JCRUZ'], 
                      'Application Name': ['.NET Framework 2.0', '.NET x64 6.0', '7-Zip 9.20.00.0', 'ASP .NET Core x64 2.2', '.NET x64 6.0', '7-Zip 9.20.00.0', 'MiKTeX 21.8', '7-Zip 9.20.00.0', 'ASP .NET Core x64 2.2'], 
                      'CVE Count': [22, 43, 64, 2, 46, 76, 45, 78, 65]}
    level_three_data = {'Computer Name': ['PCXNW-JCAPOZZA', 'PCXNW-JCAPOZZA', 'PCXNW-JCAPOZZA', 'PCXNW-JCAPOZZA', 'PCXNW-JCAPOZZA', 'PCXNW-BCAPOZZA', 'PCXNW-BCAPOZZA', 'PCXNW-BCAPOZZA', 'PCXNW-BCAPOZZA', 'PCXNW-BCAPOZZA', 'PCXNW-BCAPOZZA', 'PCXNW-JCRUZ', 'PCXNW-JCRUZ', 'PCXNW-JCRUZ', 'PCXNW-JCRUZ'], 
                        'Application Name': ['.NET Framework 2.0', '.NET Framework 2.0', '.NET x64 6.0', '.NET x64 6.0', '7-Zip 9.20.00.0', 'ASP .NET Core x64 2.2', 'ASP .NET Core x64 2.2', 'ASP .NET Core x64 2.2', '.NET x64 6.0', '7-Zip 9.20.00.0', 'MiKTeX 21.8', '7-Zip 9.20.00.0', '7-Zip 9.20.00.0', '7-Zip 9.20.00.0', 'ASP .NET Core x64 2.2'], 
                        'CVE ID': ['CVE-2021-1', 'CVE-2021-2', 'CVE-2021-3', 'CVE-2021-4', 'CVE-2021-5', 'CVE-2021-6', 'CVE-2021-7', 'CVE-2021-8', 'CVE-2021-9', 'CVE-2021-10', 'CVE-2021-11', 'CVE-2021-12', 'CVE-2021-13', 'CVE-2021-14', 'CVE-2021-15'], 
                        'CVE Count': [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]}
    data = {'Computer Name': pd.DataFrame(level_one_data), 'Application Name': pd.DataFrame(level_two_data), 'CVE ID': pd.DataFrame(level_three_data)}

    total_row = {'Computer Name': ['Total Row'], 'CVE Count': [8888]}


    my_excel = ExcelWriter(excel_path=excel_path)

    my_excel.add_worksheet('test_sheet')

    my_table = CollapsibleTable(title='Collapsible Table Test',
                                header=['Computer Name', 'Application Name', 'CVE ID', 'CVE Count'],
                                data=data,
                                total_row='top')
    my_table.add_total(line=pd.DataFrame(total_row))
    my_table.add_total_formats(columns='CVE Count', format='Bold')
    my_excel.add_collapsible_table(my_table, 'test_sheet')
    

    #my_chart = Chart('chart', 'Chart', type='line', x_axis_text_rotation=315)
    #my_chart.add_value(my_table, 'FiscalMonth', 'Efficiency', marker='circle')

    #my_excel.add_chart(my_chart, 'test_sheet')

    my_excel.close_workbook()


if __name__ == '__main__':
    main()