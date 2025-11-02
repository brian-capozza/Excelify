# Excelify

A powerful Python library for creating complex, formatted Excel workbooks with advanced features including dynamic tables, charts, conditional formatting, and collapsible hierarchical structures.

## 🚀 Features

### Tables
- **Dynamic Table Generation**: Create formatted Excel tables from pandas DataFrames
- **Conditional Formatting**: Apply formatting based on cell values using formula-based conditions
- **Aggregate Functions**: Built-in support for SUM, AVERAGE, COUNT, MIN, MAX, and more
- **Data Validation**: Add dropdown menus and data validation rules
- **Hyperlinks**: Automatic hyperlink generation with custom mapping
- **Custom Operations**: Define complex formulas using multiple aggregate functions

### Collapsible Tables
- **Hierarchical Data**: Create multi-level collapsible tables
- **Automatic Grouping**: Recursive data processing for nested relationships
- **Drill-Down Reporting**: Perfect for detailed analysis with summary views
- **Customizable Visibility**: Control default collapse states for each level

### Charts
- **Multiple Chart Types**: Line, bar, column, and more
- **Dual Y-Axis**: Support for secondary y-axis
- **Chart Combination**: Overlay multiple chart types
- **Dynamic Series**: Automatically generate series from table data
- **Custom Styling**: Full control over colors, markers, and formatting

### Formatting
- **Pre-built Formats**: Bold, currency, accounting, percentages, dates, and more
- **Conditional Application**: Apply formats based on cell values
- **Custom Number Formats**: Support for all Excel number format codes
- **Background Colors**: Red, green, yellow backgrounds for highlighting

## 📦 Installation

```bash
pip install xlsxwriter pandas
```

## 🔧 Quick Start

### Basic Table Example

```python
from Excel import ExcelWriter, Table
import pandas as pd

# Create sample data
data = pd.DataFrame({
    'Product': ['Widget A', 'Widget B', 'Widget C'],
    'Sales': [1000, 1500, 2000],
    'Profit': [200, 300, 400]
})

# Initialize Excel writer
excel = ExcelWriter('output.xlsx')
excel.add_worksheet('Sales Report')

# Create and configure table
table = Table(
    name='SalesTable',
    title='Q4 Sales Report',
    data=data,
    total_row='bottom'
)

# Add aggregate functions
table.add_total_operation(
    columns=['Sales', 'Profit'],
    agg_function='SUM'
)

# Add formatting
table.add_column_formats(
    columns=['Sales', 'Profit'],
    format='Accounting'
)

# Add table to worksheet
excel.add_table(table, 'Sales Report')
excel.close_workbook()
```

### Conditional Formatting Example

```python
# Apply green background when profit > 300
table.add_column_formats(
    columns='Profit',
    format='GreenBackground',
    condition='_column_ > 300'
)

# Multiple conditions
table.add_column_formats(
    columns='Status',
    format=['GreenBackground', 'RedBackground'],
    condition=[
        "_column_ == 'Complete'",
        "_column_ == 'Failed'"
    ]
)
```

### Collapsible Table Example

```python
from Excel import CollapsibleTable

# Hierarchical data structure
data = {
    'Department': pd.DataFrame({
        'Department': ['Engineering', 'Sales'],
        'Budget': [500000, 300000]
    }),
    'Team': pd.DataFrame({
        'Department': ['Engineering', 'Engineering', 'Sales'],
        'Team': ['Frontend', 'Backend', 'Enterprise'],
        'Headcount': [5, 8, 6]
    }),
    'Employee': pd.DataFrame({
        'Department': ['Engineering', 'Engineering', 'Sales'],
        'Team': ['Frontend', 'Backend', 'Enterprise'],
        'Employee': ['John Doe', 'Jane Smith', 'Bob Johnson'],
        'Salary': [120000, 130000, 110000]
    })
}

# Create collapsible table
collapsible = CollapsibleTable(
    title='Organization Structure',
    header=['Department', 'Team', 'Employee', 'Headcount', 'Salary', 'Budget'],
    data=data
)

excel.add_collapsible_table(collapsible, 'Organization')
```

### Chart Example

```python
from Excel import Chart

# Create chart
chart = Chart(
    name='SalesChart',
    title='Sales Performance',
    type='line',
    x_axis_name='Product',
    y_axis_name='Revenue'
)

# Add data series
chart.add_value(
    table=table,
    category='Product',
    column='Sales',
    marker='circle',
    color='#0066CC'
)

# Add chart to worksheet
excel.add_chart(chart, 'Sales Report')
```

## 📚 Available Formats

| Format | Description |
|--------|-------------|
| `Bold` | Bold text |
| `Num` | Number with thousands separator |
| `TwoDecimalNum` | Number with 2 decimal places |
| `Percent` | Percentage format |
| `Accounting` | Accounting number format with $ |
| `AccountingNoSign` | Accounting format without $ |
| `Date` | Date format (m/d/yyyy) |
| `Datetime` | Datetime format |
| `GreenBackground` | Green background highlight |
| `RedBackground` | Red background highlight |
| `YellowBackground` | Yellow background highlight |

## 🔑 Key Classes

### Table
Main class for creating Excel tables with formatting and formulas.

**Key Methods:**
- `add_column_formats()`: Apply formatting to columns
- `add_total_operation()`: Add aggregate functions to total row
- `add_links()`: Add hyperlinks to cells
- `add_data_validation()`: Add dropdown menus
- `hide_columns()`: Hide specific columns

### CollapsibleTable
Create hierarchical, collapsible tables for drill-down analysis.

**Key Methods:**
- `add_column_formats()`: Apply formatting by level
- `add_total()`: Add total row
- `add_links()`: Add hyperlinks by level

### Chart
Create and customize Excel charts.

**Key Methods:**
- `add_value()`: Add data series
- `combine()`: Combine multiple charts

### ExcelWriter
Main workbook manager.

**Key Methods:**
- `add_worksheet()`: Create new worksheet
- `add_table()`: Add table to worksheet
- `add_chart()`: Add chart to worksheet
- `hide_worksheet()`: Hide worksheets
- `close_workbook()`: Save and close

## 🎯 Use Cases

- **Financial Reporting**: Automated P&L statements, balance sheets, and KPI dashboards
- **Data Analysis**: Complex data analysis with conditional formatting and charts
- **Inventory Management**: Multi-level bill of materials and stock reports
- **Security Reporting**: CVE tracking and vulnerability assessment reports
- **Project Management**: Resource allocation and timeline tracking
- **Sales Analytics**: Territory performance and product analysis

## 👤 Author

Brian Capozza

## 🙏 Acknowledgments

Built with:
- [xlsxwriter](https://xlsxwriter.readthedocs.io/) - Excel file creation
- [pandas](https://pandas.pydata.org/) - Data manipulation

---

**Note**: This library requires Python 3.7+ and is designed for generating Excel files, not reading them.
