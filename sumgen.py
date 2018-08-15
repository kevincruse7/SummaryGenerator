"""
Processes for transferring data from proprietary portfolio review spreadsheet
to summary spreadsheet.

Author: Kevin Cruse
Version: 1.3
"""

from json import loads
from openpyxl import load_workbook


class SecurityReturn:
    """Return data aggregate for securities"""
    def __init__(self, name='', year_to_date=0.0, month_to_date=0.0,
                 quarter_to_date=0.0, one_year=0.0, three_year=0.0,
                 five_year=0.0):
        self.name = name
        self.year_to_date = year_to_date
        self.month_to_date = month_to_date
        self.quarter_to_date = quarter_to_date
        self.one_year = one_year
        self.three_year = three_year
        self.five_year = five_year


def read_workbook(name, read):
    """
    Reads in Excel workbook that user enters.

    Args:
        name: String prompt to give user to specify type of workbook.
        read: Boolean indicating whether to use read only mode or not.
        data: Boolean indicating whether to use data only mode or not.

    Returns:
        Tuple with openpyxl workbook object and string workbook filename.
    """
    typo = True
    while typo:
        workbook_filename = (input('Enter the ' + name + ' filename: ') +
                             '.xlsx')

        # Check if filename is valid
        try:
            workbook = load_workbook(report_path + workbook_filename,
                                     read_only=read, data_only=True)
            typo = False
        except FileNotFoundError:
            print('File does not exist.\n')

    return workbook, workbook_filename


def get_security_info(worksheet, names, name_column, start_column):
    """
    Gets security return data from ticker.

    Args:
        worksheet: Tuple of rows of review worksheet.
        names: Tickers or names of securities to research.
        name_column: Zero-indexed column number of tickers or names in review
                     worksheet.

    Returns:
        List of SecurityReturn objects with return data filled in.
    """
    # Get list of all names in review worksheet for connecting tickers to data
    all_names = tuple(row[name_column].value.lower()
                      if type(row[name_column].value) is str
                      else row[name_column].value for row in worksheet)

    # Connect ticker or name list to corresponding return data
    securities = []
    for name in names:
        row = all_names.index(name.lower())
        security = SecurityReturn(name.upper())

        security.year_to_date = (worksheet[row][start_column].value if
                                 type(worksheet[row][start_column].value) is
                                 int or
                                 type(worksheet[row][start_column].value) is
                                 float else 0.0)
        security.month_to_date = (worksheet[row][start_column + 1].value if
                                  type(worksheet[row][start_column + 1].value)
                                  is int or
                                  type(worksheet[row][start_column + 1].value)
                                  is float else 0.0)
        security.quarter_to_date = (worksheet[row][start_column + 2].value if
                                    type(worksheet[row][start_column + 2].value
                                         ) is int or
                                    type(worksheet[row][start_column + 2].value
                                         ) is float else 0.0)
        security.one_year = (worksheet[row][start_column + 3].value if
                             type(worksheet[row][start_column + 3].value) is
                             int or
                             type(worksheet[row][start_column + 3].value) is
                             float else 0.0)
        security.three_year = (worksheet[row][start_column + 4].value if
                               type(worksheet[row][start_column + 4].value) is
                               int or
                               type(worksheet[row][start_column + 4].value) is
                               float else 0.0)
        security.five_year = (worksheet[row][start_column + 5].value if
                              type(worksheet[row][start_column + 5].value) is
                              int or
                              type(worksheet[row][start_column + 5].value) is
                              float else 0.0)

        securities.append(security)

    return securities


def write_securities(worksheet, securities, positions):
    """
    Write individual security returns to summary worksheet at given positions.

    Args:
        worksheet: openpyxl worksheet object of summary worksheet with write
                   permissions.
        securities: List of SecurityReturn objects to write.
        positions: Corresponding list of security position dictionaries.
    """
    for security in securities:
        # Extract row and column from position dictionary
        position = positions[securities.index(security)]
        row = position['row']
        column = position['column']

        worksheet[column + str(row)].value = security.year_to_date
        worksheet[column + str(row + 1)].value = security.month_to_date
        worksheet[column + str(row + 2)].value = security.quarter_to_date
        worksheet[column + str(row + 3)].value = security.one_year
        worksheet[column + str(row + 4)].value = security.three_year
        worksheet[column + str(row + 5)].value = security.five_year


def write_comparison(worksheet, securities, position):
    """
    Write best/worst security comparison at given position.

    Args:
        worksheet: openpyxl worksheet object of summary worksheet with write
                   permissions.
        securities: List of SecurityReturn objects to compare.
        position: Position dictionary indicating where to write comparison.
    """
    def write(name, attribute, row, column):
        """
        Write security name and attribute at specified position.

        Args:
            name: String name or ticker of security.
            attribute: Return attribute of security.
            row: One-indexed row number of position to write at.
            column: A-indexed column character of position to write at.
        """
        worksheet[column + str(row)].value = name
        worksheet[chr(ord(column) + 1) + str(row)].value = attribute

    row = position['row']
    column = position['column']

    # Find best/worst security by year-to-date return
    securities.sort(key=lambda security: security.year_to_date)
    write(securities[len(securities) - 1].name,
          securities[len(securities) - 1].year_to_date, row, column)
    write(securities[0].name, securities[0].year_to_date, row,
          chr(ord(column) + 2))

    # Find best/worst security by month-to-date return
    securities.sort(key=lambda security: security.month_to_date)
    write(securities[len(securities) - 1].name,
          securities[len(securities) - 1].month_to_date, row + 1, column)
    write(securities[0].name, securities[0].month_to_date, row + 1,
          chr(ord(column) + 2))

    # Find best/worst security by quarter-to-date return
    securities.sort(key=lambda security: security.quarter_to_date)
    write(securities[len(securities) - 1].name,
          securities[len(securities) - 1].quarter_to_date, row + 2, column)
    write(securities[0].name, securities[0].quarter_to_date, row + 2,
          chr(ord(column) + 2))

    # Find best/worst security by one-year return
    securities.sort(key=lambda security: security.one_year)
    write(securities[len(securities) - 1].name,
          securities[len(securities) - 1].one_year, row + 3, column)
    write(securities[0].name, securities[0].one_year, row + 3,
          chr(ord(column) + 2))

    # Find best/worst security by three-year return
    securities.sort(key=lambda security: security.three_year)
    write(securities[len(securities) - 1].name,
          securities[len(securities) - 1].three_year, row + 4, column)
    write(securities[0].name, securities[0].three_year, row + 4,
          chr(ord(column) + 2))

    # Find best/worst security by five-year return
    securities.sort(key=lambda security: security.five_year)
    write(securities[len(securities) - 1].name,
          securities[len(securities) - 1].five_year, row + 5, column)
    write(securities[0].name, securities[0].five_year, row + 5,
          chr(ord(column) + 2))


# Read in settings from JSON file
print('Reading settings file... ', end='')
settings_file = open('settings.json')
settings = loads(settings_file.read())
settings_file.close()

# Assign settings to variables
report_path = settings['report_path']
return_start_column = ord(settings['return_start_column'].lower()) - ord('a')
bond_tickers = settings['bond_tickers']
international_security_tickers = settings['international_security_tickers']
domestic_security_tickers = settings['domestic_security_tickers']
total_return_positions = settings['total_return_positions']
index_positions = settings['index_positions']
comparison_positions = settings['comparison_positions']
print('Done.\n')

# Load in workbooks
review_workbook = read_workbook('portfolio review', True)[0]
summary_workbook, summary_workbook_filename = read_workbook('portfolio ' +
                                                            'summary', False)

# Read data from review worksheet
print('\nReading review workbook... ', end='')
review_worksheet = tuple(review_workbook.active.rows)
review_workbook.close()

# Convert worksheet rows to SecurityReturn objects
total_returns = get_security_info(review_worksheet,
                                  list(position['name'] for position in
                                       total_return_positions), 3,
                                  return_start_column)
indexes = get_security_info(review_worksheet,
                            list(position['ticker'] for position in
                                 index_positions), 2, return_start_column)
bonds = get_security_info(review_worksheet, bond_tickers, 2,
                          return_start_column)
international_securities = get_security_info(review_worksheet,
                                             international_security_tickers, 2,
                                             return_start_column)
domestic_securities = get_security_info(review_worksheet,
                                        domestic_security_tickers, 2,
                                        return_start_column)
print('Done.')

# Read data from summary worksheet
print('Writing to summary workbook... ', end='')
summary_worksheet = summary_workbook.active

# Write in individual securities or total returns
write_securities(summary_worksheet, total_returns, total_return_positions)
write_securities(summary_worksheet, indexes, index_positions)

# Write in security comparisons
write_comparison(summary_worksheet, bonds,
                 next(position for position in comparison_positions
                      if position['security_set'] == 'bonds'))
write_comparison(summary_worksheet, international_securities,
                 next(position for position in comparison_positions
                      if position['security_set'] ==
                      'international_securities'))
write_comparison(summary_worksheet, domestic_securities,
                 next(position for position in comparison_positions
                      if position['security_set'] == 'domestic_securities'))

summary_workbook.save(report_path + summary_workbook_filename)
print('Done!')
