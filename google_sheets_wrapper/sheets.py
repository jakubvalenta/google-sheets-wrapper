import os
import re
import time

import httplib2
from apiclient import discovery

import oauth2client
import oauth2client.file
from oauth2client import client, tools

SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API'
CREDENTIAL_DIR = '.credentials'


def get_credentials(flags=None):
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    if not os.path.exists(CREDENTIAL_DIR):
        os.makedirs(CREDENTIAL_DIR)
    credential_path = os.path.join(
        CREDENTIAL_DIR, 'sheets.googleapis.com.json'
    )

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run_flow(flow, store, flags)
        print('Storing credentials to ' + credential_path)
    return credentials


def build_service_oauth(flags):
    credentials = get_credentials(flags)
    http = credentials.authorize(httplib2.Http())
    discovery_url = 'https://sheets.googleapis.com/$discovery/rest?version=v4'
    service = discovery.build(
        'sheets', 'v4', http=http, discoveryServiceUrl=discovery_url
    )
    return service


def build_service_api_key(developer_key):
    discovery_url = 'https://sheets.googleapis.com/$discovery/rest?version=v4'
    service = discovery.build(
        'sheets',
        'v4',
        discoveryServiceUrl=discovery_url,
        developerKey=developer_key,
    )
    return service


def a1(row_index, column_index):
    """Get an A1 notation for a cell specified by its row index
    and column index"""
    ord_first = ord('A')
    ord_last = ord('Z')
    letters_count = ord_last - ord_first + 1
    level, letter_index = divmod(column_index, letters_count)
    letter = chr(ord_first + letter_index) * (1 + level)
    number = row_index + 1
    return '{letter}{number}'.format(letter=letter, number=number)


def a1_all(service, spreadsheet_id, sheet_id=0):
    """Get an A1 notation to target all cells in the sheet"""
    row_count = get_row_count(service, spreadsheet_id, sheet_id=sheet_id)
    return '1:{}'.format(row_count)


def format_color(rgba):
    """Convert color tuple to a dict for the Sheets API"""
    color = {'red': 0, 'green': 0, 'blue': 0}
    if len(rgba) >= 4:
        color['alpha'] = rgba[3]
    else:
        color['alpha'] = 1
    return color


def format_all_sides(value):
    """Convert a single value (padding or border) to a dict
    with keys 'top', 'bottom', 'left' and 'right'"""
    all = {}
    for pos in ('top', 'bottom', 'left', 'right'):
        all[pos] = value
    return all


n_requests_made = 0

REQUESTS_LIMIT = 100
REQUESTS_SLEEP = 60


def _wait():
    global n_requests_made
    n_requests_made += 1
    if n_requests_made > REQUESTS_LIMIT:
        print(
            'Reached limit of {} requests. Waiting for {} seconds'.format(
                REQUESTS_LIMIT, REQUESTS_SLEEP
            )
        )
        time.sleep(REQUESTS_SLEEP)
        n_requests_made = 1


def _exec(service, spreadsheet_id, requests):
    """Execute a batch of Sheets API `batchUpdate` requests"""
    print(requests)
    batch_update_request = {'requests': requests}
    _wait()
    (
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=batch_update_request)
        .execute()
    )


def _read(service, spreadsheet_id, sheet_id=0, cell_range=None):
    """Read a cell range"""
    if cell_range is None:
        cell_range = a1_all(service, spreadsheet_id, sheet_id=sheet_id)
    _wait()
    result = (
        service.spreadsheets()
        .values()
        .get(
            spreadsheetId=spreadsheet_id,
            range=cell_range,
            valueRenderOption='FORMULA',
        )
        .execute()
    )
    values = result.get('values', [])
    return values


def read_cell(service, spreadsheet_id, row_index, column_index, sheet_id=0):
    """Read a cell value"""
    cell_range = a1(row_index, column_index)
    values = _read(
        service, spreadsheet_id, sheet_id=sheet_id, cell_range=cell_range
    )
    try:
        return values[0][0]
    except IndexError:
        return None


def update(service, spreadsheet_id, rows, begin=1):
    """Update rows (overwrite existing content)

    Put the rows on the first row of the sheet by default or to another row
    if parameter `begin` is passed."""
    rows_len = len(rows)
    print('Updating {} rows.'.format(rows_len))
    _wait()
    (
        service.spreadsheets()
        .values()
        .update(
            spreadsheetId=spreadsheet_id,
            range='{begin}:{end}'.format(begin=begin, end=begin + rows_len),
            valueInputOption='USER_ENTERED',
            body={'values': rows},
        )
        .execute()
    )


def set_properties(
    service,
    spreadsheet_id,
    title=None,
    locale=None,
    time_zone=None,
    sheet_id=0,
):
    """Set spreadsheet title, locale and timeZone.

    Currently locale and timeZone are not supported by the Sheets API."""
    args = {'title': title, 'locale': locale, 'timeZone': time_zone}
    properties = {}
    fields = []
    for name, value in args.items():
        if value is not None:
            # Empty string sets only `fields`, resulting in field reset
            if value != '':
                properties[name] = value
            fields.append(name)
    requests = [
        {
            'updateSpreadsheetProperties': {
                'properties': properties,
                'fields': ','.join(fields),
            }
        }
    ]
    print('Initializing spreadsheet')
    _exec(service, spreadsheet_id, requests)


def resize_grid(service, spreadsheet_id, row_count, column_count):
    """Set number of rows in a sheet"""
    requests = [
        {
            'updateSheetProperties': {
                'properties': {
                    'gridProperties': {
                        'rowCount': row_count,
                        'columnCount': column_count,
                    }
                },
                'fields': 'gridProperties.rowCount,gridProperties.columnCount',
            }
        }
    ]
    print('Resizing grid')
    _exec(service, spreadsheet_id, requests)


def is_first_cell_empty(service, spreadsheet_id, sheet_id=0):
    """Check if the cell A1:A1 is empty"""
    rows = _read(
        service, spreadsheet_id, sheet_id=sheet_id, cell_range="A1:A1"
    )
    return not rows


def get_row_count(service, spreadsheet_id, sheet_id=0):
    """Get total number of rows in a sheet"""
    print('Reading total row count')
    result = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = result.get('sheets')
    for sheet in sheets:
        if sheet['properties']['sheetId'] == sheet_id:
            row_count = sheet['properties']['gridProperties']['rowCount']
            print('Total row count {}'.format(row_count))
            return row_count


def get_filled_rows_count(
    service, spreadsheet_id, start_row_index=0, sheet_id=0
):
    """Get total number of rows in a sheet that have at least one cell
    filled"""
    print('Reading filled rows count')
    rows = _read(service, spreadsheet_id, sheet_id=sheet_id)
    count = len(rows[start_row_index:])
    print('Filled rows count {}'.format(count))
    return count


def move(service, spreadsheet_id, row_count, start_row_index=0, sheet_id=0):
    """Move rows down by `row_count` steps"""
    requests = [
        {
            'cutPaste': {
                'source': {
                    'sheetId': sheet_id,
                    'startRowIndex': start_row_index,
                    'startColumnIndex': 0,
                },
                'destination': {
                    'sheetId': sheet_id,
                    'rowIndex': start_row_index + row_count,
                    'columnIndex': 0,
                },
                'pasteType': 'PASTE_NORMAL',
            }
        }
    ]
    print('Moving existing rows')
    _exec(service, spreadsheet_id, requests)


def auto_resize(
    service, spreadsheet_id, start_index=0, end_index=None, sheet_id=0
):
    "Auto resize columns"
    dimensions = {
        'sheetId': sheet_id,
        'dimension': 'COLUMNS',
        'startIndex': start_index,
    }
    if end_index is not None:
        dimensions['endIndex'] = end_index
    requests = [{'autoResizeDimensions': {'dimensions': dimensions}}]
    print('Auto resizing all cells')
    _exec(service, spreadsheet_id, requests)


def resize_column(service, spreadsheet_id, column_index, size, sheet_id=0):
    "Resize a column"
    requests = [
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'COLUMNS',
                    'startIndex': column_index,
                    'endIndex': column_index + 1,
                },
                'properties': {'pixelSize': size},
                'fields': 'pixelSize',
            }
        }
    ]
    print('Resizing column {}, size: {} px'.format(column_index, size))
    _exec(service, spreadsheet_id, requests)


def resize_rows(
    service,
    spreadsheet_id,
    size,
    start_row_index=0,
    end_row_index=None,
    sheet_id=0,
):
    """Resize all rows, optionally only those from `start_row_index`
    or to `end_row_index`"""
    cell_range = {
        'sheetId': sheet_id,
        'dimension': 'ROWS',
        'startIndex': start_row_index,
    }
    if end_row_index is not None:
        cell_range['endIndex'] = end_row_index
    requests = [
        {
            'updateDimensionProperties': {
                'range': cell_range,
                'properties': {'pixelSize': size},
                'fields': 'pixelSize',
            }
        }
    ]
    print(
        'Resizing rows from {} to {}, size: {} px'.format(
            start_row_index, end_row_index, size
        )
    )
    _exec(service, spreadsheet_id, requests)


def highlight_cell(
    service, spreadsheet_id, row_index, column_index, sheet_id=0
):
    border = {'style': 'DASHED', 'width': 2, 'color': format_color((0, 0, 0))}
    borders = format_all_sides(border)

    """Set cells background color to yellow and font weight to bold"""
    format_cell(
        service,
        spreadsheet_id,
        row_index,
        column_index,
        borders=borders,
        bold=True,
        sheet_id=sheet_id,
    )


def format_row(
    service,
    spreadsheet_id,
    row_index=None,
    start_column_index=0,
    end_column_index=None,
    number_format=None,
    horizontal_alignment=None,
    vertical_alignment=None,
    font_family=None,
    font_size=None,
    bold=None,
    italic=None,
    underline=None,
    background_color=None,
    borders=None,
    padding=None,
    wrap_strategy=None,
    sheet_id=0,
):
    """Format all cells in a row"""
    cell_range = {
        'sheetId': sheet_id,
        'startRowIndex': row_index,
        'endRowIndex': row_index + 1,
        'startColumnIndex': start_column_index,
    }
    if end_column_index is not None:
        cell_range['endColumnIndex'] = end_column_index
    format_range(
        service,
        spreadsheet_id,
        cell_range,
        number_format=number_format,
        horizontal_alignment=horizontal_alignment,
        vertical_alignment=vertical_alignment,
        font_family=font_family,
        font_size=font_size,
        bold=bold,
        italic=italic,
        underline=underline,
        background_color=background_color,
        borders=borders,
        padding=padding,
        wrap_strategy=wrap_strategy,
    )


def format_column(
    service,
    spreadsheet_id,
    column_index,
    start_row_index=0,
    end_row_index=None,
    number_format=None,
    horizontal_alignment=None,
    vertical_alignment=None,
    font_family=None,
    font_size=None,
    bold=None,
    italic=None,
    underline=None,
    background_color=None,
    borders=None,
    padding=None,
    wrap_strategy=None,
    sheet_id=0,
):
    """Format all cells in a column"""
    cell_range = {
        'sheetId': sheet_id,
        'startRowIndex': start_row_index,
        'startColumnIndex': column_index,
        'endColumnIndex': column_index + 1,
    }
    if end_row_index:
        cell_range['endRowIndex'] = end_row_index
    format_range(
        service,
        spreadsheet_id,
        cell_range,
        number_format=number_format,
        horizontal_alignment=horizontal_alignment,
        vertical_alignment=vertical_alignment,
        font_family=font_family,
        font_size=font_size,
        bold=bold,
        italic=italic,
        underline=underline,
        background_color=background_color,
        borders=borders,
        padding=padding,
        wrap_strategy=wrap_strategy,
    )


def format_columns(
    service,
    spreadsheet_id,
    start_column_index,
    end_column_index,
    start_row_index=0,
    number_format=None,
    horizontal_alignment=None,
    vertical_alignment=None,
    font_family=None,
    font_size=None,
    bold=None,
    italic=None,
    underline=None,
    background_color=None,
    borders=None,
    padding=None,
    wrap_strategy=None,
    sheet_id=0,
):
    """Format all cells in several columns"""
    cell_range = {
        'sheetId': sheet_id,
        'startRowIndex': start_row_index,
        'endRowIndex': 1000,
        'startColumnIndex': start_column_index,
        'endColumnIndex': end_column_index,
    }
    format_range(
        service,
        spreadsheet_id,
        cell_range,
        number_format=number_format,
        horizontal_alignment=horizontal_alignment,
        vertical_alignment=vertical_alignment,
        font_family=font_family,
        font_size=font_size,
        bold=bold,
        italic=italic,
        underline=underline,
        background_color=background_color,
        borders=borders,
        padding=padding,
        wrap_strategy=wrap_strategy,
    )


def format_cell(
    service,
    spreadsheet_id,
    row_index,
    column_index,
    number_format=None,
    horizontal_alignment=None,
    vertical_alignment=None,
    font_family=None,
    font_size=None,
    bold=None,
    italic=None,
    underline=None,
    background_color=None,
    borders=None,
    padding=None,
    wrap_strategy=None,
    sheet_id=0,
):
    """Format a cell"""
    cell_range = {
        'sheetId': sheet_id,
        'startRowIndex': row_index,
        'endRowIndex': row_index + 1,
        'startColumnIndex': column_index,
        'endColumnIndex': column_index + 1,
    }
    format_range(
        service,
        spreadsheet_id,
        cell_range,
        number_format=number_format,
        horizontal_alignment=horizontal_alignment,
        vertical_alignment=vertical_alignment,
        font_family=font_family,
        font_size=font_size,
        bold=bold,
        italic=italic,
        underline=underline,
        background_color=background_color,
        borders=borders,
        padding=padding,
        wrap_strategy=wrap_strategy,
    )


def format_range(
    service,
    spreadsheet_id,
    cell_range=None,
    number_format=None,
    horizontal_alignment=None,
    vertical_alignment=None,
    font_family=None,
    font_size=None,
    bold=None,
    italic=None,
    underline=None,
    background_color=None,
    borders=None,
    padding=None,
    wrap_strategy=None,
):
    """Set various formatting options for all cells in a range"""

    if background_color is not None and background_color != '':
        background_color = format_color(background_color)
    if padding is not None and padding != '':
        padding = format_all_sides(padding)

    args = {
        'numberFormat': number_format,
        'horizontalAlignment': horizontal_alignment,
        'verticalAlignment': vertical_alignment,
        'backgroundColor': background_color,
        'borders': borders,
        'padding': padding,
        'wrap_strategy': wrap_strategy,
    }
    cell_format = {}
    fields = []
    for name, value in args.items():
        if value is not None:
            # Empty string sets only `fields`, resulting in field reset
            if value != '':
                cell_format[name] = value
            fields.append(name)

    args_text_format = {
        'fontFamily': font_family,
        'fontSize': font_size,
        'bold': bold,
        'italic': italic,
        'underline': underline,
    }
    for name, value in args_text_format.items():
        if value is not None:
            if 'textFormat' not in cell_format:
                cell_format['textFormat'] = {}
            # Empty string sets only `fields`, resulting in field reset
            if value != '':
                cell_format['textFormat'][name] = value
            fields.append('textFormat.' + name)

    _format_range(service, spreadsheet_id, cell_range, cell_format, fields)


def _format_range(service, spreadsheet_id, cell_range, cell_format, fields):
    """Set userEnteredFormat for all cells in a range"""
    fields_formatted = ','.join('userEnteredFormat.' + x for x in fields)
    requests = [
        {
            'repeatCell': {
                'range': cell_range,
                'cell': {'userEnteredFormat': cell_format},
                'fields': fields_formatted,
            }
        }
    ]
    print('Formatting range')
    _exec(service, spreadsheet_id, requests)


def clear_formatting(
    service, spreadsheet_id, start_row_index=0, end_row_index=0, sheet_id=0
):
    cell_range = {
        'sheetId': sheet_id,
        'startRowIndex': start_row_index,
        'startColumnIndex': end_row_index,
    }
    requests = [
        {
            'repeatCell': {
                'range': cell_range,
                'cell': {},
                'fields': 'userEnteredFormat',
            }
        }
    ]
    print('Clearing formatting')
    _exec(service, spreadsheet_id, requests)


def delete_all_rows(service, spreadsheet_id, sheet_id=0):
    """Delete all rows in the sheet except for the first one, because
    Google Sheets require at least one row in a sheet

    Currently the Sheets API throws error 503 'The service is currently
    unavailable.'"""
    requests = [
        {
            'deleteDimension': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'ROWS',
                    'startIndex': 1,
                }
            }
        }
    ]
    print('Deleting all rows')
    _exec(service, spreadsheet_id, requests)


def format_formula_image(url):
    return '=IMAGE("{}")'.format(url)


def parse_formula_image(formula):
    m = re.match(r'^=IMAGE\("([^"]+)".*\)$', formula)
    if m:
        return m.group(1)
