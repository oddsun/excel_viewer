import pyxlsb
import openpyxl
import os


# from validation_utils import setup_logging, logging
#
# setup_logging()
# logger = logging.getLogger(__name__)


# TODO: split into separate files, make tabs a href

def escape_html(text):
    # .replace('\'', '&#39').replace('\"', '&quot;') for attributes
    if not isinstance(text, str):
        return str(text)
    return text.replace('&', '&amp;').replace('>', '&gt;').replace('<', '&lt;').encode('ascii',
                                                                                       'xmlcharrefreplace').decode(
        'ascii')


def create_row(row_values):
    return '<tr><td>' + '</td><td>'.join((escape_html(x or '') for x in row_values)) + '</td></tr>'


def create_table(id, table_class='xlsx-tables', row_html=''):
    return '<table id="%s" class="%s">\n' % (id, table_class) + row_html + '\n</table>\n'


def create_tab_table(sheet_names):
    return '<table><tr id="tabs"><td class="not-selected">' + '</td><td class="not-selected">'.join(
        sheet_names) + '</td></tr></table>\n'


def read_excel(excel_file, XLRB):
    if XLRB:
        wb = pyxlsb.open_workbook(excel_file)
        return wb, wb.sheets
    else:
        wb = openpyxl.load_workbook(excel_file, data_only=True, read_only=True)
        return wb, wb.sheetnames


def excel_sheet_gen(wb, sheet_names, XLRB):
    for sheet_name in sheet_names:
        # logger.info(sheet_name)
        sheet = wb.get_sheet(sheet_name) if XLRB else wb.get_sheet_by_name(sheet_name)
        if isinstance(sheet, openpyxl.chartsheet.Chartsheet):
            continue
        yield sheet_name, sheet


def excel_row_gen(sheet, XLRB):
    if XLRB:
        for row_cells in sheet.rows():
            if all(not x.v for x in row_cells):
                continue
            yield (x.v for x in row_cells)
    else:
        for row_values in sheet.values:
            if all(not x for x in row_values):
                continue
            yield row_values


def parse_excel(excel_file):
    XLRB = True if excel_file.endswith('.xlsb') else False
    wb, sheet_names = read_excel(excel_file, XLRB)

    tab_html = create_tab_table(sheet_names)
    data_html = ''
    for sheet_name, sheet in excel_sheet_gen(wb, sheet_names, XLRB):
        rows = ''
        for row_values in excel_row_gen(sheet, XLRB):
            rows += create_row(row_values)
        data_html += create_table(id=sheet_name, table_class='xlsx-tables', row_html=rows)
    return tab_html, data_html


def make_html(tab_html, data_html):
    with open('templates/excel_viewer_template.html', 'r') as f:
        template = f.read()
    # with open('test/excel_output.html', 'w') as f:
    #     f.write(template.replace('{{tab_html}}', tab_html).replace('{{data_html}}', data_html))
    # print(template)
    return template.replace('{{tab_html}}', tab_html).replace('{{data_html}}', data_html)


def excel2html(excel_file):
    return make_html(*parse_excel(excel_file))


if __name__ == '__main__':
    excel_file = input('xlsx location:').strip('\"')
    excel2html(excel_file)
