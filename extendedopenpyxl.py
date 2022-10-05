from openpyxl.reader.excel import ExcelReader, _validate_archive
from openpyxl.xml.constants import SHEET_MAIN_NS, SHARED_STRINGS, ARC_SHARED_STRINGS, ARC_APP, ARC_CORE, ARC_THEME, ARC_STYLE, ARC_ROOT_RELS, ARC_WORKBOOK, ARC_WORKBOOK_RELS
from openpyxl.xml.functions import iterparse, xmlfile, tostring
from openpyxl.utils import coordinate_to_tuple
import openpyxl.cell._writer
from zipfile import ZipFile, ZIP_DEFLATED
from openpyxl.writer.excel import ExcelWriter
from io import BytesIO
from xml.etree.ElementTree import register_namespace
from xml.etree.ElementTree import tostring as xml_tostring
from lxml.etree import fromstring as lxml_fromstring
from openpyxl.worksheet._writer import WorksheetWriter
from openpyxl.workbook._writer import WorkbookWriter
from openpyxl.packaging.extended import ExtendedProperties
from openpyxl.styles.stylesheet import write_stylesheet
from openpyxl.packaging.relationship import Relationship
from openpyxl.cell._writer import write_cell
from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing
from openpyxl import LXML
from openpyxl.packaging.manifest import DEFAULT_OVERRIDE, Override, Manifest
import openpyxl

DEFAULT_OVERRIDE.append(Override("/" + ARC_SHARED_STRINGS, SHARED_STRINGS))


def to_integer(value):
    if type(value) == int:
        return value
    if type(value) == str:
        try:
            num = int(value)
            return num
        except ValueError:
            num = float(value)
            if num.is_integer():
                return int(num)
    raise ValueError('Value {} is not an integer.'.format(value))
    return

def parse_cell(cell):
    VALUE_TAG = '{%s}v' % SHEET_MAIN_NS
    data_type = cell.get('t', 'n')
    value = None
    if data_type == 's':
        value = cell.findtext(VALUE_TAG, None) or None
        if value is not None:
            value = int(value)
    return value

def get_coordinates(cell, row_counter, col_counter):
    coordinate = cell.get('r')
    if coordinate:
        row, column = coordinate_to_tuple(coordinate)
    else:
        row, column = row_counter, col_counter
    return row, column

def parse_row(row, row_counter):
    row_counter = to_integer(row.get('r', row_counter + 1))
    col_counter = 0
    data = dict()
    for cell in row:
        col_counter += 1
        value = parse_cell(cell)
        if value is not None:
            coordinates = get_coordinates(cell, row_counter, col_counter)
            data[coordinates] = value
            col_counter = coordinates[1]
    return data, row_counter

def parse_sheet(xml_source):
    ROW_TAG = '{%s}row' % SHEET_MAIN_NS
    row_counter = 0
    it = iterparse(xml_source)
    data = dict()
    for _, element in it:
        tag_name = element.tag
        if tag_name == ROW_TAG:
            pass
            row_data, row_counter = parse_row(element, row_counter)
            data.update(row_data)
            element.clear()
    return data

def extended_archive_open(archive, name):
    with archive.open(name,) as src:
        namespaces = {node[0]: node[1] for _, node in
                      iterparse(src, events=['start-ns'])}
    for key, value in namespaces.items():
        register_namespace(key, value)
    return archive.open(name,)

def get_data_strings(xml_source):
    STRING_TAG = '{%s}si' % SHEET_MAIN_NS
    strings = []
    for _, node in iterparse(xml_source):
        if node.tag == STRING_TAG:
            strings.append(node)
    return strings

def load_workbook(filename, read_only=False, keep_vba=False,
                  data_only=False, keep_links=True):
    reader = ExcelReader(filename, read_only, keep_vba,
                        data_only, keep_links)
    reader.read()

    archive = _validate_archive(filename)

    workbook_data = dict()
    for sheet, rel in reader.parser.find_sheets():
        if rel.target not in reader.valid_files or "chartsheet" in rel.Type:
            continue
        fh = archive.open(rel.target)
        sheet_data = parse_sheet(fh)
        workbook_data[sheet.name] = sheet_data

    data_strings = []
    ct = reader.package.find(SHARED_STRINGS)
    if ct is not None:
        strings_path = ct.PartName[1:]
        with extended_archive_open(archive, strings_path) as src:
            data_strings = get_data_strings(src)

    archive.close()
    workbook = reader.wb
    workbook._extended_value_workbook_data = workbook_data
    workbook._extended_value_data_strings = data_strings
    return workbook

def check_if_lxml(element):
    if type(element).__module__ == 'xml.etree.ElementTree':
        string = xml_tostring(element)
        el = lxml_fromstring(string)
        return el
    return element

def write_string_table(workbook):
    workbook_data = workbook._extended_value_workbook_data
    data_strings = workbook._extended_value_data_strings
    out = BytesIO()
    with xmlfile(out) as xf:
        with xf.element("sst", xmlns=SHEET_MAIN_NS, uniqueCount="%d" % len(data_strings)):
            for i in range(0, len(data_strings)):
                xml_el = data_strings[i]
                el = check_if_lxml(xml_el)
                xf.write(el)
    return out.getvalue()

def check_cell(cell):
    if cell.data_type != 's':
        return False
    if cell._comment is not None:
        return False
    if cell.hyperlink:
        return False
    return True

def extended_write_cell(xf, worksheet, cell, styled=None):
    if worksheet.title not in worksheet.parent._extended_value_workbook_data:
        return
    sheet = worksheet.parent._extended_value_workbook_data[worksheet.title]
    if (cell.row, cell.column) in sheet and check_cell(cell):
        attributes = {'r': cell.coordinate, 't': cell.data_type}
        if styled:
            attributes['s'] = '%d' % cell.style_id
        if LXML:
            with xf.element('c', attributes):
                with xf.element('v'):
                    xf.write('%.16g' % sheet[(cell.row, cell.column)])
        else:
            el = Element('c', attributes)
            cell_content = SubElement(el, 'v')
            cell_content.text = '%.16g' % sheet[(cell.row, cell.column)]
            xf.write(el)
    else:
        write_cell(xf, worksheet, cell, styled)
    return

class ExtendedWorksheetWriter(WorksheetWriter):

    def write_row(self, xf, row, row_idx):
        attrs = {'r': f"{row_idx}"}
        dims = self.ws.row_dimensions
        attrs.update(dims.get(row_idx, {}))

        with xf.element("row", attrs):

            for cell in row:
                if cell._comment is not None:
                    comment = openpyxl.comments.comment_sheet.CommentRecord.from_cell(cell)
                    self.ws._comments.append(comment)
                if (
                    cell._value is None
                    and not cell.has_style
                    and not cell._comment
                    ):
                    continue
                extended_write_cell(xf, self.ws, cell, cell.has_style)
        return


class ExtendedWorkbookWriter(WorkbookWriter):

    def write_rels(self, *args, **kwargs):
        styles =  Relationship(type='sharedStrings', Target='sharedStrings.xml')
        self.rels.append(styles)
        return super().write_rels(*args, **kwargs)

class ExtendedExcelWriter(ExcelWriter):

    def __init__(self, workbook, archive):
        self._archive = archive
        self.workbook = workbook
        self.manifest = Manifest(Override = DEFAULT_OVERRIDE)
        self.vba_modified = set()
        self._tables = []
        self._charts = []
        self._images = []
        self._drawings = []
        self._comments = []
        self._pivots = []
        return

    def write_data(self):
        archive = self._archive
        props = ExtendedProperties()
        archive.writestr(ARC_APP, tostring(props.to_tree()))
        archive.writestr(ARC_CORE, tostring(self.workbook.properties.to_tree()))
        if self.workbook.loaded_theme:
            archive.writestr(ARC_THEME, self.workbook.loaded_theme)
        else:
            archive.writestr(ARC_THEME, theme_xml)
        self._write_worksheets()
        self._write_chartsheets()
        self._write_images()
        self._write_charts()

        if self.workbook._extended_value_workbook_data \
                and self.workbook._extended_value_data_strings:
            string_table_out = write_string_table(self.workbook)
            self._archive.writestr(ARC_SHARED_STRINGS, string_table_out)

        self._write_external_links()
        stylesheet = write_stylesheet(self.workbook)
        archive.writestr(ARC_STYLE, tostring(stylesheet))

        writer = ExtendedWorkbookWriter(self.workbook)

        archive.writestr(ARC_ROOT_RELS, writer.write_root_rels())
        archive.writestr(ARC_WORKBOOK, writer.write())
        archive.writestr(ARC_WORKBOOK_RELS, writer.write_rels())
        self._merge_vba()
        self.manifest._write(archive, self.workbook)
        return

    def write_worksheet(self, ws):
        ws._drawing = SpreadsheetDrawing()
        ws._drawing.charts = ws._charts
        ws._drawing.images = ws._images
        if self.workbook.write_only:
            if not ws.closed:
                ws.close()
            writer = ws._writer
        else:
            writer = ExtendedWorksheetWriter(ws)
            writer.write()

        ws._rels = writer._rels
        self._archive.write(writer.out, ws.path[1:])
        self.manifest.append(ws)
        writer.cleanup()
        return

def save_workbook(workbook, filename):
    archive = ZipFile(filename, 'w', ZIP_DEFLATED, allowZip64=True)
    writer = ExtendedExcelWriter(workbook, archive)
    writer.save()
    return True