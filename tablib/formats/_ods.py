# -*- coding: utf-8 -*-

""" Tablib - ODF Support.
"""

import sys


if sys.version_info[0] > 2:
    from io import BytesIO
else:
    from cStringIO import StringIO as BytesIO

from tablib.compat import opendocument, style, table, text, unicode
from xml.dom import Node

title = 'ods'
extensions = ('ods',)

bold = style.Style(name="bold", family="paragraph")
bold.addElement(style.TextProperties(fontweight="bold", fontweightasian="bold", fontweightcomplex="bold"))

def export_set(dataset):
    """Returns ODF representation of Dataset."""

    wb = opendocument.OpenDocumentSpreadsheet()
    wb.automaticstyles.addElement(bold)

    ws = table.Table(name=dataset.title if dataset.title else 'Tablib Dataset')
    wb.spreadsheet.addElement(ws)
    dset_sheet(dataset, ws)

    stream = BytesIO()
    wb.save(stream)
    return stream.getvalue()


def export_book(databook):
    """Returns ODF representation of DataBook."""

    wb = opendocument.OpenDocumentSpreadsheet()
    wb.automaticstyles.addElement(bold)

    for i, dset in enumerate(databook._datasets):
        ws = table.Table(name=dset.title if dset.title else 'Sheet%s' % (i))
        wb.spreadsheet.addElement(ws)
        dset_sheet(dset, ws)


    stream = BytesIO()
    wb.save(stream)
    return stream.getvalue()


def dset_sheet(dataset, ws):
    """Completes given worksheet from given Dataset."""
    _package = dataset._package(dicts=False)

    for i, sep in enumerate(dataset._separators):
        _offset = i
        _package.insert((sep[0] + _offset), (sep[1],))

    for i, row in enumerate(_package):
        row_number = i + 1
        odf_row = table.TableRow(stylename=bold, defaultcellstylename='bold')
        for j, col in enumerate(row):
            try:
                col = unicode(col, errors='ignore')
            except TypeError:
                ## col is already unicode
                pass
            ws.addElement(table.TableColumn())

            # bold headers
            if (row_number == 1) and dataset.headers:
                odf_row.setAttribute('stylename', bold)
                ws.addElement(odf_row)
                cell = table.TableCell()
                p = text.P()
                p.addElement(text.Span(text=col, stylename=bold))
                cell.addElement(p)
                odf_row.addElement(cell)

            # wrap the rest
            else:
                try:
                    if '\n' in col:
                        ws.addElement(odf_row)
                        cell = table.TableCell()
                        cell.addElement(text.P(text=col))
                        odf_row.addElement(cell)
                    else:
                        ws.addElement(odf_row)
                        cell = table.TableCell()
                        cell.addElement(text.P(text=col))
                        odf_row.addElement(cell)
                except TypeError:
                    ws.addElement(odf_row)
                    cell = table.TableCell()
                    cell.addElement(text.P(text=col))
                    odf_row.addElement(cell)

def _import_set(dset, sht, headers=True):
    def getvalue(cell):
        NUMERIC_TYPES = ('float', 'percentage', 'currency')

        TYPE_VALUE_MAP = {
            'string':     'stringvalue',
            'float':      'value',
            'percentage': 'value',
            'currency':   'value',
            'date':       'datevalue',
            'time':       'timevalue',
            'boolean':    'booleanvalue',
            }

        def convert(value, value_type):
            if value is None:
                pass
            elif value_type in NUMERIC_TYPES:
                value = float(value)
            elif value_type == 'boolean':
                value = True if value == 'true' else False
            return value

        def plaintext(cell):
            t = []
            for p in cell.getElementsByType(text.P):
                t.extend([unicode(n.data) for n in p.childNodes if n.nodeType == Node.TEXT_NODE])
            return '\n'.join(t)

        t = cell.getAttribute('valuetype')
        if  t is None:
            result = None
        elif t == 'string':
            result = plaintext(cell)
        else:
            result = convert(cell.getAttribute(TYPE_VALUE_MAP[t]), t)
        return result

    dset.title = sht.getAttribute('name')
    for i, row in enumerate(sht.getElementsByType(table.TableRow)):
        values = []
        for cell in row.getElementsByType(table.TableCell):
            repeat = cell.getAttribute("numbercolumnsrepeated") or 1
            values.extend([getvalue(cell)] * repeat)
        if (i == 0) and headers:
            dset.headers = values
        else:
            dset.append(values)

def import_set(dset, in_stream, headers=True):
    dset.wipe()

    doc = opendocument.load(in_stream)
    sht = doc.getElementsByType(table.Table)[0]
    _import_set(dset, sht, headers)

def import_book(dbook, in_stream, headers=True):
    dbook.wipe()

    doc = opendocument.load(in_stream)

    xls_book = xlrd.open_workbook(file_contents=in_stream)
    for sht in doc.getElementsByType(table.Table):
        data = tablib.Dataset()
        _import_set(data, sht, headers)
        dbook.add_sheet(data)

def detect(stream):
    try:
        opendocument.load(stream)
        return True
    except:
        return False
