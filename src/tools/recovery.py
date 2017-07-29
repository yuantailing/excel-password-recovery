#!/usr/bin/env python
# -*- coding: UTF-8 -*-

from __future__ import absolute_import
from __future__ import division
from __future__ import print_function

import sys
import xml.dom.minidom
import zipfile


def xlsx_remove_protections(zipin, zipout):
    books = set()
    sheets = set()

    content_types = zipin.read('[Content_Types].xml')
    dom = xml.dom.minidom.parseString(content_types)
    type_map = {'application/vnd.openxmlformats-officedocument.'
                'spreadsheetml.sheet.main+xml': books,
                'application/vnd.openxmlformats-officedocument.'
                'spreadsheetml.worksheet+xml': sheets}
    root = dom.documentElement
    for node in root.childNodes:
        if node.hasAttribute('ContentType') and \
                node.getAttribute('ContentType') in type_map:
            assert(node.nodeName == 'Override' and
                   node.hasAttribute('PartName'))
            part_name = node.getAttribute('PartName')
            assert(part_name.startswith('/'))
            part_name = part_name[1:]
            type_map[node.getAttribute('ContentType')].add(part_name)

    for zinfo in zipin.infolist():
        content = zipin.read(zinfo)
        if zinfo.filename in books:
            dom = xml.dom.minidom.parseString(content)
            root = dom.documentElement
            protections = root.getElementsByTagName('workbookProtection')
            for protection in protections:
                root.removeChild(protection)
            protections = root.getElementsByTagName('fileSharing')
            for protection in protections:
                root.removeChild(protection)
            content = dom.toxml(encoding=dom.encoding)
        if zinfo.filename in sheets:
            dom = xml.dom.minidom.parseString(content)
            root = dom.documentElement
            protections = root.getElementsByTagName('sheetProtection')
            for protection in protections:
                root.removeChild(protection)
            content = dom.toxml(encoding=dom.encoding)
        zipout.writestr(zipfile.ZipInfo(zinfo.filename, zinfo.date_time),
                        content, compress_type=zipfile.ZIP_DEFLATED)


if __name__ == '__main__':
    filenamein, filenameout = sys.argv[1:]
    with zipfile.ZipFile(filenamein, 'r') as zipin, \
            zipfile.ZipFile(filenameout, 'w') as zipout:
        xlsx_remove_protections(zipin, zipout)
