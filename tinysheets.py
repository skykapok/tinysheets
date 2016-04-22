# -*- coding: utf-8 –*-

import os
import shutil
import argparse
import tempfile
import csv

VBSCRIPT = """
Set objFSO = CreateObject("Scripting.FileSystemObject")
src_file = objFSO.GetAbsolutePathName(Wscript.Arguments.Item(0))
dest_file = objFSO.GetAbsolutePathName(WScript.Arguments.Item(1))
Dim oExcel
Set oExcel = CreateObject("Excel.Application")
Dim oBook
Set oBook = oExcel.Workbooks.Open(src_file)
oBook.SaveAs dest_file, 6
oBook.Close False
oExcel.Quit
"""

def convert(fin, fout, format):
    reader = csv.reader(open(fin))
    writer = open(fout, "w")

    writer.write("return {\n")

    # 表头
    field_types = reader.next()
    field_keys = reader.next()

    # 解析数据
    for row in reader:
        if not row[0]: continue
        itemid = format_itemid(row[0].strip(), field_types[0].strip())
        writer.write("\t%s = {\n" % itemid)
        for i in xrange(1, len(row)):
            if not row[i]: continue
            t = field_types[i]
            if t:
                k = field_keys[i]
                v = format_value(row[i], field_types[i])
                writer.write("\t\t%s = %s,\n" % (k, v))
        writer.write("\t},\n")

    writer.write("}")

def format_itemid(itemid, tp):
    if tp.startswith("int"):
        return "[%s]" % itemid
    elif tp.startswith("string"):
        return itemid

def format_value(v, tp):
    parts = tp.split(":")
    tp = parts[0].strip()
    arg = parts[1].strip() if len(parts) == 2 else None
    if tp.endswith("[]"):
        elements = v.split("|")
        rets = []
        for e in elements:
            rets.append(format_value_element(e, tp, arg))
        return "{ %s }" % ", ".join(rets)
    else:
        return format_value_element(v, tp, arg)

def format_value_element(v, tp, arg):
    if tp.startswith("int"):
        return str(v)
    elif tp.startswith("float"):
        return str(v)
    elif tp.startswith("bool"):
        return "true" if v else "false"
    elif tp.startswith("string"):
        return "[[%s]]" % v.decode("gbk").encode("utf8")

def parse_args():
    parser = argparse.ArgumentParser(description='Convert xlsx/xls to py/lua/json.')
    parser.add_argument('filename',
                        nargs='+',
                        help='specifies the excel file(s) to be convert.')
    parser.add_argument('-o',
                        dest='destination',
                        help='specifies the directory for the generated script file(s).')
    parser.add_argument('-f',
                        dest='format',
                        choices=['py', 'lua', 'json'],
                        required=True,
                        help='specifies the script file format')
    return parser.parse_args()

def main():
    args = parse_args()
    outdir = args.destination if args.destination is not None and os.path.isdir(args.destination) else None
    tmpdir = tempfile.mkdtemp()

    pvbs = os.path.join(tmpdir, "exceltocsv.vbs")
    fvbs = open(pvbs, "w")
    fvbs.write(VBSCRIPT)
    fvbs.close()

    for p in args.filename:
        pxls = os.path.abspath(p)
        pdir, pfile = os.path.split(pxls)
        pbase, pext = os.path.splitext(pfile)
        pcsv = os.path.join(tmpdir, pbase + ".csv")
        plua = os.path.join(outdir or pdir, pbase + ".lua")

        if os.path.exists(plua) and os.path.getmtime(pxls) < os.path.getmtime(plua):
            print "%s skipped" % p
            continue

        os.popen("%s %s %s" % (pvbs, pxls, pcsv))
        convert(pcsv, plua, args.format)
        print "%s converted" % p

    shutil.rmtree(tmpdir)

if __name__ == "__main__":
    main()
