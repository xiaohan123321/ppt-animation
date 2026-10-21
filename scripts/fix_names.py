# -*- coding: utf-8 -*-
"""修复 pptx 中形状缺失的 name 属性。
editor_sdk 生成的 cNvPr 只有 id，没有 name；PowerPoint 宽容，
但 LibreOffice / WPS 解析动画时会失败。这里给每个缺 name 的
cNvPr 补上稳定名字（Shape1、Shape2 ...）。
"""
import re
import shutil
import sys
import zipfile

from lxml import etree

P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def fix(src, dst=None):
    dst = dst or src
    tmp = src + ".tmp"
    zin = zipfile.ZipFile(src)
    zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    fixed = 0
    for item in zin.infolist():
        data = zin.read(item.filename)
        if re.match(r"ppt/slides/slide\d+\.xml$", item.filename):
            root = etree.fromstring(data)
            for k, nv in enumerate(root.iter(f"{{{P}}}cNvPr"), 1):
                if not (nv.get("name") or "").strip():
                    nv.set("name", f"Shape{k}")
                    fixed += 1
            data = etree.tostring(root, xml_declaration=True,
                                  encoding="UTF-8", standalone=True)
        zout.writestr(item, data)
    zin.close()
    zout.close()
    shutil.move(tmp, dst)
    return fixed


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python fix_names.py <pptx>")
        sys.exit(1)
    path = sys.argv[1]
    n = fix(path)
    print(f"已补齐 {n} 个形状 name -> {path}")
