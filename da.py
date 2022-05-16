#! /usr/bin/env python3
# -*- coding: utf-8 -*-
# @Author  : muyan.wang@centurygame.com
# @Date    : 2022/5/9
# @Desc    :

from docx import Document
document = Document("da_docx/460218311《Python编程案例教程（第2版）》（高登）786-3项目实训和项目考核答案.docx")
import re

da_data = {}

one_index = 0

three_index = 0

for text in document.paragraphs:
    if text.text:
        if text.style.name == "Heading 1":
            one_index = text.text
        elif text.style.name == "Heading 3":
            three_index = text.text
        else:
            data = re.split(r"(（[\d]）)", text.text)
            if not data[0]:
                data = data[1:]
                key_biaoti = data[::2]
                key_text = data[1::2]
                for key, value in zip(key_biaoti, key_text):
                    key_name = f"{one_index}_{three_index}_{key}"
                    da_data.setdefault(key_name, []).append(value)
            else:
                da_data.setdefault(key_name, []).extend(data)

