#! /usr/bin/env python3
# -*- coding: utf-8 -*-
# @Author  : muyan.wang@centurygame.com
# @Date    : 2022/4/29
# @Desc    :
import re

from docx import Document
from collections import defaultdict
import yaml

with open("config.yaml", 'r', encoding='utf-8') as yamlFile:
    config_data = yaml.load(yamlFile, Loader=yaml.FullLoader)


def find_index(content: str, index, tag):
    if content.find(str(index)) == 0:
        return f"{content[:2]}【{tag}】{content[2:]}"
    elif content.find(str(index)) == 1:
        return f"{content[:3]}【{tag}】{content[3:]}"
    else:
        print(f"无法识别{content}首位")


def get_daan(file_name):
    document = Document(f"da_docx/{file_name}")
    import re

    da_data = {}

    one_index = 0

    three_index = 0

    for text in document.paragraphs:
        if text.text:
            if re.findall(config_data['big_title'], text.text):
                one_index = re.findall(config_data['big_title'], text.text)[0]
            elif re.match(config_data['填空题']['biaotiguanjianzi'], text.text) or re.match(config_data['简答题']['biaotiguanjianzi'], text.text) or re.match(config_data['判断题']['biaotiguanjianzi'], text.text) or re.match(config_data['多项选择题']['biaotiguanjianzi'], text.text) or re.match(config_data['单项选择题']['biaotiguanjianzi'], text.text):
                three_index = text.text
            else:
                if one_index and three_index:
                    data = re.split(r"(（\d{0,9}）|\d{0,9}\．)", text.text)
                    if not data[0]:
                        data = data[1:]
                        key_biaoti = data[::2]
                        key_text = data[1::2]
                        for key, value in zip(key_biaoti, key_text):
                            key_name = f"{one_index}_{three_index}_{key}"
                            da_data.setdefault(key_name, []).append(value)
                    else:
                        da_data.setdefault(key_name, []).extend(data)
    return da_data


def new_docx(file_name):
    da_data = get_daan(file_name)

    document = Document(f"old_docx/{file_name}")

    book_index = 0
    book_tag = ""
    one_text = ""
    three_text = ""
    local_index = 0
    for text in document.paragraphs:
        if text.text:
            if book_index - 1 > 0 and local_index != book_index:
                if re.findall("（\d{0,9}）|\d{0,9}\．", text.text) or re.match(config_data['big_title'], text.text) or re.match(config_data['填空题']['biaotiguanjianzi'], text.text) or re.match(config_data['简答题']['biaotiguanjianzi'], text.text) or re.match(config_data['判断题']['biaotiguanjianzi'], text.text) or re.match(config_data['多项选择题']['biaotiguanjianzi'], text.text) or re.match(config_data['单项选择题']['biaotiguanjianzi'], text.text):
                    key_name = f"{one_text}_{three_text}_{f'（{book_index - 1}）'}"
                    two_key = f"{one_text}_{three_text}_{f'{book_index - 1}．'}"
                    ddd = da_data.get(key_name) or da_data.get(two_key) or [""]
                    text.insert_paragraph_before(text=f'【答案】{ddd[0]}')
                    for index in ddd[1:]:
                        text.insert_paragraph_before(text=index)
                    text.insert_paragraph_before(text='【解析】')
                    tag = re.sub(r"[一二三四五六七八九十\d．、]", "", three_text)
                    text.insert_paragraph_before(text=f'【标签】{tag}')
                    local_index = book_index

            if re.match(config_data['big_title'], text.text):
                one_text = re.findall(config_data['big_title'], text.text)[0]
                book_index = 0

            elif re.match(config_data['填空题']['biaotiguanjianzi'], text.text):
                content = text.text
                p = text.clear()
                p.add_run("【题组】" + content)
                book_tag = "填空题"
                book_index = 1
                three_text = content

            elif re.match(config_data['简答题']['biaotiguanjianzi'], text.text):
                content = text.text
                p = text.clear()
                p.add_run("【题组】" + content)
                book_tag = "简答题"
                book_index = 1
                three_text = content

            elif re.match(config_data['判断题']['biaotiguanjianzi'], text.text):
                content = text.text
                p = text.clear()
                p.add_run("【题组】" + content)
                book_tag = "判断题"
                book_index = 1
                three_text = content

            elif re.match(config_data['多项选择题']['biaotiguanjianzi'], text.text):
                content = text.text
                p = text.clear()
                p.add_run("【题组】" + content)
                book_tag = "多项选择题"
                book_index = 1
                three_text = content

            elif re.match(config_data['单项选择题']['biaotiguanjianzi'], text.text):
                content = text.text
                p = text.clear()
                p.add_run("【题组】" + content)
                book_tag = "单项选择题"
                book_index = 1
                three_text = content

            elif book_tag == "填空题":
                content = re.sub("_+|\s+", "（ ）", text.text)
                if content:
                    new_content = find_index(content, book_index, "填空题")
                    if new_content:
                        p = text.clear()
                        p.add_run(new_content)
                        book_index += 1

            elif book_tag in ["简答题", "判断题", "多项选择题", "单项选择题"]:
                content = text.text
                if content:
                    new_content = find_index(content, book_index, book_tag)
                    if new_content:
                        p = text.clear()
                        p.add_run(new_content)
                        book_index += 1

            if book_tag and re.findall("（\d{0,9}）|\d{0,9}\．", text.text) and "1" in re.findall("（\d{0,9}）|\d{0,9}\．", text.text)[0]:
                text.insert_paragraph_before(text=config_data[book_tag]['shuoming'])

    document.save(f'new_docx/{file_name}')


from pathlib import Path

old_docx_path = Path("old_docx")
for file in old_docx_path.glob('**/*.docx'):
    print(file.name)
    new_docx(file.name)
