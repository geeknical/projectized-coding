# -*- coding: utf-8 -*-
# @UpdateTime    : 2021/4/1 00:08
# @Author    : 27
# @File    : geeknical_excel.py
# @Description    : 说明示例，不具有直接运行的保证， 依赖pydantic， xlsxwriter，读源码后可自行安装尝试
# 建议阅读或者了解xlsxwriter基础api之后再来看本文件
from typing import Dict, Optional, List, Tuple, Union
import tempfile
import os.path
from .....lib import datetime_util
import xlsxwriter
from pydantic import BaseModel
import shutil



class ExcelTitleVO(BaseModel):
    title: str
    title_format_dict: Optional[Dict]
    title_range: str  # "B3:D4"


class SingleSheetContent(BaseModel):
    work_sheet_name: str
    titles: List[ExcelTitleVO]
    contents: List[List]
    write_in_by_row: bool = True
    content_start_col: int = 0
    content_start_row: int = 1

    @classmethod
    def create_single_sheet_content(cls, sheet_name, titles: List[ExcelTitleVO],
                                    contents: List[List], write_in_by_row=True,
                                    content_start_col=0, content_start_row=1
                                    ) -> 'SingleSheetContent':
        return SingleSheetContent(
            work_sheet_name=sheet_name,
            titles=titles,
            contents=contents,
            write_in_by_row=write_in_by_row,
            content_start_col=content_start_col,
            content_start_row=content_start_row
        )


class ExcelContents:

    @classmethod
    def build_simple_contents(cls, title: str, list_of_datas: list, title_range="A1:Z1"):
        return cls(
            [ExcelTitleVO(title=title, title_range=title_range)],
            list_of_datas
        )

    @classmethod
    def build_with_titles(cls, titles: List[ExcelTitleVO], list_of_datas: list):
        return cls(
            titles, list_of_datas
        )

    def __init__(self, titles: List[ExcelTitleVO], datas: List[Optional[List[str]]]):
        self.titles = titles
        self.datas = datas
        self.workbook = None
        self.worksheet = None

    def close_work_book(self):
        self.workbook.close()

    def set_workbook(self, target_file):
        self.workbook = xlsxwriter.Workbook(target_file)

    def write_in_excel_with_mul_sheets(self, worksheet_name, write_in_by_row=True,
                                       content_start_row=1, content_start_col=0):
        self.worksheet = self.workbook.add_worksheet(worksheet_name)
        self._write_titles()
        if write_in_by_row:
            self._write_content_by_row(content_start_row, content_start_col)
        else:
            self._write_content_by_col(content_start_row, content_start_col)

    def write_in_excel(self, target_file, worksheet_name="train_record_1", write_in_by_row=True,
                       content_start_row=1, content_start_col=0):
        self.workbook = xlsxwriter.Workbook(target_file)
        self.worksheet = self.workbook.add_worksheet(worksheet_name)
        self._write_titles()
        if write_in_by_row:
            self._write_content_by_row(content_start_row, content_start_col)
        else:
            self._write_content_by_col(content_start_row, content_start_col)

        self.workbook.close()

    def _write_content_by_row(self, content_start_row=1, content_start_col=0, content_format_dict: Dict = None):
        """
        写内容，非标题行的所有内容
        """
        row = content_start_row
        col = content_start_col
        content_format = None
        if content_format_dict is not None:
            content_format = self.workbook.add_format(content_format_dict)
        for row_data in self.datas:
            for field_data in row_data:
                if content_format:
                    self.worksheet.write(row, col, field_data, content_format)
                else:
                    self.worksheet.write(row, col, field_data)
                col += 1
            row += 1
            col = content_start_col

    def _write_content_by_col(self, content_start_row=1, content_start_col=0, content_format_dict: Dict = None):
        """
        写内容，非标题行的所有内容, 竖向写入
        """
        row = content_start_row
        col = content_start_col
        content_format = None
        if content_format_dict is not None:
            content_format = self.workbook.add_format(content_format_dict)
        for col_data in self.datas:
            for field_data in col_data:
                # print("data:{}, type:{}".format(field_data, type(field_data)))
                if content_format:
                    self.worksheet.write(row, col, field_data, content_format)
                else:
                    self.worksheet.write(row, col, field_data)
                row += 1
            col += 1
            row = content_start_row

    def _write_titles(self):
        for excel_title_vo in self.titles:
            title = excel_title_vo.title
            title_range = excel_title_vo.title_range
            title_format_dict = excel_title_vo.title_format_dict
            title_format = None
            if title_format_dict:
                title_format = self.workbook.add_format(title_format_dict)
            if ":" in title_range:
                self.worksheet.merge_range(
                    title_range, title, cell_format=title_format)
            else:
                if title_format:
                    self.worksheet.write(title_range, title, title_format)
                else:
                    self.worksheet.write(title_range, title)


def build_2D_titles(two_dimension_list: List[List], title_format_dict=None) -> List[ExcelTitleVO]:
    """ 按照二维数组位置给title位置 比如列就是abcd , 行就是1234类似这样, 占位补空字符串 只考虑单格 不考虑合并单元格问题, 也暂时默认title格式一样 """
    title_infos = []
    for x in range(len(two_dimension_list)):
        row = two_dimension_list[x]
        row_value = x + 1
        for i in range(len(row)):
            title_str = row[i]
            col_value = i + 1
            letter = get_col_letter_position_from_num(col_value)
            d = {
                "title": title_str,
                "title_range": "{}{}".format(letter, row_value)
            }
            if title_format_dict:
                d["title_format_dict"] = title_format_dict
            title_infos.append(d)
    return build_excel_titles(title_infos)


def build_single_line_titles(single_line_titles: List[str], title_format_dict=None) -> List[ExcelTitleVO]:
    """
    Args:
        single_line_titles:
        title_format_dict: Optional[{
                    'bold': True,
                    'align': 'center',
                    'border': 2,
                    'valign': 'vcenter',
                    'fg_color': '#DDDDDD', 等等等 查看xlsxwriter官方文档
                }]

    Returns:
    """
    return build_2D_titles([single_line_titles], title_format_dict)


def build_excel_titles(title_infos: List[Dict]) -> List[ExcelTitleVO]:
    """
    Args:
        title_infos:
        [
            {
                title:
                title_range:
                title_format_dict: Optional[{
                    'bold': True,
                    'align': 'center',
                    'border': 2,
                    'valign': 'vcenter',
                    'fg_color': '#DDDDDD', 等等等 查看xlsxwriter官方文档
                }]
            }, ...
        ]

    Returns:

    """
    titles = []
    for title_info in title_infos:
        title = title_info['title']
        title_range = title_info['title_range']

        title_format_dict = title_info.get('title_format_dict')
        excel_title_vo = build_excel_title(
            title, title_range, title_format_dict)
        titles.append(excel_title_vo)
    return titles


def build_excel_title(title, title_range, title_format_dict=None) -> ExcelTitleVO:
    title_info = {
        "title": title,
        "title_range": title_range
    }
    if title_format_dict is None:
        title_format_dict = {
            'bold': True,
            'align': 'center',
            'border': 2,
            'valign': 'vcenter',
            'fg_color': '#DDDDDD',
        }
    title_info['title_format_dict'] = title_format_dict
    return ExcelTitleVO(**title_info)


def read_from_excel(target_file):
    pass

# 这里好像有现成的库有这个逻辑
letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N",
           "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]

num_map_letter = dict([(i + 1, letters[i]) for i in range(26)])


def get_col_letter_position_from_num(col_num):
    if col_num > 702:  # 暂无需超过两位字母的列
        raise Exception("exceed length")
    else:
        start_base_num, remainder = divmod(col_num, 26)
        if start_base_num == 0:
            return num_map_letter[remainder]
        else:
            if remainder == 0:
                if start_base_num == 1:
                    return "Z"
                return "{}{}".format(num_map_letter[start_base_num - 1], "Z")
            else:
                return "{}{}".format(num_map_letter[start_base_num], num_map_letter[remainder])


def upload_excel_file_for_url(excel_file_name, excel_content: ExcelContents, worksheet_name="train_record_1",
                              content_start_row=1, content_start_col=0, write_in_by_row=True):
    file_path, tmpdir = _build_tmp_excel_file(excel_file_name)
    excel_content.write_in_excel(file_path, worksheet_name, content_start_row=content_start_row,
                                 content_start_col=content_start_col, write_in_by_row=write_in_by_row)

    return _get_excel_file_url(file_path, tmpdir)


def build_local_excel_with_multi_sheets(excel_file_name, multiple_sheet_contents: List[SingleSheetContent]):
    e_c = ExcelContents([], [])
    e_c.set_workbook(excel_file_name)

    for single_sheet_content_vo in multiple_sheet_contents:
        titles = single_sheet_content_vo.titles
        datas = single_sheet_content_vo.contents
        write_in_by_row = single_sheet_content_vo.write_in_by_row
        content_start_col = single_sheet_content_vo.content_start_col
        content_start_row = single_sheet_content_vo.content_start_row
        worksheet_name = single_sheet_content_vo.work_sheet_name
        e_c.titles = titles
        e_c.datas = datas
        e_c.write_in_excel_with_mul_sheets(worksheet_name=worksheet_name, write_in_by_row=write_in_by_row,
                                           content_start_col=content_start_col, content_start_row=content_start_row)
    e_c.close_work_book()


def upload_excel_file_with_multi_sheets(excel_file_name, multiple_sheet_contents: List[SingleSheetContent]) -> str:
    """ return excel url """
    e_c, file_path, tmpdir = _get_init_excel_content(excel_file_name)

    for single_sheet_content_vo in multiple_sheet_contents:
        titles = single_sheet_content_vo.titles
        datas = single_sheet_content_vo.contents
        write_in_by_row = single_sheet_content_vo.write_in_by_row
        content_start_col = single_sheet_content_vo.content_start_col
        content_start_row = single_sheet_content_vo.content_start_row
        worksheet_name = single_sheet_content_vo.work_sheet_name
        e_c.titles = titles
        e_c.datas = datas
        e_c.write_in_excel_with_mul_sheets(worksheet_name=worksheet_name, write_in_by_row=write_in_by_row,
                                           content_start_col=content_start_col, content_start_row=content_start_row)
    e_c.close_work_book()

    return _get_excel_file_url(file_path, tmpdir)


def _get_init_excel_content(excel_file_name) -> Tuple[ExcelContents, str, str]:
    file_path, tmpdir = _build_tmp_excel_file(excel_file_name)
    e_c = ExcelContents([], [])
    e_c.set_workbook(file_path)
    return e_c, file_path, tmpdir


def _get_excel_file_url(file_path, tmpdir) -> str:
    """
    利用oss生成链接
    """
    # key = oss_util.generate_xxxx_key(
    #     file_path, "project_key"
    # )
    # oss_util.save_file_to_oss(
    #     key, file_path
    # )
    # shutil.rmtree(tmpdir)
    # return oss_util.build_visit_url(key)
    pass


def _build_tmp_excel_file(excel_file_name: str) -> str:
    """
    生成临时excel文件
    """
    tmpdir = tempfile.mkdtemp()
    excel_file_name = '%s_%s' % (
        # datetime_util.build_datetime_str(),  # 另外一个时间util 生成时间串防止oss上有某个文件无法重写。
        excel_file_name)
    return os.path.join(tmpdir, excel_file_name), tmpdir
