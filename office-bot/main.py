# -*- coding: utf-8 -*-

import os
import win32com.client
from win32com.client import constants
from mcp.server.fastmcp import FastMCP

mcp = FastMCP(name="office-bot")


@mcp.tool(name="word_create")
async def word_create(file_path: str, file_name: str = "新建文档.docx"):
    """
    新建空白Word文档。

    参数:
        file_path (str): Word文档存储文件夹
        file_name (str): Word文档文件名，默认为"新建文档.docx"

    返回:
        dict: 操作结果描述
    """
    try:
        # 拼接完整路径
        full_path = os.path.join(file_path, file_name)
        # 确保目录存在
        os.makedirs(file_path, exist_ok=True)
        word = win32com.client.gencache.EnsureDispatch("Word.Application")
        word.Visible = True
        doc = word.Documents.Add()
        doc.SaveAs(full_path)
        return {"success": True, "message": f"Word文档已创建: {full_path}"}
    except Exception as e:
        return {"success": False, "message": f"创建Word文档失败: {e}"}


@mcp.tool(name="file_open",title="打开文件",description="根据绝对路径打开文件")
async def file_open(file_path: str):
    if not isinstance(file_path, str) or not file_path:
        return {"success": False, "message": "file_path 不能为空且必须为字符串"}
    if not os.path.isabs(file_path):
        return {"success": False, "message": "file_path 必须为绝对路径"}
    if not os.path.exists(file_path):
        return {"success": False, "message": f"文件不存在: {file_path}"}

    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        cmd = f'"{file_path}"'
        shell.Run(cmd)
        return f"已打开文件: {cmd}"
    except Exception as e:
        return f"打开文件失败: {e}"


@mcp.tool(name="word_read")
async def word_read(file_path: str):
    # 启动Word应用
    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    doc = word.Documents.Open(file_path)

    # 读取全文本
    full_text = doc.Content.Text

    return full_text


@mcp.tool(name="word_insert")
async def word_insert(file_path: str, text: str, insert_flag: int, target: dict):
    """
    向指定Word文档插入内容。

    参数:
        file_path (str): Word文档路径
        text (str): 要插入的文本内容
        insert_flag (int): 插入位置标志，-1=文首，0=指定行/文本，1=文末
        target (dict): 定位目标，insert_flag=0时需包含
            - line_num (int): 行号
            - tar_text (str): 目标文本
            - flag (int): 1=插入在目标文本后，0=插入在目标文本前

    返回:
        str: 操作结果描述
    """
    if insert_flag not in [-1, 0, 1]:
        insert_flag = 1
    if not text:
        return "text为必填项"
    if not file_path:
        return "file_path为必填项"

    file_path = os.path.abspath(file_path)
    if not os.path.exists(file_path):
        return "文件不存在，请检查路径。"

    try:
        word = win32com.client.gencache.EnsureDispatch("Word.Application")
        word.Visible = True  # 优化：插入时不弹窗
        doc = word.Documents.Open(file_path)
        selection = word.Selection

        if insert_flag == -1:
            # 插入到文首
            selection.HomeKey(Unit=constants.wdStory)
            selection.TypeParagraph()
            selection.TypeText(text)
        elif insert_flag == 1:
            # 插入到文末
            selection.EndKey(Unit=constants.wdStory)
            selection.TypeParagraph()
            selection.TypeText(text)
        elif insert_flag == 0:
            # 插入到指定行的目标文本前/后
            if (
                not target
                or "line_num" not in target
                or "tar_text" not in target
                or "flag" not in target
            ):
                return "target参数不完整"
            selection.GoTo(
                What=constants.wdGoToLine,
                Which=constants.wdGoToAbsolute,
                Count=target["line_num"],
            )
            found = selection.Find.Execute(
                FindText=target["tar_text"],
                Forward=True,
                MatchWholeWord=False,
                MatchCase=False,
            )
            if found:
                if target["flag"] == 1:
                    # 插入在目标文本后
                    selection.MoveRight(
                        Unit=constants.wdCharacter, Count=len(target["tar_text"])
                    )
                # 插入文本
                selection.TypeText(text)
            else:
                return f"未在第{target['line_num']}行找到“{target['tar_text']}”"
        else:
            return "insert_flag只能为[0, -1, 1]"

        doc.SaveAs(file_path)
        return "word文档内容写入成功"
    except Exception as e:
        return f"插入失败: {str(e)}"


@mcp.tool(name="word_edit")
async def word_edit(file_path: str, text: str, target: dict):
    """
    替换Word文档第target["line_num"]行中的“target["tar_text"]”为text参数内容（仅第5行，非第5段）。
    """
    file_path = os.path.abspath(file_path)
    if not os.path.exists(file_path):
        return "文件不存在，请检查路径。"
    try:
        word = win32com.client.gencache.EnsureDispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(file_path)
        selection = word.Selection

        selection.GoTo(
            What=constants.wdGoToLine,
            Which=constants.wdGoToAbsolute,
            Count=target["line_num"],
        )

        selection.HomeKey(Unit=constants.wdLine)
        selection.EndKey(Unit=constants.wdLine, Extend=1)
        line_text = selection.Text

        if target["tar_text"] not in line_text:
            return f"第{target['line_num']}行未找到“{target['tar_text']}”"

        new_line = line_text.replace(target["tar_text"], text)
        # 直接赋值替换选区内容，避免多余换行
        selection.Text = new_line
        doc.SaveAs(file_path)
        return f"第{target['line_num']}行“{target['tar_text']}”已替换为指定内容"
    except Exception as e:
        return f"发生错误: {e}"


if __name__ == "__main__":
    mcp.run(transport="stdio")
