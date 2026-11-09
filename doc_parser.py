"""文档内容提取模块。

负责把 docx / xlsx / pptx / pdf 以及常见纯文本格式的二进制数据
转换成可供 LLM 使用或直接回复给用户的纯文本。

设计说明：
- 各解析器均为同步实现，且对大文件可能较耗时。调用方（main.py）会通过
  ``run_in_executor`` 把解析放到线程池中执行，避免阻塞事件循环。
- 解析过程中的限制（页数、行数、字符数等）通过 ``limits`` 字典传入，
  防止恶意/超大文档拖垮机器人。
- 所有可预期的失败都抛出 ``DocumentParseError``，异常信息可以直接
  回复给用户，不会导致插件崩溃。
"""

from __future__ import annotations

import io
from dataclasses import dataclass, field

# 支持解析的扩展名分组
DOCX_EXTS = {"docx"}
XLSX_EXTS = {"xlsx", "xlsm"}
PPTX_EXTS = {"pptx"}
PDF_EXTS = {"pdf"}
TEXT_EXTS = {
    "txt",
    "md",
    "markdown",
    "csv",
    "tsv",
    "log",
    "json",
    "html",
    "htm",
    "xml",
    "yaml",
    "yml",
    "ini",
    "py",
    "sql",
}
# 旧版 Office 二进制格式：开源库无法稳定解析，只做友好提示，引导用户转存新格式
LEGACY_EXTS = {"doc", "xls", "ppt", "wps", "et", "dps"}

SUPPORTED_EXTS = DOCX_EXTS | XLSX_EXTS | PPTX_EXTS | PDF_EXTS | TEXT_EXTS

# 旧格式 -> 建议转存的目标格式
LEGACY_HINT = {
    "doc": "docx",
    "xls": "xlsx",
    "ppt": "pptx",
    "wps": "docx",
    "et": "xlsx",
    "dps": "pptx",
}

# xlsx 每行最多读取的单元格数，防止个别超宽表格拖垮解析
_MAX_COLS = 60


@dataclass
class ParseResult:
    """解析结果。

    ok 为 False 时，error 携带可直接回复用户的错误说明；
    stats 为人类可读的统计信息（如页数、工作表数），用于拼进回复。
    """

    ok: bool
    text: str = ""
    error: str = ""
    stats: dict = field(default_factory=dict)


class DocumentParseError(Exception):
    """解析失败，异常信息可直接展示给用户。"""


class _MissingDependency(DocumentParseError):
    """缺少第三方解析库。"""

    def __init__(self, package: str) -> None:
        super().__init__(
            f"缺少依赖库 {package}，请在 AstrBot 所在的 Python 环境中执行: "
            f"pip install {package}（或重装本插件以自动安装依赖）"
        )


def _finalize(lines: list[str], stats: dict, limits: dict) -> tuple[str, dict]:
    """把若干行合并为最终文本，并按 max_chars 截断。"""
    max_chars = int(limits.get("max_chars", 60000))
    text = "\n".join(line for line in lines if line is not None).strip()
    if len(text) > max_chars:
        text = text[:max_chars].rstrip()
        text += "\n\n…（内容过长，已按配置截断）"
        stats["截断"] = "是"
    return text, stats


# ---------------------------------------------------------------------------
# docx
# ---------------------------------------------------------------------------


def _parse_docx(data: bytes, limits: dict) -> tuple[str, dict]:
    """提取 docx 的段落与表格。"""
    try:
        import docx
    except ImportError:
        raise _MissingDependency("python-docx") from None

    try:
        document = docx.Document(io.BytesIO(data))
    except Exception as e:
        raise DocumentParseError(f"docx 解析失败（文件可能损坏或已加密）：{e}") from e

    lines: list[str] = []
    para_count = 0
    for para in document.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        para_count += 1
        # 标题样式加 "##" 前缀，保留文档层级，方便 LLM 理解结构
        style_name = ""
        try:
            if para.style is not None:
                style_name = para.style.name or ""
        except Exception:
            pass
        if style_name.lower().startswith("heading") or style_name.startswith("标题"):
            lines.append(f"## {text}")
        else:
            lines.append(text)

    table_count = len(document.tables)
    for idx, table in enumerate(document.tables, 1):
        lines.append(f"[表格 {idx}]")
        for row in table.rows:
            cells = [cell.text.strip().replace("\n", " ") for cell in row.cells]
            lines.append(" | ".join(cells))

    return _finalize(lines, {"段落": para_count, "表格": table_count}, limits)


# ---------------------------------------------------------------------------
# xlsx / xlsm
# ---------------------------------------------------------------------------


def _fmt_cell(value) -> str:
    """把单元格值转为简洁文本。"""
    if value is None:
        return ""
    # 浮点数 3.0 -> 3，避免整列数字带小数点噪音
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    return text.replace("\n", " ")[:200]


def _parse_xlsx(data: bytes, limits: dict) -> tuple[str, dict]:
    """逐工作表、逐行提取 xlsx 内容，单元格用 " | " 连接。"""
    try:
        from openpyxl import load_workbook
    except ImportError:
        raise _MissingDependency("openpyxl") from None

    max_sheets = int(limits.get("max_sheets", 10))
    max_rows = int(limits.get("max_rows_per_sheet", 200))

    try:
        # data_only=True：读取公式的计算结果而不是公式本身
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    except Exception as e:
        raise DocumentParseError(f"xlsx 打开失败（文件可能损坏或已加密）：{e}") from e

    lines: list[str] = []
    total_rows = 0
    try:
        sheet_names = list(wb.sheetnames)
        picked = sheet_names[:max_sheets]
        for title in picked:
            ws = wb[title]
            lines.append(f"=== 工作表：{title} ===")
            count = 0
            for row in ws.iter_rows(values_only=True):
                # 跳过整行皆空的数据行
                if all(v is None or str(v).strip() == "" for v in row):
                    continue
                cells = [_fmt_cell(v) for v in row[:_MAX_COLS]]
                while cells and cells[-1] == "":
                    cells.pop()
                lines.append(" | ".join(cells))
                count += 1
                total_rows += 1
                if count >= max_rows:
                    lines.append(f"…（工作表「{title}」仅读取前 {max_rows} 行）")
                    break
        if len(sheet_names) > len(picked):
            lines.append(f"…（其余 {len(sheet_names) - len(picked)} 个工作表未读取）")
    finally:
        try:
            wb.close()
        except Exception:
            pass

    return _finalize(
        lines,
        {
            "工作表": f"{min(len(sheet_names), max_sheets)}/{len(sheet_names)}",
            "数据行": total_rows,
        },
        limits,
    )


# ---------------------------------------------------------------------------
# pptx
# ---------------------------------------------------------------------------


def _parse_pptx(data: bytes, limits: dict) -> tuple[str, dict]:
    """提取 pptx 每页幻灯片的文本、表格与演讲者备注。"""
    try:
        from pptx import Presentation
    except ImportError:
        raise _MissingDependency("python-pptx") from None

    max_slides = int(limits.get("max_slides", 100))

    try:
        prs = Presentation(io.BytesIO(data))
    except Exception as e:
        raise DocumentParseError(f"pptx 打开失败（文件可能损坏或已加密）：{e}") from e

    slides = list(prs.slides)
    lines: list[str] = []
    for idx, slide in enumerate(slides[:max_slides], 1):
        lines.append(f"=== 幻灯片 {idx} ===")
        for shape in slide.shapes:
            # 文本框
            if getattr(shape, "has_text_frame", False):
                text = shape.text_frame.text.strip()
                if text:
                    lines.append(text)
            # 表格
            if getattr(shape, "has_table", False):
                try:
                    for row in shape.table.rows:
                        cells = [
                            cell.text.strip().replace("\n", " ") for cell in row.cells
                        ]
                        lines.append(" | ".join(cells))
                except Exception:
                    pass
        # 演讲者备注
        try:
            if slide.has_notes_slide:
                notes = slide.notes_slide.notes_text_frame.text.strip()
                if notes:
                    lines.append(f"[备注] {notes}")
        except Exception:
            pass

    if len(slides) > max_slides:
        lines.append(f"…（共 {len(slides)} 页，仅读取前 {max_slides} 页）")

    return _finalize(
        lines, {"幻灯片": f"{min(len(slides), max_slides)}/{len(slides)}"}, limits
    )


# ---------------------------------------------------------------------------
# pdf
# ---------------------------------------------------------------------------


def _parse_pdf(data: bytes, limits: dict) -> tuple[str, dict]:
    """逐页提取 pdf 文本层（扫描件无文本层时会给出明确提示）。"""
    try:
        from pypdf import PdfReader
    except ImportError:
        try:
            from PyPDF2 import PdfReader  # 兼容仍安装着旧包名的环境
        except ImportError:
            raise _MissingDependency("pypdf") from None

    max_pages = int(limits.get("max_pages", 100))

    try:
        reader = PdfReader(io.BytesIO(data))
    except Exception as e:
        raise DocumentParseError(f"pdf 打开失败（文件可能损坏）：{e}") from e

    # 加密 PDF：尝试空密码（多数仅设了权限密码的文件可以解开），失败则明确告知
    if getattr(reader, "is_encrypted", False):
        try:
            if not reader.decrypt(""):
                raise DocumentParseError("PDF 已加密，请先解除密码保护后再发送")
        except DocumentParseError:
            raise
        except Exception as e:
            raise DocumentParseError(f"PDF 已加密且无法自动解密：{e}") from e

    total = len(reader.pages)
    lines: list[str] = []
    shown = 0
    empty_pages = 0
    for i, page in enumerate(reader.pages, 1):
        if i > max_pages:
            lines.append(f"…（共 {total} 页，仅读取前 {max_pages} 页）")
            break
        try:
            text = page.extract_text() or ""
        except Exception:
            text = ""
        text = text.strip()
        if not text:
            empty_pages += 1
            lines.append(f"--- 第 {i} 页 ---")
            lines.append("（本页未提取到文本，可能为扫描图片）")
        else:
            lines.append(f"--- 第 {i} 页 ---")
            lines.append(text)
        shown += 1

    # 所有页都没有文本层：基本可判定为扫描件，直接报错并解释原因
    if shown > 0 and empty_pages == shown:
        raise DocumentParseError(
            "未能从 PDF 中提取到任何文本：该文件很可能是纯扫描件（图片型 PDF），"
            "提取正文需要 OCR 支持，请改用含文字层的 PDF 或先进行 OCR 处理"
        )

    return _finalize(lines, {"页数": f"{shown}/{total}"}, limits)


# ---------------------------------------------------------------------------
# 纯文本类
# ---------------------------------------------------------------------------


def _parse_text(data: bytes, limits: dict) -> tuple[str, dict]:
    """按常见编码依次尝试解码纯文本文件。"""
    text = None
    used_enc = ""
    for enc in ("utf-8", "utf-8-sig", "gb18030", "big5"):
        try:
            text = data.decode(enc)
            used_enc = enc
            break
        except (UnicodeDecodeError, LookupError):
            continue
    if text is None:
        # latin-1 永远不会解码失败，作为最后兜底
        text = data.decode("latin-1", errors="replace")
        used_enc = "latin-1"

    # 纯文本扩展名但内容含大量空字节：多半是伪装的二进制文件
    if "\x00" in text[:4096]:
        raise DocumentParseError("该文件看起来是二进制文件而非文本，暂不支持解析")

    return _finalize([text], {"编码": used_enc}, limits)


# 扩展名 -> 解析器 分发表
_PARSER_MAP = {
    **{ext: _parse_docx for ext in DOCX_EXTS},
    **{ext: _parse_xlsx for ext in XLSX_EXTS},
    **{ext: _parse_pptx for ext in PPTX_EXTS},
    **{ext: _parse_pdf for ext in PDF_EXTS},
    **{ext: _parse_text for ext in TEXT_EXTS},
}


def sniff_ext(data: bytes) -> str:
    """通过文件头魔数嗅探真实文档类型，返回小写扩展名（无法识别返回空串）。

    用于处理文件名缺少扩展名、或被改名为 .zip / .bin 的场景：
    - ``%PDF-`` 开头 → pdf
    - ``PK\\x03\\x04``（zip 容器）→ 按压缩包内部目录判断 docx / xlsx / pptx
    """
    if data[:5] == b"%PDF-":
        return "pdf"
    if data[:4] == b"PK\x03\x04":
        try:
            import zipfile

            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                names = zf.namelist()
            if any(n.startswith("word/") for n in names):
                return "docx"
            if any(n.startswith("xl/") for n in names):
                return "xlsx"
            if any(n.startswith("ppt/") for n in names):
                return "pptx"
        except Exception:
            return ""
    return ""


def parse_document(
    name: str, data: bytes, limits: dict | None = None, ext_hint: str = ""
) -> ParseResult:
    """根据文件扩展名（或 ext_hint 嗅探结果）分发到对应解析器。

    Args:
        name: 文件名（用于识别扩展名）
        data: 文件的完整二进制内容
        limits: 解析限制，可包含 max_chars / max_sheets /
                max_rows_per_sheet / max_pages / max_slides
        ext_hint: 扩展名提示。文件名缺少扩展名或不可信时，
                  传入 sniff_ext() 的嗅探结果。

    Returns:
        ParseResult: 永不抛出异常，失败信息放在 result.error 中。
    """
    limits = limits or {}
    ext = (ext_hint or "").strip().lstrip(".").lower()
    if not ext:
        ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""

    if ext in LEGACY_EXTS:
        target = LEGACY_HINT.get(ext, "新版格式")
        return ParseResult(
            ok=False,
            error=f"暂不支持旧版 Office 格式 .{ext}，请在 Office/WPS 中另存为 .{target} 后重新发送。",
        )

    parser = _PARSER_MAP.get(ext)
    if parser is None:
        return ParseResult(ok=False, error=f"暂不支持的文件格式：.{ext or '未知'}")

    try:
        text, stats = parser(data, limits)
    except DocumentParseError as e:
        return ParseResult(ok=False, error=str(e))
    except Exception as e:  # 防御：解析库的未预期异常不能炸掉插件
        return ParseResult(ok=False, error=f"解析出错：{type(e).__name__}: {e}")

    return ParseResult(ok=True, text=text, stats=stats)
