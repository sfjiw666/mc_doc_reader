"""AstrBot 文档解析助手插件。

在飞书（lark）等平台的会话中接收 docx / xlsx / pptx / pdf 等文档，
自动提取正文并提供 AI 总结、文档问答能力。

工作流程：
1. 监听所有消息，发现文件组件（File）后按扩展名判断是否为支持的文档；
2. 通过 File 组件获取本地文件（飞书适配器已把文件下载到本地临时目录）；
3. 在线程池中调用 doc_parser.parse_document 提取正文（避免阻塞事件循环）；
4. 回复解析统计与内容预览，并把全文缓存到会话（多文档、落盘持久化，
   插件重载 / AstrBot 重启后仍然保留）；
5. 通过 on_llm_request 钩子把文档正文注入后续 LLM 请求的 system_prompt，
   让用户用日常聊天即可询问文档内容（无需 /doc 指令）。
"""

from __future__ import annotations

import asyncio
import hashlib
import json
import os
import re
import time
from dataclasses import asdict, dataclass
from functools import partial

from astrbot.api import AstrBotConfig, logger
from astrbot.api.event import AstrMessageEvent, filter
from astrbot.api.star import Context, Star, register

try:  # 兼容不同 AstrBot 版本中消息组件的导出位置
    from astrbot.api.message_components import File
except ImportError:  # pragma: no cover
    from astrbot.core.message.components import File

try:  # LLM 请求实体，仅用于类型标注
    from astrbot.api.provider import ProviderRequest
except ImportError:  # pragma: no cover
    ProviderRequest = None  # type: ignore[assignment,misc]

try:  # 以 data.plugins 包形式加载时使用相对导入，单独调试时回退为直接导入
    from .doc_parser import (
        LEGACY_EXTS,
        LEGACY_HINT,
        SUPPORTED_EXTS,
        parse_document,
        sniff_ext,
    )
except ImportError:  # pragma: no cover
    from doc_parser import (
        LEGACY_EXTS,
        LEGACY_HINT,
        SUPPORTED_EXTS,
        parse_document,
        sniff_ext,
    )

PLUGIN_NAME = "astrbot_plugin_doc_reader"
# 持久化数据放在 data 目录下（而非插件目录），避免更新插件时被覆盖
DATA_DIR = os.path.join("data", "plugin_data", PLUGIN_NAME)
SESSIONS_DIR = os.path.join(DATA_DIR, "sessions")  # 会话文档记忆持久化目录
CACHE_KEEP = 20  # 提取文本缓存文件最多保留的份数
MAX_SESSION_FILES = 100  # 会话记忆文件最多保留的份数（按修改时间淘汰）

# 注入 system_prompt 的文档上下文标记，防止同一请求被重复注入
INJECT_MARKER = "<doc_reader_doc_context>"

# 插件配置缺省值。正常情况下 AstrBot 会按 _conf_schema.json 自动填充默认值，
# 这里作为无 schema 环境（如极旧版本 AstrBot）下的兜底。
DEFAULTS = {
    "auto_summary": False,
    "inject_to_llm": True,
    "send_preview": True,
    "preview_chars": 1200,
    "max_chars": 60000,
    "llm_max_chars": 30000,
    "max_file_mb": 30,
    "max_docs_per_session": 3,
    "max_sheets": 10,
    "max_rows_per_sheet": 200,
    "max_pages": 100,
    "max_slides": 100,
    "hint_legacy": True,
    "only_wake_in_group": False,
    "summary_prompt": "请用中文总结以下文档的主要内容：先给出一段总体概述，再分点列出关键信息，语言简洁准确。",
}


@dataclass
class DocSession:
    """一份成功解析的文档。"""

    name: str  # 原始文件名
    ext: str  # 小写扩展名
    size: int  # 文件字节数
    text: str  # 提取出的正文
    stats: dict  # 解析统计（页数/工作表数等）
    path: str  # 原始文件在服务器上的本地路径（临时目录，可能被清理）
    ts: float  # 解析完成时间戳


@register(
    "astrbot_plugin_doc_reader",
    "sfjiw",
    "接收并解析 docx / xlsx / pptx / pdf 等文档，多文档持久记忆，支持 AI 总结与文档问答",
    "1.1.0",
)
class DocReaderPlugin(Star):
    def __init__(self, context: Context, config: AstrBotConfig = None):
        super().__init__(context)
        # 无 _conf_schema.json 的环境下 AstrBot 只传 context，config 为 None
        self.config = config if config is not None else {}
        # unified_msg_origin -> 该会话已记住的文档列表（最新在末尾）
        self._sessions: dict[str, list[DocSession]] = {}
        self._locks: dict[str, asyncio.Lock] = {}  # 同一会话串行处理，防止并发解析

    async def initialize(self):
        """插件加载后的异步初始化：从磁盘恢复各会话的文档记忆。"""
        try:
            os.makedirs(DATA_DIR, exist_ok=True)
        except Exception as e:
            logger.warning(f"[doc_reader] 数据目录创建失败（将以仅内存模式运行）: {e}")
        self._load_sessions()
        logger.info(
            f"[doc_reader] 文档解析助手已加载，支持格式: {' '.join(sorted(SUPPORTED_EXTS))}；"
            f"已恢复 {len(self._sessions)} 个会话的文档记忆"
        )

    # ------------------------------------------------------------------
    # 配置与工具函数
    # ------------------------------------------------------------------

    def _get(self, key: str):
        """读取配置项，带回退默认值与类型保护。"""
        default = DEFAULTS[key]
        value = None
        try:
            value = self.config.get(key)
        except Exception:
            pass
        if value is None or value == "":
            return default
        if isinstance(default, int) and not isinstance(default, bool):
            try:
                return int(value)
            except (TypeError, ValueError):
                return default
        return value

    def _limits(self) -> dict:
        """构造传给解析器的限制字典。"""
        return {
            "max_chars": int(self._get("max_chars")),
            "max_sheets": int(self._get("max_sheets")),
            "max_rows_per_sheet": int(self._get("max_rows_per_sheet")),
            "max_pages": int(self._get("max_pages")),
            "max_slides": int(self._get("max_slides")),
        }

    @staticmethod
    def _file_ext(name: str) -> str:
        return name.rsplit(".", 1)[-1].lower() if "." in name else ""

    @staticmethod
    def _human_size(n: float) -> str:
        for unit in ("B", "KB", "MB", "GB"):
            if n < 1024 or unit == "GB":
                return f"{int(n)} B" if unit == "B" else f"{n:.1f} {unit}"
            n /= 1024
        return f"{n:.1f} GB"

    @staticmethod
    def _fit_text(text: str, limit: int) -> str:
        """把超长文本裁剪到 limit 附近，保留头尾，省略中间。"""
        if limit <= 0 or len(text) <= limit:
            return text
        head = text[: int(limit * 0.6)]
        tail = text[-int(limit * 0.25) :] if limit > 500 else ""
        omitted = len(text) - len(head) - len(tail)
        return f"{head}\n\n…[已省略中间 {omitted} 字符]…\n\n{tail}"

    @staticmethod
    def _strip_think(text: str) -> str:
        """去掉部分推理模型输出中的 <think> 标签内容。"""
        return re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL | re.IGNORECASE)

    # ------------------------------------------------------------------
    # 文件获取与缓存
    # ------------------------------------------------------------------

    @staticmethod
    def _read_sync(path: str) -> bytes:
        with open(path, "rb") as f:
            return f.read()

    @staticmethod
    async def _read_file(comp: File) -> tuple[bytes, str]:
        """从 File 组件取得文件内容。

        注意：await comp.get_file() 返回的是本地文件路径（必要时会自动下载），
        而不是字节内容，因此这里读取磁盘。飞书适配器在收到文件消息时已经
        通过开放平台 API 把文件下载到了本地临时目录。
        """
        path = await comp.get_file()
        if not path:
            raise RuntimeError("无法获取文件内容（下载失败或文件已过期），请重新发送")
        loop = asyncio.get_running_loop()
        return await loop.run_in_executor(None, DocReaderPlugin._read_sync, path), path

    def _save_extracted(self, name: str, text: str) -> str:
        """把提取出的全文落盘，方便用户事后取用；失败不影响主流程。"""
        try:
            os.makedirs(DATA_DIR, exist_ok=True)
        except Exception as e:
            logger.debug(f"[doc_reader] 创建数据目录失败: {e}")
            return ""
        safe = re.sub(r"[\\/:*?\"<>|\s]+", "_", name)[:60] or "doc"
        digest = hashlib.md5(text.encode("utf-8", "ignore")).hexdigest()[:8]
        path = os.path.join(
            DATA_DIR, f"{time.strftime('%Y%m%d-%H%M%S')}_{digest}_{safe}.txt"
        )
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(text)
        except Exception as e:
            logger.debug(f"[doc_reader] 保存提取文本失败: {e}")
            return ""
        self._prune_cache()
        return path

    def _prune_cache(self, keep: int = CACHE_KEEP) -> None:
        """只保留最近 keep 份提取文本，避免数据目录无限增长。"""
        try:
            files = [
                os.path.join(DATA_DIR, f)
                for f in os.listdir(DATA_DIR)
                if os.path.isfile(os.path.join(DATA_DIR, f))
            ]
            files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
            for old in files[keep:]:
                os.remove(old)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # 会话文档记忆的持久化（插件重载 / 重启后仍保留）
    # ------------------------------------------------------------------

    @staticmethod
    def _session_file(umo: str) -> str:
        """UMO 含 ':'、'/' 等分隔符，用 sha1 作为会话记忆文件名。"""
        digest = hashlib.sha1(umo.encode("utf-8")).hexdigest()
        return os.path.join(SESSIONS_DIR, f"{digest}.json")

    def _save_session(self, umo: str) -> None:
        """把会话的文档列表写入磁盘；失败仅降级为内存模式，不影响主流程。"""
        docs = self._sessions.get(umo) or []
        try:
            os.makedirs(SESSIONS_DIR, exist_ok=True)
            payload = {
                "umo": umo,
                "updated": time.time(),
                "docs": [asdict(d) for d in docs],
            }
            with open(self._session_file(umo), "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False)
            self._prune_session_files()
        except Exception as e:
            logger.warning(f"[doc_reader] 会话记忆写入磁盘失败（仅内存模式）: {e}")

    def _delete_session(self, umo: str) -> None:
        """清除会话记忆：内存 + 磁盘。"""
        self._sessions.pop(umo, None)
        try:
            path = self._session_file(umo)
            if os.path.exists(path):
                os.remove(path)
        except Exception as e:
            logger.warning(f"[doc_reader] 会话记忆文件删除失败: {e}")

    def _load_sessions(self) -> None:
        """启动时从磁盘恢复各会话的文档记忆。"""
        self._sessions = {}
        if not os.path.isdir(SESSIONS_DIR):
            return
        total_docs = 0
        for fname in os.listdir(SESSIONS_DIR):
            if not fname.endswith(".json"):
                continue
            path = os.path.join(SESSIONS_DIR, fname)
            try:
                with open(path, encoding="utf-8") as f:
                    payload = json.load(f)
                umo = str(payload.get("umo") or "").strip()
                raw_docs = payload.get("docs")
                if not umo or not isinstance(raw_docs, list):
                    continue
                docs = []
                for d in raw_docs:
                    if not isinstance(d, dict) or not d.get("text"):
                        continue
                    d.setdefault("stats", {})
                    d.setdefault("path", "")
                    d.setdefault("ts", 0.0)
                    docs.append(DocSession(**d))
                if docs:
                    self._sessions[umo] = docs
                    total_docs += len(docs)
            except Exception as e:
                logger.warning(f"[doc_reader] 跳过损坏的会话记忆文件 {fname}: {e}")
        if total_docs:
            logger.info(f"[doc_reader] 已从磁盘恢复 {total_docs} 份文档记忆")

    @staticmethod
    def _prune_session_files(keep: int = MAX_SESSION_FILES) -> None:
        """按修改时间淘汰最旧的会话记忆文件，控制磁盘占用。"""
        try:
            files = [
                os.path.join(SESSIONS_DIR, f)
                for f in os.listdir(SESSIONS_DIR)
                if f.endswith(".json")
            ]
            files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
            for old in files[keep:]:
                os.remove(old)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # LLM 调用
    # ------------------------------------------------------------------

    async def _get_provider(self, umo: str):
        """获取当前可用的文本生成 LLM 提供商。

        优先使用新版异步 API（get_using_provider_async），
        旧版本 AstrBot 上回退到同步 API（get_using_provider）。
        """
        getter_async = getattr(self.context, "get_using_provider_async", None)
        if getter_async is not None:
            try:
                return await getter_async(umo)
            except Exception:
                pass
        getter_sync = getattr(self.context, "get_using_provider", None)
        if getter_sync is not None:
            try:
                return getter_sync(umo)
            except Exception:
                pass
        return None

    def _combined_doc_text(self, umo: str) -> str:
        """把会话中记住的多份文档拼成一段带标题的上下文，总量不超过 llm_max_chars。"""
        docs = self._sessions.get(umo) or []
        if not docs:
            return ""
        total_budget = max(int(self._get("llm_max_chars")), 1000)
        per_doc = max(total_budget // len(docs), 1000)
        blocks = [
            f"【文档 {i}：{d.name}】\n{self._fit_text(d.text, per_doc)}"
            for i, d in enumerate(docs, 1)
        ]
        return "\n\n".join(blocks)

    async def _llm_doc_task(self, umo: str, instruction: str) -> str:
        """把「指令 + 当前会话全部文档内容」发给 LLM 并返回纯文本结果。"""
        docs = self._sessions.get(umo) or []
        if not docs:
            raise RuntimeError("当前会话还没有已解析的文档，请先直接发送一个文档文件")
        provider = await self._get_provider(umo)
        if provider is None:
            raise RuntimeError(
                "未配置 LLM 提供商，请先在 WebUI「配置页 → 服务提供商」中启用一个对话模型"
            )
        prompt = f"{instruction}\n\n{self._combined_doc_text(umo)}"
        resp = await provider.text_chat(prompt=prompt, session_id=umo, contexts=[])
        text = getattr(resp, "completion_text", "") or ""
        return self._strip_think(text).strip()

    # ------------------------------------------------------------------
    # 消息监听：自动解析文件
    # ------------------------------------------------------------------

    @filter.event_message_type(filter.EventMessageType.ALL)
    async def on_message(self, event: AstrMessageEvent):
        """监听所有消息，发现文件后自动解析（任何文件消息都会给出反馈）。"""
        chain = event.get_messages()
        if not chain:
            return

        # 忽略机器人自己发出的消息（两侧都需非空，防止 None == None 误判）
        try:
            sender_id = event.get_sender_id()
            self_id = event.get_self_id()
            if sender_id and self_id and str(sender_id) == str(self_id):
                return
        except Exception:
            pass

        # 可选：群聊中仅在被唤醒（@机器人/唤醒词）时才解析文件
        if self._get("only_wake_in_group") and not event.is_at_or_wake_command:
            try:
                if event.get_group_id():
                    return
            except Exception:
                pass

        file_comps = [c for c in chain if isinstance(c, File)]
        if not file_comps:
            return

        # info 级：便于在现场日志中直接确认文件事件到达了本插件
        logger.info(
            "[doc_reader] 检测到文件消息: "
            f"{[getattr(c, 'name', '') for c in file_comps]}"
        )

        umo = event.unified_msg_origin
        lock = self._locks.setdefault(umo, asyncio.Lock())
        async with lock:
            try:
                replies = await self._handle_file(event, file_comps)
                event.stop_event()  # 文件已由本插件接管，避免 LLM 等其他处理器响应
                for msg in replies:
                    yield event.plain_result(msg)
            except Exception as e:
                logger.error(f"[doc_reader] 处理文档失败: {e}")
                event.stop_event()
                yield event.plain_result(f"❌ 文档处理失败：{e}")

    async def _handle_file(
        self, event: AstrMessageEvent, comps: list[File]
    ) -> list[str]:
        """选取并解析消息中的文件，返回要回复的消息列表。

        选取优先级：扩展名支持的文件 > 魔数嗅探可识别的文件；
        若消息中没有可解析的文件，也会给出明确反馈
        （证明文件确实送达了机器人，而不是「没收到」）。
        """
        umo = event.unified_msg_origin
        max_bytes = int(self._get("max_file_mb")) * 1024 * 1024

        chosen: tuple[str, str, bytes, str] | None = None  # (name, ext, data, path)
        legacy_name = ""
        fallback_name, fallback_ext = "", ""

        for comp in comps:
            name = (getattr(comp, "name", "") or "").strip() or "未命名文件"
            ext = self._file_ext(name)

            # 1) 扩展名直接支持：最优选择
            if ext in SUPPORTED_EXTS:
                data, path = await self._read_file(comp)
                chosen = (name, ext, data, path)
                break

            # 2) 旧版 Office 格式：记住文件名，稍后统一提示转存
            if ext in LEGACY_EXTS and not legacy_name:
                legacy_name = name
                continue

            # 3) 扩展名缺失/不可识别：下载后用文件头嗅探真实类型
            if ext in ("", "zip", "bin"):
                try:
                    data, path = await self._read_file(comp)
                    if len(data) <= max_bytes:
                        sniffed = sniff_ext(data)
                        if sniffed:
                            chosen = (name, sniffed, data, path)
                            break
                except Exception as e:
                    logger.warning(f"[doc_reader] 读取文件 {name} 失败: {e}")

            if not fallback_name:
                fallback_name, fallback_ext = name, ext

        if chosen is None:
            # 没有可解析的文件：给出明确反馈，避免「机器人说没收到文件」的困惑
            if legacy_name and self._get("hint_legacy"):
                ext = self._file_ext(legacy_name)
                to_fmt = LEGACY_HINT.get(ext, "新版格式")
                return [
                    f"❌ 已收到「{legacy_name}」，但暂不支持旧版 Office 格式 .{ext}，"
                    f"请在 Office/WPS 中另存为 .{to_fmt} 后重新发送。"
                ]
            if fallback_name:
                ext_disp = fallback_ext or "未知"
                return [
                    f"❌ 已收到文件「{fallback_name}」，但格式 .{ext_disp} 暂不支持解析。\n"
                    f"当前支持：docx / xlsx / pptx / pdf / txt / md / csv / json / html 等。"
                ]
            return []

        name, ext, data, path = chosen
        size = len(data)

        if size > max_bytes:
            return [
                f"❌ 文件过大（{self._human_size(size)}），"
                f"当前处理上限为 {self._get('max_file_mb')} MB，可在插件配置中调整。"
            ]

        # 解析可能较耗时，放到线程池执行，避免阻塞事件循环
        loop = asyncio.get_running_loop()
        result = await loop.run_in_executor(
            None, partial(parse_document, name, data, self._limits(), ext_hint=ext)
        )
        if not result.ok:
            return [f"❌ {result.error}"]

        cache_path = self._save_extracted(name, result.text)
        docs = self._sessions.setdefault(umo, [])
        docs.append(
            DocSession(
                name=name,
                ext=ext,
                size=size,
                text=result.text,
                stats=result.stats,
                path=path,
                ts=time.time(),
            )
        )
        # 只保留最近 N 份文档（可配置），超出淘汰最旧的
        max_docs = max(int(self._get("max_docs_per_session")), 1)
        dropped = len(docs) - max_docs
        if dropped > 0:
            del docs[:-max_docs]
        self._save_session(umo)  # 落盘，重载/重启后记忆仍在
        logger.info(
            f"[doc_reader] 已解析文档 {name}（{self._human_size(size)}），"
            f"提取 {len(result.text)} 字符，stats={result.stats}，"
            f"会话已记住 {len(docs)} 份文档"
        )

        # 组装回复
        lines = [f"📄 已解析文档：{name}"]
        summary_line = f"大小：{self._human_size(size)}"
        if result.stats:
            summary_line += "　" + "　".join(
                f"{k}：{v}" for k, v in result.stats.items()
            )
        lines.append(summary_line)
        lines.append(f"提取字符数：{len(result.text)}")
        if len(comps) > 1:
            lines.append(
                f"（本条消息包含 {len(comps)} 个文件，仅处理第一个支持的文件）"
            )
        if cache_path:
            lines.append(f"完整文本已保存：{cache_path}")

        if self._get("send_preview"):
            preview = result.text[: int(self._get("preview_chars"))].rstrip()
            if preview:
                lines.append("\n———————— 内容预览 ————————")
                lines.append(preview)
                lines.append("———————— 预览结束 ————————")

        if dropped > 0:
            lines.append(f"📚 已记住 {len(docs)} 份文档（淘汰了最旧的 {dropped} 份）")
        else:
            lines.append(f"📚 已记住 {len(docs)} 份文档")
        lines.append("现在直接向我提问即可，无需任何指令；也可用 /doc 总结 · /doc 提问")
        replies = ["\n".join(lines)]

        # 可选：解析完立即让 AI 总结
        if self._get("auto_summary"):
            try:
                summary = await self._llm_doc_task(umo, self._get("summary_prompt"))
                replies.append(f"🤖 AI 总结：\n\n{summary}")
            except Exception as e:
                logger.warning(f"[doc_reader] 自动总结失败: {e}")
                replies.append(f"⚠️ 自动总结失败：{e}")

        return replies

    # ------------------------------------------------------------------
    # LLM 请求钩子：把当前文档注入普通对话
    # ------------------------------------------------------------------

    @filter.on_llm_request()
    async def on_llm_request(
        self, event: AstrMessageEvent, req: ProviderRequest
    ) -> None:
        """LLM 请求前，把会话记住的全部文档正文注入 system_prompt。

        这样用户发送文档后，直接用日常聊天提问（如「刚才那份报告的重点是什么」），
        模型也能看到文档内容，而不必使用 /doc 提问。
        """
        try:
            if not self._get("inject_to_llm"):
                return
            docs = self._sessions.get(event.unified_msg_origin) or []
            if not docs:
                return
            base = req.system_prompt or ""
            if INJECT_MARKER in base:
                return
            body = self._combined_doc_text(event.unified_msg_origin)
            names = "、".join(d.name for d in docs)
            req.system_prompt = (
                f"{base}\n{INJECT_MARKER}\n"
                f"用户在当前会话发送过 {len(docs)} 份文档（{names}），"
                f"以下是提取出的正文。当用户的问题与这些文档相关时，"
                f"请依据此内容回答；超长文档可能被截断：\n\n{body}\n"
            ).strip()
            # info 级：在日志中确认注入确实发生（排查「机器人不记得文档」时先看这里）
            logger.info(
                f"[doc_reader] 已为 LLM 请求注入文档上下文：{len(docs)} 份，"
                f"共 {len(body)} 字符"
            )
        except Exception as e:
            logger.warning(f"[doc_reader] 注入文档上下文失败: {e}")

    # ------------------------------------------------------------------
    # /doc 指令
    # ------------------------------------------------------------------

    @filter.command("doc", alias={"文档"})
    async def doc_cmd(self, event: AstrMessageEvent):
        """文档解析管理指令：doc 总结/提问/状态/清除/帮助"""
        try:
            # handler 中 message_str 仍包含指令名本身（如 "doc 总结"），自行解析
            tokens = event.message_str.replace("/", " ").split()
            if tokens and tokens[0].strip("/").lower() in ("doc", "文档"):
                tokens = tokens[1:]
            sub = tokens[0].lower() if tokens else ""
            query = " ".join(tokens[1:]).strip()
            umo = event.unified_msg_origin
            docs = self._sessions.get(umo) or []

            if sub in ("", "help", "帮助"):
                yield event.plain_result(self._help_text(docs))
            elif sub in ("status", "状态"):
                yield event.plain_result(self._status_text(docs))
            elif sub in ("clear", "清除", "清空", "forget", "删除"):
                self._delete_session(umo)
                yield event.plain_result("✅ 已清除当前会话记住的全部文档。")
            elif sub in ("summary", "总结", "摘要", "概要"):
                for msg in await self._do_summary(event, docs):
                    yield event.plain_result(msg)
            elif sub in ("ask", "提问", "问答", "问", "qa"):
                if not query:
                    yield event.plain_result(
                        "用法：/doc 提问 <你的问题>\n"
                        "例如：/doc 提问 这份报告的核心结论是什么？"
                    )
                else:
                    for msg in await self._do_ask(event, docs, query):
                        yield event.plain_result(msg)
            else:
                yield event.plain_result(
                    f"未知的子指令：{sub}\n可用子指令：总结 / 提问 / 状态 / 清除 / 帮助"
                )
        except Exception as e:
            logger.error(f"[doc_reader] /doc 指令执行失败: {e}")
            yield event.plain_result(f"❌ 指令执行失败：{e}")

    async def _do_summary(
        self, event: AstrMessageEvent, docs: list[DocSession]
    ) -> list[str]:
        """执行 AI 总结（基于会话全部文档），返回多条待回复消息。"""
        if not docs:
            return ["当前会话还没有已解析的文档，请先直接发送一个文档文件。"]
        try:
            summary = await self._llm_doc_task(
                event.unified_msg_origin, self._get("summary_prompt")
            )
        except Exception as e:
            logger.error(f"[doc_reader] AI 总结失败: {e}")
            return [f"❌ AI 总结失败：{e}"]
        scope = docs[-1].name if len(docs) == 1 else f"共 {len(docs)} 份文档"
        return ["⏳ 正在阅读文档，请稍候…", f"🤖 文档总结（{scope}）：\n\n{summary}"]

    async def _do_ask(
        self, event: AstrMessageEvent, docs: list[DocSession], question: str
    ) -> list[str]:
        """基于会话全部文档回答问题。"""
        if not docs:
            return ["当前会话还没有已解析的文档，请先直接发送一个文档文件。"]
        instruction = (
            f"请根据下面的文档内容回答问题：{question}\n"
            "只依据文档内容回答；如果文档中没有相关信息，请明确说明。\n\n"
        )
        try:
            answer = await self._llm_doc_task(event.unified_msg_origin, instruction)
        except Exception as e:
            logger.error(f"[doc_reader] 文档问答失败: {e}")
            return [f"❌ 文档问答失败：{e}"]
        return ["⏳ 正在查阅文档，请稍候…", f"🤖 关于「{question}」：\n\n{answer}"]

    def _help_text(self, docs: list[DocSession]) -> str:
        return (
            "📖 文档解析助手\n"
            "直接发送文档文件即可自动解析，支持：docx / xlsx / pptx / pdf / "
            "txt / md / csv / json / html 等格式。\n"
            "解析成功后，直接用日常聊天提问即可引用文档内容，也可使用以下指令：\n\n"
            "可用指令：\n"
            "• /doc 总结 — AI 总结会话中的文档\n"
            "• /doc 提问 <你的问题> — 基于会话文档问答\n"
            "• /doc 状态 — 查看已记住的文档\n"
            "• /doc 清除 — 清除当前会话记住的全部文档\n\n"
            f"{self._status_text(docs)}\n"
            f"提示：每个会话最多记住最近 {max(int(self._get('max_docs_per_session')), 1)} 份文档，"
            "记忆持久化在服务器上，机器人重启后依然保留。"
        )

    def _status_text(self, docs: list[DocSession]) -> str:
        if not docs:
            return "当前会话暂无缓存文档，直接发送文件即可解析。"
        lines = [f"📄 当前会话已记住 {len(docs)} 份文档："]
        for i, d in enumerate(docs, 1):
            lines.append(
                f"{i}. {d.name}（{(d.ext or '未知').upper()}，"
                f"{self._human_size(d.size)}，提取 {len(d.text)} 字符，"
                f"{time.strftime('%Y-%m-%d %H:%M', time.localtime(d.ts))}）"
            )
        return "\n".join(lines)

    async def terminate(self):
        """插件被卸载/停用时清理内存缓存（磁盘记忆保留，重启后自动恢复）。"""
        self._sessions.clear()
        self._locks.clear()
        logger.info("[doc_reader] 文档解析助手已停止（磁盘记忆已保留）")
