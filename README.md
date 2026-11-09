# astrbot_plugin_doc_reader

AstrBot 文档解析助手 —— 让机器人读懂你发来的文档。

在飞书（Lark）等平台的对话中直接发送 **docx / xlsx / pptx / pdf** 等文档，
机器人会自动下载并提取正文内容，回复解析统计与内容预览，并支持基于文档的
**AI 总结** 与 **文档问答**。

> 本插件基于 [AstrBot](https://github.com/AstrBotDevs/AstrBot) 插件模板
> [helloworld](https://github.com/Soulter/helloworld) 开发，文档解析使用了
> python-docx、openpyxl、python-pptx、pypdf 等开源库。

# Supports

- [AstrBot Repo](https://github.com/AstrBotDevs/AstrBot)
- [AstrBot Plugin Development Docs (Chinese)](https://docs.astrbot.app/dev/star/plugin-new.html)

## 功能

- **自动解析**：会话中收到文档文件后自动下载并提取正文，无需任何指令；
- **持久记忆**：每个会话最多记住最近 N 份文档（默认 3，可配置），**落盘持久化**——插件重载、机器人重启后记忆依然在；
- **对话直问**：解析成功后，**用日常聊天直接问**（如「刚才那份报告的重点是什么」）即可引用文档内容，无需任何指令——全部已记住文档的正文会自动注入对话上下文；
- **内容预览**：回复文件类型、大小、页数/工作表数等统计信息与正文预览；
- **AI 总结**：`/doc 总结` 一键让当前配置的 LLM 总结会话文档；
- **文档问答**：`/doc 提问 <问题>` 基于会话全部文档回答问题；
- **必反馈**：任何文件消息都会得到明确反馈（解析成功 / 格式不支持 / 文件过大），不会出现「机器人说没收到文件」的困惑；
- **格式嗅探**：文件名没有扩展名或被改名为 .zip/.bin 时，通过文件头魔数自动识别 pdf / docx / xlsx / pptx；
- **安全防护**：文件大小、页数、行数、字符数等多重限制，解析在线程池中执行，不会阻塞机器人。

## 支持的格式

| 格式 | 说明 |
| --- | --- |
| `.docx` | 正文段落（保留标题层级）+ 表格 |
| `.xlsx` / `.xlsm` | 各工作表按行输出（读取公式计算结果） |
| `.pptx` | 每页幻灯片文本 + 表格 + 演讲者备注 |
| `.pdf` | 逐页提取文本层；纯扫描件会明确提示需 OCR |
| `.txt` `.md` `.csv` `.log` `.json` `.html` 等纯文本 | 自动识别 UTF-8 / GBK 等编码 |
| `.doc` `.xls` `.ppt` 旧版二进制格式 | 暂不支持，会提示另存为新格式 |

## 快速开始（飞书）

1. 在 AstrBot WebUI 中配置好 [aiocqhttp/飞书(lark)](https://docs.astrbot.app/platforms/lark.html) 平台适配器；
2. 通过插件市场安装本插件，或手动放置到 `data/plugins/astrbot_plugin_doc_reader`；
3. 在飞书对话中把一个 docx / xlsx / pptx / pdf 文件直接拖进聊天窗口；
4. 机器人回复解析结果，随后直接用聊天提问，或使用 `/doc 总结`、`/doc 提问` 等指令。

> AI 总结与问答需要在 WebUI「服务提供商」中配置并启用至少一个 LLM。

## 指令

| 指令 | 说明 |
| --- | --- |
| `/doc`（或 `/文档`） | 查看帮助与当前文档状态 |
| `/doc 总结` | AI 总结当前会话最近一次解析的文档 |
| `/doc 提问 <问题>` | 基于当前文档内容回答问题 |
| `/doc 状态` | 查看当前缓存的文档信息 |
| `/doc 清除` | 清除当前会话记住的全部文档 |

每个会话（私聊/各群聊）独立记忆，默认保留最近 3 份文档（`max_docs_per_session` 可调），
全部落盘持久化，机器人重启后依然有效。

## 配置项

安装后可在 WebUI 插件页的「配置」中修改：

| 配置项 | 默认值 | 说明 |
| --- | --- | --- |
| `auto_summary` | `false` | 收到文档后自动调用 AI 总结 |
| `inject_to_llm` | `true` | 解析后把文档正文注入对话上下文，普通聊天即可直接询问文档内容 |
| `send_preview` | `true` | 回复中附带内容预览 |
| `preview_chars` | `1200` | 预览的最大字符数 |
| `max_chars` | `60000` | 提取正文的最大字符数，超出截断 |
| `llm_max_chars` | `30000` | 发送给 AI 的文档最大字符数（超长省略中间部分） |
| `max_docs_per_session` | `3` | 每个会话最多记住的文档数量（超出淘汰最旧的） |
| `max_file_mb` | `30` | 允许处理的文件大小上限（MB） |
| `max_sheets` | `10` | xlsx 最多读取的工作表数 |
| `max_rows_per_sheet` | `200` | xlsx 每表最多读取的行数 |
| `max_pages` | `100` | pdf 最多读取的页数 |
| `max_slides` | `100` | pptx 最多读取的幻灯片页数 |
| `hint_legacy` | `true` | 收到 doc/xls/ppt 旧格式时提示转换 |
| `only_wake_in_group` | `false` | 群聊中仅 @机器人 时才解析文件 |
| `summary_prompt` | 见配置页 | AI 总结使用的提示词 |

## 工作原理

```
用户发送文件 (飞书 IM)
        │
        ▼
飞书平台适配器：调用开放平台 API 下载文件到本地临时目录，
构造 File(name=文件名, file=本地路径) 消息组件
        │
        ▼
本插件监听消息，识别支持的扩展名
        │
        ▼
读取文件 → 线程池中解析 (doc_parser.py)
   ├─ docx  → python-docx
   ├─ xlsx  → openpyxl
   ├─ pptx  → python-pptx
   ├─ pdf   → pypdf
   └─ 纯文本 → 编码探测
        │
        ▼
回复统计 + 预览，全文缓存到会话 & data/plugin_data/
        │
        ▼
on_llm_request 钩子：把文档正文注入后续对话的 system_prompt
（普通聊天直接问即可）+ /doc 总结 · /doc 提问 → 调用当前 LLM 提供商
```

## 常见问题

- **机器人说「没收到文件」/「请把内容粘贴过来」**：这是裸 LLM 的回复，说明文件消息没有被插件接管。按顺序排查：
  1. WebUI「插件」页确认本插件**已启用且不在「加载失败」列表**，日志中有
     `[doc_reader] 文档解析助手已加载`；
  2. 发送文件后机器人是否有任何回复？若回复「📄 已解析文档」说明链路正常，是别的问题；
  3. **群聊限制**：飞书机器人默认只能收到群里 @它 的消息，而发文件时无法 @——
     请先在**私聊**中测试，或在群聊中发文件后由管理员确认事件订阅正常；
  4. **飞书应用权限**：需要开通读取消息与下载消息资源的权限（`im:message` 相关 scope），
     日志中若出现 `[Lark] 下载消息资源失败` 即为权限问题；
  5. **云文档链接不是文件**：发送飞书云文档分享卡片只是链接消息，不是文件消息，
     本插件无法解析，请下载为本地文件后发送。
- **日志出现「飞书 暂时不支持消息段: ComponentType.Reply」**：AstrBot 核心
  「平台设置 → 回复时引用消息」（`reply_with_quote`）开启时，会给所有平台的回复
  统一插入 Reply 消息段；飞书适配器不支持用消息段表达引用（飞书的引用走 API
  参数实现，适配器发送时已自动携带），于是丢弃该段并警告。文字照发、功能无损，
  与本插件无关。想消除警告：WebUI → 配置 → 平台设置 → 关闭「回复时引用消息」。
- **提示缺少依赖库**：在 AstrBot 所在的 Python 环境手动执行
  `pip install python-docx openpyxl python-pptx pypdf`，或重装本插件触发依赖安装。
- **PDF 提示“纯扫描件”**：扫描版 PDF 没有文字层，需要 OCR 支持，请先转换/OCR 后再发送。
- **提示“未配置 LLM 提供商”**：总结与问答功能需要在 WebUI 中配置并启用对话模型；
  纯解析与预览功能不依赖 LLM。
- **机器人「不记得」已发送的文档**：先确认 `inject_to_llm` 已开启（默认开），再在日志中
  搜索 `[doc_reader] 已为 LLM 请求注入文档上下文`——提问时出现这行说明注入正常；
  若没有，检查是否安装了同样挂载 LLM 请求钩子的其他插件（如 office_assistant 类插件）
  并尝试暂时停用它排查冲突。记忆按会话隔离：私聊与群聊、不同群之间互不相通。
- **大文件没有反应**：检查是否超过 `max_file_mb` 限制；飞书平台本身对消息资源也有约 30MB 的限制。

## 许可

本项目继承模板的 [LICENSE](LICENSE)。使用的解析库均为各自协议下的开源项目：
python-docx (MIT)、openpyxl (MIT)、python-pptx (MIT)、pypdf (BSD-3-Clause)。
