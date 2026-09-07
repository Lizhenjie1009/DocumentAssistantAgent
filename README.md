# docx-formatter

把源 Word 文档「套用」成模板文档的格式。以**匹配到的模板为底板**（继承其样式、分节、页眉页脚、自动编号），再把源文档正文按顺序搬入并套模板样式，最后重建目录、另存新文件并生成处理报告。**原文件只读，绝不改动。**

## 能力对照

| 需求 | 实现 |
|---|---|
| 多种模板 + 自动识别 | 按「章节指纹」（1/2 级标题结构）相似度匹配，可配关键词加权 |
| 章节结构 + 标题层级 | 源标题层级 → 模板 `Heading 1~5` |
| 编号 | 剥离源手打编号，套模板多级自动编号（1 / 1.1 / 1.1.1） |
| 目录 | 静态条目 + 占位页码（默认 `00`）+ 点线引导 + 与正文分页（沿用模板分节） |
| 页眉页脚 | 直接继承模板（正文前罗马页码 / 正文阿拉伯页码自动生效） |
| 表格 | 复制结构、边框、合并单元格、列宽、单元格图片 |
| 图片 | 原位复制，按模板规范缩放（默认：正文可用宽度为上限、超出才缩小、不放大、居中） |
| 报告 + 原文件不动 | 每文件另存为 `<源名>_<模板名>.docx`，输出 `report.md` |

## 安装

```bash
pip install -r requirements.txt   # python-docx  PyYAML  Pillow
```

## 使用

```bash
# 单个文件
python convert.py -i 某系统测试报告.docx -t "软件文档模板\软件文档模板" -o output

# 批量（整个文件夹）
python convert.py -i input -t "软件文档模板\软件文档模板" -o output

# 带配置文件
python convert.py -i input -t "软件文档模板\软件文档模板" -o output -c config.yaml

# 检查生成的 docx（打开校验 + 结构汇总）
python check_output.py output\某系统测试报告_测试报告模板.docx   # 单个文件
python check_output.py output                                     # 整个文件夹

# 查看运行历史（SQLite，默认库 run_history.db）
python convert.py --history                # 最近 10 次
python convert.py --history 5              # 最近 5 次
python convert.py -i ... -o ... --db my.db # 自定义历史库位置

# AI 起草文档并自动转换（功能2）
python convert.py --generate -t "软件文档模板\软件文档模板" -o output
```

参数：

| 参数 | 说明 |
|---|---|
| `-i/--input` | 源 `.docx` 文件或文件夹（必填） |
| `-t/--templates` | 模板文件夹（必填） |
| `-o/--output` | 输出文件夹（默认 `output`） |
| `-c/--config` | 可选 YAML 配置 |
| `-r/--report` | 报告路径（默认 `<output>/report.md`） |

## 配置（config.yaml，可选）

不提供时用内置默认值即可跑通。需要微调时参考 `config.example.yaml`。

```yaml
defaults:
  numbering: auto            # auto=剥离源编号套模板自动编号；keep=保留源编号
  image:
    max_width_ratio: 1.0     # 图片最大宽度 = 正文可用宽度 × 比例
    enlarge: false           # 不放大
    center: true             # 独占一行的图片居中
  toc:
    levels: [1, 2]           # 目录收录的标题层级
    placeholder: "00"        # 占位页码
    tab_leader: true         # 点线引导
  table_style: null          # 强制表格样式名；null=原样复制源表格格式
  title:
    placeholder_regex: '[X×Ｘx]{1,3}分?系统'   # 封面上要替换为源标题的占位符
    source: auto             # auto=取源 Title 样式标题，否则用文件名；filename；null 关闭
templates:                   # 按模板文件名单独覆盖
  "测试报告模板.docx":
    match:
      keywords: [测试报告, 测试总结]   # 命中任一关键词则相似度加分
      min_similarity: 0.35
```

## 输出

- `output/<源名>_<模板名>.docx` —— 转换结果
- `output/report.md` —— 处理报告（匹配模板、相似度、标题/段落/表格/图片统计、备注）

## AI 起草文档并转换（功能 2，可选）

`--generate` 模式：命令行问答（选模板 → 填标题/系统名/素材要点）→ 调用 LLM 按**所选模板的章节大纲**起草结构化内容 → 渲染成带真实标题样式的 docx 放进 `input\` → 自动走转换流水线输出到 `output\`。

```bash
# 方式一：用环境变量提供 key（一次设置后免配置）
set DEEPSEEK_API_KEY=sk-xxxx            # Windows CMD
$env:DEEPSEEK_API_KEY = 'sk-xxxx'       # PowerShell
python convert.py --generate -t "软件文档模板\软件文档模板"

# 方式二：配置文件（参考 llm.example.yaml）
python convert.py --generate -t "软件文档模板\软件文档模板" --llm llm.yaml
```

说明：
- LLM 默认 DeepSeek API（`base_url/model` 可在 llm 配置里改，兼容其他 OpenAI 接口）。
- 素材要点**留空**时 AI 会写合理占位内容，转换后请人工核对替换。
- 图片无法由 AI 直接生成，相关位置会写红色占位提示「【此处插入XXX图片】」，转换后自行补图。

## 运行历史（SQLite）

每次运行 `convert.py` 都会自动把详细信息记录到 SQLite 库（默认 `run_history.db`，可用 `--db` 改位置）：

- `runs` 表：每次运行的时间、输入/模板/输出路径、报告路径、成功/失败/总数、命令行参数
- `files` 表：每个文件的输入/输出路径、状态、匹配模板、相似度、标题/段落/表格/图片统计、警告、生成的目录条目明细、错误信息

用 `python convert.py --history [N]` 查看最近 N 次运行及最近一次的文件明细。历史库为追加式，多次运行都保留，方便对比匹配结果和排查失败原因。

## 目录结构

```text
convert.py               CLI 入口（转换）
check_output.py          检查生成的 docx（打开校验 + 打印标题/目录/表格/图片汇总）
formatter/
  config.py              模板发现 + 配置
  matcher.py             章节指纹匹配
  mapper.py              核心：模板底板 + 正文映射 + 标题填充
  tables.py              表格复制（结构/边框/合并/列宽）
  images.py              图片提取 / 缩放重插
  content.py             段落内文本+图片复制
  toc.py                 目录重建
  report.py              处理报告
  history.py             SQLite 运行历史记录
  generate.py            AI 起草（功能2：问答→结构化JSON→渲染源docx）
  util.py                公共工具（样式查找/标题识别/编号剥离）
```

## 已知边界（python-docx 限制）

- 目录页码为**占位符**，需在 Word 里手工补；若想真实页码，可改用 Word 目录域（`toc.tab_leader` 相关方案后续可加）。
- 浮动（锚定环绕）图片默认转成行内图片。
- 源文档里的超链接只保留文字，不保留链接。
- 模板前置的「更改记录」等固定页不在自动生成的静态目录中（目录只反映源文档标题）。
- 表格复杂底纹/边框尽量原样复制，但极复杂的表格样式可能有细微差异。
