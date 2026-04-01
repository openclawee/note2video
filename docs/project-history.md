# 项目历史归档

## 文档目的

这份文档用于归档 `Note2Video / 备注成片` 项目的关键背景、设计共识、实现进展和后续计划。

它的目标不是替代详细设计文档，而是作为一个可迁移、可快速接手的项目上下文入口。即使后续修改本地目录名、迁移到新电脑，或者 Cursor 工作区历史发生变化，也可以依靠这份文档快速恢复上下文。

## 项目概览

- 项目名：`Note2Video`
- 中文名：`备注成片`
- 当前形态：开源、免费、`CLI-first`
- 核心目标：把 PowerPoint 备注自动转换为配音、字幕和讲解视频

项目的基本定位已经记录在：

- `docs/decisions/001-product-positioning.md`
- `docs/decisions/002-cli-first-strategy.md`
- `docs/decisions/003-mvp-scope.md`

## 项目起点

这个项目最初不是作为一个普通本地工具来思考的，而是从“做成可复用 skill / 智能体工具”的方向出发。

逐步收敛后，确定了以下产品判断：

- 与其先做 UI、插件或编辑器，不如先做一个可自动化、可调试、可复用的 CLI。
- 与其直接操作剪映等编辑器，不如先把从 `pptx` 到 `mp4` 的核心流水线打通。
- 项目的核心竞争力不在“特效编辑”，而在“把已有 PPT 内容资产高效变成讲解视频”。
- 未来如果封装成 skill，应该直接复用 CLI，而不是重写一套逻辑。

## 已确定的产品方向

### 主要价值

- 复用现有 PPT 与备注内容
- 自动完成导图、提取备注、生成脚本、配音、字幕和视频
- 让整个流程可以脚本化、可重复执行、可被智能体调用

### 明确不做的事情

- 不做完整视频编辑器
- 不与剪映、CapCut 直接竞争编辑能力
- 不以复杂动效或多角色配音为第一优先级
- 第一阶段不做 PowerPoint 插件和 Web SaaS

## 当前 CLI 范围

当前主命令和子命令如下：

- `note2video build`
- `note2video extract`
- `note2video voice`
- `note2video voices`
- `note2video subtitle`
- `note2video render`

核心原则：

- 一条命令可跑完整流程
- 每一步也可以单独执行
- 中间产物必须可落盘、可检查、可调试

## 当前实现状态

截至目前，项目已经具备从 `pptx` 到 `mp4` 的最小闭环。

### 已完成能力

- Python 项目骨架与 CLI 入口
- `extract` 真实实现
- 通过 PowerPoint COM 导出页面图片
- 提取每页备注、生成 `raw_notes`、`speaker_notes`、`script`
- 输出结构化 JSON 与逐页 / 汇总 `.txt` 文件
- `voice` 配音流程
- 支持 `pyttsx3` 与 `edge-tts` 两个 TTS provider
- `voices` 命令用于列出可用音色
- `subtitle` 生成 `srt` 和 `json`
- `render` 使用 ffmpeg 合成最终视频
- `build` 串联完整流程
- 单元测试覆盖核心命令主路径

### 当前输出结构

典型产物包括：

- `manifest.json`
- `slides/*.png`
- `notes/notes.json`
- `notes/raw/*.txt`
- `notes/speaker/*.txt`
- `notes/all.txt`
- `scripts/script.json`
- `scripts/txt/*.txt`
- `scripts/all.txt`
- `audio/*.wav`
- `audio/merged.wav`
- `audio/timings.json`
- `subtitles/subtitles.srt`
- `subtitles/subtitles.json`
- `video/output.mp4`

## 关键设计演进

### 1. 文本分层

项目中间文本现在分为三层：

- `raw_notes`：从 PPT 备注页中直接提取的原始文本
- `speaker_notes`：做过基础清洗后的可读备注
- `script`：进一步整理成更适合口播与字幕切分的脚本

这样做的原因是：

- 方便排查提取问题
- 方便手工审阅和复制
- 为后续更高级的脚本清洗和配音优化留出空间

### 2. 配音 provider 可插拔

TTS 采用 provider 结构，而不是把逻辑写死在 CLI 里。

当前已接入：

- `pyttsx3`：本地可用，便于离线验证
- `edge-tts`：在线能力更强，音质更适合公开演示

### 3. 字幕时间轴从“估算”升级到“真实句级时长”

项目一开始的字幕时间分配，是按句子长度比例估算的。

后续为提升字幕同步质量，已经升级为：

- `voice` 按 `script` 逐句生成音频
- 再拼成整页音频与整片音频
- 同时写出 `audio/timings.json`
- `subtitle` 优先读取这份真实句级时间轴

这意味着字幕不再单纯按文本长度分配时间，而是依据每句 TTS 的真实时长来生成。

## 已处理的重要问题

### 1. pytest 无法导入项目模块

问题：

- `pytest` 运行时找不到 `note2video`

处理：

- 在 `pyproject.toml` 中为 pytest 增加 `src` 路径

### 2. PowerPoint 隐藏窗口兼容性问题

问题：

- 某些 PowerPoint 版本对 `app.Visible = 0` 兼容性差

处理：

- 改为显式可见模式，提高兼容性

### 3. `Slide.Export` 使用相对路径不稳定

问题：

- PowerPoint COM 在导出图片时对相对路径不稳定

处理：

- 改为传入绝对路径

### 4. 备注提取误把占位内容和页码当备注

问题：

- 没有真实备注时，仍然可能提取出类似 `1 2 3` 之类无效文本

处理：

- 收紧备注区块识别逻辑
- 区分备注占位符和真实备注主体
- 增加文本清洗与过滤

### 5. `raw_notes` 与 `script` 内容过于相近

问题：

- 初期实现中，脚本层没有体现出与原始备注层的差异

处理：

- 增加 `_to_speaker_notes` 与 `_to_script` 两层转换
- 让 `script` 更接近口播和字幕使用场景

### 6. 字幕只显示第一页的第一句或每页只显示第一句

问题：

- 虽然 SRT 文件中有多句字幕，但最终视频里字幕更新异常

原因：

- 图片拼接视频时使用了不利于字幕逐句更新的轨道生成方式

处理：

- 在渲染阶段先生成固定帧率视频，再进行字幕烧录

### 7. Edge TTS 音频格式不统一

问题：

- `edge-tts` 输出结果与后续 WAV 处理链不完全兼容

处理：

- 先输出临时音频，再统一转换为标准 WAV

## 当前验证状态

目前已完成以下验证：

- `tests/test_cli.py` 已通过
- `build` 完整链路已可在样例 `pptx` 上跑通
- 实际产物中已生成：
  - `audio/timings.json`
  - `subtitles/subtitles.srt`
  - `video/output.mp4`
- `manifest.json` 已能正确记录：
  - `merged_audio`
  - `timings`
  - `subtitle`
  - `video`
  - `video_subtitles_burned`

## 当前最值得继续做的方向

### 方向 1：继续提升脚本清洗质量

候选优化包括：

- 数字读法处理
- 英文缩写处理
- 括号内容清洗
- 长句进一步切分
- 更贴近中文口播节奏的断句

### 方向 2：继续优化字幕与口播节奏

在已有逐句 TTS 基础上，还可以继续做：

- 句间停顿控制
- 最短字幕时长限制
- 过短句自动合并
- 过长句自动二次切分
- 页内句间时间微调

### 方向 3：为公开开源做准备

在打算发布到 GitHub 前，建议继续补齐：

- `.gitignore`
- 初版开源仓库说明
- 示例输入输出说明
- 安装依赖与环境要求
- 对 Windows / PowerPoint 依赖的明确说明

## 迁移建议

如果后续要改本地目录名、换电脑或迁移 Cursor 工作区，建议至少保留以下内容：

- 本文档 `docs/project-history.md`
- `README.md`
- `docs/cli.md`
- `docs/decisions/*.md`

如果需要保留更完整的对话背景，可以把关键讨论继续沉淀为新的决策文档，而不要只依赖 IDE 内部聊天记录。

## 后续维护建议

以后每完成一个关键阶段，建议都更新本文档中的以下部分：

- 当前实现状态
- 已处理的重要问题
- 当前验证状态
- 当前最值得继续做的方向

这样项目历史就能持续沉淀，而不是散落在聊天记录里。
