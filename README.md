# Note2Video / 备注成片

`备注成片` 是一个将 PPT 备注自动转换为配音、字幕和讲解视频的开源命令行工具。

项目英文名为 `Note2Video`，当前以 `CLI-first` 方式开发，后续可继续封装为 skill 或其他自动化能力。

## 项目价值

很多培训、教学、知识分享的工作流，本来就是从 PPT 和备注开始的。但把它们变成可以发布的视频，往往还要重复做很多手工操作：

- 导出页面图片
- 把备注整理成讲稿
- 生成配音
- 生成字幕
- 把图片、音频、字幕拼成最终视频

`Note2Video` 聚焦的就是这条流水线，让它变得可脚本化、可重复执行、可被智能体调用。

## 项目目标

- 将 `pptx` 导出为逐页图片
- 提取每页备注
- 将备注整理为适合口播的脚本
- 生成配音音频
- 生成字幕文件
- 渲染最终讲解视频
- 提供稳定的 CLI，方便自动化与 skill 封装

## 非目标

- 不做完整视频编辑器
- 不做复杂动效系统
- 不追求完整保留 PowerPoint 动画
- 不直接与剪映等完整编辑器比拼编辑能力

## 目标用户

- 企业培训和内部赋能团队
- 教师、讲师、课程制作者
- 以 PPT 为内容源的知识创作者
- 需要把这条流程嵌入自动化系统的开发者

## MVP 范围

首个可用版本计划支持：

- 输入：`pptx`
- 输出：页面图片、备注 JSON、逐页文本、脚本 JSON、音频文件、`srt` 与 `mp4`
- 一个主语言流程
- 一个或两个 TTS Provider
- 支持 `16:9` 与 `9:16`
- 支持逐步命令，便于调试

## CLI 概览

主命令：

```bash
note2video build input.pptx --out ./dist
```

当前 `build` 已经会串联执行：

- `extract`
- `voice`
- `subtitle`
- `render`

子命令：

- `build`：执行完整流程
- `extract`：导出页面图片、备注和脚本
- `voice`：根据脚本生成配音
- `voices`：列出可用音色
- `subtitle`：生成字幕文件
- `render`：根据准备好的素材渲染最终视频

详细命令说明见 `docs/cli.md`。

## 平台与幻灯片导出

- **Windows**：优先使用 Microsoft PowerPoint COM 导出真实页面图片；失败时回退为 OpenXML + 占位图。
- **Linux / macOS**：若 `PATH` 上同时能找到 `soffice`（或 `libreoffice`）以及 `pdftoppm`（Poppler），则通过「LibreOffice 无界面转 PDF → `pdftoppm` 切 PNG」导出真实页面图片；否则回退为 OpenXML + 占位图。

Debian / Ubuntu 示例：

```bash
sudo apt install libreoffice-nogui poppler-utils
```

可选环境变量：

- `NOTE2VIDEO_USE_LIBREOFFICE`：设为 `0`、`false` 或 `off` 可禁用 LibreOffice 路径（例如测试或强制占位图）。
- `NOTE2VIDEO_LIBREOFFICE`：`soffice` 可执行文件的绝对路径。
- `NOTE2VIDEO_PDF_RENDER_DPI`：`pdftoppm` 分辨率，默认 `150`。

**Windows**：无需安装 LibreOffice；程序会优先走 PowerPoint COM（与 Linux 路径无关）。

### Docker（仅 Linux 镜像）

容器内预装 `libreoffice-nogui` 与 `poppler-utils`，用于在无桌面环境下导出真实幻灯片图。Windows 上请直接本机安装 Python 运行，不要用此镜像替代 PowerPoint 路径。

```bash
docker build -t note2video .
docker run --rm -v "%CD%:/work" -w /work note2video extract ./deck.pptx --out ./dist
```

（PowerShell 下将卷挂载改为 `-v "${PWD}:/work"`。）

CI（`.github/workflows/ci.yml`）在 **Ubuntu** 上安装上述系统依赖以便与 Docker 一致；在 **Windows** 上**不会**安装 LibreOffice，与本地行为一致。测试任务统一设置 `NOTE2VIDEO_USE_LIBREOFFICE=0`，避免对极简测试用 `.pptx` 做转换导致不稳定；真实 Linux 环境不设该变量即可自动走 LibreOffice。

## 输出目录示例

```text
dist/
  manifest.json
  slides/
    001.png
    002.png
  notes/
    notes.json
    all.txt
    raw/
      001.txt
      002.txt
    speaker/
      001.txt
      002.txt
  scripts/
    script.json
    all.txt
    txt/
      001.txt
      002.txt
  audio/
    001.wav
    002.wav
    merged.wav
  subtitles/
    subtitles.srt
    subtitles.json
  video/
    output.mp4
  logs/
    build.log
```

## 输出说明

当前 `extract` 除了写出结构化 JSON，还会额外生成便于人工复制粘贴的 `.txt` 文件：

- `notes/raw/*.txt`：逐页原始备注文本
- `notes/speaker/*.txt`：逐页可读备注文本
- `scripts/txt/*.txt`：逐页脚本文本
- `notes/all.txt`：整套备注汇总文本
- `scripts/all.txt`：整套脚本汇总文本

这样在剪映等工具里手工复制时，会保留真实换行，而不是看到字面量 `\n`。

当前文本层次分为：

- `raw_notes`：从 PPT 备注页提取出的原始文本
- `speaker_notes`：做过基础清洗后的可读备注
- `script`：进一步做了基础分句后的口播脚本

## 规划中的架构

```text
src/note2video/
  cli/
  parser/
  script/
  tts/
  subtitle/
  render/
  schemas/
```

模块职责：

- `parser`：解析 `pptx`、导出页面图片、提取备注
- `script`：清洗并规范化备注文本，生成口播脚本
- `tts`：配音能力的 Provider 抽象层
- `subtitle`：字幕切分与时间轴生成
- `render`：合成图片、音频和字幕为最终视频
- `schemas`：定义中间产物与清单结构

## Roadmap

### 第一阶段

- 搭建 CLI 项目骨架
- 实现 `extract`
- 定义备注、脚本、清单的 JSON 结构
- 接入第一个 TTS Provider
- 渲染简单的讲解视频

### 第二阶段

- 改进口播脚本清洗
- 支持按页重生成
- 增强 Provider 配置与字幕时间对齐
- 改进 `9:16` 的画面处理

### 第三阶段

- 封装为可复用 skill（OpenClaw / Agent Skills 形态见 `skills/note2video/`，安装说明见其中 `README.md`）
- 支持批量处理
- 支持 `pptx` 之外的更多输入源

## 当前待确认问题

- Windows 下默认应采用哪条渲染路径以保证页面保真？
- 第一个公开演示版默认选哪个 TTS Provider？
- 字幕时间轴应优先基于文本估算，还是基于音频对齐？

## 配置

配置样例见 `config.example.yaml`。

## 当前状态

当前仓库已完成：

- CLI 基础骨架
- Windows PowerPoint COM 导图；Linux / macOS 可选 LibreOffice + Poppler 真实导图
- 备注提取（OpenXML）
- 备注页误提取过滤
- 按页与汇总文本导出
- `voice` 第一版骨架与本地 TTS 接口
- `voice` 已支持 `edge` 与 `pyttsx3` 两种 provider
- `subtitle` 第一版生成与 `srt/json` 输出
- `render` 第一版视频合成

当前已经具备从 `pptx` 到 `mp4` 的最小闭环，后续重点将放在效果质量和可控性增强。
