# CLI 设计草案

## 命令格式

```bash
note2video <command> [options]
```

## 设计原则

- 普通用户可以用一条命令跑完整流程
- 开发阶段可以逐步执行，便于调试
- 输出路径和结构尽量稳定，便于 skill 封装
- JSON 输出优先面向自动化调用
- 额外导出 `.txt` 文本，方便人工复制与编辑

## 主命令

### `build`

从 `pptx` 一次性执行到最终视频的完整流程。

当前实现会按顺序执行：

- `extract`
- `voice`
- `subtitle`
- `render`

```bash
note2video build input.pptx --out ./dist
```

典型输出：

- `slides/`
- `notes/notes.json`
- `notes/raw/*.txt`
- `notes/speaker/*.txt`
- `scripts/script.json`
- `scripts/all.txt`
- `scripts/txt/*.txt`
- `audio/*.wav`
- `subtitles/subtitles.srt`
- `video/output.mp4`
- `manifest.json`

## 子命令

### `extract`

从 PowerPoint 文件中提取页面图片、备注和脚本。

```bash
note2video extract input.pptx --out ./work
```

输出：

- `slides/*.png`
- `notes/notes.json`
- `notes/all.txt`
- `notes/raw/*.txt`
- `notes/speaker/*.txt`
- `scripts/script.json`
- `scripts/all.txt`
- `scripts/txt/*.txt`
- `manifest.json`

### `voice`

根据脚本逐页生成配音音频。

当前支持：

- `edge`：更适合正式成片，推荐优先使用
- `pyttsx3`：本地兜底方案，适合离线打通流程

```bash
note2video voice ./work/scripts/script.json --out ./work
note2video voice ./work/scripts/script.json --out ./work --tts-provider edge --voice zh-CN-XiaoxiaoNeural
```

输出：

- `audio/001.wav`
- `audio/002.wav`
- `audio/merged.wav`
- `audio/timings.json`
- 更新 `manifest.json` 中的音频路径与时长信息

### `voices`

列出可用音色，便于先挑选 `voice id` 再执行 `voice` 或 `build`。

```bash
note2video voices --tts-provider edge --keyword zh-CN
note2video voices --tts-provider edge --keyword Xiaoxiao --json
```

### `subtitle`

根据脚本和时间信息生成字幕文件。

当前实现采用句级切分，并按每页音频时长分配字幕时间。
如果存在 `audio/timings.json`，则优先使用逐句配音得到的真实时间轴。

```bash
note2video subtitle ./work/scripts/script.json --audio ./work/audio --out ./work/subtitles
```

输出：

- `subtitles/subtitles.srt`
- `subtitles/subtitles.json`

### `render`

根据页面图片、音频和字幕渲染最终视频。

当前实现使用 `imageio-ffmpeg` 提供的 ffmpeg 可执行文件路径，不要求系统预先安装全局 `ffmpeg`。
渲染成功后会自动清理 `video_only.mp4` 和 `slides.ffconcat` 这类中间文件。

```bash
note2video render ./work --out ./dist/output.mp4
```

输出：

- `video/output.mp4`

## 通用参数

- `--out <dir>`：输出目录
- `--temp <dir>`：临时工作目录
- `--config <file>`：配置文件路径
- `--overwrite`：覆盖已有输出
- `--verbose`：详细日志
- `--quiet`：最少日志
- `--json`：输出机器可读的 JSON 摘要

## 输入相关参数

- `--notes-source <auto|speaker-notes|file>`
- `--script-file <file>`
- `--pages <range>`

示例：

```bash
note2video build input.pptx --pages 1-5
note2video build input.pptx --script-file scripts.json
```

## 配音相关参数

- `--tts-provider <name>`
- `--voice <id>`
- `--rate <value>`
- `--pitch <value>`
- `--style <value>`
- `--voice-config <file>`

示例：

```bash
note2video build input.pptx --tts-provider edge --voice zh-CN-XiaoxiaoNeural
note2video build input.pptx --voice narrator_female --rate 1.05
```

## 脚本相关参数

- `--script-mode <raw|clean|spoken>`
- `--max-sentence-length <n>`
- `--pause-ms <n>`
- `--normalize-numbers`
- `--remove-brackets`

示例：

```bash
note2video build input.pptx --script-mode spoken --normalize-numbers
```

## 视频相关参数

- `--ratio <16:9|9:16|1:1>`
- `--resolution <720p|1080p|custom>`
- `--fps <n>`
- `--transition <none|fade>`
- `--slide-padding-ms <n>`

示例：

```bash
note2video build input.pptx --ratio 9:16 --resolution 1080p
```

## 字幕相关参数

- `--subtitle <on|off|soft|burn>`
- `--subtitle-style <name>`
- `--subtitle-max-chars <n>`

示例：

```bash
note2video build input.pptx --subtitle burn
note2video subtitle ./work/notes/script.json --subtitle-max-chars 18
```

## JSON 输出约定

当启用 `--json` 时，命令应向标准输出打印一个摘要对象。

示例：

```json
{
  "command": "build",
  "status": "ok",
  "input": "input.pptx",
  "output_dir": "./dist",
  "slide_count": 12,
  "artifacts": {
    "manifest": "./dist/manifest.json",
    "notes": "./dist/notes/notes.json",
    "script": "./dist/scripts/script.json",
    "subtitle": "./dist/subtitles/subtitles.srt",
    "video": "./dist/video/output.mp4"
  }
}
```

## 退出码

- `0`：成功
- `1`：一般运行时错误
- `2`：CLI 参数错误
- `3`：输入文件不存在或格式不支持
- `4`：PowerPoint 解析或导出失败
- `5`：TTS 生成失败
- `6`：字幕生成失败
- `7`：视频渲染失败

## 对 skill 封装的建议

推荐暴露的高层动作：

- `ppt-to-video`
- `ppt-to-assets`

封装时建议优先：

- 显式指定输出目录
- 使用 `--json` 获取机器可读结果
- 保持产物路径稳定，方便二次调用
