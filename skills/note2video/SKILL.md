---
name: note2video
description: 通过 note2video 命令行将 PowerPoint（.pptx）备注转为逐页图、脚本、配音、字幕与 MP4；需在本机一次性 pip 安装。
metadata: {"openclaw":{"emoji":"🎬","homepage":"https://github.com/openclawee/note2video","os":["win32","linux","darwin"],"requires":{"anyBins":["python","python3"]}}}
---

# Note2Video（备注成片）

指导智能体调用 **note2video**：把 PPTX 转成素材目录或完整 MP4 的命令行工具。

## 重要说明：OpenClaw 不会自动安装本包

**Skill 里不包含 Python 源码包。** OpenClaw 通常**不会**替你执行 `pip install`（沙箱、审批或未接入安装流程）。请把安装视为 **用户在本机终端完成**，或 **用户明确同意且环境允许时** 再通过 `exec` 执行安装。

**在运行任何 `extract` / `build` 之前**，智能体应当：

1. 若允许 `exec`，先做**探测**：  
   `python -m note2video --help`  
   或  
   `python3 -m note2video --help`
2. 若失败（如 `No module named note2video`、找不到 `python`、非零退出码）：**不要假定已安装。**  
   - 把下面 **用户安装（一次性）** 整段命令发给用户，请用户在**自己的终端**里执行（若沙箱禁止安装，勿在沙箱内硬装）。  
   - 仅当**用户明确要求**且环境允许对所选 venv **联网 + 写入**时，才用 `exec` 跑 `pip install`。

## 用户安装（一次性）——请原样复制执行

任选一种方式。需要 **Python 3.10+**。务必使用 **venv**，避免污染系统 Python。

### Windows（PowerShell）

```powershell
cd $env:USERPROFILE
python -m venv note2video-venv
.\note2video-venv\Scripts\Activate.ps1
python -m pip install -U pip
python -m pip install "git+https://github.com/openclawee/note2video.git"
python -m note2video --help
```

### Linux / macOS（bash）

```bash
cd ~
python3 -m venv note2video-venv
source note2video-venv/bin/activate
python -m pip install -U pip
python -m pip install "git+https://github.com/openclawee/note2video.git"
python -m note2video --help
```

安装完成后，每次运行都要用 **该 venv 里的 Python**，例如：

- Windows：`~\note2video-venv\Scripts\python.exe -m note2video ...`
- 类 Unix：`~/note2video-venv/bin/python -m note2video ...`

### 单仓库（Skill 在已克隆的 note2video 仓库里）

仓库根目录 = `skills` 文件夹的上一级（其中有 `pyproject.toml`）：

```bash
python -m venv .venv
# 激活 venv 后：
python -m pip install -U pip && python -m pip install -e .
python -m note2video --help
```

## 何时使用本 Skill

- 用户有 `.pptx`，需要 **extract**（导出逐页图 + 备注 + 脚本），或 **build**（一条龙生成 MP4）。

## 命令（安装完成后；Python 用法同上）

```bash
python -m note2video extract "/path/to/deck.pptx" --out "./dist/deck01"
python -m note2video build "/path/to/deck.pptx" --out "./dist/deck01"
```

需要机器可读输出时加上 `--json`。

## 平台与幻灯片图像

- **Windows**：已装 Office 时优先用 **PowerPoint COM** 导出真实 PNG。
- **Linux / macOS**：`PATH` 上同时有 **`soffice` 或 `libreoffice`** 以及 **`pdftoppm`** 时导出真实 PNG；否则使用 OpenXML **占位图**。  
  环境变量 `NOTE2VIDEO_USE_LIBREOFFICE=0` 可强制走占位图。

## Linux 系统依赖（仅「要真实幻灯片图」时需要）

```bash
sudo apt install -y libreoffice-nogui poppler-utils
```

## Docker（带 LibreOffice 与 Poppler 的 Linux 镜像）

```bash
docker build -t note2video <含 Dockerfile 的仓库路径>
docker run --rm -v "$PWD:/work" -w /work note2video build ./deck.pptx --out ./dist
```

## 故障排除

- **`No module named note2video`**：未安装或用了错误的解释器 —— 执行上文 **用户安装**，或显式指定 venv 里的 `python`。
- **PowerPoint 相关错误（Windows）**：先在 PowerPoint 里打开该文件；需要带 Office 的图形界面 Windows。
- **`LibreOffice … failed`**：安装 Impress 与 Poppler；必要时设置 `NOTE2VIDEO_LIBREOFFICE` 指向 `soffice` 可执行文件。

## 安全

路径请加引号；勿将不可信用户输入直接拼进 shell 命令。
