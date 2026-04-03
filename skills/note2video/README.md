# OpenClaw Skill：note2video

本目录是符合 **[Agent Skills](https://agentskills.io)** 规范的 **[OpenClaw](https://github.com/openclaw/openclaw)** 用 Skill，说明如何安装并运行 **note2video** 命令行（PPTX → 素材 / MP4）。

## 装入 OpenClaw 工作区

1. 将整个目录复制到工作区的 skills 根下，例如：  
   `<workspace>/skills/note2video/`  
   使得 `SKILL.md` 的路径为 `<workspace>/skills/note2video/SKILL.md`。

2. 新建会话或重启 Gateway，以便重新加载 Skill（如 `/new` 或 `openclaw gateway restart`）。

3. 验证：执行 `openclaw skills list`，应能看到 `note2video`。

**OpenClaw 不会自动安装 Python 包。** Skill 加载后，请在本机终端按 `SKILL.md` 中的 **用户安装** 执行一次（在 venv 里 `pip install git+https://github.com/openclawee/note2video.git`）；仅当你的环境允许带网络的 `exec` 且你明确授权时，才可让智能体代执行安装。

若该 Skill 已发布到 **ClawHub**，也可使用 `openclaw skills install …`（以当时平台说明为准）。

## 与本仓库的关系

克隆本仓库后，Skill 路径为仓库根下的 `skills/note2video/`。`SKILL.md` 正文中的 `{baseDir}` 由 OpenClaw 解析为 Skill 所在目录，便于智能体定位相对路径。

## 关于 ClawHub 许可

在 ClawHub 上发布的 Skill 按注册表策略为 **MIT-0**；请勿在 `SKILL.md` 中加入与之冲突的许可条款。
