from __future__ import annotations

import os
import sys
import tempfile
import traceback
import faulthandler
from dataclasses import dataclass
from functools import partial
from pathlib import Path
from typing import Any

PREVIEW_SAMPLE_TEXT = "你好，这是一段音色试听。"

# locale_combo itemData: 大陆普通话（含 Edge zh-CN 与 MiniMax Chinese (Mandarin)…）
LOCALE_FILTER_MAINLAND = "__MAINLAND_ZH__"
LOCALE_FILTER_ALL = "__ALL__"


def _voice_locale_key(item: dict[str, Any]) -> str:
    """Group voices by locale / region for UI filtering."""
    loc = str(item.get("locale", "") or "").strip()
    if loc:
        return loc
    name = str(item.get("name", "") or "").strip()
    if "_" in name:
        return name.split("_", 1)[0].strip()
    return name or "(未标注)"


def _is_mainland_mandarin_locale_key(key: str) -> bool:
    if key == "zh-CN" or key.startswith("zh-CN-"):
        return True
    if key == "Chinese (Mandarin)" or key.startswith("Chinese (Mandarin)"):
        return True
    return False


def _voice_matches_locale_filter(filter_data: Any, item: dict[str, Any]) -> bool:
    if filter_data is None:
        return True
    if filter_data == LOCALE_FILTER_ALL:
        return True
    if filter_data == LOCALE_FILTER_MAINLAND:
        return _is_mainland_mandarin_locale_key(_voice_locale_key(item))
    return _voice_locale_key(item) == filter_data


def _locale_key_label_zh(key: str) -> str:
    if key == "(未标注)":
        return "未标注地区 / 语言"
    table = {
        "zh-CN": "中国大陆 · 普通话（zh-CN）",
        "zh-TW": "台湾 · 繁体中文（zh-TW）",
        "zh-HK": "香港 · 粤语（zh-HK）",
        "yue-CN": "粤语（中国大陆）",
        "en-US": "美国 · 英语",
        "en-GB": "英国 · 英语",
        "ja-JP": "日本 · 日语",
        "ko-KR": "韩国 · 韩语",
        "fr-FR": "法国 · 法语",
        "de-DE": "德国 · 德语",
        "es-ES": "西班牙 · 西班牙语",
        "pt-BR": "巴西 · 葡萄牙语",
        "ru-RU": "俄罗斯 · 俄语",
        "it-IT": "意大利 · 意大利语",
        "hi-IN": "印度 · 印地语",
        "th-TH": "泰国 · 泰语",
        "vi-VN": "越南 · 越南语",
        "id-ID": "印尼 · 印尼语",
    }
    if key in table:
        return table[key]
    if key.startswith("Chinese "):
        return f"中国大陆 · 普通话（{key}）"
    if "-" in key and len(key) <= 16:
        return f"其他 · {key}"
    return f"其他 · {key}"


def _run_tts_settings_dialog(parent, *, QtWidgets, append_log) -> None:
    from note2video.user_config import (
        default_tts_provider,
        load_user_config,
        minimax_host_ui_index_from_provider_cfg,
        normalize_user_config,
        save_user_config,
        tts_provider_config,
        user_config_path,
    )

    merged = normalize_user_config(load_user_config())
    dlg = QtWidgets.QDialog(parent)
    dlg.setWindowTitle("TTS Provider 设置")
    dlg.setMinimumWidth(520)

    layout = QtWidgets.QVBoxLayout(dlg)
    path_label = QtWidgets.QLabel(f"配置文件：{user_config_path()}")
    path_label.setWordWrap(True)
    path_label.setStyleSheet("color: gray;")
    layout.addWidget(path_label)

    form = QtWidgets.QFormLayout()

    provider_combo = QtWidgets.QComboBox()
    provider_combo.addItem("edge（本地/微软）", "edge")
    provider_combo.addItem("minimax_cn（在线 / 国内）", "minimax_cn")
    provider_combo.addItem("minimax_global（在线 / 国际）", "minimax_global")
    form.addRow("配置 Provider", provider_combo)

    default_combo = QtWidgets.QComboBox()
    default_combo.addItem("（不设置默认，仍按界面/命令行选择）", "")
    default_combo.addItem("edge", "edge")
    default_combo.addItem("minimax_cn", "minimax_cn")
    default_combo.addItem("minimax_global", "minimax_global")
    form.addRow("默认 Provider", default_combo)

    current_default = default_tts_provider(merged) or ""
    idx = default_combo.findData(current_default)
    default_combo.setCurrentIndex(idx if idx >= 0 else 0)

    # Provider-specific container
    stack = QtWidgets.QStackedWidget()
    layout.addLayout(form)
    layout.addWidget(stack)

    # Edge page (currently no persistent settings)
    edge_page = QtWidgets.QWidget()
    edge_layout = QtWidgets.QVBoxLayout(edge_page)
    edge_layout.addWidget(QtWidgets.QLabel("edge 暂无需要写入配置文件的参数。"))
    edge_layout.addStretch(1)
    stack.addWidget(edge_page)

    # MiniMax CN page
    minimax_page = QtWidgets.QWidget()
    mm_form = QtWidgets.QFormLayout(minimax_page)

    key_edit = QtWidgets.QLineEdit()
    key_edit.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)
    key_edit.setPlaceholderText("留空保留已保存的密钥；填写则覆盖")
    mm_form.addRow("API Key", key_edit)
    clear_key_chk = QtWidgets.QCheckBox("删除已保存的 API Key")
    mm_form.addRow("", clear_key_chk)

    mm_form.addRow("API 主机", QtWidgets.QLabel("固定：https://api.minimax.chat"))

    mm_cfg = tts_provider_config(merged, "minimax_cn")
    model_edit = QtWidgets.QLineEdit(str(mm_cfg.get("model") or "").strip() or "speech-2.8-hd")
    mm_form.addRow("T2A 模型", model_edit)

    timeout_spin = QtWidgets.QDoubleSpinBox()
    timeout_spin.setMinimum(0)
    timeout_spin.setMaximum(600)
    timeout_spin.setDecimals(0)
    timeout_spin.setSpecialValueText("默认（合成 60s / 列音色 30s）")
    raw_to = mm_cfg.get("timeout_s")
    if raw_to is not None and str(raw_to).strip() != "":
        timeout_spin.setValue(float(raw_to))
    else:
        timeout_spin.setValue(0)
    mm_form.addRow("请求超时（秒，0=内置默认）", timeout_spin)

    stack.addWidget(minimax_page)

    # MiniMax Global page
    minimax_global_page = QtWidgets.QWidget()
    mg_form = QtWidgets.QFormLayout(minimax_global_page)
    mg_form.addRow("API 主机", QtWidgets.QLabel("固定：https://api.minimaxi.chat"))
    g_key_edit = QtWidgets.QLineEdit()
    g_key_edit.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)
    g_key_edit.setPlaceholderText("留空保留已保存的密钥；填写则覆盖")
    mg_form.addRow("API Key", g_key_edit)
    g_clear_key_chk = QtWidgets.QCheckBox("删除已保存的 API Key")
    mg_form.addRow("", g_clear_key_chk)

    mg_cfg = tts_provider_config(merged, "minimax_global")
    g_model_edit = QtWidgets.QLineEdit(str(mg_cfg.get("model") or "").strip() or "speech-2.8-hd")
    mg_form.addRow("T2A 模型", g_model_edit)

    g_timeout_spin = QtWidgets.QDoubleSpinBox()
    g_timeout_spin.setMinimum(0)
    g_timeout_spin.setMaximum(600)
    g_timeout_spin.setDecimals(0)
    g_timeout_spin.setSpecialValueText("默认（合成 60s / 列音色 30s）")
    g_raw_to = mg_cfg.get("timeout_s")
    if g_raw_to is not None and str(g_raw_to).strip() != "":
        g_timeout_spin.setValue(float(g_raw_to))
    else:
        g_timeout_spin.setValue(0)
    mg_form.addRow("请求超时（秒，0=内置默认）", g_timeout_spin)

    stack.addWidget(minimax_global_page)

    def _sync_stack() -> None:
        data = provider_combo.currentData()
        if data == "edge":
            stack.setCurrentIndex(0)
        elif data == "minimax_cn":
            stack.setCurrentIndex(1)
        else:
            stack.setCurrentIndex(2)

    provider_combo.currentIndexChanged.connect(lambda _i: _sync_stack())
    provider_combo.setCurrentIndex(0)
    _sync_stack()

    def _apply() -> None:
        # default provider
        tts = dict(merged.get("tts") or {})
        if default_combo.currentData():
            tts["default_provider"] = str(default_combo.currentData())
        else:
            tts.pop("default_provider", None)
        merged["tts"] = tts

        # provider-specific updates (MiniMax only for now)
        providers = dict((tts.get("providers") or {}) if isinstance(tts.get("providers"), dict) else {})
        mm_cn = dict((providers.get("minimax_cn") or {}) if isinstance(providers.get("minimax_cn"), dict) else {})
        mm_gl = dict((providers.get("minimax_global") or {}) if isinstance(providers.get("minimax_global"), dict) else {})

        if clear_key_chk.isChecked():
            mm_cn["api_key"] = ""
        elif key_edit.text().strip():
            mm_cn["api_key"] = key_edit.text().strip()

        mm_cn_model = model_edit.text().strip()
        if mm_cn_model:
            mm_cn["model"] = mm_cn_model
        else:
            mm_cn.pop("model", None)
        if timeout_spin.value() > 0:
            mm_cn["timeout_s"] = int(timeout_spin.value())
        else:
            mm_cn.pop("timeout_s", None)

        if g_clear_key_chk.isChecked():
            mm_gl["api_key"] = ""
        elif g_key_edit.text().strip():
            mm_gl["api_key"] = g_key_edit.text().strip()

        mm_gl_model = g_model_edit.text().strip()
        if mm_gl_model:
            mm_gl["model"] = mm_gl_model
        else:
            mm_gl.pop("model", None)
        if g_timeout_spin.value() > 0:
            mm_gl["timeout_s"] = int(g_timeout_spin.value())
        else:
            mm_gl.pop("timeout_s", None)

        providers["minimax_cn"] = mm_cn
        providers["minimax_global"] = mm_gl
        tts["providers"] = providers
        merged["tts"] = tts

        save_user_config(merged)
        append_log(f"已保存用户配置：{user_config_path()}")
        dlg.accept()

    buttons = QtWidgets.QDialogButtonBox(
        QtWidgets.QDialogButtonBox.StandardButton.Ok | QtWidgets.QDialogButtonBox.StandardButton.Cancel
    )
    buttons.accepted.connect(_apply)
    buttons.rejected.connect(dlg.reject)
    layout.addWidget(buttons)

    dlg.exec()


def _require_pyside6():
    try:
        from PySide6 import QtCore, QtWidgets  # type: ignore
    except Exception as exc:  # pragma: no cover
        raise RuntimeError(
            "未安装 GUI 依赖。请先执行：\n"
            "  python -m pip install -e .[gui]\n"
            "或：\n"
            "  python -m pip install .[gui]\n"
        ) from exc
    return QtCore, QtWidgets


@dataclass(frozen=True)
class JobConfig:
    mode: str  # "extract" | "build"
    pptx_path: Path
    out_dir: Path
    pages: str
    tts_provider: str
    voice_id: str
    tts_rate: float
    minimax_base_url: str | None = None
    subtitle_color: str | None = None
    subtitle_highlight_mode: str | None = None
    subtitle_highlight_color: str | None = None
    subtitle_fade_in_ms: int | None = None
    subtitle_fade_out_ms: int | None = None
    subtitle_scale_from: int | None = None
    subtitle_scale_to: int | None = None
    subtitle_outline: int | None = None
    subtitle_shadow: int | None = None
    subtitle_font: str | None = None
    subtitle_size: int | None = None
    bgm_path: str | None = None
    bgm_volume: float = 0.18
    narration_volume: float = 1.0
    bgm_fade_in_s: float = 0.0
    bgm_fade_out_s: float = 0.0


def main(argv: list[str] | None = None) -> int:
    QtCore, QtWidgets = _require_pyside6()
    faulthandler.enable()

    # Only pass the program path to Qt. Full sys.argv (e.g. `python -m ...`) can confuse Qt's
    # argument parser and has been observed to break the event loop on some Windows setups.
    if argv is None:
        qt_argv = [sys.argv[0]] if sys.argv else ["note2video-gui"]
    else:
        qt_argv = list(argv) if argv else ["note2video-gui"]

    app = QtWidgets.QApplication(qt_argv)
    app.setApplicationName("Note2Video")

    window = MainWindow(QtCore=QtCore, QtWidgets=QtWidgets)
    window.resize(860, 560)
    window.show()

    inner = window._window
    inner.setMinimumSize(520, 380)

    def _position_main_window() -> None:
        if not inner.isVisible():
            inner.show()
        screen = QtWidgets.QApplication.primaryScreen()
        if screen is not None:
            avail = screen.availableGeometry()
            frame = inner.frameGeometry()
            if frame.width() <= 0 or frame.height() <= 0:
                inner.adjustSize()
                frame = inner.frameGeometry()
            frame.moveCenter(avail.center())
            inner.move(frame.topLeft())
        inner.raise_()
        inner.activateWindow()

    QtCore.QTimer.singleShot(0, _position_main_window)

    return app.exec()


def _run_extract_or_build(config: JobConfig, emit_log) -> int:
    from note2video.parser.extract import extract_project
    from note2video.render.video import render_video
    from note2video.subtitle.generate import generate_subtitles
    from note2video.tts.voice import generate_voice_assets

    emit_log(f"mode: {config.mode}")
    emit_log(f"pptx: {config.pptx_path}")
    emit_log(f"out:  {config.out_dir}")
    emit_log(f"pages: {config.pages}")
    emit_log(f"tts_rate: {config.tts_rate}")

    emit_log("阶段：extract")
    manifest = extract_project(str(config.pptx_path), str(config.out_dir), pages=config.pages)
    try:
        emit_log(f"细节：slides={getattr(manifest, 'slide_count', '?')}")
    except Exception:
        pass

    if config.mode == "extract":
        emit_log("完成：extract")
        return 0

    emit_log("阶段：voice")
    script_path = config.out_dir / "scripts" / "script.json"
    voice_result = generate_voice_assets(
        str(script_path),
        str(config.out_dir),
        provider_name=(config.tts_provider or "pyttsx3"),
        voice_id=config.voice_id,
        tts_rate=config.tts_rate,
        minimax_base_url=config.minimax_base_url,
    )
    try:
        emit_log(
            "细节："
            + f"provider={voice_result.get('provider')}, "
            + f"voice={voice_result.get('voice')}, "
            + f"tts_rate={voice_result.get('tts_rate')}"
        )
    except Exception:
        pass

    emit_log("阶段：subtitle")
    sub_result = generate_subtitles(str(script_path), str(config.out_dir))
    try:
        emit_log(f"细节：segments={sub_result.get('segment_count')}, slides={sub_result.get('slide_count')}")
    except Exception:
        pass

    emit_log("阶段：render")
    result = render_video(
        str(config.out_dir),
        subtitle_color=config.subtitle_color,
        subtitle_highlight_mode=config.subtitle_highlight_mode,
        subtitle_highlight_color=config.subtitle_highlight_color,
        subtitle_fade_in_ms=config.subtitle_fade_in_ms,
        subtitle_fade_out_ms=config.subtitle_fade_out_ms,
        subtitle_scale_from=config.subtitle_scale_from,
        subtitle_scale_to=config.subtitle_scale_to,
        subtitle_outline=config.subtitle_outline,
        subtitle_shadow=config.subtitle_shadow,
        subtitle_font=config.subtitle_font,
        subtitle_size=config.subtitle_size,
        bgm_path=config.bgm_path,
        bgm_volume=float(config.bgm_volume),
        narration_volume=float(config.narration_volume),
        bgm_fade_in_s=float(config.bgm_fade_in_s),
        bgm_fade_out_s=float(config.bgm_fade_out_s),
    )
    try:
        emit_log(f"细节：subtitles_burned={result.get('subtitles_burned')}, mixed_audio={bool(result.get('mixed_audio'))}")
    except Exception:
        pass
    emit_log(f"输出视频：{result.get('video')}")

    emit_log("完成：build")
    return 0


def _build_worker(QtCore, config: JobConfig):
    class Worker(QtCore.QObject):
        log = QtCore.Signal(str)
        done = QtCore.Signal(int)

        def run(self) -> None:
            try:
                exit_code = _run_extract_or_build(config, self.log.emit)
            except Exception:
                self.log.emit(traceback.format_exc())
                exit_code = 1
            self.done.emit(exit_code)

    return Worker()


def _run_pipeline_with_log(config: JobConfig, emit_log) -> int:
    try:
        return _run_extract_or_build(config, emit_log)
    except Exception:
        emit_log(traceback.format_exc())
        return 1


def _build_ui(QtWidgets):
    class MainWindow(QtWidgets.QMainWindow):
        def __init__(self) -> None:
            super().__init__()
            self.setWindowTitle("Note2Video / 备注成片")

            menu_bar = QtWidgets.QMenuBar(self)
            settings_menu = menu_bar.addMenu("设置")
            settings_menu.addAction("TTS Provider…").triggered.connect(self._open_tts_settings)
            self.setMenuBar(menu_bar)

            self._all_voice_items: list[dict[str, Any]] = []
            self._preview_player = None
            self._preview_audio_out = None
            self._preview_thread: Any = None
            self._preview_worker: Any = None
            self._preview_temp_path: Path | None = None
            self._preview_proc: Any = None
            self._preview_timer: Any = None
            self._preview_log_path: Path | None = None
            self._preview_log_fh: Any = None
            self._last_preview_path: Path | None = None
            self._preview_web_view: Any = None
            self._preview_web_available: bool = False
            self._pipeline_busy = False
            self._pipeline_stage_label: str = ""

            central = QtWidgets.QWidget(self)
            self.setCentralWidget(central)

            layout = QtWidgets.QVBoxLayout(central)
            top = QtWidgets.QHBoxLayout()
            layout.addLayout(top)

            left = QtWidgets.QVBoxLayout()
            right = QtWidgets.QVBoxLayout()
            top.addLayout(left, 3)
            top.addLayout(right, 2)

            # --- Project / IO ---
            io_group = QtWidgets.QGroupBox("项目")
            io_grid = QtWidgets.QGridLayout(io_group)
            io_grid.setColumnStretch(1, 1)
            io_grid.setColumnStretch(3, 1)
            left.addWidget(io_group)

            self.pptx_edit = QtWidgets.QLineEdit()
            self.pptx_btn = QtWidgets.QPushButton("选择…")
            pptx_row = QtWidgets.QHBoxLayout()
            pptx_row.addWidget(self.pptx_edit, 1)
            pptx_row.addWidget(self.pptx_btn)
            io_grid.addWidget(QtWidgets.QLabel("PPTX"), 0, 0)
            io_grid.addLayout(pptx_row, 0, 1, 1, 3)

            self.out_edit = QtWidgets.QLineEdit(str(Path("dist").resolve()))
            self.out_btn = QtWidgets.QPushButton("选择…")
            out_row = QtWidgets.QHBoxLayout()
            out_row.addWidget(self.out_edit, 1)
            out_row.addWidget(self.out_btn)
            io_grid.addWidget(QtWidgets.QLabel("输出"), 1, 0)
            io_grid.addLayout(out_row, 1, 1, 1, 3)

            self.pages_edit = QtWidgets.QLineEdit("all")
            self.pages_edit.setToolTip("页码范围：all 或 1-3,5")
            io_grid.addWidget(QtWidgets.QLabel("页码"), 2, 0)
            io_grid.addWidget(self.pages_edit, 2, 1)

            # Spacer to keep grid compact
            io_grid.addWidget(QtWidgets.QLabel(""), 2, 2)
            io_grid.addWidget(QtWidgets.QLabel(""), 2, 3)

            # --- TTS ---
            tts_group = QtWidgets.QGroupBox("配音（TTS）")
            tts_grid = QtWidgets.QGridLayout(tts_group)
            tts_grid.setColumnStretch(1, 1)
            tts_grid.setColumnStretch(3, 1)
            left.addWidget(tts_group)

            self.tts_combo = QtWidgets.QComboBox()
            self.tts_combo.addItems(["pyttsx3", "edge", "minimax_cn", "minimax_global"])
            tts_grid.addWidget(QtWidgets.QLabel("Provider"), 0, 0)
            tts_grid.addWidget(self.tts_combo, 0, 1)

            self.locale_combo = QtWidgets.QComboBox()
            self.locale_combo.setMinimumWidth(320)
            self.locale_combo.addItem("（请先刷新音色列表）", None)
            self.locale_combo.setToolTip("刷新音色后，默认「中国大陆 · 普通话」；也可选其他地区或「全部」。")
            tts_grid.addWidget(QtWidgets.QLabel("地区"), 0, 2)
            tts_grid.addWidget(self.locale_combo, 0, 3)

            voice_row = QtWidgets.QHBoxLayout()
            self.voice_combo = QtWidgets.QComboBox()
            self.voice_combo.setEditable(True)
            self.voice_combo.setMinimumWidth(240)
            self.voice_combo.setMaxVisibleItems(18)
            self._voice_combo_reset_default()
            self.voice_refresh_btn = QtWidgets.QPushButton("刷新音色")
            self.voice_refresh_btn.setToolTip("从当前 Provider 拉取可用音色（edge/minimax 需网络；minimax 需 API Key）。")
            self.voice_preview_btn = QtWidgets.QPushButton("试听")
            self.voice_preview_btn.setToolTip(f"用当前 Provider、音色与语速合成一句试听：{PREVIEW_SAMPLE_TEXT}")
            voice_row.addWidget(self.voice_combo, 1)
            voice_row.addWidget(self.voice_refresh_btn)
            voice_row.addWidget(self.voice_preview_btn)
            tts_grid.addWidget(QtWidgets.QLabel("声音"), 1, 0)
            tts_grid.addLayout(voice_row, 1, 1, 1, 3)

            self.tts_rate_spin = QtWidgets.QDoubleSpinBox()
            self.tts_rate_spin.setRange(0.5, 2.0)
            self.tts_rate_spin.setSingleStep(0.1)
            self.tts_rate_spin.setDecimals(2)
            self.tts_rate_spin.setValue(1.0)
            self.tts_rate_spin.setToolTip("1.0 为正常语速；在合成阶段生效，字幕与配音时长一致。")
            tts_grid.addWidget(QtWidgets.QLabel("语速"), 2, 0)
            tts_grid.addWidget(self.tts_rate_spin, 2, 1)

            # --- Mix ---
            mix_group = QtWidgets.QGroupBox("混音（BGM）")
            mix_grid = QtWidgets.QGridLayout(mix_group)
            mix_grid.setColumnStretch(1, 1)
            mix_grid.setColumnStretch(3, 1)
            left.addWidget(mix_group)

            # BGM mixer controls
            self.bgm_path_edit = QtWidgets.QLineEdit()
            self.bgm_path_btn = QtWidgets.QPushButton("选择…")
            bgm_row = QtWidgets.QHBoxLayout()
            bgm_row.addWidget(self.bgm_path_edit, 1)
            bgm_row.addWidget(self.bgm_path_btn)
            mix_grid.addWidget(QtWidgets.QLabel("BGM"), 0, 0)
            mix_grid.addLayout(bgm_row, 0, 1, 1, 3)

            self.bgm_volume_spin = QtWidgets.QDoubleSpinBox()
            self.bgm_volume_spin.setRange(0.0, 2.0)
            self.bgm_volume_spin.setSingleStep(0.05)
            self.bgm_volume_spin.setDecimals(2)
            self.bgm_volume_spin.setValue(0.18)
            self.bgm_volume_spin.setToolTip("背景音乐音量，默认 0.18。")
            mix_grid.addWidget(QtWidgets.QLabel("BGM 音量"), 1, 0)
            mix_grid.addWidget(self.bgm_volume_spin, 1, 1)

            self.narration_volume_spin = QtWidgets.QDoubleSpinBox()
            self.narration_volume_spin.setRange(0.0, 3.0)
            self.narration_volume_spin.setSingleStep(0.05)
            self.narration_volume_spin.setDecimals(2)
            self.narration_volume_spin.setValue(1.0)
            self.narration_volume_spin.setToolTip("旁白音量，默认 1.0。")
            mix_grid.addWidget(QtWidgets.QLabel("旁白音量"), 1, 2)
            mix_grid.addWidget(self.narration_volume_spin, 1, 3)

            self.bgm_fade_in_spin = QtWidgets.QDoubleSpinBox()
            self.bgm_fade_in_spin.setRange(0.0, 30.0)
            self.bgm_fade_in_spin.setSingleStep(0.5)
            self.bgm_fade_in_spin.setDecimals(1)
            self.bgm_fade_in_spin.setValue(0.0)
            self.bgm_fade_in_spin.setToolTip("背景音乐淡入时长（秒）。")
            mix_grid.addWidget(QtWidgets.QLabel("淡入(s)"), 2, 0)
            mix_grid.addWidget(self.bgm_fade_in_spin, 2, 1)

            self.bgm_fade_out_spin = QtWidgets.QDoubleSpinBox()
            self.bgm_fade_out_spin.setRange(0.0, 30.0)
            self.bgm_fade_out_spin.setSingleStep(0.5)
            self.bgm_fade_out_spin.setDecimals(1)
            self.bgm_fade_out_spin.setValue(0.0)
            self.bgm_fade_out_spin.setToolTip("背景音乐淡出时长（秒）。")
            mix_grid.addWidget(QtWidgets.QLabel("淡出(s)"), 2, 2)
            mix_grid.addWidget(self.bgm_fade_out_spin, 2, 3)

            # --- Subtitles ---
            subtitle_group = QtWidgets.QGroupBox("字幕（样式 / 特效）")
            render_grid = QtWidgets.QGridLayout(subtitle_group)
            render_grid.setColumnStretch(1, 1)
            render_grid.setColumnStretch(3, 1)
            left.addWidget(subtitle_group)

            # Subtitle color picker
            self.subtitle_color_value = QtWidgets.QLineEdit()
            self.subtitle_color_value.setReadOnly(True)
            self.subtitle_color_value.setPlaceholderText("默认")
            self.subtitle_color_btn = QtWidgets.QPushButton("选择…")
            self.subtitle_color_clear_btn = QtWidgets.QPushButton("清除")
            self.subtitle_color_preview = QtWidgets.QLabel("      ")
            self.subtitle_color_preview.setToolTip("字幕颜色预览")
            self.subtitle_color_preview.setStyleSheet("background: transparent; border: 1px solid #999;")
            sub_row = QtWidgets.QHBoxLayout()
            sub_row.addWidget(self.subtitle_color_preview)
            sub_row.addWidget(self.subtitle_color_value, 1)
            sub_row.addWidget(self.subtitle_color_btn)
            sub_row.addWidget(self.subtitle_color_clear_btn)
            render_grid.addWidget(QtWidgets.QLabel("字幕颜色"), 0, 0)
            render_grid.addLayout(sub_row, 0, 1, 1, 3)

            # Subtitle font / size
            self.subtitle_font_edit = QtWidgets.QComboBox()
            self.subtitle_font_edit.setEditable(True)
            self.subtitle_font_edit.setToolTip("烧录字幕时使用的字体名（FontName）。示例：Microsoft YaHei")
            try:
                from PySide6.QtGui import QFontDatabase

                fonts = list(QFontDatabase.families())
            except Exception:
                fonts = []
            if fonts:
                self.subtitle_font_edit.addItem("（默认）", "")
                for f in fonts:
                    self.subtitle_font_edit.addItem(str(f), str(f))
                # Prefer a common CJK font on Windows if available.
                idx = self.subtitle_font_edit.findText("Microsoft YaHei")
                if idx >= 0:
                    self.subtitle_font_edit.setCurrentIndex(idx)
            render_grid.addWidget(QtWidgets.QLabel("字幕字体"), 1, 0)
            render_grid.addWidget(self.subtitle_font_edit, 1, 1, 1, 3)

            self.subtitle_size_spin = QtWidgets.QSpinBox()
            self.subtitle_size_spin.setRange(8, 200)
            self.subtitle_size_spin.setValue(48)
            self.subtitle_size_spin.setToolTip("烧录字幕时字体大小（FontSize）。默认 48。")
            render_grid.addWidget(QtWidgets.QLabel("字体大小"), 2, 0)
            render_grid.addWidget(self.subtitle_size_spin, 2, 1)

            # Subtitle effects / highlight
            self.subtitle_highlight_mode_combo = QtWidgets.QComboBox()
            self.subtitle_highlight_mode_combo.addItem("无", "none")
            self.subtitle_highlight_mode_combo.addItem("整句高亮", "line")
            self.subtitle_highlight_mode_combo.addItem("逐词高亮（Edge）", "word")
            self.subtitle_highlight_mode_combo.setToolTip("高亮模式：整句/逐词。逐词需要 edge-tts 并生成 word_timings。")
            render_grid.addWidget(QtWidgets.QLabel("高亮模式"), 3, 0)
            render_grid.addWidget(self.subtitle_highlight_mode_combo, 3, 1)

            self.subtitle_highlight_color_value = QtWidgets.QLineEdit()
            self.subtitle_highlight_color_value.setReadOnly(True)
            self.subtitle_highlight_color_value.setPlaceholderText("默认")
            self.subtitle_highlight_color_btn = QtWidgets.QPushButton("选择…")
            self.subtitle_highlight_color_clear_btn = QtWidgets.QPushButton("清除")
            self.subtitle_highlight_color_preview = QtWidgets.QLabel("      ")
            self.subtitle_highlight_color_preview.setToolTip("高亮颜色预览")
            self.subtitle_highlight_color_preview.setStyleSheet("background: transparent; border: 1px solid #999;")
            hi_row = QtWidgets.QHBoxLayout()
            hi_row.addWidget(self.subtitle_highlight_color_preview)
            hi_row.addWidget(self.subtitle_highlight_color_value, 1)
            hi_row.addWidget(self.subtitle_highlight_color_btn)
            hi_row.addWidget(self.subtitle_highlight_color_clear_btn)
            render_grid.addWidget(QtWidgets.QLabel("高亮颜色"), 4, 0)
            render_grid.addLayout(hi_row, 4, 1, 1, 3)

            self.subtitle_fade_in_spin = QtWidgets.QSpinBox()
            self.subtitle_fade_in_spin.setRange(0, 5000)
            self.subtitle_fade_in_spin.setValue(80)
            self.subtitle_fade_in_spin.setToolTip("逐句出现（淡入）时长，单位 ms。")
            self.subtitle_fade_out_spin = QtWidgets.QSpinBox()
            self.subtitle_fade_out_spin.setRange(0, 5000)
            self.subtitle_fade_out_spin.setValue(120)
            self.subtitle_fade_out_spin.setToolTip("逐句消失（淡出）时长，单位 ms。")
            render_grid.addWidget(QtWidgets.QLabel("淡入/淡出(ms)"), 5, 0)
            fade_row = QtWidgets.QHBoxLayout()
            fade_row.addWidget(self.subtitle_fade_in_spin)
            fade_row.addWidget(QtWidgets.QLabel("/"))
            fade_row.addWidget(self.subtitle_fade_out_spin)
            render_grid.addLayout(fade_row, 5, 1)

            self.subtitle_scale_from_spin = QtWidgets.QSpinBox()
            self.subtitle_scale_from_spin.setRange(50, 200)
            self.subtitle_scale_from_spin.setValue(100)
            self.subtitle_scale_to_spin = QtWidgets.QSpinBox()
            self.subtitle_scale_to_spin.setRange(50, 200)
            self.subtitle_scale_to_spin.setValue(104)
            render_grid.addWidget(QtWidgets.QLabel("缩放(%)"), 6, 0)
            scale_row = QtWidgets.QHBoxLayout()
            scale_row.addWidget(self.subtitle_scale_from_spin)
            scale_row.addWidget(QtWidgets.QLabel("→"))
            scale_row.addWidget(self.subtitle_scale_to_spin)
            render_grid.addLayout(scale_row, 6, 1)

            self.subtitle_outline_spin = QtWidgets.QSpinBox()
            self.subtitle_outline_spin.setRange(0, 20)
            self.subtitle_outline_spin.setValue(1)
            self.subtitle_shadow_spin = QtWidgets.QSpinBox()
            self.subtitle_shadow_spin.setRange(0, 20)
            self.subtitle_shadow_spin.setValue(0)
            render_grid.addWidget(QtWidgets.QLabel("描边/阴影"), 7, 0)
            os_row = QtWidgets.QHBoxLayout()
            os_row.addWidget(self.subtitle_outline_spin)
            os_row.addWidget(QtWidgets.QLabel("/"))
            os_row.addWidget(self.subtitle_shadow_spin)
            render_grid.addLayout(os_row, 7, 1)

            left.addStretch(1)

            # --- Buttons / Progress / Log on right ---
            buttons = QtWidgets.QHBoxLayout()
            right.addLayout(buttons)
            self.extract_btn = QtWidgets.QPushButton("一键 Extract")
            self.build_btn = QtWidgets.QPushButton("一键 Build")
            self.stop_btn = QtWidgets.QPushButton("停止（当前任务）")
            self.stop_btn.setEnabled(False)
            buttons.addWidget(self.extract_btn)
            buttons.addWidget(self.build_btn)
            buttons.addWidget(self.stop_btn)
            buttons.addStretch(1)

            self.progress = QtWidgets.QProgressBar()
            self.progress.setRange(0, 4)
            self.progress.setValue(0)
            self.progress.setTextVisible(True)
            self.progress.setFormat("就绪")
            right.addWidget(self.progress)

            # --- Preview player (embedded browser) ---
            preview_group = QtWidgets.QGroupBox("试听播放器（内嵌）")
            preview_v = QtWidgets.QVBoxLayout(preview_group)
            preview_hint = QtWidgets.QLabel(
                "若系统默认播放器不稳定，可用内嵌浏览器播放（需要 QtWebEngine）。"
            )
            preview_hint.setWordWrap(True)
            preview_hint.setStyleSheet("color: gray;")
            preview_v.addWidget(preview_hint)

            self.preview_open_btn = QtWidgets.QPushButton("在内嵌播放器中打开最近一次试听")
            self.preview_open_btn.setEnabled(False)
            preview_v.addWidget(self.preview_open_btn)

            preview_actions = QtWidgets.QHBoxLayout()
            self.preview_play_system_btn = QtWidgets.QPushButton("播放（系统默认）")
            self.preview_reveal_btn = QtWidgets.QPushButton("定位文件")
            self.preview_play_system_btn.setEnabled(False)
            self.preview_reveal_btn.setEnabled(False)
            preview_actions.addWidget(self.preview_play_system_btn)
            preview_actions.addWidget(self.preview_reveal_btn)
            preview_actions.addStretch(1)
            preview_v.addLayout(preview_actions)

            try:
                from PySide6.QtWebEngineWidgets import QWebEngineView  # type: ignore

                self._preview_web_view = QWebEngineView()
                self._preview_web_available = True
                self._preview_web_view.setMinimumHeight(120)
                preview_v.addWidget(self._preview_web_view)
            except Exception:
                self._preview_web_view = None
                self._preview_web_available = False
                missing = QtWidgets.QLabel(
                    "当前环境未安装/无法加载 QtWebEngine，内嵌播放不可用；仍可用系统播放器播放。"
                )
                missing.setWordWrap(True)
                missing.setStyleSheet("color: gray;")
                preview_v.addWidget(missing)

            right.addWidget(preview_group)

            self.log = QtWidgets.QPlainTextEdit()
            self.log.setReadOnly(True)
            right.addWidget(self.log, 1)

            self._thread = None
            self._worker = None

            self.pptx_btn.clicked.connect(self._pick_pptx)
            self.out_btn.clicked.connect(self._pick_out_dir)
            self.extract_btn.clicked.connect(lambda: self._start("extract"))
            self.build_btn.clicked.connect(lambda: self._start("build"))
            self.stop_btn.clicked.connect(self._stop)
            self.tts_combo.currentTextChanged.connect(self._on_tts_provider_changed)
            self.locale_combo.currentIndexChanged.connect(self._repopulate_voice_combo)
            self.voice_refresh_btn.clicked.connect(self._refresh_voice_list)
            self.voice_preview_btn.clicked.connect(self._preview_voice)
            self.preview_open_btn.clicked.connect(self._open_last_preview_in_web)
            self.preview_play_system_btn.clicked.connect(self._play_last_preview_system)
            self.preview_reveal_btn.clicked.connect(self._reveal_last_preview)
            self.subtitle_color_btn.clicked.connect(self._pick_subtitle_color)
            self.subtitle_color_clear_btn.clicked.connect(self._clear_subtitle_color)
            self.subtitle_highlight_color_btn.clicked.connect(self._pick_subtitle_highlight_color)
            self.subtitle_highlight_color_clear_btn.clicked.connect(self._clear_subtitle_highlight_color)
            self.bgm_path_btn.clicked.connect(self._pick_bgm)

            self._restore_gui_state_from_config()

        def closeEvent(self, event) -> None:  # noqa: N802
            try:
                self._persist_gui_state_to_config()
            except Exception:
                # Never block window close due to config write errors.
                self._append_log(traceback.format_exc())
            super().closeEvent(event)

        def _restore_gui_state_from_config(self) -> None:
            from note2video.user_config import gui_state, load_user_config, normalize_user_config

            cfg = normalize_user_config(load_user_config())
            st = gui_state(cfg)

            def _get_str(key: str, default: str = "") -> str:
                v = st.get(key)
                return str(v) if v is not None else default

            self.pptx_edit.setText(_get_str("pptx_path", self.pptx_edit.text()))
            self.out_edit.setText(_get_str("out_dir", self.out_edit.text()))
            self.pages_edit.setText(_get_str("pages", self.pages_edit.text()) or "all")

            prov = _get_str("tts_provider", "").strip()
            if prov:
                idx = self.tts_combo.findText(prov)
                if idx >= 0:
                    self.tts_combo.setCurrentIndex(idx)

            voice = _get_str("voice_id", "").strip()
            if voice:
                self.voice_combo.setCurrentText(voice)

            try:
                rate = float(st.get("tts_rate", self.tts_rate_spin.value()))
                self.tts_rate_spin.setValue(rate)
            except Exception:
                pass

            subc = _get_str("subtitle_color", "").strip()
            self._set_subtitle_color(subc or None)
            hic = _get_str("subtitle_highlight_color", "").strip()
            self._set_subtitle_highlight_color(hic or None)
            try:
                hm = _get_str("subtitle_highlight_mode", "none").strip().lower() or "none"
                idx = self.subtitle_highlight_mode_combo.findData(hm)
                if idx >= 0:
                    self.subtitle_highlight_mode_combo.setCurrentIndex(idx)
            except Exception:
                pass
            try:
                self.subtitle_fade_in_spin.setValue(int(st.get("subtitle_fade_in_ms", self.subtitle_fade_in_spin.value())))
            except Exception:
                pass
            try:
                self.subtitle_fade_out_spin.setValue(
                    int(st.get("subtitle_fade_out_ms", self.subtitle_fade_out_spin.value()))
                )
            except Exception:
                pass
            try:
                self.subtitle_scale_from_spin.setValue(
                    int(st.get("subtitle_scale_from", self.subtitle_scale_from_spin.value()))
                )
            except Exception:
                pass
            try:
                self.subtitle_scale_to_spin.setValue(int(st.get("subtitle_scale_to", self.subtitle_scale_to_spin.value())))
            except Exception:
                pass
            try:
                self.subtitle_outline_spin.setValue(int(st.get("subtitle_outline", self.subtitle_outline_spin.value())))
            except Exception:
                pass
            try:
                self.subtitle_shadow_spin.setValue(int(st.get("subtitle_shadow", self.subtitle_shadow_spin.value())))
            except Exception:
                pass
            font = _get_str("subtitle_font", "").strip()
            try:
                if font:
                    idx = self.subtitle_font_edit.findText(font)
                    if idx >= 0:
                        self.subtitle_font_edit.setCurrentIndex(idx)
                    else:
                        self.subtitle_font_edit.setCurrentText(font)
            except Exception:
                pass
            try:
                raw = int(st.get("subtitle_size", self.subtitle_size_spin.value()) or 0)
                self.subtitle_size_spin.setValue(raw if raw > 0 else 48)
            except Exception:
                pass

            self.bgm_path_edit.setText(_get_str("bgm_path", "").strip())
            try:
                self.bgm_volume_spin.setValue(float(st.get("bgm_volume", self.bgm_volume_spin.value())))
            except Exception:
                pass
            try:
                self.narration_volume_spin.setValue(
                    float(st.get("narration_volume", self.narration_volume_spin.value()))
                )
            except Exception:
                pass
            try:
                self.bgm_fade_in_spin.setValue(float(st.get("bgm_fade_in_s", self.bgm_fade_in_spin.value())))
            except Exception:
                pass
            try:
                self.bgm_fade_out_spin.setValue(float(st.get("bgm_fade_out_s", self.bgm_fade_out_spin.value())))
            except Exception:
                pass

            # window geometry
            try:
                w = int(st.get("window_w", 0) or 0)
                h = int(st.get("window_h", 0) or 0)
                if w > 200 and h > 200:
                    self.resize(w, h)
            except Exception:
                pass

        def _persist_gui_state_to_config(self) -> None:
            from note2video.user_config import load_user_config, normalize_user_config, save_user_config

            cfg = normalize_user_config(load_user_config())
            gui = dict(cfg.get("gui") or {})
            gui.update(
                {
                    "pptx_path": self.pptx_edit.text().strip(),
                    "out_dir": self.out_edit.text().strip(),
                    "pages": self.pages_edit.text().strip() or "all",
                    "tts_provider": self.tts_combo.currentText().strip(),
                    "voice_id": self._current_voice_id(),
                    "tts_rate": float(self.tts_rate_spin.value()),
                    "subtitle_color": self.subtitle_color_value.text().strip(),
                    "subtitle_highlight_mode": str(self.subtitle_highlight_mode_combo.currentData() or "none"),
                    "subtitle_highlight_color": self.subtitle_highlight_color_value.text().strip(),
                    "subtitle_fade_in_ms": int(self.subtitle_fade_in_spin.value()),
                    "subtitle_fade_out_ms": int(self.subtitle_fade_out_spin.value()),
                    "subtitle_scale_from": int(self.subtitle_scale_from_spin.value()),
                    "subtitle_scale_to": int(self.subtitle_scale_to_spin.value()),
                    "subtitle_outline": int(self.subtitle_outline_spin.value()),
                    "subtitle_shadow": int(self.subtitle_shadow_spin.value()),
                    "subtitle_font": self.subtitle_font_edit.currentText().strip(),
                    "subtitle_size": int(self.subtitle_size_spin.value()),
                    "bgm_path": self.bgm_path_edit.text().strip(),
                    "bgm_volume": float(self.bgm_volume_spin.value()),
                    "narration_volume": float(self.narration_volume_spin.value()),
                    "bgm_fade_in_s": float(self.bgm_fade_in_spin.value()),
                    "bgm_fade_out_s": float(self.bgm_fade_out_spin.value()),
                    "window_w": int(self.size().width()),
                    "window_h": int(self.size().height()),
                }
            )
            cfg["gui"] = gui
            save_user_config(cfg)

        def _set_subtitle_color(self, color_hex: str | None) -> None:
            c = (color_hex or "").strip()
            if not c:
                self.subtitle_color_value.setText("")
                self.subtitle_color_preview.setStyleSheet("background: transparent; border: 1px solid #999;")
                return
            if not c.startswith("#"):
                c = "#" + c
            self.subtitle_color_value.setText(c)
            self.subtitle_color_preview.setStyleSheet(f"background: {c}; border: 1px solid #333;")

        def _pick_subtitle_color(self) -> None:
            QtWidgets = self._QtWidgets
            try:
                from PySide6.QtGui import QColor
            except Exception:
                QColor = None  # type: ignore
            current = self.subtitle_color_value.text().strip() or "#FFFFFF"
            dlg = QtWidgets.QColorDialog(self)
            dlg.setOption(QtWidgets.QColorDialog.ColorDialogOption.ShowAlphaChannel, False)
            if QColor is not None:
                dlg.setCurrentColor(QColor(current))
            if dlg.exec():
                col = dlg.currentColor()
                self._set_subtitle_color(col.name().upper())

        def _clear_subtitle_color(self) -> None:
            self._set_subtitle_color(None)

        def _set_subtitle_highlight_color(self, color_hex: str | None) -> None:
            c = (color_hex or "").strip()
            if not c:
                self.subtitle_highlight_color_value.setText("")
                self.subtitle_highlight_color_preview.setStyleSheet("background: transparent; border: 1px solid #999;")
                return
            if not c.startswith("#"):
                c = "#" + c
            self.subtitle_highlight_color_value.setText(c)
            self.subtitle_highlight_color_preview.setStyleSheet(f"background: {c}; border: 1px solid #333;")

        def _pick_subtitle_highlight_color(self) -> None:
            QtWidgets = self._QtWidgets
            try:
                from PySide6.QtGui import QColor
            except Exception:
                QColor = None  # type: ignore
            current = self.subtitle_highlight_color_value.text().strip() or "#FFD400"
            dlg = QtWidgets.QColorDialog(self)
            dlg.setOption(QtWidgets.QColorDialog.ColorDialogOption.ShowAlphaChannel, False)
            if QColor is not None:
                dlg.setCurrentColor(QColor(current))
            if dlg.exec():
                col = dlg.currentColor()
                self._set_subtitle_highlight_color(col.name().upper())

        def _clear_subtitle_highlight_color(self) -> None:
            self._set_subtitle_highlight_color(None)

        def _pick_bgm(self) -> None:
            QtWidgets = self._QtWidgets
            path, _ = QtWidgets.QFileDialog.getOpenFileName(
                self,
                "选择背景音乐",
                self.bgm_path_edit.text() or str(Path.cwd()),
                "Audio (*.mp3 *.wav *.m4a *.aac *.flac);;All files (*)",
            )
            if path:
                self.bgm_path_edit.setText(path)

        def _voice_combo_reset_default(self) -> None:
            self.voice_combo.blockSignals(True)
            self.voice_combo.clear()
            self.voice_combo.addItem("（默认，由引擎决定）", "")
            self.voice_combo.setCurrentIndex(0)
            self.voice_combo.blockSignals(False)

        def _on_tts_provider_changed(self, _text: str) -> None:
            self._all_voice_items = []
            self.locale_combo.blockSignals(True)
            self.locale_combo.clear()
            self.locale_combo.addItem("（请先刷新音色列表）", None)
            self.locale_combo.setCurrentIndex(0)
            self.locale_combo.blockSignals(False)
            self._voice_combo_reset_default()

        def _current_voice_id(self) -> str:
            idx = self.voice_combo.currentIndex()
            if idx == 0:
                return ""
            if idx > 0:
                data = self.voice_combo.itemData(idx)
                if data is not None and str(data).strip():
                    return str(data).strip()
            text = self.voice_combo.currentText().strip()
            if text.startswith("（默认"):
                return ""
            return text

        def _repopulate_voice_combo(self) -> None:
            filter_data = self.locale_combo.currentData()
            self.voice_combo.blockSignals(True)
            self.voice_combo.clear()
            self.voice_combo.addItem("（默认，由引擎决定）", "")
            for item in self._all_voice_items:
                if not _voice_matches_locale_filter(filter_data, item):
                    continue
                name = str(item.get("name", "") or "")
                display = str(item.get("display_name", "") or "") or name
                gender = str(item.get("gender", "") or "").strip()
                bits = [display]
                if name and name != display:
                    bits.append(name)
                if gender:
                    bits.append(gender)
                label = " — ".join(bits)
                self.voice_combo.addItem(label, name)
            self.voice_combo.setCurrentIndex(0)
            self.voice_combo.blockSignals(False)

            if (
                filter_data == LOCALE_FILTER_MAINLAND
                and self.voice_combo.count() <= 1
                and self._all_voice_items
            ):
                idx_all = self.locale_combo.findData(LOCALE_FILTER_ALL)
                if idx_all >= 0:
                    self.locale_combo.blockSignals(True)
                    self.locale_combo.setCurrentIndex(idx_all)
                    self.locale_combo.blockSignals(False)
                    self._repopulate_voice_combo()

        def _refresh_voice_list(self) -> None:
            from note2video.tts.voice import VoiceGenerationError, list_available_voices

            provider = self.tts_combo.currentText().strip() or "pyttsx3"
            self._append_log(f"正在加载音色列表：{provider} …")
            try:
                voices = list_available_voices(provider_name=provider, keyword="")
            except (VoiceGenerationError, ValueError) as exc:
                self._append_log(f"音色列表加载失败：{exc}")
                QtWidgets = self._QtWidgets
                QtWidgets.QMessageBox.warning(self, "音色列表", str(exc))
                return
            self._all_voice_items = list(voices[:800])
            raw_keys = sorted(
                {_voice_locale_key(v) for v in self._all_voice_items},
                key=lambda k: (k.startswith("("), k.lower()),
            )
            other_keys = [k for k in raw_keys if not _is_mainland_mandarin_locale_key(k)]

            self.locale_combo.blockSignals(True)
            self.locale_combo.clear()
            self.locale_combo.addItem("中国大陆 · 普通话", LOCALE_FILTER_MAINLAND)
            self.locale_combo.addItem("全部地区 / 语言", LOCALE_FILTER_ALL)
            for k in other_keys:
                self.locale_combo.addItem(_locale_key_label_zh(k), k)
            self.locale_combo.setCurrentIndex(0)
            self.locale_combo.blockSignals(False)
            self._repopulate_voice_combo()
            self._append_log(
                f"已加载 {len(self._all_voice_items)} 条音色；默认「中国大陆 · 普通话」，再选具体声音；可手动编辑 Voice ID。"
            )

        def _preview_voice(self) -> None:
            QtCore = self._QtCore
            QtWidgets = self._QtWidgets
            if self._preview_thread is not None or self._preview_proc is not None:
                self._append_log("试听正在进行中，请稍候…")
                return
            if self._pipeline_busy or self._thread is not None:
                self._append_log("当前有任务在运行，请稍后再试听。")
                return

            if self._preview_temp_path and self._preview_temp_path.exists():
                try:
                    self._preview_temp_path.unlink()
                except OSError:
                    pass

            provider = self.tts_combo.currentText().strip() or "pyttsx3"
            voice_id = self._current_voice_id()
            rate = float(self.tts_rate_spin.value())

            fd, raw_path = tempfile.mkstemp(suffix=".wav", prefix="note2video_preview_")
            os.close(fd)
            out_path = Path(raw_path)

            self._preview_temp_path = out_path
            self.voice_preview_btn.setEnabled(False)
            self._append_log("正在生成试听音频…")

            # Run preview in a subprocess. Avoid QThread/Qt signals here: on some Windows setups
            # PySide/Qt + threads can end up in native crashes (0xC0000409). We poll the child
            # process state from the main thread via QTimer instead.
            import subprocess

            cmd = [
                sys.executable,
                "-m",
                "note2video.tts.preview_worker",
                "--provider",
                provider,
                "--voice",
                voice_id,
                "--tts-rate",
                str(rate),
                "--text",
                PREVIEW_SAMPLE_TEXT,
                "--out",
                str(out_path),
            ]

            fd_log, raw_log_path = tempfile.mkstemp(suffix=".log", prefix="note2video_preview_")
            os.close(fd_log)
            log_path = Path(raw_log_path)
            self._preview_log_path = log_path
            try:
                self._preview_log_fh = open(log_path, "wb")
            except OSError:
                self._preview_log_fh = None

            creationflags = 0
            if sys.platform == "win32":
                creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

            try:
                self._preview_proc = subprocess.Popen(
                    cmd,
                    stdout=(self._preview_log_fh if self._preview_log_fh is not None else subprocess.DEVNULL),
                    stderr=(self._preview_log_fh if self._preview_log_fh is not None else subprocess.DEVNULL),
                    creationflags=creationflags,
                )
            except Exception as exc:
                if self._preview_log_fh is not None:
                    try:
                        self._preview_log_fh.close()
                    except Exception:
                        pass
                self._preview_log_fh = None
                self._preview_proc = None
                self._preview_log_path = None
                self.voice_preview_btn.setEnabled(True)
                self._append_log(f"试听失败：无法启动子进程：{type(exc).__name__}: {exc}")
                QtWidgets.QMessageBox.warning(self, "试听失败", f"无法启动试听子进程：{exc}")
                return

            self._append_log(f"cmd: {cmd}")
            timer = QtCore.QTimer(self)
            timer.setInterval(120)
            self._preview_timer = timer

            def _finish_preview() -> None:
                proc = self._preview_proc
                if proc is None:
                    return
                rc = proc.poll()
                if rc is None:
                    return

                timer.stop()
                self._preview_timer = None
                self._preview_proc = None
                self.voice_preview_btn.setEnabled(True)

                if self._preview_log_fh is not None:
                    try:
                        self._preview_log_fh.close()
                    except Exception:
                        pass
                self._preview_log_fh = None

                log_text = ""
                if self._preview_log_path and self._preview_log_path.exists():
                    try:
                        raw = self._preview_log_path.read_bytes()
                        log_text = raw.decode("utf-8", errors="replace").strip()
                    except Exception:
                        log_text = ""

                p = self._preview_temp_path
                if rc != 0:
                    self._append_log(f"试听子进程退出码：{rc}")
                    if log_text:
                        self._append_log("preview subprocess log:\n" + log_text)
                    QtWidgets.QMessageBox.warning(self, "试听失败", f"试听子进程退出码：{rc}\n请查看下方日志。")
                    return

                if not p or not p.exists():
                    self._append_log("试听失败：未生成输出文件。")
                    if log_text:
                        self._append_log("preview subprocess log:\n" + log_text)
                    QtWidgets.QMessageBox.warning(self, "试听失败", "试听未生成输出音频文件，请查看下方日志。")
                    return

                try:
                    size = p.stat().st_size
                except OSError:
                    size = -1
                self._append_log(f"试听音频已生成：{p}  ({size} bytes)")
                self._last_preview_path = p
                self.preview_open_btn.setEnabled(bool(self._preview_web_available))
                self.preview_play_system_btn.setEnabled(True)
                self.preview_reveal_btn.setEnabled(True)
                if self._preview_web_available and self._preview_web_view is not None:
                    self._load_preview_in_web(p, autoplay=False)
                    # Player is ready; user can hit play or "打开最近一次试听" to autoplay.
                    self._append_log("内嵌播放器已加载试听文件，可在右侧播放器中播放。")
                else:
                    self._append_log("提示：QtWebEngine 不可用，可用右侧按钮用系统默认播放器播放。")

            timer.timeout.connect(_finish_preview)
            timer.start()

        def _safe_play_preview(self, path: Path) -> None:
            # Default to system player. QtMultimedia/QMediaPlayer has been observed to crash on
            # some Windows setups (e.g. exit code 0xC0000409) depending on codecs/backends.
            mode = (os.getenv("NOTE2VIDEO_GUI_PREVIEW_PLAYER") or "").strip().lower()
            if mode in {"qt", "qtmultimedia"}:
                QtWidgets = self._QtWidgets
                try:
                    self._play_preview_file(path)
                except Exception:
                    self._append_log(traceback.format_exc())
                    self._reveal_preview_file(path)
                    QtWidgets.QMessageBox.warning(
                        self,
                        "试听",
                        "内置播放器播放失败，已改为在资源管理器中定位试听文件（见日志）。",
                    )
                return

            # Default: do NOT auto-open the file with a media player, because some file
            # associations / codecs may crash the current process on Windows. Reveal it instead.
            if mode in {"open", "system", "startfile"}:
                self._open_preview_file(path)
                return
            self._reveal_preview_file(path)

        def _reveal_preview_file(self, path: Path) -> None:
            """Reveal the generated WAV in file explorer (safer than auto-playing)."""
            QtWidgets = self._QtWidgets
            try:
                if sys.platform == "win32":
                    import subprocess

                    subprocess.Popen(
                        ["explorer.exe", "/select,", str(path)],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                    )
                    self._append_log("已在资源管理器中定位试听音频文件。")
                    return
                raise RuntimeError("Reveal is only implemented on Windows.")
            except Exception as exc:
                self._append_log(f"无法在资源管理器中定位试听文件：{exc}")
                QtWidgets.QMessageBox.information(
                    self,
                    "试听音频已生成",
                    f"已生成试听文件：\n{path}\n\n当前环境无法自动播放，请手动打开该文件。",
                )

        def _open_preview_file(self, path: Path) -> None:
            """Open the generated WAV with the system default player (may be unsafe on some setups)."""
            QtWidgets = self._QtWidgets
            try:
                if sys.platform == "win32":
                    os.startfile(str(path))  # type: ignore[attr-defined]
                    self._append_log("已调用系统默认播放器打开试听音频。")
                    return
                raise RuntimeError("os.startfile is only available on Windows.")
            except Exception as exc:
                self._append_log(f"无法打开试听文件：{exc}")
                QtWidgets.QMessageBox.information(
                    self,
                    "试听音频已生成",
                    f"已生成试听文件：\n{path}\n\n当前环境无法自动播放，请手动打开该文件。",
                )

        def _play_preview_file(self, path: Path) -> None:
            QtWidgets = self._QtWidgets
            try:
                from PySide6.QtCore import QUrl
                from PySide6.QtMultimedia import QAudioOutput, QMediaPlayer
            except ImportError as exc:
                self._append_log(f"无法播放：缺少 QtMultimedia（{exc}）")
                QtWidgets.QMessageBox.warning(
                    self,
                    "试听",
                    "当前环境无法加载 QtMultimedia，请确认已安装完整 PySide6（pip install -U PySide6）。",
                )
                self._reveal_preview_file(path)
                return

            if self._preview_audio_out is None:
                self._preview_audio_out = QAudioOutput(self)
                self._preview_player = QMediaPlayer(self)
                self._preview_player.setAudioOutput(self._preview_audio_out)
            assert self._preview_player is not None
            try:
                self._preview_player.stop()
                self._preview_player.setSource(QUrl.fromLocalFile(str(path.resolve())))
                self._preview_audio_out.setVolume(1.0)
                self._preview_player.play()
            except Exception as exc:
                self._reveal_preview_file(path)
                raise RuntimeError(f"QMediaPlayer 播放失败：{exc}") from exc
            self._append_log("开始播放试听（请检查系统音量）。")

        def _append_log(self, text: str) -> None:
            self.log.appendPlainText(text.rstrip("\n"))

        def _open_last_preview_in_web(self) -> None:
            QtWidgets = self._QtWidgets
            p = self._last_preview_path
            if not p or not p.exists():
                QtWidgets.QMessageBox.information(self, "试听", "还没有可播放的试听文件。")
                return
            if not self._preview_web_available or self._preview_web_view is None:
                QtWidgets.QMessageBox.information(self, "试听", "内嵌浏览器播放不可用（缺少 QtWebEngine）。")
                return
            self._load_preview_in_web(p, autoplay=True)

        def _play_last_preview_system(self) -> None:
            QtWidgets = self._QtWidgets
            p = self._last_preview_path
            if not p or not p.exists():
                QtWidgets.QMessageBox.information(self, "试听", "还没有可播放的试听文件。")
                return
            self._open_preview_file(p)

        def _reveal_last_preview(self) -> None:
            QtWidgets = self._QtWidgets
            p = self._last_preview_path
            if not p or not p.exists():
                QtWidgets.QMessageBox.information(self, "试听", "还没有可播放的试听文件。")
                return
            self._reveal_preview_file(p)

        def _load_preview_in_web(self, path: Path, *, autoplay: bool) -> None:
            view = self._preview_web_view
            if view is None:
                return
            try:
                from PySide6.QtCore import QUrl
                from PySide6.QtWebEngineCore import QWebEngineSettings  # type: ignore
            except Exception:
                return

            # Make the HTML page same-origin with the WAV file directory so the <audio> element
            # can read it reliably on Windows.
            base_dir = path.parent.resolve()
            base_url = QUrl.fromLocalFile(str(base_dir) + os.sep)
            auto = " autoplay" if autoplay else ""
            safe_path = str(path).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            safe_name = path.name.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("'", "&#39;")
            html = (
                "<!doctype html><html><head><meta charset='utf-8'></head><body>"
                "<div style='font-family:Segoe UI,Arial,sans-serif;font-size:12px;color:#444;'>"
                f"试听文件：{safe_path}"
                "</div>"
                f"<audio controls{auto} style='width:100%; margin-top:6px;' src='{safe_name}'></audio>"
                "</body></html>"
            )

            try:
                settings = view.settings()
                settings.setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessFileUrls, True)
                settings.setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessRemoteUrls, False)
            except Exception:
                pass

            view.setHtml(html, base_url)

        def _pick_pptx(self) -> None:
            QtWidgets = self._QtWidgets
            path, _ = QtWidgets.QFileDialog.getOpenFileName(
                self, "选择 PPTX", str(Path.cwd()), "PowerPoint (*.pptx)"
            )
            if path:
                self.pptx_edit.setText(path)

        def _pick_out_dir(self) -> None:
            QtWidgets = self._QtWidgets
            path = QtWidgets.QFileDialog.getExistingDirectory(
                self, "选择输出目录", self.out_edit.text() or str(Path.cwd())
            )
            if path:
                self.out_edit.setText(path)

        def _validate(self) -> JobConfig:
            pptx = Path(self.pptx_edit.text().strip().strip('"'))
            out_dir = Path(self.out_edit.text().strip().strip('"'))
            pages = self.pages_edit.text().strip() or "all"
            tts_provider = self.tts_combo.currentText().strip()
            voice_id = self._current_voice_id()
            tts_rate = float(self.tts_rate_spin.value())
            subtitle_color = self.subtitle_color_value.text().strip() or None
            subtitle_highlight_mode = str(self.subtitle_highlight_mode_combo.currentData() or "none").strip() or "none"
            subtitle_highlight_color = self.subtitle_highlight_color_value.text().strip() or None
            subtitle_fade_in_ms = int(self.subtitle_fade_in_spin.value())
            subtitle_fade_out_ms = int(self.subtitle_fade_out_spin.value())
            subtitle_scale_from = int(self.subtitle_scale_from_spin.value())
            subtitle_scale_to = int(self.subtitle_scale_to_spin.value())
            subtitle_outline = int(self.subtitle_outline_spin.value())
            subtitle_shadow = int(self.subtitle_shadow_spin.value())
            subtitle_font = self.subtitle_font_edit.currentText().strip() or None
            subtitle_size = int(self.subtitle_size_spin.value())
            bgm_path = self.bgm_path_edit.text().strip() or None
            bgm_volume = float(self.bgm_volume_spin.value())
            narration_volume = float(self.narration_volume_spin.value())
            bgm_fade_in_s = float(self.bgm_fade_in_spin.value())
            bgm_fade_out_s = float(self.bgm_fade_out_spin.value())

            if not pptx.exists() or pptx.suffix.lower() != ".pptx":
                raise ValueError("请选择有效的 .pptx 文件。")
            if not out_dir:
                raise ValueError("请选择输出目录。")

            return JobConfig(
                mode="extract",
                pptx_path=pptx,
                out_dir=out_dir,
                pages=pages,
                tts_provider=tts_provider,
                voice_id=voice_id,
                tts_rate=tts_rate,
                minimax_base_url=None,
                subtitle_color=subtitle_color,
                subtitle_highlight_mode=subtitle_highlight_mode,
                subtitle_highlight_color=subtitle_highlight_color,
                subtitle_fade_in_ms=subtitle_fade_in_ms,
                subtitle_fade_out_ms=subtitle_fade_out_ms,
                subtitle_scale_from=subtitle_scale_from,
                subtitle_scale_to=subtitle_scale_to,
                subtitle_outline=subtitle_outline,
                subtitle_shadow=subtitle_shadow,
                subtitle_font=subtitle_font,
                subtitle_size=subtitle_size,
                bgm_path=bgm_path,
                bgm_volume=bgm_volume,
                narration_volume=narration_volume,
                bgm_fade_in_s=bgm_fade_in_s,
                bgm_fade_out_s=bgm_fade_out_s,
            )

        def _set_running(self, running: bool) -> None:
            self.extract_btn.setEnabled(not running)
            self.build_btn.setEnabled(not running)
            self.stop_btn.setEnabled(running)
            self.tts_combo.setEnabled(not running)
            self.locale_combo.setEnabled(not running)
            self.voice_combo.setEnabled(not running)
            self.voice_refresh_btn.setEnabled(not running)
            self.voice_preview_btn.setEnabled(not running and self._preview_thread is None and self._preview_proc is None)
            self.tts_rate_spin.setEnabled(not running)
            self.progress.setEnabled(True)
            if not running:
                # Keep the last progress visible; mark idle explicitly.
                if self.progress.value() >= 4:
                    self.progress.setFormat("完成")
                else:
                    self.progress.setFormat("就绪")

        def _handle_pipeline_log(self, text: str) -> None:
            self._append_log(text)
            t = (text or "").strip()
            if t == "阶段：extract":
                self.progress.setValue(1)
                self._pipeline_stage_label = "extract (1/4)"
                self.progress.setFormat(f"进度：{self._pipeline_stage_label}")
            elif t == "阶段：voice":
                self.progress.setValue(2)
                self._pipeline_stage_label = "voice (2/4)"
                self.progress.setFormat(f"进度：{self._pipeline_stage_label}")
            elif t == "阶段：subtitle":
                self.progress.setValue(3)
                self._pipeline_stage_label = "subtitle (3/4)"
                self.progress.setFormat(f"进度：{self._pipeline_stage_label}")
            elif t == "阶段：render":
                self.progress.setValue(4)
                self._pipeline_stage_label = "render (4/4)"
                self.progress.setFormat(f"进度：{self._pipeline_stage_label}")
            elif t.startswith("细节："):
                detail = t[len("细节：") :].strip()
                if self._pipeline_stage_label:
                    self.progress.setFormat(f"进度：{self._pipeline_stage_label} · {detail}")
                else:
                    self.progress.setFormat(f"进度：{detail}")
            elif t.startswith("完成："):
                # Keep as-is; final "退出码" will come next.
                pass

        def _start(self, mode: str) -> None:
            QtCore = self._QtCore
            QtWidgets = self._QtWidgets

            if self._preview_thread is not None or self._preview_proc is not None:
                self._append_log("试听进行中，请等待试听完成后再执行任务。")
                return

            if self._thread is not None or self._pipeline_busy:
                return

            try:
                config = self._validate()
                config = JobConfig(
                    mode=mode,
                    pptx_path=config.pptx_path,
                    out_dir=config.out_dir,
                    pages=config.pages,
                    tts_provider=config.tts_provider,
                    voice_id=config.voice_id,
                    tts_rate=config.tts_rate,
                    minimax_base_url=config.minimax_base_url,
                    subtitle_color=config.subtitle_color,
                    subtitle_highlight_mode=config.subtitle_highlight_mode,
                    subtitle_highlight_color=config.subtitle_highlight_color,
                    subtitle_fade_in_ms=config.subtitle_fade_in_ms,
                    subtitle_fade_out_ms=config.subtitle_fade_out_ms,
                    subtitle_scale_from=config.subtitle_scale_from,
                    subtitle_scale_to=config.subtitle_scale_to,
                    subtitle_outline=config.subtitle_outline,
                    subtitle_shadow=config.subtitle_shadow,
                    subtitle_font=config.subtitle_font,
                    subtitle_size=config.subtitle_size,
                    bgm_path=config.bgm_path,
                    bgm_volume=config.bgm_volume,
                    narration_volume=config.narration_volume,
                    bgm_fade_in_s=config.bgm_fade_in_s,
                    bgm_fade_out_s=config.bgm_fade_out_s,
                )
            except Exception as exc:
                QtWidgets.QMessageBox.warning(self, "参数错误", str(exc))
                return

            self.log.clear()
            self._append_log("开始运行…")
            self._pipeline_busy = True
            self._set_running(True)
            self.progress.setValue(0)
            self.progress.setFormat("进度：准备中 (0/4)")

            # PowerPoint COM on Windows must run on the GUI (STA) thread; a QThread
            # background worker often fails or hangs during extract.
            if sys.platform == "win32":
                QtCore.QTimer.singleShot(0, partial(self._finish_pipeline_main_thread, config))
                return

            thread = QtCore.QThread()
            worker = _build_worker(QtCore, config)
            self._thread = thread
            self._worker = worker
            worker.moveToThread(thread)
            worker.log.connect(self._handle_pipeline_log)

            thread.started.connect(worker.run)

            def _done(exit_code: int) -> None:
                self._append_log(f"退出码：{exit_code}")
                self._pipeline_busy = False
                self._set_running(False)
                thread.quit()
                thread.wait(3000)
                self._thread = None
                self._worker = None
                if exit_code != 0:
                    QtWidgets.QMessageBox.warning(self, "运行失败", "任务未成功完成，请查看下方日志。")

            worker.done.connect(_done)
            thread.start()

        def _finish_pipeline_main_thread(self, config: JobConfig) -> None:
            QtWidgets = self._QtWidgets
            exit_code = _run_pipeline_with_log(config, self._handle_pipeline_log)
            self._append_log(f"退出码：{exit_code}")
            self._pipeline_busy = False
            self._set_running(False)
            if exit_code != 0:
                QtWidgets.QMessageBox.warning(self, "运行失败", "任务未成功完成，请查看下方日志。")

        def _stop(self) -> None:
            # Best-effort: we can stop the thread, but subprocesses / COM calls may not be interruptible.
            if self._thread is None:
                return
            self._append_log("请求停止（best-effort）…")
            self._thread.requestInterruption()

        def _open_tts_settings(self) -> None:
            _run_tts_settings_dialog(self, QtWidgets=self._QtWidgets, append_log=self._append_log)

    return MainWindow


class MainWindow:  # thin wrapper used by main()
    def __init__(self, *, QtCore, QtWidgets) -> None:
        self._QtCore = QtCore
        self._QtWidgets = QtWidgets
        WindowCls = _build_ui(QtWidgets)
        window = WindowCls()
        window._QtCore = QtCore
        window._QtWidgets = QtWidgets
        self._window = window

    def resize(self, w: int, h: int) -> None:
        self._window.resize(w, h)

    def show(self) -> None:
        self._window.show()


if __name__ == "__main__":
    raise SystemExit(main())
