from __future__ import annotations

from dataclasses import dataclass

from note2video.gui.preview_model import PreviewData
from note2video.subtitle.wrap import subtitle_wrap_layout_from_canvas, wrap_subtitle_text


@dataclass(frozen=True)
class PreviewStyle:
    subtitle_color: str | None = None
    subtitle_font: str = ""
    subtitle_size: int = 48
    subtitle_outline: int = 1
    subtitle_shadow: int = 0
    subtitle_y_ratio: float | None = None


class SubtitlePreviewWidget:  # imported lazily from app.py after PySide6 is available
    def __new__(cls, *args, **kwargs):
        from PySide6 import QtWidgets

        class _SubtitlePreviewWidget(QtWidgets.QWidget):
            def __init__(self, *widget_args, **widget_kwargs) -> None:
                super().__init__(*widget_args, **widget_kwargs)
                self._preview_data: PreviewData | None = None
                self._preview_style = PreviewStyle()
                self.setMinimumSize(320, 220)
                self.setAutoFillBackground(False)

            def set_preview(self, *, data: PreviewData | None, style: PreviewStyle) -> None:
                self._preview_data = data
                self._preview_style = style
                self.update()

            def paintEvent(self, event) -> None:  # noqa: N802
                from PySide6 import QtCore, QtGui

                painter = QtGui.QPainter(self)
                painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing, True)
                painter.setRenderHint(QtGui.QPainter.RenderHint.TextAntialiasing, True)
                painter.fillRect(self.rect(), QtGui.QColor("#1E1E1E"))

                data = self._preview_data
                if data is None:
                    self._draw_placeholder(painter, self.rect(), "暂无预览")
                    return

                frame = self.rect().adjusted(12, 12, -12, -12)
                target = self._fit_rect(frame, data.canvas_w, data.canvas_h)
                self._draw_background(painter, target, data)
                self._draw_header(painter, target, data)
                if data.active_text.strip():
                    self._draw_subtitle(painter, target, data)
                else:
                    self._draw_placeholder(painter, target, data.status_text)

                painter.setPen(QtGui.QColor("#D0D0D0"))
                painter.setFont(self._status_font())
                painter.drawText(
                    QtCore.QRectF(frame.left(), target.bottom() + 8, frame.width(), max(24, frame.bottom() - target.bottom())),
                    int(QtCore.Qt.AlignmentFlag.AlignLeft | QtCore.Qt.AlignmentFlag.AlignVCenter),
                    data.status_text,
                )

            def _fit_rect(self, outer, width: int, height: int):
                from PySide6 import QtCore

                if width <= 0 or height <= 0:
                    return QtCore.QRectF(outer)
                scale = min(float(outer.width()) / float(width), float(outer.height()) / float(height))
                w = max(1.0, float(width) * scale)
                h = max(1.0, float(height) * scale)
                x = float(outer.left()) + (float(outer.width()) - w) / 2.0
                y = float(outer.top()) + (float(outer.height()) - h) / 2.0
                return QtCore.QRectF(x, y, w, h)

            def _draw_background(self, painter, target, data: PreviewData) -> None:
                from PySide6 import QtCore, QtGui

                painter.save()
                painter.setClipRect(target)
                painter.fillRect(target, QtGui.QColor("black"))
                if data.image_path is not None:
                    pixmap = QtGui.QPixmap(str(data.image_path))
                else:
                    pixmap = QtGui.QPixmap()
                if not pixmap.isNull():
                    image_rect = self._fit_rect(target, pixmap.width(), pixmap.height())
                    painter.drawPixmap(
                        image_rect,
                        pixmap,
                        QtCore.QRectF(0.0, 0.0, float(pixmap.width()), float(pixmap.height())),
                    )
                else:
                    gradient = QtGui.QLinearGradient(target.topLeft(), target.bottomRight())
                    gradient.setColorAt(0.0, QtGui.QColor("#2B2B2B"))
                    gradient.setColorAt(1.0, QtGui.QColor("#3E4A5B"))
                    painter.fillRect(target, gradient)
                    pen = QtGui.QPen(QtGui.QColor("#707070"))
                    pen.setStyle(QtCore.Qt.PenStyle.DashLine)
                    painter.setPen(pen)
                    painter.drawRect(target)
                painter.restore()

            def _draw_header(self, painter, target, data: PreviewData) -> None:
                from PySide6 import QtCore, QtGui

                title = data.title.strip() or f"第 {data.page} 页"
                badge = f"P{data.page} · {self._source_label(data.text_source)}"
                if data.cue_count > 1 and data.active_cue_index >= 0:
                    badge += f" {data.active_cue_index + 1}/{data.cue_count}"

                painter.save()
                painter.setFont(self._header_font())
                fm = QtGui.QFontMetrics(painter.font())
                title_rect = QtCore.QRectF(target.left() + 12, target.top() + 10, target.width() - 180, fm.height() + 10)
                badge_w = max(88, fm.horizontalAdvance(badge) + 18)
                badge_rect = QtCore.QRectF(target.right() - badge_w - 12, target.top() + 10, badge_w, fm.height() + 10)

                painter.setPen(QtCore.Qt.PenStyle.NoPen)
                painter.setBrush(QtGui.QColor(0, 0, 0, 120))
                painter.drawRoundedRect(title_rect, 8, 8)
                painter.drawRoundedRect(badge_rect, 8, 8)

                painter.setPen(QtGui.QColor("white"))
                painter.drawText(title_rect, int(QtCore.Qt.AlignmentFlag.AlignCenter), title)
                painter.drawText(badge_rect, int(QtCore.Qt.AlignmentFlag.AlignCenter), badge)
                painter.restore()

            def _draw_subtitle(self, painter, target, data: PreviewData) -> None:
                from PySide6 import QtCore, QtGui

                style = self._preview_style
                layout = subtitle_wrap_layout_from_canvas(
                    canvas_w=data.canvas_w,
                    canvas_h=data.canvas_h,
                    font_size=max(8, int(style.subtitle_size)),
                    margin_l=max(24, int(round(80 * (float(data.canvas_w) / 1920.0)))),
                    margin_r=max(24, int(round(80 * (float(data.canvas_w) / 1920.0)))),
                    outline=max(0, int(style.subtitle_outline)),
                    max_lines=4,
                )
                wrapped = wrap_subtitle_text(data.active_text, layout=layout, font_name=style.subtitle_font)
                lines = [line.strip() for line in wrapped.splitlines() if line.strip()]
                if not lines:
                    return

                scale = min(float(target.width()) / float(data.canvas_w), float(target.height()) / float(data.canvas_h))
                font = QtGui.QFont(self.font())
                if style.subtitle_font.strip():
                    font.setFamily(style.subtitle_font.strip())
                font.setPixelSize(max(10, int(round(float(style.subtitle_size) * scale))))
                painter.setFont(font)
                fm = QtGui.QFontMetrics(font)
                line_height = max(fm.lineSpacing(), fm.height())
                margin_v = max(
                    int(round(float(style.subtitle_size) * 1.1)),
                    int(round(60.0 * (float(data.canvas_h) / 1080.0))),
                )
                outline_px = float(max(0, int(style.subtitle_outline))) * scale
                shadow_px = float(max(0, int(style.subtitle_shadow))) * scale
                bottom_y = (
                    target.top() + target.height() * float(style.subtitle_y_ratio)
                    if style.subtitle_y_ratio is not None
                    else target.bottom() - float(margin_v) * scale
                )
                first_baseline = bottom_y - float((len(lines) - 1) * line_height) - float(fm.descent())
                color = QtGui.QColor(style.subtitle_color or "#FFFFFF")
                outline_color = QtGui.QColor(0, 0, 0)
                shadow_color = QtGui.QColor(0, 0, 0, 150)

                for index, line in enumerate(lines):
                    width = fm.horizontalAdvance(line)
                    x = float(target.center().x()) - float(width) / 2.0
                    baseline = first_baseline + index * line_height
                    path = QtGui.QPainterPath()
                    path.addText(x, baseline, font, line)
                    if shadow_px > 0:
                        shadow_path = QtGui.QPainterPath(path)
                        shadow_path.translate(shadow_px, shadow_px)
                        painter.fillPath(shadow_path, shadow_color)
                    if outline_px > 0:
                        pen = QtGui.QPen(outline_color, outline_px, QtCore.Qt.PenStyle.SolidLine, QtCore.Qt.PenCapStyle.RoundCap, QtCore.Qt.PenJoinStyle.RoundJoin)
                        painter.strokePath(path, pen)
                    painter.fillPath(path, color)

            def _draw_placeholder(self, painter, rect, text: str) -> None:
                from PySide6 import QtCore, QtGui

                painter.save()
                painter.setPen(QtGui.QColor("#D0D0D0"))
                painter.setFont(self._status_font())
                painter.drawText(rect, int(QtCore.Qt.AlignmentFlag.AlignCenter | QtCore.Qt.TextFlag.TextWordWrap), text)
                painter.restore()

            def _header_font(self):
                from PySide6 import QtGui

                font = QtGui.QFont(self.font())
                font.setPointSize(max(9, font.pointSize()))
                return font

            def _status_font(self):
                from PySide6 import QtGui

                font = QtGui.QFont(self.font())
                font.setPointSize(max(9, font.pointSize() - 1))
                return font

            def _source_label(self, source: str) -> str:
                return {
                    "subtitle": "字幕",
                    "script": "脚本",
                    "sample": "示例",
                    "empty": "空白",
                }.get(source, source)

        return _SubtitlePreviewWidget(*args, **kwargs)
