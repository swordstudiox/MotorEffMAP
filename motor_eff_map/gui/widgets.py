import logging

from PySide6.QtCore import QObject, Qt, Signal
from PySide6.QtWidgets import QHBoxLayout, QLabel, QWidget

from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg


class _LogEmitter(QObject):
    message = Signal(str)


class QTextEditLogger(logging.Handler):
    def __init__(self, widget):
        super().__init__()
        self.widget = widget
        self.emitter = _LogEmitter()
        self.emitter.message.connect(self.widget.append)

    def emit(self, record):
        if self.widget is None:
            return
        msg = self.format(record)
        self.emitter.message.emit(msg)
        if record.levelno >= logging.ERROR:
            pass

    def close(self):
        try:
            self.emitter.message.disconnect()
        except (TypeError, RuntimeError):
            pass
        self.widget = None
        super().close()


class AspectRatioWidget(QWidget):
    """
    一个包含 FigureCanvas 并保持固定长宽比（居中）的小部件，
    不管容器大小如何变化。
    """

    def __init__(self, figure, aspect_ratio=1.25, parent=None):
        super().__init__(parent)
        self.aspect_ratio = aspect_ratio

        self.canvas = FigureCanvasQTAgg(figure)
        self.canvas.setParent(self)

        from PySide6.QtWidgets import QSizePolicy

        self.canvas.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)

    def resizeEvent(self, event):
        w = event.size().width()
        h = event.size().height()

        if w <= 0 or h <= 0:
            return

        if w / h > self.aspect_ratio:
            target_h = h
            target_w = int(h * self.aspect_ratio)
        else:
            target_w = w
            target_h = int(w / self.aspect_ratio)

        x = (w - target_w) // 2
        y = (h - target_h) // 2

        self.canvas.setGeometry(x, y, target_w, target_h)


class SignatureWidget(QWidget):
    def __init__(self, author="sword", data="2026-01-15"):
        super().__init__()
        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(10, 0, 10, 0)

        self.repo_label = QLabel(
            '<a href="https://github.com/swordstudiox/MotorEffMAP">'
            '项目地址：swordstudiox/MotorEffMAP</a>'
        )
        self.repo_label.setOpenExternalLinks(True)
        self.repo_label.setTextInteractionFlags(Qt.TextBrowserInteraction)
        self.repo_label.setStyleSheet("font-size: 9px; color: #3366cc;")
        self.layout.addWidget(self.repo_label)
        self.layout.addStretch()

        self.lbl = QLabel(f"Author: {author} | Date: {data}")
        self.lbl.setStyleSheet(
            "background-color: #EEE; border: 1px solid #CCC; "
            "padding: 1px; font-size: 9px; color: #333;"
        )
        self.lbl.hide()

        self.layout.addWidget(self.lbl)

        self.setMouseTracking(True)
        self.setFixedHeight(20)

    def enterEvent(self, event):
        self.lbl.show()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.lbl.hide()
        super().leaveEvent(event)
