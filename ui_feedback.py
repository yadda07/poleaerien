# -*- coding: utf-8 -*-
"""
UI Feedback Module - Visual states for processing operations
Provides animated indicators for loading, success, and error states.
"""

from qgis.PyQt.QtCore import QTimer, QPropertyAnimation, QEasingCurve, Qt, QSize
from qgis.PyQt.QtGui import QIcon, QMovie, QPixmap, QPainter, QColor
from qgis.PyQt.QtWidgets import QLabel, QPushButton, QWidget, QGraphicsOpacityEffect
from qgis.PyQt.QtSvg import QSvgRenderer
import os


class StatusIndicator(QLabel):
    """
    Animated status indicator widget.
    States: idle, loading, success, error
    """
    
    COLORS = {
        'idle': '#718096',
        'loading': '#3182ce',
        'success': '#38a169',
        'error': '#e53e3e'
    }
    
    def __init__(self, parent=None, size=24):
        super().__init__(parent)
        self.size = size
        self.setFixedSize(size, size)
        self.setAlignment(Qt.AlignCenter)
        
        self._state = 'idle'
        self._rotation = 0
        self._pulse_value = 1.0
        
        # Animation timers
        self._spin_timer = QTimer(self)
        self._spin_timer.timeout.connect(self._rotate)
        
        self._pulse_timer = QTimer(self)
        self._pulse_timer.timeout.connect(self._pulse)
        
        # Plugin dir for icons
        self.plugin_dir = os.path.dirname(__file__)
        
        self.set_idle()
    
    def set_idle(self):
        """Reset to idle state"""
        self._stop_animations()
        self._state = 'idle'
        self._draw_pole_icon()
    
    def set_loading(self, message="Traitement..."):
        """Start loading animation"""
        self._stop_animations()
        self._state = 'loading'
        self._rotation = 0
        self._spin_timer.start(50)
        self.setToolTip(message)
    
    def set_success(self, message="Terminé"):
        """Show success state with pulse animation"""
        self._stop_animations()
        self._state = 'success'
        self._draw_check_icon()
        self._pulse_value = 1.0
        self._pulse_timer.start(50)
        self.setToolTip(message)
        
        # Auto-reset after 3s
        QTimer.singleShot(3000, self.set_idle)
    
    def set_error(self, message="Erreur"):
        """Show error state"""
        self._stop_animations()
        self._state = 'error'
        self._draw_error_icon()
        self.setToolTip(message)
        
        # Auto-reset after 5s
        QTimer.singleShot(5000, self.set_idle)
    
    def _stop_animations(self):
        self._spin_timer.stop()
        self._pulse_timer.stop()
    
    def _rotate(self):
        """Rotation animation for loading"""
        self._rotation = (self._rotation + 15) % 360
        self._draw_loading_icon()
    
    def _pulse(self):
        """Pulse animation for success"""
        self._pulse_value -= 0.02
        if self._pulse_value <= 0.6:
            self._pulse_value = 1.0
        self._draw_check_icon()
    
    def _draw_pole_icon(self):
        """Draw electric pole icon (idle state)"""
        pixmap = QPixmap(self.size, self.size)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        
        color = QColor(self.COLORS['idle'])
        painter.setPen(Qt.NoPen)
        painter.setBrush(color)
        
        # Pole vertical
        s = self.size
        painter.drawRect(int(s*0.42), int(s*0.15), int(s*0.16), int(s*0.75))
        
        # Crossarms
        painter.drawRect(int(s*0.15), int(s*0.2), int(s*0.7), int(s*0.08))
        painter.drawRect(int(s*0.25), int(s*0.4), int(s*0.5), int(s*0.06))
        
        # Insulators
        painter.setBrush(QColor('#3182ce'))
        painter.drawEllipse(int(s*0.18), int(s*0.18), int(s*0.1), int(s*0.1))
        painter.drawEllipse(int(s*0.45), int(s*0.18), int(s*0.1), int(s*0.1))
        painter.drawEllipse(int(s*0.72), int(s*0.18), int(s*0.1), int(s*0.1))
        
        painter.end()
        self.setPixmap(pixmap)
    
    def _draw_loading_icon(self):
        """Draw spinning loading indicator"""
        pixmap = QPixmap(self.size, self.size)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        
        painter.translate(self.size/2, self.size/2)
        painter.rotate(self._rotation)
        painter.translate(-self.size/2, -self.size/2)
        
        color = QColor(self.COLORS['loading'])
        from qgis.PyQt.QtGui import QPen
        pen = QPen(color)
        pen.setWidth(3)
        pen.setCapStyle(Qt.RoundCap)
        painter.setPen(pen)
        painter.setBrush(Qt.NoBrush)
        
        # Arc spinner
        s = self.size
        margin = int(s * 0.15)
        rect = pixmap.rect().adjusted(margin, margin, -margin, -margin)
        painter.drawArc(rect, 0, 270 * 16)
        
        painter.end()
        self.setPixmap(pixmap)
    
    def _draw_check_icon(self):
        """Draw success checkmark"""
        pixmap = QPixmap(self.size, self.size)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        
        color = QColor(self.COLORS['success'])
        color.setAlphaF(self._pulse_value)
        
        from qgis.PyQt.QtGui import QPen
        
        # Circle
        pen = QPen(color)
        pen.setWidth(2)
        painter.setPen(pen)
        painter.setBrush(Qt.NoBrush)
        s = self.size
        margin = int(s * 0.1)
        painter.drawEllipse(margin, margin, s - 2*margin, s - 2*margin)
        
        # Checkmark
        pen.setWidth(3)
        pen.setCapStyle(Qt.RoundCap)
        pen.setJoinStyle(Qt.RoundJoin)
        painter.setPen(pen)
        
        from qgis.PyQt.QtCore import QPointF
        from qgis.PyQt.QtGui import QPainterPath
        path = QPainterPath()
        path.moveTo(s*0.25, s*0.5)
        path.lineTo(s*0.42, s*0.67)
        path.lineTo(s*0.75, s*0.33)
        painter.drawPath(path)
        
        painter.end()
        self.setPixmap(pixmap)
    
    def _draw_error_icon(self):
        """Draw error X icon"""
        pixmap = QPixmap(self.size, self.size)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        
        color = QColor(self.COLORS['error'])
        from qgis.PyQt.QtGui import QPen
        
        # Circle
        pen = QPen(color)
        pen.setWidth(2)
        painter.setPen(pen)
        painter.setBrush(Qt.NoBrush)
        s = self.size
        margin = int(s * 0.1)
        painter.drawEllipse(margin, margin, s - 2*margin, s - 2*margin)
        
        # X mark
        pen.setWidth(3)
        pen.setCapStyle(Qt.RoundCap)
        painter.setPen(pen)
        painter.drawLine(int(s*0.3), int(s*0.3), int(s*0.7), int(s*0.7))
        painter.drawLine(int(s*0.7), int(s*0.3), int(s*0.3), int(s*0.7))
        
        painter.end()
        self.setPixmap(pixmap)


class AnimatedButton(QPushButton):
    """
    Button with visual feedback during operations.
    Changes appearance based on state: ready, processing, done, error
    """
    
    STYLES = {
        'ready': """
            QPushButton {{
                background-color: {bg};
                color: {fg};
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: {hover};
            }}
            QPushButton:pressed {{
                background-color: {pressed};
            }}
            QPushButton:disabled {{
                background-color: #a0aec0;
                color: #718096;
            }}
        """,
        'processing': """
            QPushButton {
                background-color: #4299e1;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
            }
        """,
        'success': """
            QPushButton {
                background-color: #48bb78;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
            }
        """,
        'error': """
            QPushButton {
                background-color: #fc8181;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
            }
        """
    }
    
    def __init__(self, text="", parent=None, bg_color="#3182ce", fg_color="#ffffff"):
        super().__init__(text, parent)
        self._original_text = text
        self._bg_color = bg_color
        self._fg_color = fg_color
        self._dots = 0
        
        self._anim_timer = QTimer(self)
        self._anim_timer.timeout.connect(self._animate_dots)
        
        self.set_ready()
    
    def set_colors(self, bg_color, fg_color="#ffffff"):
        """Set button colors"""
        self._bg_color = bg_color
        self._fg_color = fg_color
        self.set_ready()
    
    def set_ready(self, text=None):
        """Reset to ready state"""
        self._anim_timer.stop()
        if text:
            self._original_text = text
        self.setText(self._original_text)
        self.setEnabled(True)
        
        # Compute hover/pressed colors
        base = QColor(self._bg_color)
        hover = base.lighter(115).name()
        pressed = base.darker(110).name()
        
        style = self.STYLES['ready'].format(
            bg=self._bg_color,
            fg=self._fg_color,
            hover=hover,
            pressed=pressed
        )
        self.setStyleSheet(style)
    
    def set_processing(self, text="Traitement"):
        """Start processing animation"""
        self._original_text = self.text() if not self._anim_timer.isActive() else self._original_text
        self._processing_text = text
        self._dots = 0
        self.setEnabled(False)
        self.setStyleSheet(self.STYLES['processing'])
        self._anim_timer.start(400)
    
    def set_success(self, text="Terminé", auto_reset=3000):
        """Show success state"""
        self._anim_timer.stop()
        self.setText(text)
        self.setEnabled(True)
        self.setStyleSheet(self.STYLES['success'])
        
        if auto_reset > 0:
            QTimer.singleShot(auto_reset, self.set_ready)
    
    def set_error(self, text="Erreur", auto_reset=5000):
        """Show error state"""
        self._anim_timer.stop()
        self.setText(text)
        self.setEnabled(True)
        self.setStyleSheet(self.STYLES['error'])
        
        if auto_reset > 0:
            QTimer.singleShot(auto_reset, self.set_ready)
    
    def _animate_dots(self):
        """Animate processing dots"""
        self._dots = (self._dots + 1) % 4
        dots = "." * self._dots
        self.setText(f"{self._processing_text}{dots}")


class ProcessingOverlay(QWidget):
    """
    Semi-transparent overlay with loading indicator.
    Use over a widget during long operations.
    """
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, False)
        self.setStyleSheet("background-color: rgba(255, 255, 255, 0.85);")
        
        from qgis.PyQt.QtWidgets import QVBoxLayout
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        
        self._indicator = StatusIndicator(self, size=48)
        self._label = QLabel("Traitement en cours...", self)
        self._label.setAlignment(Qt.AlignCenter)
        self._label.setStyleSheet("color: #2d3748; font-size: 14px; font-weight: bold;")
        
        layout.addWidget(self._indicator)
        layout.addWidget(self._label)
        
        self.hide()
    
    def show_loading(self, message="Traitement en cours..."):
        """Show overlay with loading animation"""
        if self.parent():
            self.setGeometry(self.parent().rect())
        self._label.setText(message)
        self._indicator.set_loading(message)
        self.show()
        self.raise_()
    
    def show_success(self, message="Opération terminée"):
        """Show success state briefly"""
        self._label.setText(message)
        self._indicator.set_success(message)
        QTimer.singleShot(2000, self.hide_overlay)
    
    def show_error(self, message="Une erreur est survenue"):
        """Show error state"""
        self._label.setText(message)
        self._indicator.set_error(message)
        QTimer.singleShot(3000, self.hide_overlay)
    
    def hide_overlay(self):
        """Hide overlay"""
        self._indicator.set_idle()
        self.hide()


def create_status_indicator(parent=None, size=24):
    """Factory function for StatusIndicator"""
    return StatusIndicator(parent, size)


def create_animated_button(text, parent=None, bg_color="#3182ce", fg_color="#ffffff"):
    """Factory function for AnimatedButton"""
    btn = AnimatedButton(text, parent, bg_color, fg_color)
    return btn
