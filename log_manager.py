# -*- coding: utf-8 -*-
"""
Log Manager - Structured logging with color coding for QGIS plugin.
Provides real-time feedback to users with visual hierarchy.
"""

from qgis.PyQt.QtWidgets import QTextBrowser, QProgressBar, QLabel
from qgis.PyQt.QtCore import QObject, pyqtSignal
from qgis.PyQt.QtGui import QTextCursor
from datetime import datetime


class LogLevel:
    """Log level constants with colors and icons."""
    INFO = ('info', '#3b82f6', 'ℹ')      # Blue
    SUCCESS = ('success', '#22c55e', '✓')  # Green
    WARNING = ('warning', '#f59e0b', '⚠')  # Orange
    ERROR = ('error', '#ef4444', '✗')      # Red
    SECTION = ('section', '#6366f1', '━')  # Purple


class LogManager(QObject):
    """Centralized log manager with color-coded messages."""
    
    # Signals for thread-safe updates
    log_added = pyqtSignal(str, str, str)  # level, icon, message
    kpi_updated = pyqtSignal(str, str, str)  # key, value, style
    progress_updated = pyqtSignal(int, str)  # percent, step_text
    
    def __init__(self, text_browser: QTextBrowser = None, 
                 progress_bar: QProgressBar = None,
                 progress_label: QLabel = None):
        super().__init__()
        self._browser = text_browser
        self._progress = progress_bar
        self._progress_label = progress_label
        self._kpi_widgets = {}
        
        # Connect signals
        self.log_added.connect(self._write_log)
        self.progress_updated.connect(self._update_progress)
        self.kpi_updated.connect(self._update_kpi_widget)
    
    def set_browser(self, browser: QTextBrowser):
        """Set the text browser widget."""
        self._browser = browser
    
    def set_progress(self, progress: QProgressBar, label: QLabel = None):
        """Set progress bar and optional label."""
        self._progress = progress
        self._progress_label = label
    
    def set_kpi_widgets(self, widgets: dict):
        """Set KPI label widgets for real-time updates."""
        self._kpi_widgets = widgets
    
    def clear(self):
        """Clear log content."""
        if self._browser:
            self._browser.clear()
        if self._progress:
            self._progress.setValue(0)
        if self._progress_label:
            self._progress_label.setText('')
    
    # Log methods
    def info(self, msg: str):
        """Info message (blue)."""
        self.log_added.emit(*LogLevel.INFO, msg)
    
    def success(self, msg: str):
        """Success message (green)."""
        self.log_added.emit(*LogLevel.SUCCESS, msg)
    
    def warning(self, msg: str):
        """Warning message (orange)."""
        self.log_added.emit(*LogLevel.WARNING, msg)
    
    def error(self, msg: str):
        """Error message (red)."""
        self.log_added.emit(*LogLevel.ERROR, msg)
    
    def section(self, title: str):
        """Section separator with title."""
        sep = '━' * 20
        self._write_section(title, sep)
    
    def progress(self, step: str, percent: int):
        """Update progress with step description."""
        self.progress_updated.emit(percent, step)
        self.info(f"{step} ({percent}%)")
    
    def update_kpi(self, key: str, value: str, style: str = 'kpi_default'):
        """Update KPI badge value."""
        self.kpi_updated.emit(key, value, style)
    
    def result(self, label: str, value, style: str = 'info'):
        """Log a result with label and value."""
        color = {
            'info': '#3b82f6',
            'success': '#22c55e', 
            'warning': '#f59e0b',
            'error': '#ef4444'
        }.get(style, '#3b82f6')
        
        html = f'<span style="color:#64748b">{label}:</span> <b style="color:{color}">{value}</b>'
        self._append_html(html)
    
    def table_start(self, headers: list):
        """Start a results table."""
        cols = ''.join(f'<th style="padding:4px 8px;background:#f1f5f9;border:1px solid #e2e8f0;text-align:left">{h}</th>' for h in headers)
        html = f'<table style="border-collapse:collapse;margin:8px 0;width:100%"><tr>{cols}</tr>'
        self._append_html(html)
    
    def table_row(self, cells: list, row_style: str = 'normal'):
        """Add a table row."""
        bg = {'normal': '#fff', 'ok': '#dcfce7', 'warn': '#fef3c7', 'error': '#fee2e2'}.get(row_style, '#fff')
        cols = ''.join(f'<td style="padding:4px 8px;border:1px solid #e2e8f0;background:{bg}">{c}</td>' for c in cells)
        html = f'<tr>{cols}</tr>'
        self._append_html(html)
    
    def table_end(self):
        """End a results table."""
        self._append_html('</table>')
    
    # Internal methods
    def _write_log(self, level: str, color: str, icon: str, msg: str):
        """Write formatted log entry.
        
        Args:
            level: Log level name (for filtering/styling)
            color: Hex color code
            icon: Icon character
            msg: Message text
        """
        _ = level  # Reserved for future filtering
        time_str = datetime.now().strftime('%d/%m %H:%M:%S')
        html = f'<p style="margin:0;padding:0;line-height:1.2">'
        html += f'<span style="color:#94a3b8;font-size:9pt">{time_str}</span> '
        html += f'<span style="color:{color};font-weight:bold">{icon}</span> '
        html += f'<span style="color:{color}">{msg}</span></p>'
        self._append_html(html)
    
    def _write_section(self, title: str, sep: str):
        """Write section header."""
        html = f'<p style="margin:4px 0;padding:0;line-height:1.2"><span style="color:#6366f1;font-weight:bold">{sep} {title.upper()} {sep}</span></p>'
        self._append_html(html)
    
    def _append_html(self, html: str):
        """Append HTML to browser."""
        if self._browser:
            self._browser.append(html)
            # Auto-scroll
            cursor = self._browser.textCursor()
            cursor.movePosition(QTextCursor.End)
            self._browser.setTextCursor(cursor)
    
    def _update_progress(self, percent: int, step: str):
        """Update progress bar and label."""
        if self._progress:
            self._progress.setValue(percent)
        if self._progress_label:
            self._progress_label.setText(step)
    
    def _update_kpi_widget(self, key: str, value: str, style: str):
        """Update KPI widget."""
        if key in self._kpi_widgets:
            widget = self._kpi_widgets[key]
            widget.setText(value)
            # Apply style from STYLES dict
            from .ui_pages import STYLES
            if style in STYLES:
                widget.setStyleSheet(STYLES[style])


class LogManagerRegistry:
    """Registry for LogManager instances per plugin dialog.
    
    Avoids global singleton, supports multiple plugin instances and testing.
    """
    _instances: dict = {}
    
    @classmethod
    def get(cls, key: str = 'default') -> LogManager:
        """Get LogManager for key, create if needed."""
        if key not in cls._instances:
            cls._instances[key] = LogManager()
        return cls._instances[key]
    
    @classmethod
    def register(cls, key: str, manager: LogManager):
        """Register a LogManager instance."""
        cls._instances[key] = manager
    
    @classmethod
    def unregister(cls, key: str):
        """Remove LogManager instance (cleanup)."""
        if key in cls._instances:
            del cls._instances[key]
    
    @classmethod
    def clear(cls):
        """Clear all instances (for testing)."""
        cls._instances.clear()


def get_log_manager(key: str = 'default') -> LogManager:
    """Get LogManager instance by key.
    
    Args:
        key: Instance identifier (default: 'default')
    
    Returns:
        LogManager instance
    """
    return LogManagerRegistry.get(key)


def init_log_manager(browser: QTextBrowser, progress: QProgressBar = None, 
                     progress_label: QLabel = None, kpis: dict = None,
                     key: str = 'default') -> LogManager:
    """Initialize and register LogManager with widgets.
    
    Args:
        browser: QTextBrowser for log output
        progress: Optional QProgressBar
        progress_label: Optional progress label
        kpis: Optional KPI widgets dict
        key: Instance identifier
    
    Returns:
        Configured LogManager instance
    """
    manager = LogManager(browser, progress, progress_label)
    if kpis:
        manager.set_kpi_widgets(kpis)
    LogManagerRegistry.register(key, manager)
    return manager


def cleanup_log_manager(key: str = 'default'):
    """Cleanup LogManager instance.
    
    Args:
        key: Instance identifier to cleanup
    """
    LogManagerRegistry.unregister(key)
