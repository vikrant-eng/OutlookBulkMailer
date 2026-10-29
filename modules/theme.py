"""
Theme / style definitions for Outlook Bulk Mailer.
Provides CSS-like QSS strings for light and dark modes.
"""

FONT_FAMILY = "Segoe UI"

# ── Colour palette ─────────────────────────────────────────────────────────
DARK = {
    "bg_primary":    "#0F1923",
    "bg_secondary":  "#1A2535",
    "bg_surface":    "#1E2D3D",
    "bg_hover":      "#243447",
    "accent":        "#0078D4",
    "accent_hover":  "#106EBE",
    "accent_light":  "#C7E0F4",
    "text_primary":  "#FFFFFF",
    "text_secondary":"#A0AEB8",
    "text_muted":    "#637080",
    "border":        "#2A3F55",
    "success":       "#107C10",
    "success_text":  "#BAD80A",
    "warning":       "#FFB900",
    "danger":        "#D83B01",
    "danger_text":   "#F4694C",
    "sidebar_bg":    "#0B1520",
    "sidebar_text":  "#A0AEB8",
    "sidebar_active":"#0078D4",
}

LIGHT = {
    "bg_primary":    "#F3F6FB",
    "bg_secondary":  "#FFFFFF",
    "bg_surface":    "#EDF2F8",
    "bg_hover":      "#E0EAF5",
    "accent":        "#0078D4",
    "accent_hover":  "#106EBE",
    "accent_light":  "#EBF3FB",
    "text_primary":  "#1B1B1B",
    "text_secondary":"#444E5A",
    "text_muted":    "#8A949E",
    "border":        "#D0DAE8",
    "success":       "#107C10",
    "success_text":  "#107C10",
    "warning":       "#986F0B",
    "danger":        "#C50F1F",
    "danger_text":   "#C50F1F",
    "sidebar_bg":    "#0F1923",
    "sidebar_text":  "#A0AEB8",
    "sidebar_active":"#0078D4",
}


def get_stylesheet(dark: bool = True) -> str:
    c = DARK if dark else LIGHT

    return f"""
/* ── Global ───────────────────────────────── */
QWidget {{
    font-family: "{FONT_FAMILY}", "Helvetica Neue", Arial, sans-serif;
    font-size: 13px;
    color: {c['text_primary']};
    background-color: {c['bg_primary']};
    outline: none;
}}

QMainWindow {{
    background-color: {c['bg_primary']};
}}

/* ── Sidebar ───────────────────────────────── */
#sidebar {{
    background-color: {c['sidebar_bg']};
    border-right: 1px solid {c['border']};
    min-width: 220px;
    max-width: 220px;
}}

#sidebar_logo_label {{
    color: #FFFFFF;
    font-size: 17px;
    font-weight: 700;
    padding: 0 16px;
}}

#sidebar_sub_label {{
    color: {c['sidebar_text']};
    font-size: 11px;
    padding: 0 16px;
}}

QPushButton#nav_button {{
    background: transparent;
    color: {c['sidebar_text']};
    border: none;
    text-align: left;
    padding: 11px 20px;
    font-size: 13px;
    border-radius: 0;
}}
QPushButton#nav_button:hover {{
    background-color: {c['bg_hover']};
    color: {c['text_primary']};
}}
QPushButton#nav_button[active="true"] {{
    background-color: {c['sidebar_active']};
    color: #FFFFFF;
    font-weight: 600;
    border-left: 3px solid #82C7FF;
}}

/* ── Content area ──────────────────────────── */
#content_area {{
    background-color: {c['bg_primary']};
}}

#page_title {{
    font-size: 22px;
    font-weight: 700;
    color: {c['text_primary']};
}}
#page_subtitle {{
    font-size: 13px;
    color: {c['text_secondary']};
}}

/* ── Cards ─────────────────────────────────── */
QFrame#card {{
    background-color: {c['bg_secondary']};
    border: 1px solid {c['border']};
    border-radius: 10px;
    padding: 16px;
}}

/* ── Buttons ────────────────────────────────── */
QPushButton {{
    background-color: {c['accent']};
    color: #FFFFFF;
    border: none;
    border-radius: 6px;
    padding: 8px 18px;
    font-weight: 600;
    font-size: 13px;
}}
QPushButton:hover  {{ background-color: {c['accent_hover']}; }}
QPushButton:pressed {{ background-color: #004E9A; }}
QPushButton:disabled {{ background-color: {c['border']}; color: {c['text_muted']}; }}

QPushButton#btn_secondary {{
    background-color: transparent;
    color: {c['accent']};
    border: 1.5px solid {c['accent']};
}}
QPushButton#btn_secondary:hover {{
    background-color: {c['accent_light']};
}}

QPushButton#btn_danger {{
    background-color: {c['danger']};
}}
QPushButton#btn_danger:hover {{
    background-color: #E74C3C;
}}

QPushButton#btn_success {{
    background-color: {c['success']};
}}

/* ── Inputs ─────────────────────────────────── */
QLineEdit, QTextEdit, QPlainTextEdit {{
    background-color: {c['bg_surface']};
    color: {c['text_primary']};
    border: 1.5px solid {c['border']};
    border-radius: 6px;
    padding: 8px 10px;
    selection-background-color: {c['accent']};
}}
QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus {{
    border: 1.5px solid {c['accent']};
}}

QComboBox {{
    background-color: {c['bg_surface']};
    color: {c['text_primary']};
    border: 1.5px solid {c['border']};
    border-radius: 6px;
    padding: 6px 10px;
}}
QComboBox::drop-down {{ border: none; width: 24px; }}
QComboBox QAbstractItemView {{
    background-color: {c['bg_secondary']};
    color: {c['text_primary']};
    border: 1px solid {c['border']};
    selection-background-color: {c['accent']};
}}

QSpinBox {{
    background-color: {c['bg_surface']};
    color: {c['text_primary']};
    border: 1.5px solid {c['border']};
    border-radius: 6px;
    padding: 6px 10px;
}}

/* ── Tables ─────────────────────────────────── */
QTableWidget {{
    background-color: {c['bg_secondary']};
    color: {c['text_primary']};
    gridline-color: {c['border']};
    border: 1px solid {c['border']};
    border-radius: 8px;
    alternate-background-color: {c['bg_surface']};
}}
QTableWidget::item {{
    padding: 6px 8px;
}}
QTableWidget::item:selected {{
    background-color: {c['accent']};
    color: #FFFFFF;
}}
QHeaderView::section {{
    background-color: {c['bg_surface']};
    color: {c['text_secondary']};
    font-weight: 600;
    font-size: 12px;
    border: none;
    border-bottom: 1px solid {c['border']};
    padding: 8px;
}}

/* ── Progress bar ───────────────────────────── */
QProgressBar {{
    background-color: {c['bg_surface']};
    border: 1px solid {c['border']};
    border-radius: 6px;
    height: 12px;
    text-align: center;
    color: {c['text_primary']};
    font-size: 11px;
}}
QProgressBar::chunk {{
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 {c['accent']}, stop:1 #50C0FF);
    border-radius: 6px;
}}

/* ── Scroll bars ────────────────────────────── */
QScrollBar:vertical {{
    background: {c['bg_primary']};
    width: 8px;
    margin: 0;
}}
QScrollBar::handle:vertical {{
    background: {c['border']};
    border-radius: 4px;
    min-height: 24px;
}}
QScrollBar::handle:vertical:hover {{
    background: {c['text_muted']};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QScrollBar:horizontal {{
    background: {c['bg_primary']};
    height: 8px;
}}
QScrollBar::handle:horizontal {{
    background: {c['border']};
    border-radius: 4px;
}}

/* ── Labels ─────────────────────────────────── */
QLabel {{ background: transparent; }}
QLabel#label_success {{
    color: {c['success_text']};
    font-weight: 600;
}}
QLabel#label_warning {{
    color: {c['warning']};
    font-weight: 600;
}}
QLabel#label_danger {{
    color: {c['danger_text']};
    font-weight: 600;
}}
QLabel#label_muted {{
    color: {c['text_muted']};
    font-size: 12px;
}}

/* ── Status bar ─────────────────────────────── */
QStatusBar {{
    background-color: {c['bg_secondary']};
    color: {c['text_secondary']};
    border-top: 1px solid {c['border']};
    font-size: 12px;
}}

/* ── Separators ─────────────────────────────── */
QFrame[frameShape="4"] {{ color: {c['border']}; }}

/* ── Tooltips ───────────────────────────────── */
QToolTip {{
    background-color: {c['bg_surface']};
    color: {c['text_primary']};
    border: 1px solid {c['border']};
    padding: 4px 8px;
    border-radius: 4px;
}}

/* ── Dialog ─────────────────────────────────── */
QDialog {{
    background-color: {c['bg_primary']};
}}
QMessageBox {{
    background-color: {c['bg_primary']};
}}

/* ── Tab bar ─────────────────────────────────── */
QTabBar::tab {{
    background: {c['bg_surface']};
    color: {c['text_secondary']};
    padding: 8px 18px;
    border: 1px solid {c['border']};
    border-bottom: none;
    border-radius: 6px 6px 0 0;
}}
QTabBar::tab:selected {{
    background: {c['accent']};
    color: #FFFFFF;
    font-weight: 600;
}}
QTabWidget::pane {{
    border: 1px solid {c['border']};
    border-radius: 0 6px 6px 6px;
    background: {c['bg_secondary']};
}}

/* ── Checkbox ───────────────────────────────── */
QCheckBox {{
    spacing: 8px;
    color: {c['text_primary']};
}}
QCheckBox::indicator {{
    width: 16px;
    height: 16px;
    border: 1.5px solid {c['border']};
    border-radius: 3px;
    background: {c['bg_surface']};
}}
QCheckBox::indicator:checked {{
    background: {c['accent']};
    border-color: {c['accent']};
}}
"""
