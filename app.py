"""
Outlook Bulk Mailer — Main Application
PySide6 desktop app that sends personalised bulk emails via Microsoft Graph API.
"""

import sys
import os
import webbrowser

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QTextEdit, QFileDialog, QTableWidget,
    QTableWidgetItem, QProgressBar, QFrame, QStackedWidget, QDialog,
    QDialogButtonBox, QMessageBox, QHeaderView, QSizePolicy, QScrollArea,
    QSpinBox, QCheckBox, QComboBox, QStatusBar, QAbstractItemView,
    QGroupBox, QFormLayout, QSplitter,
)
from PySide6.QtCore import (
    Qt, QSize, QTimer, Signal, QThread, QPropertyAnimation,
    QEasingCurve, QRect, QPoint,
)
from PySide6.QtGui import (
    QFont, QIcon, QColor, QPalette, QPixmap, QPainter, QPen,
    QBrush, QLinearGradient, QAction,
)

# Project modules
from modules import database, auth, excel_reader, exporter
from modules.workers import SendWorker, AuthWorker
from modules.theme import get_stylesheet


# ──────────────────────────────────────────────────────────────────────────────
# Helper widgets
# ──────────────────────────────────────────────────────────────────────────────

class StatCard(QFrame):
    """A small card showing a metric."""

    def __init__(self, title: str, value: str = "0", icon: str = "📧",
                 color: str = "#0078D4", parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setFixedHeight(110)

        layout = QVBoxLayout(self)
        layout.setSpacing(6)
        layout.setContentsMargins(16, 14, 16, 14)

        top = QHBoxLayout()
        icon_lbl = QLabel(icon)
        icon_lbl.setFont(QFont("Segoe UI Emoji", 22))
        top.addWidget(icon_lbl)
        top.addStretch()

        self.value_lbl = QLabel(value)
        self.value_lbl.setFont(QFont("Segoe UI", 26, QFont.Bold))
        self.value_lbl.setStyleSheet(f"color: {color};")

        title_lbl = QLabel(title)
        title_lbl.setObjectName("label_muted")

        layout.addLayout(top)
        layout.addWidget(self.value_lbl)
        layout.addWidget(title_lbl)

    def set_value(self, v):
        self.value_lbl.setText(str(v))


class SectionTitle(QLabel):
    def __init__(self, text: str, parent=None):
        super().__init__(text, parent)
        self.setFont(QFont("Segoe UI", 14, QFont.Bold))
        self.setObjectName("page_title")
        self.setContentsMargins(0, 0, 0, 4)


class Divider(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.HLine)
        self.setFrameShadow(QFrame.Sunken)


def make_button(text: str, obj_name: str = "", icon_text: str = "") -> QPushButton:
    btn = QPushButton(f"{icon_text}  {text}" if icon_text else text)
    if obj_name:
        btn.setObjectName(obj_name)
    btn.setCursor(Qt.PointingHandCursor)
    return btn


# ──────────────────────────────────────────────────────────────────────────────
# Page: Dashboard
# ──────────────────────────────────────────────────────────────────────────────

class DashboardPage(QWidget):
    go_to_compose = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(24)
        layout.setContentsMargins(32, 28, 32, 28)

        # Title
        title = QLabel("Dashboard")
        title.setObjectName("page_title")
        title.setFont(QFont("Segoe UI", 22, QFont.Bold))
        layout.addWidget(title)

        subtitle = QLabel("Overview of your email campaigns")
        subtitle.setObjectName("label_muted")
        layout.addWidget(subtitle)

        layout.addWidget(Divider())

        # Stat cards
        cards_row = QHBoxLayout()
        cards_row.setSpacing(14)
        self.card_total = StatCard("Total Campaigns", "0", "📁", "#0078D4")
        self.card_sent = StatCard("Emails Sent", "0", "✅", "#107C10")
        self.card_failed = StatCard("Failed", "0", "❌", "#D83B01")
        self.card_pending = StatCard("Pending", "0", "⏳", "#FFB900")
        for c in [self.card_total, self.card_sent, self.card_failed, self.card_pending]:
            cards_row.addWidget(c)
        layout.addLayout(cards_row)

        # Quick actions
        qa_label = SectionTitle("Quick Actions")
        layout.addWidget(qa_label)

        qa_row = QHBoxLayout()
        qa_row.setSpacing(12)
        btn_new = make_button("New Campaign", icon_text="✉️")
        btn_new.setFixedHeight(44)
        btn_new.clicked.connect(self.go_to_compose)
        btn_history = make_button("View History", "btn_secondary", "📋")
        btn_history.setFixedHeight(44)
        qa_row.addWidget(btn_new)
        qa_row.addWidget(btn_history)
        qa_row.addStretch()
        layout.addLayout(qa_row)

        # Recent campaigns table
        rec_label = SectionTitle("Recent Campaigns")
        layout.addWidget(rec_label)

        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["Name", "Subject", "Status", "Sent", "Date"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.verticalHeader().hide()
        layout.addWidget(self.table)

        layout.addStretch()

    def refresh(self):
        campaigns = database.get_campaigns()
        total = len(campaigns)
        sent = sum(c["sent"] for c in campaigns)
        failed = sum(c["failed"] for c in campaigns)
        pending = sum(max(c["total"] - c["sent"] - c["failed"], 0) for c in campaigns)

        self.card_total.set_value(total)
        self.card_sent.set_value(sent)
        self.card_failed.set_value(failed)
        self.card_pending.set_value(pending)

        self.table.setRowCount(0)
        for c in campaigns[:20]:
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(c["name"]))
            self.table.setItem(row, 1, QTableWidgetItem(c["subject"]))

            status_item = QTableWidgetItem(c["status"].capitalize())
            color_map = {
                "completed": QColor("#107C10"),
                "failed":    QColor("#D83B01"),
                "pending":   QColor("#FFB900"),
                "paused":    QColor("#0078D4"),
                "draft":     QColor("#637080"),
            }
            fg = color_map.get(c["status"], QColor("#FFFFFF"))
            status_item.setForeground(QBrush(fg))
            self.table.setItem(row, 2, status_item)
            self.table.setItem(row, 3, QTableWidgetItem(f"{c['sent']}/{c['total']}"))
            self.table.setItem(row, 4, QTableWidgetItem((c["created_at"] or "")[:10]))


# ──────────────────────────────────────────────────────────────────────────────
# Page: Compose
# ──────────────────────────────────────────────────────────────────────────────

class ComposePage(QWidget):
    campaign_ready = Signal(dict)   # emitted when user clicks Send

    def __init__(self, parent=None):
        super().__init__(parent)
        self.recipients: list[dict] = []
        self.attachments: list[dict] = []
        self._build_ui()

    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setSpacing(0)
        outer.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        outer.addWidget(scroll)

        container = QWidget()
        scroll.setWidget(container)

        layout = QVBoxLayout(container)
        layout.setSpacing(20)
        layout.setContentsMargins(32, 28, 32, 28)

        title = QLabel("New Campaign")
        title.setObjectName("page_title")
        title.setFont(QFont("Segoe UI", 22, QFont.Bold))
        layout.addWidget(title)
        subtitle = QLabel("Compose your email and import recipients")
        subtitle.setObjectName("label_muted")
        layout.addWidget(subtitle)
        layout.addWidget(Divider())

        # ── Step 1: Campaign name & subject ──────────────────────────
        step1 = QFrame()
        step1.setObjectName("card")
        s1l = QVBoxLayout(step1)
        s1l.setSpacing(10)

        s1l.addWidget(SectionTitle("1. Campaign Details"))

        form1 = QFormLayout()
        form1.setSpacing(10)
        form1.setLabelAlignment(Qt.AlignRight)

        self.campaign_name = QLineEdit()
        self.campaign_name.setPlaceholderText("e.g. October Newsletter")
        form1.addRow("Campaign Name:", self.campaign_name)

        self.subject_edit = QLineEdit()
        self.subject_edit.setPlaceholderText("Email subject — use {{Name}} to personalise")
        form1.addRow("Subject:", self.subject_edit)

        s1l.addLayout(form1)
        layout.addWidget(step1)

        # ── Step 2: Message body ─────────────────────────────────────
        step2 = QFrame()
        step2.setObjectName("card")
        s2l = QVBoxLayout(step2)
        s2l.setSpacing(10)

        s2l.addWidget(SectionTitle("2. Message Body"))

        hint = QLabel("Use <b>{{Name}}</b> anywhere to insert the recipient's name. HTML is supported.")
        hint.setObjectName("label_muted")
        hint.setWordWrap(True)
        s2l.addWidget(hint)

        self.body_edit = QTextEdit()
        self.body_edit.setPlaceholderText(
            "Hi {{Name}},\n\nYour message here...\n\nBest regards,\nYour Name"
        )
        self.body_edit.setMinimumHeight(200)
        s2l.addWidget(self.body_edit)

        layout.addWidget(step2)

        # ── Step 3: Attachments ──────────────────────────────────────
        step3 = QFrame()
        step3.setObjectName("card")
        s3l = QVBoxLayout(step3)
        s3l.setSpacing(10)

        s3l.addWidget(SectionTitle("3. Attachments (Optional)"))

        att_row = QHBoxLayout()
        btn_add_att = make_button("Add Files", "btn_secondary", "📎")
        btn_add_att.clicked.connect(self._add_attachments)
        btn_clear_att = make_button("Clear All", "btn_danger")
        btn_clear_att.clicked.connect(self._clear_attachments)
        att_row.addWidget(btn_add_att)
        att_row.addWidget(btn_clear_att)
        att_row.addStretch()
        s3l.addLayout(att_row)

        self.att_label = QLabel("No attachments added.")
        self.att_label.setObjectName("label_muted")
        self.att_label.setWordWrap(True)
        s3l.addWidget(self.att_label)

        layout.addWidget(step3)

        # ── Step 4: Recipients ───────────────────────────────────────
        step4 = QFrame()
        step4.setObjectName("card")
        s4l = QVBoxLayout(step4)
        s4l.setSpacing(10)

        s4l.addWidget(SectionTitle("4. Recipients"))

        imp_row = QHBoxLayout()
        btn_import = make_button("Import Excel (.xlsx)", icon_text="📂")
        btn_import.clicked.connect(self._import_excel)
        btn_clear_rec = make_button("Clear List", "btn_danger")
        btn_clear_rec.clicked.connect(self._clear_recipients)
        self.rec_count_lbl = QLabel("0 recipients loaded")
        self.rec_count_lbl.setObjectName("label_muted")
        imp_row.addWidget(btn_import)
        imp_row.addWidget(btn_clear_rec)
        imp_row.addStretch()
        imp_row.addWidget(self.rec_count_lbl)
        s4l.addLayout(imp_row)

        self.rec_table = QTableWidget(0, 2)
        self.rec_table.setHorizontalHeaderLabels(["Email", "Name"])
        self.rec_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.rec_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.rec_table.setAlternatingRowColors(True)
        self.rec_table.setMaximumHeight(220)
        self.rec_table.verticalHeader().hide()
        s4l.addWidget(self.rec_table)

        layout.addWidget(step4)

        # ── Send button ──────────────────────────────────────────────
        send_row = QHBoxLayout()
        send_row.addStretch()
        self.btn_send = make_button("Review & Send  →", icon_text="🚀")
        self.btn_send.setFixedHeight(44)
        self.btn_send.setMinimumWidth(200)
        self.btn_send.clicked.connect(self._on_send_clicked)
        send_row.addWidget(self.btn_send)
        layout.addLayout(send_row)

        layout.addStretch()

    # ── Slots ────────────────────────────────────────────────────────

    def _import_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)"
        )
        if not path:
            return
        try:
            result = excel_reader.load_excel(path)
        except Exception as e:
            QMessageBox.critical(self, "Import Error", str(e))
            return

        if result["errors"]:
            details = "\n".join(result["errors"][:20])
            QMessageBox.warning(self, "Import Warnings",
                                f"Some rows were skipped:\n\n{details}")

        if not result["recipients"]:
            QMessageBox.warning(self, "No Data",
                                "No valid recipients found in the file.")
            return

        self.recipients, added = excel_reader.merge_recipients(
            self.recipients, result["recipients"]
        )
        self._refresh_rec_table()
        QMessageBox.information(
            self, "Import Complete",
            f"✅ Added {added} recipients.  "
            f"({len(result['duplicates'])} duplicates skipped, "
            f"{len(result['invalid'])} invalid.)"
        )

    def _clear_recipients(self):
        self.recipients.clear()
        self.rec_table.setRowCount(0)
        self.rec_count_lbl.setText("0 recipients loaded")

    def _refresh_rec_table(self):
        self.rec_table.setRowCount(0)
        for r in self.recipients:
            row = self.rec_table.rowCount()
            self.rec_table.insertRow(row)
            self.rec_table.setItem(row, 0, QTableWidgetItem(r["email"]))
            self.rec_table.setItem(row, 1, QTableWidgetItem(r.get("name", "")))
        self.rec_count_lbl.setText(f"{len(self.recipients)} recipient(s) loaded")

    def _add_attachments(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select Attachments", "", "All Files (*)"
        )
        for p in paths:
            fname = os.path.basename(p)
            if not any(a["filepath"] == p for a in self.attachments):
                self.attachments.append({"filename": fname, "filepath": p})
        self._refresh_att_label()

    def _clear_attachments(self):
        self.attachments.clear()
        self.att_label.setText("No attachments added.")

    def _refresh_att_label(self):
        if self.attachments:
            names = ", ".join(a["filename"] for a in self.attachments)
            self.att_label.setText(f"📎 {len(self.attachments)} file(s): {names}")
        else:
            self.att_label.setText("No attachments added.")

    def _on_send_clicked(self):
        # Validate
        if not self.campaign_name.text().strip():
            QMessageBox.warning(self, "Missing Field", "Please enter a campaign name.")
            return
        if not self.subject_edit.text().strip():
            QMessageBox.warning(self, "Missing Field", "Please enter an email subject.")
            return
        if not self.body_edit.toPlainText().strip():
            QMessageBox.warning(self, "Missing Field", "Message body is empty.")
            return
        if not self.recipients:
            QMessageBox.warning(self, "No Recipients",
                                "Please import at least one recipient.")
            return

        data = {
            "name":        self.campaign_name.text().strip(),
            "subject":     self.subject_edit.text().strip(),
            "body":        self.body_edit.toHtml(),
            "recipients":  self.recipients,
            "attachments": self.attachments,
        }
        self.campaign_ready.emit(data)

    def reset(self):
        self.campaign_name.clear()
        self.subject_edit.clear()
        self.body_edit.clear()
        self.recipients.clear()
        self.attachments.clear()
        self.rec_table.setRowCount(0)
        self.rec_count_lbl.setText("0 recipients loaded")
        self.att_label.setText("No attachments added.")


# ──────────────────────────────────────────────────────────────────────────────
# Page: Send Progress
# ──────────────────────────────────────────────────────────────────────────────

class SendPage(QWidget):
    send_completed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.campaign_id: int | None = None
        self.worker: SendWorker | None = None
        self._access_token: str = ""
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.setContentsMargins(32, 28, 32, 28)

        title = QLabel("Send Progress")
        title.setObjectName("page_title")
        title.setFont(QFont("Segoe UI", 22, QFont.Bold))
        layout.addWidget(title)
        layout.addWidget(Divider())

        # ── Campaign info card ───────────────────────────────────────
        info_card = QFrame()
        info_card.setObjectName("card")
        ic_layout = QVBoxLayout(info_card)

        self.campaign_info_lbl = QLabel("No active campaign")
        self.campaign_info_lbl.setFont(QFont("Segoe UI", 13))
        ic_layout.addWidget(self.campaign_info_lbl)

        layout.addWidget(info_card)

        # ── Options ──────────────────────────────────────────────────
        opts_card = QFrame()
        opts_card.setObjectName("card")
        oc = QHBoxLayout(opts_card)

        self.test_mode_chk = QCheckBox("Test mode (first 3 recipients only)")
        self.delay_spin = QSpinBox()
        self.delay_spin.setRange(1, 60)
        self.delay_spin.setValue(2)
        self.delay_spin.setSuffix(" s delay")
        self.delay_spin.setToolTip("Seconds between each email send (avoids rate limits)")

        oc.addWidget(self.test_mode_chk)
        oc.addStretch()
        oc.addWidget(QLabel("Delay between emails:"))
        oc.addWidget(self.delay_spin)

        layout.addWidget(opts_card)

        # ── Confirmation ─────────────────────────────────────────────
        self.confirm_card = QFrame()
        self.confirm_card.setObjectName("card")
        cc = QVBoxLayout(self.confirm_card)
        self.confirm_lbl = QLabel("")
        self.confirm_lbl.setWordWrap(True)
        cc.addWidget(self.confirm_lbl)

        btn_row = QHBoxLayout()
        self.btn_start = make_button("Start Sending  →", icon_text="🚀")
        self.btn_start.clicked.connect(self._start_sending)
        self.btn_cancel_before = make_button("Cancel", "btn_secondary")
        self.btn_cancel_before.clicked.connect(self.send_completed)
        btn_row.addStretch()
        btn_row.addWidget(self.btn_cancel_before)
        btn_row.addWidget(self.btn_start)
        cc.addLayout(btn_row)
        layout.addWidget(self.confirm_card)

        # ── Progress card ────────────────────────────────────────────
        self.progress_card = QFrame()
        self.progress_card.setObjectName("card")
        self.progress_card.setVisible(False)
        pc = QVBoxLayout(self.progress_card)
        pc.setSpacing(12)

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFixedHeight(18)
        pc.addWidget(self.progress_bar)

        stat_row = QHBoxLayout()
        self.stat_sent = QLabel("✅ Sent: 0")
        self.stat_failed = QLabel("❌ Failed: 0")
        self.stat_current = QLabel("Current: —")
        self.stat_current.setObjectName("label_muted")
        stat_row.addWidget(self.stat_sent)
        stat_row.addWidget(self.stat_failed)
        stat_row.addStretch()
        stat_row.addWidget(self.stat_current)
        pc.addLayout(stat_row)

        # Live log
        self.log_table = QTableWidget(0, 3)
        self.log_table.setHorizontalHeaderLabels(["#", "Email", "Status"])
        self.log_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.log_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.log_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.log_table.setAlternatingRowColors(True)
        self.log_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.log_table.verticalHeader().hide()
        self.log_table.setMaximumHeight(260)
        pc.addWidget(self.log_table)

        ctrl_row = QHBoxLayout()
        self.btn_pause = make_button("Pause", "btn_secondary", "⏸")
        self.btn_pause.clicked.connect(self._toggle_pause)
        self.btn_stop = make_button("Stop", "btn_danger", "⏹")
        self.btn_stop.clicked.connect(self._stop_sending)
        ctrl_row.addStretch()
        ctrl_row.addWidget(self.btn_pause)
        ctrl_row.addWidget(self.btn_stop)
        pc.addLayout(ctrl_row)

        layout.addWidget(self.progress_card)
        layout.addStretch()

    # ── Public API ───────────────────────────────────────────────────

    def prepare(self, campaign_id: int, access_token: str, summary: str):
        self.campaign_id = campaign_id
        self._access_token = access_token
        self.campaign_info_lbl.setText(summary)
        self.confirm_lbl.setText(
            f"<b>Ready to send.</b><br><br>{summary}<br><br>"
            "Please review the details above.  Emails will be sent individually and saved to your Sent folder."
        )
        self.confirm_card.setVisible(True)
        self.progress_card.setVisible(False)
        self.log_table.setRowCount(0)
        self.progress_bar.setValue(0)
        self.stat_sent.setText("✅ Sent: 0")
        self.stat_failed.setText("❌ Failed: 0")
        self.stat_current.setText("Current: —")

    # ── Slots ────────────────────────────────────────────────────────

    def _start_sending(self):
        if self.campaign_id is None:
            return
        self.confirm_card.setVisible(False)
        self.progress_card.setVisible(True)

        recipients = database.get_recipients(self.campaign_id, status="pending")
        total = len(recipients)
        self.progress_bar.setMaximum(max(total, 1))
        self.progress_bar.setValue(0)
        self._sent = 0
        self._failed = 0

        self.worker = SendWorker(
            campaign_id=self.campaign_id,
            access_token=self._access_token,
            delay=float(self.delay_spin.value()),
            test_mode=self.test_mode_chk.isChecked(),
            test_limit=3,
        )
        self.worker.progress.connect(self._on_progress)
        self.worker.finished.connect(self._on_finished)
        self.worker.error.connect(self._on_error)
        self.worker.token_expired.connect(self._on_token_expired)
        self.worker.start()

        self.btn_pause.setEnabled(True)
        self.btn_stop.setEnabled(True)

    def _on_progress(self, current: int, total: int, email: str, success: bool, err: str):
        self.progress_bar.setValue(current)
        self.progress_bar.setMaximum(total)
        if success:
            self._sent += 1
        else:
            self._failed += 1
        self.stat_sent.setText(f"✅ Sent: {self._sent}")
        self.stat_failed.setText(f"❌ Failed: {self._failed}")
        self.stat_current.setText(f"Last: {email}")

        row = self.log_table.rowCount()
        self.log_table.insertRow(row)
        self.log_table.setItem(row, 0, QTableWidgetItem(str(current)))
        self.log_table.setItem(row, 1, QTableWidgetItem(email))

        if success:
            status_item = QTableWidgetItem("✅ Sent")
            status_item.setForeground(QBrush(QColor("#107C10")))
        else:
            status_item = QTableWidgetItem(f"❌ {err[:60]}")
            status_item.setForeground(QBrush(QColor("#D83B01")))
        self.log_table.setItem(row, 2, status_item)
        self.log_table.scrollToBottom()

    def _on_finished(self, sent: int, failed: int):
        self.btn_pause.setEnabled(False)
        self.btn_stop.setEnabled(False)
        QMessageBox.information(
            self, "Sending Complete",
            f"Campaign finished!\n\n✅ Sent: {sent}\n❌ Failed: {failed}"
        )
        self.send_completed.emit()

    def _on_error(self, msg: str):
        QMessageBox.critical(self, "Error", f"An error occurred:\n\n{msg}")
        self.send_completed.emit()

    def _on_token_expired(self):
        QMessageBox.warning(
            self, "Session Expired",
            "Your Microsoft access token has expired.\n"
            "Please log in again and restart the campaign."
        )
        self.send_completed.emit()

    def _toggle_pause(self):
        if not self.worker:
            return
        if self.worker._paused:
            self.worker.resume()
            self.btn_pause.setText("⏸  Pause")
        else:
            self.worker.pause()
            self.btn_pause.setText("▶  Resume")

    def _stop_sending(self):
        if self.worker:
            ret = QMessageBox.question(
                self, "Stop Sending",
                "Stop sending emails? Progress is saved — you can resume later.",
                QMessageBox.Yes | QMessageBox.No
            )
            if ret == QMessageBox.Yes:
                self.worker.stop()


# ──────────────────────────────────────────────────────────────────────────────
# Page: History
# ──────────────────────────────────────────────────────────────────────────────

class HistoryPage(QWidget):
    resume_campaign = Signal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.setContentsMargins(32, 28, 32, 28)

        title = QLabel("Campaign History")
        title.setObjectName("page_title")
        title.setFont(QFont("Segoe UI", 22, QFont.Bold))
        layout.addWidget(title)
        subtitle = QLabel("View past campaigns and export results")
        subtitle.setObjectName("label_muted")
        layout.addWidget(subtitle)
        layout.addWidget(Divider())

        # Table
        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(
            ["ID", "Name", "Subject", "Status", "Sent", "Failed", "Date"]
        )
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.verticalHeader().hide()
        self.table.selectionModel().selectionChanged.connect(self._on_selection)
        layout.addWidget(self.table)

        # Action bar
        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        self.btn_refresh = make_button("Refresh", "btn_secondary", "🔄")
        self.btn_refresh.clicked.connect(self.refresh)
        self.btn_export_csv = make_button("Export CSV", "btn_secondary", "📄")
        self.btn_export_csv.clicked.connect(lambda: self._export("csv"))
        self.btn_export_xlsx = make_button("Export Excel", "btn_secondary", "📊")
        self.btn_export_xlsx.clicked.connect(lambda: self._export("xlsx"))
        self.btn_delete = make_button("Delete", "btn_danger", "🗑")
        self.btn_delete.clicked.connect(self._delete)
        self.btn_details = make_button("View Details", icon_text="🔍")
        self.btn_details.clicked.connect(self._view_details)

        for b in [self.btn_refresh, self.btn_export_csv, self.btn_export_xlsx,
                  self.btn_details, self.btn_delete]:
            btn_row.addWidget(b)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        self._set_action_buttons_enabled(False)

    def _set_action_buttons_enabled(self, enabled: bool):
        for b in [self.btn_export_csv, self.btn_export_xlsx,
                  self.btn_delete, self.btn_details]:
            b.setEnabled(enabled)

    def _on_selection(self, selected, _):
        self._set_action_buttons_enabled(bool(selected.indexes()))

    def _selected_campaign_id(self) -> int | None:
        rows = self.table.selectedItems()
        if not rows:
            return None
        return int(self.table.item(rows[0].row(), 0).text())

    def refresh(self):
        campaigns = database.get_campaigns()
        self.table.setRowCount(0)
        for c in campaigns:
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(str(c["id"])))
            self.table.setItem(row, 1, QTableWidgetItem(c["name"]))
            self.table.setItem(row, 2, QTableWidgetItem(c["subject"]))
            status_item = QTableWidgetItem(c["status"].capitalize())
            color_map = {
                "completed": QColor("#107C10"),
                "failed":    QColor("#D83B01"),
                "pending":   QColor("#FFB900"),
                "paused":    QColor("#0078D4"),
            }
            status_item.setForeground(QBrush(color_map.get(c["status"], QColor("#FFFFFF"))))
            self.table.setItem(row, 3, status_item)
            self.table.setItem(row, 4, QTableWidgetItem(str(c["sent"])))
            self.table.setItem(row, 5, QTableWidgetItem(str(c["failed"])))
            self.table.setItem(row, 6, QTableWidgetItem((c["created_at"] or "")[:10]))

    def _export(self, fmt: str):
        cid = self._selected_campaign_id()
        if cid is None:
            return
        ext = "CSV Files (*.csv)" if fmt == "csv" else "Excel Files (*.xlsx)"
        default = f"campaign_{cid}.{fmt}"
        path, _ = QFileDialog.getSaveFileName(self, "Export Results", default, ext)
        if not path:
            return
        try:
            if fmt == "csv":
                n = exporter.export_csv(cid, path)
            else:
                n = exporter.export_excel(cid, path)
            QMessageBox.information(self, "Exported", f"✅ {n} rows exported to:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Export Error", str(e))

    def _delete(self):
        cid = self._selected_campaign_id()
        if cid is None:
            return
        ret = QMessageBox.question(
            self, "Delete Campaign",
            f"Permanently delete campaign #{cid} and all its records?",
            QMessageBox.Yes | QMessageBox.No
        )
        if ret == QMessageBox.Yes:
            database.delete_campaign(cid)
            self.refresh()

    def _view_details(self):
        cid = self._selected_campaign_id()
        if cid is None:
            return
        dlg = CampaignDetailDialog(cid, self)
        dlg.exec()


class CampaignDetailDialog(QDialog):
    def __init__(self, campaign_id: int, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Campaign #{campaign_id} — Details")
        self.resize(720, 500)
        layout = QVBoxLayout(self)

        campaign = database.get_campaign(campaign_id)
        recipients = database.get_recipients(campaign_id)

        info = QLabel(
            f"<b>Name:</b> {campaign['name']}<br>"
            f"<b>Subject:</b> {campaign['subject']}<br>"
            f"<b>Status:</b> {campaign['status']}  |  "
            f"<b>Sent:</b> {campaign['sent']}  |  "
            f"<b>Failed:</b> {campaign['failed']}  |  "
            f"<b>Total:</b> {campaign['total']}"
        )
        info.setWordWrap(True)
        layout.addWidget(info)
        layout.addWidget(Divider())

        table = QTableWidget(len(recipients), 4)
        table.setHorizontalHeaderLabels(["Name", "Email", "Status", "Sent At"])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.setAlternatingRowColors(True)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.verticalHeader().hide()

        color_map = {
            "sent":    QColor("#107C10"),
            "failed":  QColor("#D83B01"),
            "pending": QColor("#FFB900"),
        }
        for i, r in enumerate(recipients):
            table.setItem(i, 0, QTableWidgetItem(r.get("name", "")))
            table.setItem(i, 1, QTableWidgetItem(r["email"]))
            si = QTableWidgetItem(r["status"].capitalize())
            si.setForeground(QBrush(color_map.get(r["status"], QColor("#FFFFFF"))))
            table.setItem(i, 2, si)
            table.setItem(i, 3, QTableWidgetItem(r.get("sent_at", "")[:16]))
        layout.addWidget(table)

        bb = QDialogButtonBox(QDialogButtonBox.Close)
        bb.rejected.connect(self.reject)
        layout.addWidget(bb)


# ──────────────────────────────────────────────────────────────────────────────
# Page: Settings / Login
# ──────────────────────────────────────────────────────────────────────────────

class SettingsPage(QWidget):
    auth_changed  = Signal(str, str, str)  # token, display_name, email
    theme_changed = Signal(bool)           # dark=True / light=False

    def __init__(self, parent=None):
        super().__init__(parent)
        self._auth_worker: AuthWorker | None = None
        self._test_worker = None
        self._current_token: str = ""
        self._pending_code: str = ""
        self._build_ui()

    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setSpacing(0)
        outer.setContentsMargins(0, 0, 0, 0)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        outer.addWidget(scroll)
        container = QWidget()
        scroll.setWidget(container)
        layout = QVBoxLayout(container)
        layout.setSpacing(20)
        layout.setContentsMargins(32, 28, 32, 28)

        title = QLabel("Settings & Account")
        title.setObjectName("page_title")
        title.setFont(QFont("Segoe UI", 22, QFont.Bold))
        layout.addWidget(title)
        subtitle = QLabel("Configure your Azure app and Microsoft account")
        subtitle.setObjectName("label_muted")
        layout.addWidget(subtitle)
        layout.addWidget(Divider())

        # Azure Client ID card
        azure_card = QFrame()
        azure_card.setObjectName("card")
        ac = QVBoxLayout(azure_card)
        ac.setSpacing(10)
        ac.addWidget(SectionTitle("\U0001f4d6  Azure App Registration"))
        az_info = QLabel(
            "Enter your <b>Application (client) ID</b> from the Microsoft Entra admin centre.<br>"
            "It is a UUID such as <code>12345678-1234-1234-1234-123456789abc</code>.<br>"
            "See the README for full step-by-step registration instructions."
        )
        az_info.setWordWrap(True)
        az_info.setObjectName("label_muted")
        ac.addWidget(az_info)
        id_row = QHBoxLayout()
        self.client_id_edit = QLineEdit()
        self.client_id_edit.setPlaceholderText("xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx")
        saved = auth.get_client_id()
        self.client_id_edit.setText("" if saved == "YOUR_AZURE_APP_CLIENT_ID" else saved)
        self.client_id_edit.textChanged.connect(self._on_client_id_changed)
        btn_save_id = make_button("Save", "btn_secondary")
        btn_save_id.setFixedWidth(80)
        btn_save_id.clicked.connect(self._save_client_id)
        btn_portal = make_button("Open Entra Portal", "btn_secondary", "\U0001f310")
        btn_portal.clicked.connect(
            lambda: webbrowser.open(
                "https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade"
            )
        )
        id_row.addWidget(self.client_id_edit)
        id_row.addWidget(btn_save_id)
        id_row.addWidget(btn_portal)
        ac.addLayout(id_row)
        self.client_id_status = QLabel("")
        self.client_id_status.setWordWrap(True)
        ac.addWidget(self.client_id_status)
        layout.addWidget(azure_card)

        # ── Setup Guide card (collapsible) ──────────────────────────
        guide_card = QFrame()
        guide_card.setObjectName("card")
        gc = QVBoxLayout(guide_card)
        gc.setSpacing(8)

        guide_header = QHBoxLayout()
        guide_title = SectionTitle("📋  Microsoft Entra App Registration Guide")
        self.btn_toggle_guide = QPushButton("▶  Show Guide")
        self.btn_toggle_guide.setObjectName("btn_secondary")
        self.btn_toggle_guide.setFixedWidth(150)
        self.btn_toggle_guide.setCursor(Qt.PointingHandCursor)
        self.btn_toggle_guide.clicked.connect(self._toggle_guide)
        guide_header.addWidget(guide_title)
        guide_header.addStretch()
        guide_header.addWidget(self.btn_toggle_guide)
        gc.addLayout(guide_header)

        self.guide_body = QWidget()
        self.guide_body.setVisible(False)
        gb = QVBoxLayout(self.guide_body)
        gb.setSpacing(12)
        gb.setContentsMargins(0, 8, 0, 0)

        def _step(num: str, title: str, body: str) -> QFrame:
            f = QFrame()
            f.setStyleSheet(
                "QFrame { border-left: 3px solid #0078D4; "
                "background: rgba(0,120,212,0.06); border-radius: 6px; padding: 8px; }"
            )
            fl = QVBoxLayout(f)
            fl.setSpacing(4)
            fl.setContentsMargins(12, 8, 8, 8)
            hdr = QLabel(f"<b style='color:#0078D4;'>Step {num}</b> — {title}")
            hdr.setFont(QFont("Segoe UI", 11, QFont.Bold))
            fl.addWidget(hdr)
            lbl = QLabel(body)
            lbl.setWordWrap(True)
            lbl.setOpenExternalLinks(True)
            lbl.setTextFormat(Qt.RichText)
            fl.addWidget(lbl)
            return f

        gb.addWidget(_step("1", "Open Microsoft Entra Admin Centre",
            'Go to <a href="https://entra.microsoft.com">https://entra.microsoft.com</a> '
            'and sign in with your Microsoft account.<br>'
            'In the left sidebar click <b>Applications</b> → <b>App registrations</b>.'))

        gb.addWidget(_step("2", "Create a New Registration",
            'Click <b>+ New registration</b>.<br>'
            'Set <b>Name</b> to <code>OutlookBulkMailer</code>.<br>'
            'For <b>Supported account types</b> choose:<br>'
            '&nbsp;&nbsp;• <b>Personal accounts only</b> → select '
            '"<i>Accounts in any organizational directory … and personal Microsoft accounts</i>"<br>'
            '&nbsp;&nbsp;• <b>Work/school only</b> → select '
            '"<i>Accounts in this organizational directory only</i>"<br>'
            'Leave Redirect URI blank, then click <b>Register</b>.'))

        gb.addWidget(_step("3", "Copy the Application (Client) ID",
            'On the <b>Overview</b> page locate the <b>Application (client) ID</b> field.<br>'
            'It looks like: <code>12345678-abcd-1234-efgh-123456789abc</code><br>'
            'Click the copy icon and paste it into the field above, then click <b>Save</b>.'))

        gb.addWidget(_step("4", "Configure the Authentication Platform",
            'Click <b>Authentication</b> in the left panel → <b>+ Add a platform</b> → '
            '<b>Mobile and desktop applications</b>.<br>'
            'Tick the checkbox for:<br>'
            '&nbsp;&nbsp;<code>https://login.microsoftonline.com/common/oauth2/nativeclient</code><br>'
            'Scroll down and set <b>"Allow public client flows"</b> to <b>Yes</b>.<br>'
            'Click <b>Save</b>.'))

        gb.addWidget(_step("5", "Add Microsoft Graph Permissions",
            'Click <b>API permissions</b> → <b>+ Add a permission</b> → <b>Microsoft Graph</b> '
            '→ <b>Delegated permissions</b>.<br>'
            'Search for and add:<br>'
            '&nbsp;&nbsp;• <b>Mail.Send</b> — allows sending email as you<br>'
            '&nbsp;&nbsp;• <b>User.Read</b> — allows reading your profile<br>'
            'Click <b>Add permissions</b>.<br>'
            '<i>Personal accounts: consent is granted interactively during sign-in. '
            'Work/school accounts: an admin may need to click "Grant admin consent".</i>'))

        gb.addWidget(_step("6", "Sign In",
            'Return to this Settings page, paste your Client ID above, click <b>Save</b>, '
            'then click <b>🔐 Sign in with Microsoft</b> below.<br>'
            'Accept the Mail.Send and User.Read permissions when prompted.<br>'
            'Use <b>🔄 Verify Now</b> to confirm the connection and <b>📧 Send Test Email</b> '
            'to verify end-to-end sending.'))

        ref_lbl = QLabel(
            '📖 Full guide: '
            '<a href="https://learn.microsoft.com/en-us/entra/identity-platform/scenario-desktop-app-registration">'
            'Microsoft Entra App Registration Docs</a> · '
            '<a href="https://learn.microsoft.com/en-us/graph/api/user-sendmail">'
            'Graph sendMail Docs</a>'
        )
        ref_lbl.setOpenExternalLinks(True)
        ref_lbl.setWordWrap(True)
        ref_lbl.setObjectName("label_muted")
        gb.addWidget(ref_lbl)

        gc.addWidget(self.guide_body)
        layout.addWidget(guide_card)


        conn_card = QFrame()
        conn_card.setObjectName("card")
        cc = QVBoxLayout(conn_card)
        cc.setSpacing(10)
        cc.addWidget(SectionTitle("\U0001f50c  Connection Status"))
        status_row = QHBoxLayout()
        self.conn_dot = QLabel("\u26ab")
        self.conn_dot.setFont(QFont("Segoe UI Emoji", 13))
        self.conn_status_lbl = QLabel("Not connected")
        self.conn_status_lbl.setFont(QFont("Segoe UI", 12))
        self.btn_verify = make_button("Verify Now", "btn_secondary", "\U0001f504")
        self.btn_verify.setFixedWidth(120)
        self.btn_verify.clicked.connect(self._verify_connection)
        status_row.addWidget(self.conn_dot)
        status_row.addWidget(self.conn_status_lbl)
        status_row.addStretch()
        status_row.addWidget(self.btn_verify)
        cc.addLayout(status_row)
        layout.addWidget(conn_card)

        # Login card
        login_card = QFrame()
        login_card.setObjectName("card")
        lc = QVBoxLayout(login_card)
        lc.setSpacing(10)
        lc.addWidget(SectionTitle("\U0001f510  Microsoft Account"))
        self.auth_status_lbl = QLabel("Not signed in")
        self.auth_status_lbl.setObjectName("label_warning")
        lc.addWidget(self.auth_status_lbl)

        # Device-code panel
        self.device_code_frame = QFrame()
        self.device_code_frame.setObjectName("card")
        self.device_code_frame.setStyleSheet(
            "QFrame#card { border: 1.5px solid #0078D4; background: rgba(0,120,212,0.07); }"
        )
        dc = QVBoxLayout(self.device_code_frame)
        dc.setSpacing(8)
        dc.addWidget(QLabel("\U0001f517 <b>Step 1</b> \u2014 Your browser opens automatically. If not, visit:"))
        self.device_code_url_lbl = QLabel("")
        self.device_code_url_lbl.setOpenExternalLinks(True)
        self.device_code_url_lbl.setWordWrap(True)
        dc.addWidget(self.device_code_url_lbl)
        dc.addWidget(QLabel("\U0001f4dd <b>Step 2</b> \u2014 Enter this one-time code on that page:"))
        self.device_code_lbl = QLabel("")
        self.device_code_lbl.setFont(QFont("Consolas", 20, QFont.Bold))
        self.device_code_lbl.setAlignment(Qt.AlignCenter)
        self.device_code_lbl.setStyleSheet(
            "color:#0078D4; background:rgba(0,120,212,0.12); border-radius:6px; padding:8px;"
        )
        dc.addWidget(self.device_code_lbl)
        btn_copy = make_button("Copy Code", "btn_secondary", "\U0001f4cb")
        btn_copy.setFixedWidth(130)
        btn_copy.clicked.connect(self._copy_device_code)
        dc.addWidget(btn_copy)
        dc.addWidget(QLabel(
            "\u2705 <b>Step 3</b> \u2014 Sign in and accept <b>Mail.Send</b> and "
            "<b>User.Read</b> permissions. This window updates automatically."
        ))
        self.device_code_frame.setVisible(False)
        lc.addWidget(self.device_code_frame)

        btn_row = QHBoxLayout()
        self.btn_login = make_button("Sign in with Microsoft", icon_text="\U0001f510")
        self.btn_login.clicked.connect(self._start_login)
        self.btn_cancel_login = make_button("Cancel", "btn_secondary", "\u2716")
        self.btn_cancel_login.clicked.connect(self._cancel_login)
        self.btn_cancel_login.setVisible(False)
        self.btn_logout = make_button("Sign Out", "btn_danger", "\U0001f6aa")
        self.btn_logout.clicked.connect(self._logout)
        btn_row.addWidget(self.btn_login)
        btn_row.addWidget(self.btn_cancel_login)
        btn_row.addWidget(self.btn_logout)
        btn_row.addStretch()
        lc.addLayout(btn_row)
        layout.addWidget(login_card)

        # Test email card
        test_card = QFrame()
        test_card.setObjectName("card")
        tc2 = QVBoxLayout(test_card)
        tc2.setSpacing(10)
        tc2.addWidget(SectionTitle("\U0001f9ea  Send Test Email"))
        te_info = QLabel(
            "Send a quick test email to verify auth \u2192 Graph API \u2192 Outlook works "
            "before bulk-sending. Leave blank to send to yourself."
        )
        te_info.setWordWrap(True)
        te_info.setObjectName("label_muted")
        tc2.addWidget(te_info)
        te_row = QHBoxLayout()
        self.test_email_edit = QLineEdit()
        self.test_email_edit.setPlaceholderText("recipient@example.com")
        self.btn_test_email = make_button("Send Test Email", "btn_secondary", "\U0001f4e7")
        self.btn_test_email.clicked.connect(self._send_test_email)
        te_row.addWidget(self.test_email_edit)
        te_row.addWidget(self.btn_test_email)
        tc2.addLayout(te_row)
        self.test_email_status = QLabel("")
        self.test_email_status.setWordWrap(True)
        tc2.addWidget(self.test_email_status)
        layout.addWidget(test_card)

        # Appearance card
        theme_card = QFrame()
        theme_card.setObjectName("card")
        th = QVBoxLayout(theme_card)
        th.addWidget(SectionTitle("\U0001f3a8  Appearance"))
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["Dark (default)", "Light"])
        self.theme_combo.currentIndexChanged.connect(
            lambda idx: self.theme_changed.emit(idx == 0)
        )
        th.addWidget(self.theme_combo)
        layout.addWidget(theme_card)
        layout.addStretch()

    # ---- Client ID -------------------------------------------------

    def _on_client_id_changed(self, text: str):
        if not text.strip():
            self.client_id_status.setText("")
            return
        if auth.is_valid_client_id(text):
            self.client_id_status.setText(
                "<span style='color:#107C10;'>\u2705 Valid UUID format</span>"
            )
        else:
            self.client_id_status.setText(
                "<span style='color:#D83B01;'>\u26a0 Must be UUID format: "
                "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx</span>"
            )

    def _save_client_id(self):
        cid = self.client_id_edit.text().strip()
        if not cid:
            QMessageBox.warning(self, "Missing", "Please paste your Application (client) ID.")
            return
        if not auth.is_valid_client_id(cid):
            QMessageBox.warning(
                self, "Invalid Format",
                "The Client ID must be UUID format:\nxxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx\n\n"
                "Copy it from the app registration Overview in the Entra admin centre."
            )
            return
        auth.set_client_id(cid)
        self.client_id_status.setText("<span style='color:#107C10;'>\u2705 Saved</span>")

    def _toggle_guide(self):
        visible = self.guide_body.isVisible()
        self.guide_body.setVisible(not visible)
        self.btn_toggle_guide.setText(
            "▼  Hide Guide" if not visible else "▶  Show Guide"
        )

    # ---- Connection verify -----------------------------------------

    def _verify_connection(self):
        if not self._current_token:
            self._set_conn("orange", "\U0001f7e1 Not signed in \u2014 please sign in first.")
            return
        self._set_conn("gray", "\u23f3 Verifying...")
        self.btn_verify.setEnabled(False)
        import threading
        token = self._current_token

        def _run():
            result = auth.verify_connection(token)
            self._verify_result = result
            QTimer.singleShot(0, self._on_verify_done)

        threading.Thread(target=_run, daemon=True).start()

    def _on_verify_done(self):
        self.btn_verify.setEnabled(True)
        r = getattr(self, "_verify_result", {})
        if r.get("ok"):
            self._set_conn("green", f"\U0001f7e2 Connected \u2014 {r['display_name']} ({r['email']})")
        elif r.get("expired"):
            self._set_conn("red", f"\U0001f534 Token expired \u2014 {r['error']}")
        else:
            self._set_conn("red", f"\U0001f534 {r.get('error', 'Unknown error')}")

    def _set_conn(self, colour: str, text: str):
        dots = {"green": "\U0001f7e2", "orange": "\U0001f7e1", "red": "\U0001f534", "gray": "\u26ab"}
        self.conn_dot.setText(dots.get(colour, "\u26ab"))
        self.conn_status_lbl.setText(text)

    # ---- Login flow ------------------------------------------------

    def _start_login(self):
        try:
            flow_info = auth.login_interactive()
        except auth.ConfigError as e:
            QMessageBox.warning(self, "Configuration Required", str(e))
            return
        except Exception as e:
            QMessageBox.critical(self, "Sign-in Error", f"Could not start sign-in flow:\n\n{e}")
            return

        flow  = flow_info["flow"]
        _app  = flow_info["app"]
        cache = flow_info["cache"]
        url   = flow.get("verification_uri", "https://microsoft.com/devicelogin")
        code  = flow.get("user_code", "")
        self._pending_code = code

        self.device_code_url_lbl.setText(f'<a href="{url}">{url}</a>')
        self.device_code_lbl.setText(code)
        self.device_code_frame.setVisible(True)
        self.btn_login.setEnabled(False)
        self.btn_cancel_login.setVisible(True)
        self.auth_status_lbl.setText("\u23f3 Waiting for Microsoft sign-in...")
        self._set_conn("gray", "\u23f3 Waiting for sign-in...")
        webbrowser.open(url)

        self._auth_worker = AuthWorker(_app, flow, cache)
        self._auth_worker.success.connect(self._on_auth_success)
        self._auth_worker.error.connect(self._on_auth_error)
        self._auth_worker.start()

    def _cancel_login(self):
        if self._auth_worker and self._auth_worker.isRunning():
            self._auth_worker.terminate()
        self._reset_login_ui()
        self.auth_status_lbl.setText("Sign-in cancelled.")
        self._set_conn("gray", "Not connected")

    def _reset_login_ui(self):
        self.device_code_frame.setVisible(False)
        self.btn_login.setEnabled(True)
        self.btn_cancel_login.setVisible(False)

    def _copy_device_code(self):
        if self._pending_code:
            QApplication.clipboard().setText(self._pending_code)

    def _on_auth_success(self, token: str, display: str, email: str):
        self._current_token = token
        self._reset_login_ui()
        self.auth_status_lbl.setText(
            f"<span style='color:#107C10;'>\u2705 Signed in as <b>{display}</b> ({email})</span>"
        )
        self._set_conn("green", f"\U0001f7e2 Connected \u2014 {display} ({email})")
        if not self.test_email_edit.text().strip():
            self.test_email_edit.setText(email)
        self.auth_changed.emit(token, display, email)

    def _on_auth_error(self, msg: str):
        self._reset_login_ui()
        msg_lower = msg.lower()
        if any(k in msg_lower for k in ("declined", "access_denied", "consent")):
            title = "Consent Denied"
            detail = (
                f"{msg}\n\nPlease try again and click 'Accept' when Microsoft "
                "asks for Mail.Send and User.Read permissions."
            )
        elif "expired" in msg_lower and "code" in msg_lower:
            title  = "Code Expired"
            detail = f"{msg}\n\nClick 'Sign in with Microsoft' to get a new code."
        elif "invalid" in msg_lower and "client" in msg_lower:
            title  = "Invalid Client ID"
            detail = (
                f"{msg}\n\nDouble-check the Application (client) ID — it must be a UUID "
                "copied from the Entra admin centre."
            )
        else:
            title  = "Sign-in Failed"
            detail = msg
        self.auth_status_lbl.setText(f"<span style='color:#D83B01;'>\u274c {title}</span>")
        self._set_conn("red", f"\U0001f534 {title}")
        QMessageBox.critical(self, title, detail)

    def _logout(self):
        auth.logout()
        self._current_token = ""
        self.auth_status_lbl.setText("Signed out.")
        self._set_conn("gray", "Not connected")
        self.test_email_edit.clear()
        self.auth_changed.emit("", "", "")

    # ---- Test email ------------------------------------------------

    def _send_test_email(self):
        if not self._current_token:
            QMessageBox.warning(self, "Not Signed In",
                                "Please sign in with Microsoft before sending a test email.")
            return
        to = self.test_email_edit.text().strip()
        if not to:
            QMessageBox.warning(self, "Email Required",
                                "Enter a recipient address for the test email.")
            return
        self.btn_test_email.setEnabled(False)
        self.test_email_status.setText("\u23f3 Sending test email...")
        self._test_worker = TestEmailWorker(self._current_token, to)
        self._test_worker.done.connect(self._on_test_done)
        self._test_worker.start()

    def _on_test_done(self, success: bool, message: str):
        self.btn_test_email.setEnabled(True)
        if success:
            self.test_email_status.setText(
                "<span style='color:#107C10;'>\u2705 Test email sent! Check your inbox.</span>"
            )
        else:
            self.test_email_status.setText(
                f"<span style='color:#D83B01;'>\u274c {message}</span>"
            )

    # ---- Public helpers --------------------------------------------

    def set_auth_status(self, token: str, display: str, email: str):
        self._current_token = token
        if display:
            self.auth_status_lbl.setText(
                f"<span style='color:#107C10;'>\u2705 Signed in as <b>{display}</b> ({email})</span>"
            )
            self._set_conn("green", f"\U0001f7e2 Connected \u2014 {display} ({email})")
            if not self.test_email_edit.text().strip():
                self.test_email_edit.setText(email)

    def get_theme_index(self) -> int:
        return self.theme_combo.currentIndex()


# ──────────────────────────────────────────────────────────────────────────────
# Background worker: test email
# ──────────────────────────────────────────────────────────────────────────────

class TestEmailWorker(QThread):
    done = Signal(bool, str)  # success, error_message

    def __init__(self, access_token: str, to_email: str, parent=None):
        super().__init__(parent)
        self._token = access_token
        self._to    = to_email

    def run(self):
        try:
            result = mailer.send_test_email(self._token, self._to)
            if result["success"]:
                self.done.emit(True, "")
            else:
                self.done.emit(False, result.get("error", "Unknown error"))
        except Exception as e:
            self.done.emit(False, str(e))

# ──────────────────────────────────────────────────────────────────────────────
# Main Window
# ──────────────────────────────────────────────────────────────────────────────

# ──────────────────────────────────────────────────────────────────────────────
# Page: About
# ──────────────────────────────────────────────────────────────────────────────

class AboutPage(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(24)
        layout.setContentsMargins(32, 28, 32, 28)

        title = QLabel("About")
        title.setObjectName("page_title")
        title.setFont(QFont("Segoe UI", 22, QFont.Bold))
        layout.addWidget(title)
        subtitle = QLabel("Outlook Bulk Mailer — Application Info")
        subtitle.setObjectName("label_muted")
        layout.addWidget(subtitle)
        layout.addWidget(Divider())

        # ── App info card ────────────────────────────────────────────
        app_card = QFrame()
        app_card.setObjectName("card")
        ac = QVBoxLayout(app_card)
        ac.setSpacing(10)

        _base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        _icon_path = os.path.join(_base, "assets", "icon.png")
        if os.path.exists(_icon_path):
            app_icon = QLabel()
            pix = QPixmap(_icon_path).scaled(72, 72, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            app_icon.setPixmap(pix)
        else:
            app_icon = QLabel("📧")
            app_icon.setFont(QFont("Segoe UI Emoji", 48))
        app_icon.setAlignment(Qt.AlignCenter)
        ac.addWidget(app_icon)

        app_name = QLabel("Outlook Bulk Mailer")
        app_name.setFont(QFont("Segoe UI", 20, QFont.Bold))
        app_name.setAlignment(Qt.AlignCenter)
        ac.addWidget(app_name)

        version_lbl = QLabel("Version 1.0.0")
        version_lbl.setObjectName("label_muted")
        version_lbl.setAlignment(Qt.AlignCenter)
        ac.addWidget(version_lbl)

        desc_lbl = QLabel(
            "A modern desktop application for sending personalised bulk emails\n"
            "via Microsoft Outlook using the Microsoft Graph API."
        )
        desc_lbl.setAlignment(Qt.AlignCenter)
        desc_lbl.setWordWrap(True)
        ac.addWidget(desc_lbl)

        layout.addWidget(app_card)

        # ── Developer card ───────────────────────────────────────────
        dev_card = QFrame()
        dev_card.setObjectName("card")
        dc = QVBoxLayout(dev_card)
        dc.setSpacing(14)

        dev_title = SectionTitle("👨‍💻  Developer")
        dc.addWidget(dev_title)
        dc.addWidget(Divider())

        # Avatar area
        avatar_row = QHBoxLayout()
        avatar_row.setSpacing(20)

        avatar_lbl = QLabel("VY")
        avatar_lbl.setFixedSize(72, 72)
        avatar_lbl.setAlignment(Qt.AlignCenter)
        avatar_lbl.setFont(QFont("Segoe UI", 22, QFont.Bold))
        avatar_lbl.setStyleSheet(
            "background: qlineargradient(x1:0,y1:0,x2:1,y2:1,"
            "stop:0 #0078D4, stop:1 #00B4D8);"
            "color: white; border-radius: 36px;"
        )
        avatar_row.addWidget(avatar_lbl)

        info_col = QVBoxLayout()
        info_col.setSpacing(6)

        name_lbl = QLabel("Vikrant Yadav")
        name_lbl.setFont(QFont("Segoe UI", 16, QFont.Bold))
        info_col.addWidget(name_lbl)

        role_lbl = QLabel("Developer & Maintainer")
        role_lbl.setObjectName("label_muted")
        info_col.addWidget(role_lbl)

        # GitHub link button
        github_btn = QPushButton("  🐙  github.com/vikrant-eng")
        github_btn.setObjectName("btn_secondary")
        github_btn.setCursor(Qt.PointingHandCursor)
        github_btn.setFixedHeight(36)
        github_btn.setFixedWidth(260)
        github_btn.setFont(QFont("Segoe UI", 12))
        github_btn.clicked.connect(
            lambda: webbrowser.open("https://github.com/vikrant-eng")
        )
        info_col.addWidget(github_btn)

        avatar_row.addLayout(info_col)
        avatar_row.addStretch()
        dc.addLayout(avatar_row)

        layout.addWidget(dev_card)

        # ── Tech stack card ──────────────────────────────────────────
        tech_card = QFrame()
        tech_card.setObjectName("card")
        tc = QVBoxLayout(tech_card)
        tc.setSpacing(10)
        tc.addWidget(SectionTitle("🛠  Technology Stack"))
        tc.addWidget(Divider())

        tech_grid = QHBoxLayout()
        tech_grid.setSpacing(14)
        techs = [
            ("🐍", "Python 3.11+",       "Core language"),
            ("🖥", "PySide6",             "GUI framework"),
            ("🔐", "MSAL",               "OAuth2 / Graph API"),
            ("📊", "openpyxl",           "Excel import/export"),
            ("🗄", "SQLite",             "Campaign history"),
        ]
        for icon, name, desc in techs:
            cell = QFrame()
            cell.setStyleSheet(
                "QFrame { background: rgba(0,120,212,0.10); "
                "border-radius: 8px; padding: 4px; }"
            )
            cl = QVBoxLayout(cell)
            cl.setSpacing(4)
            cl.setContentsMargins(12, 10, 12, 10)
            il = QLabel(icon)
            il.setFont(QFont("Segoe UI Emoji", 20))
            il.setAlignment(Qt.AlignCenter)
            nl = QLabel(name)
            nl.setFont(QFont("Segoe UI", 11, QFont.Bold))
            nl.setAlignment(Qt.AlignCenter)
            dl = QLabel(desc)
            dl.setObjectName("label_muted")
            dl.setAlignment(Qt.AlignCenter)
            dl.setFont(QFont("Segoe UI", 10))
            cl.addWidget(il)
            cl.addWidget(nl)
            cl.addWidget(dl)
            tech_grid.addWidget(cell)
        tc.addLayout(tech_grid)

        layout.addWidget(tech_card)

        # ── Footer ───────────────────────────────────────────────────
        footer = QLabel("MIT License · © 2026 Vikrant Yadav")
        footer.setObjectName("label_muted")
        footer.setAlignment(Qt.AlignCenter)
        layout.addWidget(footer)

        layout.addStretch()


NAV_ITEMS = [
    ("🏠  Dashboard",    0),
    ("✉️  Compose",      1),
    ("📤  Send",         2),
    ("📋  History",      3),
    ("⚙️  Settings",     4),
    ("ℹ️  About",        5),
]


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Outlook Bulk Mailer")
        self.resize(1180, 760)
        self.setMinimumSize(900, 600)

        # Icon
        _base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        _icon_path = os.path.join(_base, "assets", "icon.png")
        if os.path.exists(_icon_path):
            self.setWindowIcon(QIcon(_icon_path))

        # State
        self._access_token = ""
        self._user_display = ""
        self._user_email = ""
        self._current_campaign_id: int | None = None
        self._dark_mode = True

        # DB
        database.init_db()
        auth.load_settings()

        # Try silent token (silently skip if Client ID not yet configured)
        try:
            token = auth.get_access_token()
        except auth.ConfigError:
            token = None
        except auth.TokenExpiredError:
            token = None  # will prompt re-login in Settings

        if token:
            self._access_token = token
            try:
                info = auth.get_user_info(token)
                self._user_display = info.get("displayName", "")
                self._user_email   = info.get("userPrincipalName", info.get("mail", ""))
            except Exception:
                pass

        self._build_ui()
        self._apply_theme()
        self._nav_buttons[0].click()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QHBoxLayout(central)
        root.setSpacing(0)
        root.setContentsMargins(0, 0, 0, 0)

        # ── Sidebar ──────────────────────────────────────────────────
        sidebar = QWidget()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(220)
        sb_layout = QVBoxLayout(sidebar)
        sb_layout.setSpacing(0)
        sb_layout.setContentsMargins(0, 0, 0, 0)

        # Logo
        logo_container = QWidget()
        logo_container.setFixedHeight(80)
        logo_container.setStyleSheet("background: transparent;")
        logo_layout = QVBoxLayout(logo_container)
        logo_layout.setContentsMargins(16, 16, 16, 8)
        logo_lbl = QLabel("📧 Bulk Mailer")
        logo_lbl.setObjectName("sidebar_logo_label")
        logo_lbl.setFont(QFont("Segoe UI", 16, QFont.Bold))
        logo_sub = QLabel("Outlook · Graph API")
        logo_sub.setObjectName("sidebar_sub_label")
        logo_layout.addWidget(logo_lbl)
        logo_layout.addWidget(logo_sub)
        sb_layout.addWidget(logo_container)

        sep = Divider()
        sep.setStyleSheet("color: #2A3F55;")
        sb_layout.addWidget(sep)
        sb_layout.addSpacing(8)

        # Nav buttons
        self._nav_buttons: list[QPushButton] = []
        self._stacked: QStackedWidget | None = None

        for label, page_idx in NAV_ITEMS:
            btn = QPushButton(label)
            btn.setObjectName("nav_button")
            btn.setFixedHeight(44)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setCheckable(False)
            btn.clicked.connect(lambda checked, idx=page_idx: self._navigate(idx))
            sb_layout.addWidget(btn)
            self._nav_buttons.append(btn)

        sb_layout.addStretch()

        # User info strip
        self._user_strip = QLabel("Not signed in")
        self._user_strip.setObjectName("sidebar_sub_label")
        self._user_strip.setWordWrap(True)
        self._user_strip.setContentsMargins(16, 8, 16, 8)
        self._update_user_strip()
        sb_layout.addWidget(self._user_strip)
        sb_layout.addSpacing(12)

        root.addWidget(sidebar)

        # ── Content ──────────────────────────────────────────────────
        self._stacked = QStackedWidget()
        self._stacked.setObjectName("content_area")

        self.dashboard_page = DashboardPage()
        self.compose_page = ComposePage()
        self.send_page = SendPage()
        self.history_page = HistoryPage()
        self.settings_page = SettingsPage()
        self.about_page = AboutPage()

        self._stacked.addWidget(self.dashboard_page)
        self._stacked.addWidget(self.compose_page)
        self._stacked.addWidget(self.send_page)
        self._stacked.addWidget(self.history_page)
        self._stacked.addWidget(self.settings_page)
        self._stacked.addWidget(self.about_page)

        root.addWidget(self._stacked)

        # ── Signals ──────────────────────────────────────────────────
        self.compose_page.campaign_ready.connect(self._on_campaign_ready)
        self.send_page.send_completed.connect(self._on_send_completed)
        self.settings_page.auth_changed.connect(self._on_auth_changed)
        self.settings_page.theme_changed.connect(self._on_theme_changed)
        self.dashboard_page.go_to_compose.connect(lambda: self._navigate(1))

        if self._user_display:
            self.settings_page.set_auth_status(
                self._access_token, self._user_display, self._user_email
            )

        # ── Status bar ───────────────────────────────────────────────
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready")

    def _navigate(self, page_idx: int):
        self._stacked.setCurrentIndex(page_idx)
        for i, btn in enumerate(self._nav_buttons):
            btn.setProperty("active", "true" if i == page_idx else "false")
            btn.style().unpolish(btn)
            btn.style().polish(btn)

        if page_idx == 0:
            self.dashboard_page.refresh()
        elif page_idx == 3:
            self.history_page.refresh()

    def _on_campaign_ready(self, data: dict):
        if not self._access_token:
            QMessageBox.warning(
                self, "Not Signed In",
                "Please sign in with your Microsoft account in Settings before sending."
            )
            self._navigate(4)
            return

        # Persist campaign to DB
        campaign_id = database.create_campaign(
            name=data["name"],
            subject=data["subject"],
            body=data["body"],
            recipients=data["recipients"],
            attachments=data.get("attachments"),
        )
        self._current_campaign_id = campaign_id

        n = len(data["recipients"])
        summary = (
            f"<b>Campaign:</b> {data['name']}<br>"
            f"<b>Subject:</b> {data['subject']}<br>"
            f"<b>Recipients:</b> {n}<br>"
            f"<b>Attachments:</b> {len(data.get('attachments', []))}<br>"
            f"<b>Account:</b> {self._user_display} ({self._user_email})"
        )
        self.send_page.prepare(campaign_id, self._access_token, summary)
        self._navigate(2)

    def _on_send_completed(self):
        self._navigate(3)  # go to history
        self.history_page.refresh()
        self.compose_page.reset()

    def _on_auth_changed(self, token: str, display: str, email: str):
        self._access_token = token
        self._user_display = display
        self._user_email   = email
        self._update_user_strip()
        if token:
            self.status_bar.showMessage(f"🟢 Signed in as {display} ({email})")
        else:
            self.status_bar.showMessage("Signed out")

    def _on_theme_changed(self, dark: bool):
        self._dark_mode = dark
        self._apply_theme()

    def _update_user_strip(self):
        if self._user_display:
            self._user_strip.setText(f"👤 {self._user_display}")
        else:
            self._user_strip.setText("🔒 Not signed in")

    def _apply_theme(self):
        dark = self._dark_mode
        self.setStyleSheet(get_stylesheet(dark))

    def closeEvent(self, event):
        if hasattr(self, 'send_page') and self.send_page.worker and self.send_page.worker.isRunning():
            ret = QMessageBox.question(
                self, "Sending in Progress",
                "Emails are still being sent. Stop and quit?",
                QMessageBox.Yes | QMessageBox.No
            )
            if ret == QMessageBox.No:
                event.ignore()
                return
            self.send_page.worker.stop()
            self.send_page.worker.wait(3000)
        event.accept()


# ──────────────────────────────────────────────────────────────────────────────
# Entry point
# ──────────────────────────────────────────────────────────────────────────────

def main():
    # Set Windows AppUserModelID so taskbar groups properly and displays the custom icon
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "vikrantyadav.outlookbulkmailer.app.1.0"
            )
        except Exception:
            pass

    app = QApplication(sys.argv)
    app.setApplicationName("Outlook Bulk Mailer")
    app.setOrganizationName("BulkMailer")
    app.setStyle("Fusion")

    # ── App icon ─────────────────────────────────────────────────────
    # Resolve icon relative to this file so it works both when running
    # from source and when bundled with PyInstaller (_MEIPASS).
    _base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    _icon_path = os.path.join(_base, "assets", "icon.png")
    app_icon = None
    if os.path.exists(_icon_path):
        app_icon = QIcon(_icon_path)
        app.setWindowIcon(app_icon)
    # ─────────────────────────────────────────────────────────────────

    window = MainWindow()
    if app_icon and not app_icon.isNull():
        window.setWindowIcon(app_icon)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
