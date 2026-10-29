"""
Background worker threads for non-blocking UI operations.
"""

import time
from PySide6.QtCore import QThread, Signal

from modules import mailer, auth, database


class SendWorker(QThread):
    """
    Sends emails to all pending recipients in a campaign.

    Signals:
        progress(current, total, email, success, error_msg)
        finished(sent_count, failed_count)
        token_expired()
        error(message)
    """

    progress = Signal(int, int, str, bool, str)
    finished = Signal(int, int)
    token_expired = Signal()
    error = Signal(str)

    def __init__(self, campaign_id: int, access_token: str,
                 delay: float = 2.0, test_mode: bool = False,
                 test_limit: int = 3, parent=None):
        super().__init__(parent)
        self.campaign_id = campaign_id
        self.access_token = access_token
        self.delay = delay
        self.test_mode = test_mode
        self.test_limit = test_limit
        self._paused = False
        self._stopped = False

    def pause(self):
        self._paused = True

    def resume(self):
        self._paused = False

    def stop(self):
        self._stopped = True
        self._paused = False

    def run(self):
        try:
            campaign = database.get_campaign(self.campaign_id)
            if not campaign:
                self.error.emit("Campaign not found.")
                return

            recipients = database.get_recipients(self.campaign_id, status="pending")
            attachments = database.get_attachments(self.campaign_id)
            attachment_paths = [a["filepath"] for a in attachments if a["filepath"]]

            if self.test_mode:
                recipients = recipients[: self.test_limit]

            total = len(recipients)
            sent = failed = 0

            for idx, recipient in enumerate(recipients, start=1):
                # ── Pause / stop handling ─────────────────────────
                while self._paused and not self._stopped:
                    time.sleep(0.3)
                if self._stopped:
                    break

                result = mailer.send_email(
                    access_token=self.access_token,
                    to_email=recipient["email"],
                    to_name=recipient.get("name", ""),
                    subject=campaign["subject"],
                    body=campaign["body"],
                    attachment_paths=attachment_paths,
                    delay=self.delay,
                )

                if result["success"]:
                    database.update_recipient_status(recipient["id"], "sent")
                    sent += 1
                    self.progress.emit(idx, total, recipient["email"], True, "")
                else:
                    err  = result.get("error", "Unknown error")
                    code = result.get("code", 0)
                    # 401 = token expired → must re-login before continuing
                    if code == 401:
                        database.update_recipient_status(recipient["id"], "pending", "")
                        self.token_expired.emit()
                        return
                    # 403 = permission denied → stop campaign, not retryable
                    if code == 403:
                        database.update_recipient_status(recipient["id"], "failed", err)
                        self.error.emit(err)
                        return
                    database.update_recipient_status(recipient["id"], "failed", err)
                    failed += 1
                    self.progress.emit(idx, total, recipient["email"], False, err)

            status = "completed" if not self._stopped else "paused"
            database.finish_campaign(self.campaign_id, status)
            self.finished.emit(sent, failed)

        except Exception as e:
            self.error.emit(str(e))


class AuthWorker(QThread):
    """
    Completes the device-code OAuth flow in the background.

    Signals:
        success(access_token, display_name, email)
        error(message)
    """

    success = Signal(str, str, str)
    error = Signal(str)

    def __init__(self, app, flow, cache, parent=None):
        super().__init__(parent)
        self._app = app
        self._flow = flow
        self._cache = cache

    def run(self):
        try:
            result = auth.complete_device_flow(self._app, self._flow, self._cache)
            token = result["access_token"]
            user = auth.get_user_info(token)
            display = user.get("displayName", "Unknown")
            email = user.get("userPrincipalName", user.get("mail", ""))
            self.success.emit(token, display, email)
        except Exception as e:
            self.error.emit(str(e))
