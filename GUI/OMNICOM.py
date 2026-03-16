#!/usr/bin/env python3
import os
import sys
from datetime import datetime

from PyQt5 import uic
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QDialog, QMessageBox

from MiSmSerial import MiSmSerial


UI_FILE = "OMNICOM.ui"

DEFAULT_PORT = "/dev/ttyACM0"
DEFAULT_DEVICE = "FF"
DEFAULT_BAUD = 9600


class OmniCom(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        uic.loadUi(UI_FILE, self)

        self.port = DEFAULT_PORT
        self.device = DEFAULT_DEVICE
        self.baud = DEFAULT_BAUD

        self.history = []
        self.last_response = ""
        self.serial = None

        self.force_buttons = {
            self.pushButton_4: 0,   # Q0
            self.pushButton_8: 1,   # Q1
            self.pushButton_5: 2,   # Q2
            self.pushButton_6: 3,   # Q3
            self.pushButton_7: 4,   # Q4
            self.pushButton_10: 5,  # Q5
            self.pushButton_9: 6,   # Q6
            self.pushButton_11: 7,  # Q7
        }

        self._apply_defaults()
        self._wire_signals()
        self._refresh_preview()

    def _apply_defaults(self):
        self.setWindowTitle("OMNICOM")

        # Model defaults
        self.radioButton.setChecked(True)      # IDEC

        # Operation defaults
        self.radioButton_17.setChecked(True)   # READ
        self.radioButton_21.setChecked(True)   # WORD

        # Static labels already exist in UI, keep them aligned with defaults
        self.label_4.setText(self.device)      # FF
        self.label_5.setText("05h")            # ENQ
        self.label_6.setText("0")              # control
        self.label_12.setText("\\0")           # terminator display

        # Initial field values
        self.plainTextEdit.setPlainText("D8005")
        self.plainTextEdit_2.setPlainText("")

        # Better response appearance
        self.textBrowser_2.setOpenExternalLinks(True)
        self.textBrowser.setOpenExternalLinks(True)

        self._append_response(
            "OMNICOM ready\n"
            f"Port: {self.port}\n"
            f"Device: {self.device}\n"
            f"Baud: {self.baud}\n"
            "Model: IDEC\n"
        )

    def _wire_signals(self):
        # Send / history / help
        self.pushButton.clicked.connect(self.send_command)
        self.pushButton_2.clicked.connect(self.show_history)
        self.pushButton_3.clicked.connect(self.show_help)

        # Clear errors
        self.pushButton_12.clicked.connect(self.clear_all_errors)

        # Preview updates
        self.plainTextEdit.textChanged.connect(self._refresh_preview)
        self.plainTextEdit_2.textChanged.connect(self._refresh_preview)

        self.radioButton_17.toggled.connect(self._refresh_preview)  # READ
        self.radioButton_18.toggled.connect(self._refresh_preview)  # WRITE
        self.radioButton_19.toggled.connect(self._refresh_preview)  # CLEAR

        self.radioButton_20.toggled.connect(self._refresh_preview)  # bit
        self.radioButton_21.toggled.connect(self._refresh_preview)  # word
        self.radioButton_22.toggled.connect(self._refresh_preview)  # float

        # Force IO toggles
        for button, qnum in self.force_buttons.items():
            button.clicked.connect(lambda checked, n=qnum, b=button: self.force_output(n, checked, b))

    def _append_response(self, text):
        stamp = datetime.now().strftime("%H:%M:%S")
        self.textBrowser_2.append(f"[{stamp}] {text}")

    def _add_history(self, entry):
        self.history.append(entry)
        if len(self.history) > 200:
            self.history = self.history[-200:]

    def _current_operation(self):
        if self.radioButton_17.isChecked():
            return "READ"
        if self.radioButton_18.isChecked():
            return "WRITE"
        if self.radioButton_19.isChecked():
            return "CLEAR"
        return "READ"

    def _current_dtype(self):
        if self.radioButton_20.isChecked():
            return "bit"
        if self.radioButton_21.isChecked():
            return "word"
        if self.radioButton_22.isChecked():
            return "float"
        return "word"

    def _current_model(self):
        if self.radioButton.isChecked():
            return "IDEC"
        if self.radioButton_11.isChecked():
            return "Rockwell Automation"
        if self.radioButton_13.isChecked():
            return "Siemens"
        if self.radioButton_2.isChecked():
            return "MetaSYS"
        if self.radioButton_10.isChecked():
            return "GE"
        if self.radioButton_14.isChecked():
            return "ABB"
        if self.radioButton_5.isChecked():
            return "WAGO"
        if self.radioButton_12.isChecked():
            return "HoneyWell"
        if self.radioButton_6.isChecked():
            return "Mitsubishi"
        return "IDEC"

    def _register_text(self):
        return self.plainTextEdit.toPlainText().strip().upper()

    def _value_text(self):
        return self.plainTextEdit_2.toPlainText().strip()

    def _compute_bcc(self, payload: str) -> int:
        """
        Simple XOR BCC over the visible payload string.
        This is for display / construction help.
        If MiSmSerial internally computes BCC differently,
        the preview is still useful for operator visibility.
        """
        bcc = 0
        for ch in payload.encode("ascii", errors="ignore"):
            bcc ^= ch
        return bcc

    def _format_preview_command(self):
        model = self._current_model()
        op = self._current_operation()
        dtype = self._current_dtype()
        reg = self._register_text()
        val = self._value_text()

        if model != "IDEC":
            payload = f"{model}|{op}|{dtype}|{reg}|{val}"
        else:
            if op == "READ":
                payload = f"ENQ {self.device} 0 {op} {dtype} {reg}"
            elif op == "WRITE":
                payload = f"ENQ {self.device} 0 {op} {dtype} {reg} {val}"
            else:
                payload = f"ENQ {self.device} 0 CLEAR_ERRORS D8005.15"

        bcc = self._compute_bcc(payload)
        return payload, bcc

    def _refresh_preview(self):
        payload, bcc = self._format_preview_command()
        self.textBrowser.setText(payload)
        self.lcdNumber.display(f"{bcc:02X}")

    def _open_serial(self):
        if self.serial is None:
            self.serial = MiSmSerial(
                self.port,
                device=self.device,
                baud=self.baud,
                debug=False,
                bcc_mode="auto",
            )
        return self.serial

    def _close_serial(self):
        if self.serial is not None:
            try:
                self.serial.close()
            except Exception:
                pass
            self.serial = None

    def _parse_value(self, text, dtype):
        if dtype == "bit":
            if text == "":
                raise ValueError("bit writes require a value of 0 or 1")
            if text not in ("0", "1"):
                raise ValueError("bit value must be 0 or 1")
            return int(text)

        if dtype == "word":
            if text == "":
                raise ValueError("word writes require a numeric value")
            return int(text, 0)

        if dtype == "float":
            if text == "":
                raise ValueError("float writes require a numeric value")
            return float(text)

        raise ValueError(f"unsupported dtype: {dtype}")

    def _ensure_idec(self):
        model = self._current_model()
        if model != "IDEC":
            raise RuntimeError(
                f"{model} is still a placeholder in this build. "
                "Only IDEC actions are wired right now."
            )

    def send_command(self):
        reg = self._register_text()
        val_text = self._value_text()
        op = self._current_operation()
        dtype = self._current_dtype()

        payload, bcc = self._format_preview_command()
        history_entry = f"{payload} | BCC={bcc:02X}"
        self._add_history(history_entry)

        try:
            self._ensure_idec()
            plc = self._open_serial()

            if op == "READ":
                if not reg:
                    raise ValueError("register field is empty")

                if dtype == "bit":
                    result = plc.read_bit(reg)
                elif dtype == "word":
                    result = plc.read(reg)
                elif dtype == "float":
                    if not hasattr(plc, "read_float"):
                        raise AttributeError("MiSmSerial has no read_float()")
                    result = plc.read_float(reg)
                else:
                    raise ValueError(f"unsupported read dtype: {dtype}")

                self.last_response = f"{reg} = {result}"
                self._append_response(self.last_response)

            elif op == "WRITE":
                if not reg:
                    raise ValueError("register field is empty")

                value = self._parse_value(val_text, dtype)

                if dtype == "bit":
                    plc.write_bit(reg, value)
                elif dtype == "word":
                    plc.write(reg, value)
                elif dtype == "float":
                    if not hasattr(plc, "write_float"):
                        raise AttributeError("MiSmSerial has no write_float()")
                    plc.write_float(reg, value)
                else:
                    raise ValueError(f"unsupported write dtype: {dtype}")

                self.last_response = f"WROTE {value} -> {reg}"
                self._append_response(self.last_response)

            elif op == "CLEAR":
                self._clear_errors_impl()
            else:
                raise ValueError(f"unknown operation: {op}")

        except Exception as e:
            self.last_response = f"ERROR: {e}"
            self._append_response(self.last_response)

    def _clear_errors_impl(self):
        plc = self._open_serial()

        # Try the documented clear-errors pulse first
        errs = []
        attempts = [
            ("D8005.15", 1),
            ("D8005.15", 0),
        ]

        for addr, value in attempts:
            try:
                plc.write_bit(addr, value)
            except Exception as e:
                errs.append(f"{addr}={value}: {e}")

        if errs:
            self._append_response(
                "Clear ALL attempted, but some writes failed:\n" + "\n".join(errs)
            )
        else:
            self._append_response("Clear ALL attempted using D8005.15 pulse")

        # Read back a few common error indicators if possible
        readback = []
        for addr in ("D8005", "D8006", "M8004", "M8005"):
            try:
                if addr.startswith("M"):
                    rv = plc.read_bit(addr)
                else:
                    rv = plc.read(addr)
                readback.append(f"{addr}={rv}")
            except Exception as e:
                readback.append(f"{addr}=READ_ERROR({e})")

        self._append_response("Post-clear status: " + ", ".join(readback))

    def clear_all_errors(self):
        try:
            self._ensure_idec()
            self._open_serial()
            self._clear_errors_impl()
        except Exception as e:
            self._append_response(f"ERROR clearing: {e}")

    def force_output(self, qnum, checked, button):
        try:
            self._ensure_idec()
            plc = self._open_serial()

            plc.output(qnum, 1 if checked else 0)
            state = "ON" if checked else "OFF"
            self._append_response(f"Forced Q{qnum} {state}")

        except Exception as e:
            button.blockSignals(True)
            button.setChecked(not checked)
            button.blockSignals(False)
            self._append_response(f"ERROR forcing Q{qnum}: {e}")

    def show_history(self):
        if not self.history:
            self._append_response("History is empty")
            return

        self._append_response("Command history:\n" + "\n".join(self.history[-20:]))

    def show_help(self):
        msg = QMessageBox(self)
        msg.setWindowTitle("OMNICOM Help")
        msg.setTextFormat(Qt.RichText)
        msg.setText(
            "<b>OMNICOM</b><br><br>"
            "This app builds and sends IDEC-oriented commands using MiSmSerial.<br><br>"
            "<b>Defaults</b><br>"
            f"Port: {self.port}<br>"
            f"Device: {self.device}<br>"
            f"Baud: {self.baud}<br><br>"
            "<b>Notes</b><br>"
            "- IDEC is the only live model right now.<br>"
            "- The other model radio buttons are placeholders.<br>"
            "- Clear ALL tries to pulse D8005.15.<br>"
            "- Force IO toggles Q0 through Q7.<br><br>"
            "Populate docs and GitHub links later in this dialog."
        )
        msg.exec_()

    def closeEvent(self, event):
        self._close_serial()
        super().closeEvent(event)


def main():
    app = QApplication(sys.argv)

    # Let user run from anywhere as long as UI sits beside script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    if not os.path.exists(UI_FILE):
        raise FileNotFoundError(
            f"Could not find {UI_FILE} in {script_dir}"
        )

    dlg = OmniCom()
    dlg.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
