#!/usr/bin/env python3
import os
import sys
from datetime import datetime

from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QButtonGroup, QDialog, QMessageBox

from MiSmSerial import MiSmSerial


UI_FILE = "OMNICOM.ui"

DEFAULT_PORT = "/dev/ttyACM0"
DEFAULT_DEVICE = "FF"
DEFAULT_BAUD = 9600


GENERAL_ERROR_BITS = {
    0: "POWER_FAIL",
    1: "TIMER_ERROR",
    2: "DATALINK_CONNECTION_ERROR",
    3: "USER_PROGRAM_ROM_CRC_ERROR",
    4: "TIMER_COUNTER_PRESET_CHANGE_ERROR",
    5: "RESERVED",
    6: "KEEP_DATA_SUM_CHECK_ERROR",
    7: "USER_PROGRAM_SYNTAX_ERROR",
    8: "USER_PROGRAM_DOWNLOAD_ERROR",
    9: "SYSTEM_ERROR",
    10: "CLOCK_ERROR",
    11: "EXPANSION_BUS_INITIALIZATION_ERROR",
    12: "SD_MEMORY_CARD_TRANSFER_ERROR",
    13: "USER_PROGRAM_EXECUTION_ERROR",
    14: "SD_MEMORY_CARD_ACCESS_ERROR",
    15: "CLEAR_ERRORS_CONTROL_BIT",
}

USER_EXECUTION_ERRORS = {
    0: "NO_ERROR",
    1: "Source/destination device exceeds range",
    2: "MUL result exceeds data type range",
    3: "DIV result exceeds data type range, or division by 0",
    4: "BCDLS has S1 or S1+1 exceeding 9999",
    5: "HTOB input too large",
    6: "BTOH has a digit exceeding 9",
    7: "HTOA/ATOH/BTOA/ATOB digit count out of range",
    8: "ATOH/ATOB has non-ASCII data",
    9: "WEEK instruction time data out of range",
    10: "YEAR instruction date data out of range",
    11: "DGRD range exceeded",
    12: "CVXTY/CVYTX executed without matching XYFS",
    13: "CVXTY/CVYTX S2 exceeds value specified in XYFS",
    14: "Label in LJMP, LCAL, or DJNZ not found",
    16: "PID/PIDA instruction execution error",
    18: "Instruction cannot be used in interrupt program",
    19: "Instruction not available for this PLC",
    20: "Pulse output instruction has invalid values",
    21: "DECO has S1 exceeding 255",
    22: "BCNT has S2 exceeding 256",
    23: "ICMP>= has S1 < S3",
    25: "BCDLS has S2 exceeding 7",
    26: "Interrupt input or timer interrupt not programmed",
    27: "Work area broken",
    28: "Instruction/source invalid",
    29: "Float/data type instruction result exceeds range",
    30: "SFTL/SFTR exceeds valid range",
    31: "FOEX/FIFO used before FIFO file registered",
    32: "TADD, TSUB, HOUR, or HTOS has invalid source data",
    33: "RNDM has invalid data",
    34: "NDSRC has invalid source data",
    35: "SUM result exceeds valid range",
    36: "CSV file exceeds maximum size",
    41: "SD memory card is write protected",
    42: "A script failed",
    46: "SCALE instruction out of range",
    48: "Pulse collisions / timing errors",
    49: "Pulse output not initialized properly",
}


class OmniCom(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        uic.loadUi(UI_FILE, self)

        self.port = DEFAULT_PORT
        self.device = DEFAULT_DEVICE
        self.baud = DEFAULT_BAUD
        self.serial = None
        self.history = []

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

        self._build_button_groups()
        self._apply_defaults()
        self._wire_signals()
        self._refresh_preview()

    def _build_button_groups(self):
        # Model group
        self.model_group = QButtonGroup(self)
        for btn in [
            self.radioButton,
            self.radioButton_11,
            self.radioButton_13,
            self.radioButton_2,
            self.radioButton_10,
            self.radioButton_14,
            self.radioButton_5,
            self.radioButton_12,
            self.radioButton_6,
        ]:
            self.model_group.addButton(btn)

        # Operation group
        self.operation_group = QButtonGroup(self)
        for btn in [self.radioButton_17, self.radioButton_18, self.radioButton_19]:
            self.operation_group.addButton(btn)

        # Data type group
        self.dtype_group = QButtonGroup(self)
        for btn in [self.radioButton_20, self.radioButton_21, self.radioButton_22]:
            self.dtype_group.addButton(btn)

    def _apply_defaults(self):
        self.setWindowTitle("OMNICOM")

        # Defaults requested
        self.radioButton.setChecked(True)       # IDEC
        self.radioButton_17.setChecked(True)    # READ
        self.radioButton_21.setChecked(True)    # WORD

        # Static labels if present
        if hasattr(self, "label_4"):
            self.label_4.setText(self.device)
        if hasattr(self, "label_5"):
            self.label_5.setText("05h")
        if hasattr(self, "label_6"):
            self.label_6.setText("0")
        if hasattr(self, "label_12"):
            self.label_12.setText("\\0")

        if hasattr(self, "plainTextEdit") and not self.plainTextEdit.toPlainText().strip():
            self.plainTextEdit.setPlainText("D8005")

        for button in self.force_buttons:
            button.setCheckable(True)
            button.setMinimumHeight(42)

        if hasattr(self, "pushButton_12"):
            self.pushButton_12.setMinimumHeight(42)

        self._append_response(
            "OMNICOM ready\n"
            f"Port   : {self.port}\n"
            f"Device : {self.device}\n"
            f"Baud   : {self.baud}\n"
            "Model  : IDEC"
        )

    def _wire_signals(self):
        self.pushButton.clicked.connect(self.send_command)
        self.pushButton_2.clicked.connect(self.show_history)
        self.pushButton_3.clicked.connect(self.show_help)
        self.pushButton_12.clicked.connect(self.clear_all_errors)

        # Check button
        self.pushButton_14.clicked.connect(self.read_checked_registers)

        self.plainTextEdit.textChanged.connect(self._refresh_preview)
        self.plainTextEdit_2.textChanged.connect(self._refresh_preview)

        self.radioButton_17.toggled.connect(self._refresh_preview)
        self.radioButton_18.toggled.connect(self._refresh_preview)
        self.radioButton_19.toggled.connect(self._refresh_preview)

        self.radioButton_20.toggled.connect(self._refresh_preview)
        self.radioButton_21.toggled.connect(self._refresh_preview)
        self.radioButton_22.toggled.connect(self._refresh_preview)

        for button, qnum in self.force_buttons.items():
            button.clicked.connect(
                lambda checked, n=qnum, b=button: self.force_output(n, checked, b)
            )

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

    def _append_response(self, text):
        stamp = datetime.now().strftime("%H:%M:%S")
        self.textBrowser_2.append(f"[{stamp}] {text}")

    def _add_history(self, line):
        self.history.append(line)
        if len(self.history) > 300:
            self.history = self.history[-300:]

    def _register_text(self):
        return self.plainTextEdit.toPlainText().strip().upper()

    def _value_text(self):
        return self.plainTextEdit_2.toPlainText().strip()

    def _current_model(self):
        mapping = {
            self.radioButton: "IDEC",
            self.radioButton_11: "Rockwell Automation",
            self.radioButton_13: "Siemens",
            self.radioButton_2: "MetaSYS",
            self.radioButton_10: "GE",
            self.radioButton_14: "ABB",
            self.radioButton_5: "WAGO",
            self.radioButton_12: "HoneyWell",
            self.radioButton_6: "Mitsubishi",
        }
        for btn, name in mapping.items():
            if btn.isChecked():
                return name
        return "IDEC"

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

    def _ensure_idec(self):
        if self._current_model() != "IDEC":
            raise RuntimeError(
                f"{self._current_model()} is still a placeholder. Only IDEC is live right now."
            )

    def _compute_bcc(self, payload: str) -> int:
        bcc = 0
        for ch in payload.encode("ascii", errors="ignore"):
            bcc ^= ch
        return bcc

    def _format_preview_command(self):
        reg = self._register_text()
        val = self._value_text()
        op = self._current_operation()
        dtype = self._current_dtype()

        if op == "READ":
            payload = f"ENQ {self.device} 0 READ {dtype} {reg}"
        elif op == "WRITE":
            payload = f"ENQ {self.device} 0 WRITE {dtype} {reg} {val}"
        else:
            payload = "ENQ FF 0 CLEAR D8005.15"

        return payload, self._compute_bcc(payload)

    def _refresh_preview(self):
        payload, bcc = self._format_preview_command()
        self.textBrowser.setText(payload)
        self.lcdNumber.display(f"{bcc:02X}")

    def _safe_read_word(self, plc, reg):
        try:
            return plc.read(reg)
        except Exception as e:
            return f"READ_ERROR: {e}"

    def _safe_read_bit(self, plc, reg):
        try:
            return plc.read_bit(reg)
        except Exception as e:
            return f"READ_ERROR: {e}"

    def _safe_write_bit(self, plc, reg, value):
        try:
            plc.write_bit(reg, value)
            return True, None
        except Exception as e:
            return False, str(e)

    def _parse_value(self, text, dtype):
        if dtype == "bit":
            if text not in ("0", "1"):
                raise ValueError("bit value must be 0 or 1")
            return int(text)

        if dtype == "word":
            return int(text, 0)

        if dtype == "float":
            return float(text)

        raise ValueError(f"unsupported type: {dtype}")

    def bits16(self, value):
        return f"{value:016b}"

    def decode_d8005(self, value):
        active = []
        for bit in range(16):
            if (value >> bit) & 1:
                active.append((bit, GENERAL_ERROR_BITS.get(bit, "UNKNOWN")))
        return active

    def user_exec_text(self, code):
        return USER_EXECUTION_ERRORS.get(code, "Unknown execution error code")

    def battery_text(self, value):
        if value == 65535:
            return "not initialized yet"
        if value == 0:
            return "measurement error or no battery"
        return f"{value} mV"

    def _report_read_result(self, reg, value):
        if reg == "D8029" and isinstance(value, int):
            self._append_response(f"{reg}: {value / 100:.2f}")
            return

        if reg == "D8005" and isinstance(value, int):
            self._append_response(f"{reg}: {value} bits:{self.bits16(value)}")
            decoded = self.decode_d8005(value)
            if decoded:
                for bit, name in decoded:
                    self._append_response(f"  bit {bit}: {name}")
            else:
                self._append_response("  no active general error bits")
            return

        if reg == "D8006" and isinstance(value, int):
            self._append_response(f"{reg}: {value}")
            self._append_response(f"  {self.user_exec_text(value)}")
            return

        if reg == "D8056" and isinstance(value, int):
            self._append_response(f"{reg}: {self.battery_text(value)}")
            return

        self._append_response(f"{reg}: {value}")

    def send_command(self):
        reg = self._register_text()
        val_text = self._value_text()
        op = self._current_operation()
        dtype = self._current_dtype()

        payload, bcc = self._format_preview_command()
        self._add_history(f"{payload} | BCC={bcc:02X}")

        try:
            self._ensure_idec()
            plc = self._open_serial()

            if op == "READ":
                if not reg:
                    raise ValueError("register is empty")

                if dtype == "bit":
                    value = plc.read_bit(reg)
                elif dtype == "word":
                    value = plc.read(reg)
                elif dtype == "float":
                    if hasattr(plc, "read_float"):
                        value = plc.read_float(reg)
                    else:
                        raise AttributeError("MiSmSerial has no read_float()")
                else:
                    raise ValueError("unsupported read type")

                self._report_read_result(reg, value)

            elif op == "WRITE":
                if not reg:
                    raise ValueError("register is empty")

                value = self._parse_value(val_text, dtype)

                if dtype == "bit":
                    plc.write_bit(reg, value)
                elif dtype == "word":
                    plc.write(reg, value)
                elif dtype == "float":
                    if hasattr(plc, "write_float"):
                        plc.write_float(reg, value)
                    else:
                        raise AttributeError("MiSmSerial has no write_float()")
                else:
                    raise ValueError("unsupported write type")

                self._append_response(f"WROTE {value} -> {reg}")

            elif op == "CLEAR":
                self._clear_errors_impl()

            else:
                raise ValueError("unknown operation")

        except Exception as e:
            self._append_response(f"ERROR: {e}")

    def _clear_errors_impl(self):
        plc = self._open_serial()

        ok1, err1 = self._safe_write_bit(plc, "D8005.15", 1)
        ok2, err2 = self._safe_write_bit(plc, "D8005.15", 0)

        if ok1 and ok2:
            self._append_response("Clear ALL attempted with D8005.15 pulse")
        else:
            self._append_response(
                "Clear ALL attempted but write failed: "
                f"{err1 or ''} {err2 or ''}".strip()
            )

        post = {}
        for addr in ("D8005", "D8006", "M8004", "M8005"):
            try:
                if addr.startswith("M"):
                    post[addr] = plc.read_bit(addr)
                else:
                    post[addr] = plc.read(addr)
            except Exception as e:
                post[addr] = f"READ_ERROR: {e}"

        for k, v in post.items():
            self._append_response(f"{k}: {v}")

    def clear_all_errors(self):
        try:
            self._ensure_idec()
            self._clear_errors_impl()
        except Exception as e:
            self._append_response(f"ERROR clearing errors: {e}")

    def force_output(self, qnum, checked, button):
        try:
            self._ensure_idec()
            plc = self._open_serial()

            if hasattr(plc, "output"):
                plc.output(qnum, 1 if checked else 0)
            else:
                plc.write_bit(f"Q{qnum}", 1 if checked else 0)

            self._append_response(f"Q{qnum} -> {'ON' if checked else 'OFF'}")

        except Exception as e:
            button.blockSignals(True)
            button.setChecked(not checked)
            button.blockSignals(False)
            self._append_response(f"ERROR forcing Q{qnum}: {e}")

    def show_history(self):
        if not self.history:
            self._append_response("History is empty")
            return

        self._append_response("---- Command History ----")
        for line in self.history[-20:]:
            self._append_response(line)

    def show_help(self):
        msg = QMessageBox(self)
        msg.setWindowTitle("OMNICOM Help")
        msg.setText(
            "OMNICOM\n\n"
            "IDEC command constructor / service panel.\n\n"
            f"Port: {self.port}\n"
            f"Device: {self.device}\n"
            f"Baud: {self.baud}\n\n"
            "Check reads the selected status items into the response box.\n"
            "Clear ALL pulses D8005.15.\n"
            "Force IO toggles Q0..Q7."
        )
        msg.exec_()

    def read_checked_registers(self):
        try:
            self._ensure_idec()
            plc = self._open_serial()

            self.textBrowser_2.clear()
            self._append_response("---- Selected Items ----")

            selected = []

            # Date
            if self.checkBox.isChecked():
                selected.append(("Date", "date", "D8015"))

            # Security
            # Placeholder mapping for now
            if self.checkBox_2.isChecked():
                selected.append(("Security", "bit", "M8004"))

            # Running
            if self.checkBox_3.isChecked():
                selected.append(("Running", "bit", "M8000"))

            # CPU
            if self.checkBox_4.isChecked():
                selected.append(("CPU", "word", "D8005"))

            # Firmware
            if self.checkBox_5.isChecked():
                selected.append(("Firmware", "word", "D8029"))

            # Battery
            if self.checkBox_6.isChecked():
                selected.append(("Battery", "word", "D8056"))

            if not selected:
                self._append_response("No checkboxes selected")
                return

            for label, dtype, reg in selected:
                if dtype == "bit":
                    value = plc.read_bit(reg)
                    self._append_response(f"{label}: {value}")
                    continue

                if dtype == "word":
                    value = plc.read(reg)

                    if reg == "D8029" and isinstance(value, int):
                        self._append_response(f"{label}: {value / 100:.2f}")
                    elif reg == "D8056" and isinstance(value, int):
                        self._append_response(f"{label}: {self.battery_text(value)}")
                    elif reg == "D8005" and isinstance(value, int):
                        self._append_response(f"{label}: {value} bits:{value:016b}")
                        decoded = self.decode_d8005(value)
                        if decoded:
                            for bit, name in decoded:
                                self._append_response(f"  bit {bit}: {name}")
                        else:
                            self._append_response("  no active CPU/general error bits")
                    else:
                        self._append_response(f"{label}: {value}")
                    continue

                if dtype == "date":
                    try:
                        year = plc.read("D8015")
                        month = plc.read("D8016")
                        day = plc.read("D8017")
                        weekday = plc.read("D8018")
                        hour = plc.read("D8019")
                        minute = plc.read("D8020")
                        second = plc.read("D8021")

                        self._append_response(
                            f"{label}: 20{year:02d}-{month:02d}-{day:02d} "
                            f"{hour:02d}:{minute:02d}:{second:02d} weekday:{weekday}"
                        )
                    except Exception as e:
                        self._append_response(f"{label}: READ_ERROR: {e}")

        except Exception as e:
            self._append_response(f"ERROR: {e}")

    def closeEvent(self, event):
        self._close_serial()
        super().closeEvent(event)


def main():
    app = QApplication(sys.argv)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    if not os.path.exists(UI_FILE):
        raise FileNotFoundError(f"Could not find {UI_FILE} in {script_dir}")

    dlg = OmniCom()
    dlg.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
