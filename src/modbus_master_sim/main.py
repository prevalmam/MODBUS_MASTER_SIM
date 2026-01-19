import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import serial
import serial.tools.list_ports
import struct
import threading
import queue
import sys
import pandas as pd
import os
from importlib import resources as importlib_resources

# --- 通信用キューとスレッド ---
serial_task_queue = queue.Queue()
root = None  # Late-initialized Tk root shared across callbacks

def serial_worker():
    while True:
        try:
            func, args, kwargs = serial_task_queue.get()
            func(*args, **kwargs)
        except Exception as e:
            print("[Worker Error]", e)
        finally:
            serial_task_queue.task_done()

threading.Thread(target=serial_worker, daemon=True).start()


def _set_window_icon(window):
    """Apply the packaged ICO to the Tk window when available."""
    try:
        icon_resource = importlib_resources.files("modbus_master_sim").joinpath("icons/RegiStar.ico")
        with importlib_resources.as_file(icon_resource) as icon_file:
            window.iconbitmap(default=str(icon_file))
    except (FileNotFoundError, tk.TclError):
        pass

# --- CRC16計算（Modbus用） ---
def calc_crc(data):
    crc = 0xFFFF
    for pos in data:
        crc ^= pos
        for _ in range(8):
            if crc & 1:
                crc = (crc >> 1) ^ 0xA001
            else:
                crc >>= 1
    return crc & 0xFFFF

# --- Excelパース処理 ---
def extract_registers_from_excel(path):
    reg_df = pd.read_excel(path, sheet_name="RegisterTable", header=None)
    len_df = pd.read_excel(path, sheet_name="LengthDefs", header=None)

    length_defs = {}
    for i in range(4, len(len_df)):
        row = len_df.iloc[i]
        if str(row[1]).strip().upper() == "EOF":
            break
        macro = row[2]
        value = row[3]
        if pd.notna(macro) and pd.notna(value):
            try:
                length_defs[str(macro).strip()] = int(value)
            except ValueError:
                continue

    for i in range(len(reg_df)):
        if str(reg_df.iloc[i, 2]).strip() == "Reg_Addr":
            start_row = i + 1
            break

    reg_list = []
    for i in range(start_row, len(reg_df)):
        row = reg_df.iloc[i]
        if str(row[1]).strip().upper() == "EOF":
            break
        if pd.isna(row[2]) or pd.isna(row[3]) or pd.isna(row[4]) or pd.isna(row[5]) or pd.isna(row[6]):
            continue

        try:
            reg_addr = int(row[2])
            var_name = str(row[3]).strip()
            var_type = str(row[4]).strip()
            array_len_raw = str(row[5]).strip()
            access = str(row[6]).strip().upper()

            try:
                length = int(array_len_raw)
            except ValueError:
                if array_len_raw in length_defs:
                    length = length_defs[array_len_raw]
                else:
                    continue
        except (ValueError, TypeError):
            continue

        reg_list.append({
            "name": var_name,
            "addr": reg_addr,
            "type": var_type,
            "length": length,
            "access": access,
            "display": f"{reg_addr} {var_name}"
        })

    return reg_list

# --- 最小限GUIクラス雛形（後で拡張） ---
class ModbusMasterGUI:
    def __init__(self, root, reg_table):  # ← 引数 reg_table を追加
        self.root = root
        self.serial_port = None
        self.reg_table = reg_table        # ← メンバに保存
        self.slave_addr = 1
        self.baudrate = 57600
        self.slave_addr_var = tk.StringVar(value=str(self.slave_addr))
        self.baudrate_var = tk.StringVar(value=str(self.baudrate))

        self.polling_widgets = []
        self.polling_index = 0
        self._polling_active = False
        self._polling_task_id = None
        self.root.title("RegiStar - レジスター GUI")
        self.build_gui()
        self.init_polling_gui()

    def build_gui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=1)
        self.root.columnconfigure(2, weight=1)
        self.root.rowconfigure(1, weight=1)

        top_frame = ttk.Frame(self.root)
        top_frame.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        top_frame.columnconfigure(0, weight=0)
        top_frame.columnconfigure(1, weight=0)
        top_frame.columnconfigure(2, weight=0)
        top_frame.columnconfigure(3, weight=1)
        top_frame.columnconfigure(4, weight=0)

        ttk.Label(top_frame, text="Slave Addr:").grid(row=0, column=0, padx=2, sticky="w")

        vcmd = (self.root.register(self._validate_slave_addr_input), "%P")
        self.slave_addr_entry = ttk.Entry(
            top_frame,
            textvariable=self.slave_addr_var,
            width=6,
            justify="right",
            validate="key",
            validatecommand=vcmd,
        )
        self.slave_addr_entry.grid(row=0, column=1, padx=2, sticky="ew")

        self.baudrate_combo = ttk.Combobox(
            top_frame,
            textvariable=self.baudrate_var,
            values=self.get_baudrate_values(),
            state="readonly",
            width=8,
        )
        self.baudrate_combo.grid(row=0, column=2, padx=2, sticky="ew")

        self.port_combo = ttk.Combobox(top_frame, values=self.get_serial_ports(), state="readonly")
        self.port_combo.grid(row=0, column=3, padx=2, sticky="ew")

        ttk.Button(top_frame, text="Connect", command=self.connect_serial).grid(
            row=0, column=4, padx=2, sticky="ew"
        )

        self.reg_listbox = tk.Listbox(self.root, height=8)
        for reg in self.reg_table:
            self.reg_listbox.insert(tk.END, reg['display'])
        self.reg_listbox.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        self.reg_listbox.bind("<<ListboxSelect>>", self.on_reg_select)

        self.value_frame = tk.Frame(self.root)
        self.value_frame.grid(row=2, column=0, columnspan=2, sticky="ew")

        self.btn_frame = tk.Frame(self.root)
        self.btn_frame.grid(row=3, column=0, columnspan=2, pady=5)
        # Readボタン
        self.read_btn = ttk.Button(self.btn_frame, text="Read", command=self.on_read_button_pressed)
        # Write Singleボタン
        self.write_single_btn = ttk.Button(self.btn_frame, text="Write (1)", command=self.on_write_single_button_pressed)
        # Write Multiボタン
        self.write_multi_btn = ttk.Button(self.btn_frame, text="Write (N)", command=self.on_write_multi_button_pressed)
        self.read_btn.grid(row=0, column=0, padx=5)
        self.write_single_btn.grid(row=0, column=1, padx=5)
        self.write_multi_btn.grid(row=0, column=2, padx=5)

        self.log_area = scrolledtext.ScrolledText(self.root, state="disabled")
        self.log_area.grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")

        ttk.Button(self.root, text="Reset", command=self.reset_app).grid(row=5, column=0, columnspan=2, pady=5)

    def reset_app(self):
        self.log("[Info] Resetting application...")
        self.root.update()
        python = sys.executable
        os.execl(python, python, *sys.argv)

    def on_reg_select(self, event):
        selection = event.widget.curselection()
        if not selection:
            return
        index = selection[0]
        self.current_reg = self.reg_table[index]
        self.update_buttons_and_inputs()

    def update_buttons_and_inputs(self):
        for widget in self.value_frame.winfo_children():
            widget.destroy()
        self.input_entries = []

        if self.current_reg is None:
            return

        access = self.current_reg['access']
        typ = self.current_reg['type']
        length = self.current_reg['length']

        is_array = length > 1

        self.read_btn.config(state=tk.NORMAL if access in ["R", "RW"] else tk.DISABLED)
        self.write_single_btn.config(state=tk.NORMAL if (access in ["W", "RW"] and typ == "uint16_t" and not is_array) else tk.DISABLED)
        self.write_multi_btn.config(state=tk.NORMAL if access in ["W", "RW"] else tk.DISABLED)

        if access in ["W", "RW"]:
            for i in range(length):
                entry = ttk.Entry(self.value_frame, width=8)
                entry.insert(0, "0")
                row = i // 10
                col = i % 10
                entry.grid(row=row, column=col, padx=2, pady=2)
                self.input_entries.append(entry)

    def get_serial_ports(self):
        self.port_display_to_device = {}
        display_list = []
        for port in serial.tools.list_ports.comports():
            device = port.device
            description = (port.description or "").strip()
            if description and description != device:
                display = f"{device} ({description})"
            else:
                display = device
            self.port_display_to_device[display] = device
            display_list.append(display)
        return display_list

    def get_baudrate_values(self):
        return ["1200", "2400", "4800", "9600", "19200", "38400", "57600", "115200"]

    def _validate_slave_addr_input(self, proposed):
        if proposed == "":
            return True
        return proposed.isdigit()

    def connect_serial(self):
        port_display = self.port_combo.get()
        if not port_display:
            messagebox.showerror("Error", "Select a serial port.")
            return
        port = self.port_display_to_device.get(port_display, port_display)
        addr_raw = self.slave_addr_var.get().strip()
        if not addr_raw:
            messagebox.showerror("Error", "Enter a slave address (0-247).")
            return
        try:
            addr = int(addr_raw)
        except ValueError:
            messagebox.showerror("Error", "Slave address must be a number.")
            return
        if addr < 0 or addr > 247:
            messagebox.showerror("Error", "Slave address must be in range 0-247.")
            return
        baud_raw = self.baudrate_combo.get().strip()
        if not baud_raw:
            messagebox.showerror("Error", "Select a baudrate.")
            return
        try:
            baud = int(baud_raw)
        except ValueError:
            messagebox.showerror("Error", "Baudrate must be a number.")
            return
        if self.serial_port and self.serial_port.is_open:
            try:
                self.serial_port.close()
            except Exception:
                pass
        try:
            self.serial_port = serial.Serial(port, baudrate=baud, timeout=1)
            self.slave_addr = addr
            self.baudrate = baud
            messagebox.showinfo("Connected", f"Connected to {port}")
        except Exception as e:
            messagebox.showerror("Connection Failed", str(e))

    def on_read_button_pressed(self):
        selection = self.reg_listbox.curselection()
        if not selection:
            self.log("[Error] No register selected.")
            return

        index = selection[0]
        self.current_reg = self.reg_table[index]
        self.update_buttons_and_inputs()  # エントリ更新＆ボタン有効化

        addr = self.current_reg['addr']
        length = self.current_reg['length'] * (2 if self.current_reg['type'] in ["float", "uint32_t"] else 1)

        self.log(f"\n[Send] → Read Holding Register: Addr=0x{addr:04X}, Count={length}")
        queue_send_read(self.serial_port, self.slave_addr, addr, length, self.handle_read_result)


    def handle_read_result(self, data):
        self.log("\n[Recv Result]")
        if not data:
            self.log("\u2192 No Response")
            return

        self.log(f"\u2192 Raw: {data.hex().upper()}")

        try:
            if data[1] & 0x80:
                err = data[2] if len(data) > 2 else None
                self.log(f"\u2192 Exception Response: Func=0x{data[1]:02X}, Code=0x{err:02X} (if present)")
                return

            byte_count = data[2]
            values = data[3:3 + byte_count]
            typ = self.current_reg['type']

            # --- フォーマット関数で整形 ---
            formatted = self.format_read_values(typ, values)

            for i, val in enumerate(formatted):
                self.log(f"\u2192 [{i}] {val}")

        except Exception as e:
            self.log(f"[Decode Error] {str(e)}")


    def format_read_values(self, typ, values):
        formatted = []
        step = 4 if typ in ["float", "uint32_t"] else 2
        for i in range(0, len(values), step):
            try:
                if typ == "uint16_t":
                    val = struct.unpack('>H', bytes(values[i:i + 2]))[0]
                    formatted.append(f"0x{val:04X} ({val})")
                elif typ == "uint32_t":
                    val = struct.unpack('>I', bytes(values[i:i + 4]))[0]
                    formatted.append(f"0x{val:08X} ({val})")
                elif typ == "float":
                    fval = struct.unpack('>f', bytes(values[i:i + 4]))[0]
                    formatted.append(f"{fval:.4f}")
            except:
                formatted.append("?")
        return formatted


    def log(self, text):
        self.log_area.config(state="normal")
        self.log_area.insert(tk.END, text + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state="disabled")

    def on_write_single_button_pressed(self):
        if not self.current_reg:
            self.log("[Error] レジスタが選択されていません。")
            return

        try:
            val = int(float(self.input_entries[0].get()))
        except ValueError:
            messagebox.showerror("Error", "無効な入力値です。")
            return

        addr = self.current_reg['addr']
        queue_send_write_single(
            self.serial_port, self.slave_addr, addr, val, self.handle_write_single_result
        )

    def handle_write_single_result(self, data):
        self.log("\n[Write Single Result]")
        if not data:
            self.log("→ No Response")
        elif data[1] & 0x80:
            # Modbus exception response: MSB set in function code
            code = data[2] if len(data) > 2 else 0
            self.log(f"→ Exception Response: Func=0x{data[1]:02X}, Code=0x{code:02X} (if present)")
        else:
            self.log(f"→ ACK: {data.hex().upper()}")

    def on_polling_read_test(self):
        reg = {"name": "TEMP", "addr": 0x0002, "length": 1, "type": "uint16_t"}
        queue_send_read_for(self.serial_port, self.slave_addr, reg, self.handle_polling_result)

    def handle_polling_result(self, reg, data):
        self.log(f"\n[Polling Result] {reg['name']}")
        if not data:
            self.log("→ No Response")
        else:
            self.log(f"→ Value: {data}")    

    def on_write_multi_button_pressed(self):
        if not self.current_reg:
            self.log("[Error] レジスタが選択されていません。")
            return

        values = []
        try:
            for e in self.input_entries:
                v = float(e.get())
                values.append(v)
        except ValueError:
            messagebox.showerror("Error", "無効な入力値が含まれています。")
            return

        addr = self.current_reg['addr']
        typ = self.current_reg['type']

        queue_send_write_multi(
            self.serial_port, self.slave_addr, addr, values, typ, self.handle_write_multi_result
        )

    def handle_write_multi_result(self, data):
        self.log("\n[Write Multi Result]")

        if not data:
            self.log("→ No Response")
            return

        # Check for Exception Response (Function code >= 0x80)
        if len(data) >= 3 and data[1] & 0x80:
            func_code = data[1]
            ex_code = data[2]
            self.log(f"→ Exception Response: Func=0x{func_code:02X}, Code=0x{ex_code:02X}")
            return

        # Normal ACK expected to be 8 bytes: slave + func + addr (2B) + count (2B) + CRC (2B)
        if len(data) >= 6:
            try:
                _, func, addr, count = struct.unpack('>B B H H', data[:6])
                self.log(f"→ ACK: Addr=0x{addr:04X}, Count={count}")
            except struct.error:
                self.log(f"→ ACK (Raw): {data.hex().upper()}")
        else:
            self.log(f"→ ACK (Raw): {data.hex().upper()}")

    def init_polling_gui(self):
        self.polling_widgets = []

        self.polling_frame = ttk.Frame(self.root, width=300, relief=tk.SUNKEN, padding=5)
        self.polling_frame.grid(row=0, column=2, rowspan=6, sticky="nsew")
        self.root.columnconfigure(2, weight=0, minsize=280)

        # --- Polling制御バー ---
        interval_frame = ttk.Frame(self.polling_frame)
        interval_frame.pack(fill=tk.X, pady=2)

        ttk.Label(interval_frame, text="間隔:").pack(side=tk.LEFT)
        self.polling_interval_entry = ttk.Entry(interval_frame, width=6)
        self.polling_interval_entry.insert(0, "1000")
        self.polling_interval_entry.pack(side=tk.LEFT, padx=2)

        self.unit_combo = ttk.Combobox(interval_frame, values=["ms", "sec"], state="readonly", width=5)
        self.unit_combo.current(0)
        self.unit_combo.pack(side=tk.LEFT, padx=2)

        self.start_btn = ttk.Button(interval_frame, text="▶ Start")
        self.start_btn.pack(side=tk.LEFT, padx=2)

        self.stop_btn = ttk.Button(interval_frame, text="■ Stop")
        self.stop_btn.pack(side=tk.LEFT, padx=2)

        self.start_btn.config(command=self.start_polling_loop)
        self.stop_btn.config(command=self.stop_polling_loop)

        # --- スクロール可能なレジスタ一覧 ---
        canvas_frame = ttk.Frame(self.polling_frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(canvas_frame)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # --- スクロール処理のバインド修正 ---
        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        # canvas全体にイベントバインド
        self.canvas.bind("<Enter>", lambda e: self.canvas.focus_set())
        self.canvas.bind("<Leave>", lambda e: self.root.focus_set())

        # スクロールが期待通りの領域で動作するように明示的に bind する
        self.scrollable_frame.bind_all("<MouseWheel>", _on_mousewheel)

        # --- Polling対象として使えるR/RW項目のみ抽出して表示 ---
        for reg in self.reg_table:
            if reg.get("access") not in ["R", "RW"]:
                continue

            reg_len = reg.get("length", 1)
            typ = reg.get("type", "uint16_t")
            word_size = 2 if typ in ["float", "uint32_t"] else 1
            is_array = reg_len > 1

            for index in range(reg_len):
                row = ttk.Frame(self.scrollable_frame)
                row.pack(fill=tk.X, pady=1)

                var = tk.BooleanVar(value=False)
                check = ttk.Checkbutton(row, variable=var)
                check.pack(side=tk.LEFT)

                addr_label = ttk.Label(row, text=str(reg["addr"] + index * word_size), width=6, anchor="e")
                addr_label.pack(side=tk.LEFT, padx=2)

                name = f"{reg['name']}[{index}]" if is_array else reg["name"]
                name_label = ttk.Label(row, text=name, width=20, anchor="w")
                name_label.pack(side=tk.LEFT, padx=2)

                val_label = ttk.Label(row, text="-不定-", relief=tk.SUNKEN, width=12, background="white")
                val_label.pack(side=tk.LEFT, padx=2)

                self.polling_widgets.append({
                    "reg": reg,
                    "index": index,
                    "word_size": word_size,
                    "var": var,
                    "value_label": val_label,
                    "prev": None
                })

        self.polling_index = 0

    def start_polling_loop(self):
        try:
            interval = int(self.polling_interval_entry.get())
            if self.unit_combo.get() == "sec":
                interval *= 1000
        except ValueError:
            messagebox.showerror("Error", "ポーリング間隔が不正です。")
            return

        self.polling_interval_entry.config(state="disabled")
        self.unit_combo.config(state="disabled")
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")

        self._polling_active = True
        self._polling_task_id = self.root.after(0, self.polling_loop, interval)

    def stop_polling_loop(self):
        self._polling_active = False
        self.polling_interval_entry.config(state="normal")
        self.unit_combo.config(state="normal")
        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")

        if hasattr(self, '_polling_task_id'):
            self.root.after_cancel(self._polling_task_id)


    def polling_loop(self, interval):
        if not self._polling_active:
            return

        for entry in self.polling_widgets:
            if not entry["var"].get():
                continue

            reg = entry["reg"]
            index = entry["index"]
            label = entry["value_label"]

            def make_cb(entry):
                def cb(reg, data):
                    length = reg["length"]
                    if not data or not isinstance(data, (list, tuple)) or len(data) < length:
                        for e in self.polling_widgets:
                            if e["reg"] == reg:
                                e["value_label"].config(text="-No Response-", background="red")
                                e["prev"] = None
                        return

                    try:
                        value = data[entry["index"]]
                        txt = str(value)
                        if entry["prev"] != value:
                            entry["value_label"].config(text=txt, background="yellow")
                        else:
                            entry["value_label"].config(text=txt, background="white")
                        entry["prev"] = value
                    except (IndexError, TypeError):
                        entry["value_label"].config(text="-No Response-", background="red")
                        entry["prev"] = None
                return cb

            queue_send_read_for(self.serial_port, self.slave_addr, reg, make_cb(entry))

        self._polling_task_id = self.root.after(interval, self.polling_loop, interval)

    def make_polling_callback(self, entry):
        def cb(reg, data):
            if not data:
                entry["label"].config(text="--", background="red")
                entry["prev"] = None
            else:
                txt = ", ".join(map(str, data))
                if entry["prev"] != data:
                    entry["label"].config(text=txt, background="yellow")
                else:
                    entry["label"].config(text=txt, background="white")
                entry["prev"] = data
        return cb
    

# --- キュー化された通信処理（Read） ---
def queue_send_read(serial_port, unit_id, addr, length, callback):
    def task():
        try:
            frame = struct.pack('>B B H H', unit_id, 0x03, addr, length)
            crc = calc_crc(frame)
            frame += struct.pack('<H', crc)
            serial_port.reset_input_buffer()
            serial_port.write(frame)
            resp = serial_port.read(5 + length * 2)
            if len(resp) < 5:
                data = None
            else:
                data = resp
        except Exception:
            data = None
        root.after(0, lambda: callback(data))

    serial_task_queue.put((task, (), {}))

# --- キュー化された通信処理（Write Single Register） ---
def queue_send_write_single(serial_port, unit_id, addr, value, callback):
    def task():
        try:
            frame = struct.pack('>B B H H', unit_id, 0x06, addr, value)
            crc = calc_crc(frame)
            frame += struct.pack('<H', crc)
            serial_port.reset_input_buffer()
            serial_port.write(frame)
            resp = serial_port.read(256)  # 長さ8固定ではなく全体を読む

            if not resp:
                result = None
            elif resp[1] & 0x80:  # Exception応答
                result = resp
            elif len(resp) >= 8:  # 正常ACK応答
                result = resp
            else:
                result = None  # それ以外は異常

        except Exception:
            result = None

        root.after(0, lambda: callback(result))

    serial_task_queue.put((task, (), {}))

# --- キュー化された通信処理（Polling用 Read） ---
def queue_send_read_for(serial_port, unit_id, reg, callback):
    def task():
        try:
            addr = reg["addr"]
            length = reg["length"]
            typ = reg["type"]
            word_count = length * (2 if typ in ["float", "uint32_t"] else 1)

            frame = struct.pack('>B B H H', unit_id, 0x03, addr, word_count)
            crc = calc_crc(frame)
            frame += struct.pack('<H', crc)
            serial_port.reset_input_buffer()
            serial_port.write(frame)
            resp = serial_port.read(5 + word_count * 2)

            if len(resp) < 5:
                parsed = None
            else:
                byte_count = resp[2]
                values = resp[3:3 + byte_count]
                if typ == "uint16_t":
                    parsed = struct.unpack('>' + 'H' * length, bytes(values))
                elif typ == "uint32_t":
                    parsed = struct.unpack('>' + 'I' * length, bytes(values))
                elif typ == "float":
                    parsed = struct.unpack('>' + 'f' * length, bytes(values))
                else:
                    parsed = None
        except Exception:
            parsed = None

        root.after(0, lambda: callback(reg, parsed))

    serial_task_queue.put((task, (), {}))

# --- キュー化された通信処理（Write Multiple Registers） ---
def queue_send_write_multi(serial_port, unit_id, addr, values, typ, callback):
    def task():
        try:
            encoded = b''
            for v in values:
                if typ == "uint16_t":
                    encoded += struct.pack('>H', int(v))
                elif typ == "uint32_t":
                    encoded += struct.pack('>I', int(v))
                elif typ == "float":
                    encoded += struct.pack('>f', float(v))
                else:
                    raise ValueError("Unsupported type")

            num_regs = len(encoded) // 2
            byte_count = len(encoded)
            frame = struct.pack('>B B H H B', unit_id, 0x10, addr, num_regs, byte_count) + encoded
            crc = calc_crc(frame)
            frame += struct.pack('<H', crc)

            serial_port.reset_input_buffer()
            serial_port.write(frame)
            resp = serial_port.read(256)  # 読み取りバッファを拡大し、Exceptionも拾えるように

            if not resp:
                result = None
            elif resp[1] & 0x80:  # 異常応答の判定（例: 0x90）
                result = resp
            elif len(resp) < 8:
                result = None
            else:
                result = resp

        except Exception:
            result = None

        root.after(0, lambda: callback(result))

    serial_task_queue.put((task, (), {}))

def main():
    """Launch the RegiStar Modbus master GUI."""
    global root
    root = tk.Tk()
    _set_window_icon(root)
    root.withdraw()

    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        messagebox.showinfo("キャンセル", "ファイル選択がキャンセルされました。処理を終了します。")
        sys.exit(0)

    reg_table = extract_registers_from_excel(file_path)
    root.deiconify()
    app = ModbusMasterGUI(root, reg_table)
    root.mainloop()


# --- エントリーポイント ---
if __name__ == "__main__":
    main()
