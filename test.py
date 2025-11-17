import tkinter as tk
from tkinter import ttk, Label, messagebox
from PIL import Image, ImageTk
import win32com.client
import pyvisa
import json
import os

# COM 객체 생성
engine = win32com.client.Dispatch("LoaderEngine.LoaderEngine")

root = tk.Tk()
root.title("RFOptimizer for QORVO")
root.geometry("960x900")

notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill="both")

# 각 탭
rf_frame = ttk.Frame(notebook)
notebook.add(rf_frame, text="RF Optimizer")
pna_frame = ttk.Frame(notebook)
notebook.add(pna_frame, text="PNA Settings")
dc_frame = ttk.Frame(notebook)
notebook.add(dc_frame, text="DC Power Settings")
program_frame = ttk.Frame(notebook)
notebook.add(program_frame, text="Program")

# 상단 로고
try:
    logo_image = Image.open(r"C:\Users\daehy\OneDrive\Desktop\Python\Python study\Project\logo.png")
    logo_image = logo_image.resize((120, 40))
    logo_photo = ImageTk.PhotoImage(logo_image)
    logo_label = Label(root, image=logo_photo)
    logo_label.pack(side="top", pady=0)
except Exception:
    pass

# --------------------- RF Optimizer 탭 내용 ---------------------
rf_log = tk.Text(rf_frame, height=7, width=95)
rf_log.pack(pady=5, padx=5)
def rf_log_message(msg):
    rf_log.insert(tk.END, msg + "\n")
    rf_log.see(tk.END)
rf_engine_initialized = False

def check_rf_initialized():
    if not rf_engine_initialized:
        rf_log_message("Please RUN Program first.")
        return False
    return True

def rf_initialize_engine():
    global rf_engine_initialized
    try:
        result = engine.InitializeEngine("")
        rf_engine_initialized = True
        rf_log_message(f"Engine Initialized: {result}")
    except Exception as e:
        rf_log_message(f"Error initializing engine: {e}")

def rf_dispose_engine():
    global rf_engine_initialized
    try:
        engine.Dispose()
        rf_engine_initialized = False
        rf_log_message("Engine Disposed")
    except Exception as e:
        rf_log_message(f"Error disposing engine: {e}")

rf_button_frame = tk.Frame(rf_frame)
rf_button_frame.pack(pady=5)
tk.Button(rf_button_frame, text="Start", command=rf_initialize_engine, bg="lightgreen", width=16).pack(side="left", padx=5)
tk.Button(rf_button_frame, text="Dispose", command=rf_dispose_engine, bg="lightyellow", width=16).pack(side="left", padx=5)
tk.Button(rf_button_frame, text="END", command=root.quit, bg="red", width=16).pack(side="left", padx=5)  # END 버튼 추가

# MIPI/Interface/Clock/Chipset/Family/USID UI
rf_setting_frame = tk.LabelFrame(rf_frame, text="RF Optimizer Setting", padx=10, pady=10)
rf_setting_frame.pack(fill="x", padx=10, pady=10)

tk.Label(rf_setting_frame, text="Interface:").grid(row=0, column=0)
rf_interface_var = tk.StringVar(value="RFMDComm2")
rf_interface_dropdown = ttk.Combobox(rf_setting_frame, textvariable=rf_interface_var, values=["RFMDComm", "RFMDComm2"], width=15)
rf_interface_dropdown.grid(row=0, column=1, padx=5)
tk.Button(rf_setting_frame, text="Set", command=lambda: set_rf_interface(), width=10).grid(row=0, column=2, padx=3)

def set_rf_interface():
    if not check_rf_initialized(): return
    try:
        result = engine.SetInterface(rf_interface_var.get())
        rf_log_message(f"Set Interface: {rf_interface_var.get()} -> {result}")
    except Exception as e:
        rf_log_message(f"Error setting interface: {e}")

tk.Label(rf_setting_frame, text="Clock rate:").grid(row=1, column=0)
rf_frequency_var = tk.StringVar(value="10MHz")
rf_frequency_dropdown = ttk.Combobox(rf_setting_frame, textvariable=rf_frequency_var, values=["1MHz", "2MHz", "5MHz", "10MHz"], width=15)
rf_frequency_dropdown.grid(row=1, column=1, padx=5)
tk.Button(rf_setting_frame, text="Set", command=lambda: set_rf_frequency(), width=10).grid(row=1, column=2, padx=3)

def set_rf_frequency():
    if not check_rf_initialized(): return
    freq_map = {"1MHz": 1000000, "2MHz": 2000000, "5MHz": 5000000, "10MHz": 10000000}
    try:
        engine.Frequency_Hz = freq_map.get(rf_frequency_var.get(), 10000000)
        rf_log_message(f"Set Frequency: {engine.Frequency_Hz} Hz")
    except Exception as e:
        rf_log_message(f"Error setting frequency: {e}")

tk.Label(rf_setting_frame, text="Chipset:").grid(row=2, column=0)
rf_chipset_var = tk.StringVar()
rf_chipset_dropdown = ttk.Combobox(rf_setting_frame, textvariable=rf_chipset_var, width=15)
rf_chipset_dropdown.grid(row=2, column=1, padx=5)
tk.Button(rf_setting_frame, text="Load", command=lambda: load_rf_chipsets(), width=10).grid(row=2, column=2, padx=3)
tk.Button(rf_setting_frame, text="Set", command=lambda: set_rf_chipset(), width=10).grid(row=2, column=3, padx=3)

def load_rf_chipsets():
    if not check_rf_initialized(): return
    try:
        chipsets = engine.GetChipsetNamesAvailable()
        rf_chipset_dropdown["values"] = list(chipsets)
        rf_log_message("Chipset list loaded.")
    except Exception as e:
        rf_log_message(f"Error loading chipsets: {e}")

def set_rf_chipset():
    if not check_rf_initialized(): return
    try:
        engine.ActiveChipset = rf_chipset_var.get()
        rf_log_message(f"Set Chipset: {rf_chipset_var.get()}")
    except Exception as e:
        rf_log_message(f"Error setting chipset: {e}")

tk.Label(rf_setting_frame, text="Family:").grid(row=3, column=0)
rf_family_var = tk.StringVar()
rf_family_dropdown = ttk.Combobox(rf_setting_frame, textvariable=rf_family_var, width=15)
rf_family_dropdown.grid(row=3, column=1, padx=5)
tk.Button(rf_setting_frame, text="Load", command=lambda: load_rf_families(), width=10).grid(row=3, column=2, padx=3)
tk.Button(rf_setting_frame, text="Set", command=lambda: set_rf_family(), width=10).grid(row=3, column=3, padx=3)

def load_rf_families():
    if not check_rf_initialized(): return
    try:
        families = engine.GetFamilyNamesOfCurrentChipset()
        rf_family_dropdown["values"] = list(families)
        rf_log_message("Family list loaded.")
    except Exception as e:
        rf_log_message(f"Error loading families: {e}")

def set_rf_family():
    if not check_rf_initialized(): return
    try:
        engine.ActiveFamily = rf_family_var.get()
        rf_log_message(f"Set Family: {rf_family_var.get()}")
    except Exception as e:
        rf_log_message(f"Error setting family: {e}")

tk.Label(rf_setting_frame, text="USID:").grid(row=4, column=0)
rf_usid_var = tk.StringVar(value="default")
rf_usid_dropdown = ttk.Combobox(rf_setting_frame, textvariable=rf_usid_var, values=["default"] + [str(i) for i in range(1, 21)], width=15)
rf_usid_dropdown.grid(row=4, column=1, padx=5)
tk.Button(rf_setting_frame, text="Set", command=lambda: set_rf_usid(), width=10).grid(row=4, column=2, padx=3)

def set_rf_usid():
    if not check_rf_initialized(): return
    try:
        if rf_usid_var.get() == "default":
            rf_log_message(f"Current Slave Address: {engine.SlaveAddress}")
        else:
            engine.SetSlaveAddress(int(rf_usid_var.get()))
            rf_log_message(f"Set Slave Address: {rf_usid_var.get()}")
    except Exception as e:
        rf_log_message(f"Error setting USID: {e}")

# 레지스터 입력 및 쓰기
register_frame = tk.LabelFrame(rf_frame, text="Register Write (Value=Dec.)", padx=10, pady=10)
register_frame.pack(fill="x", padx=10, pady=10)
register_entries = []
register_values = []
for i in range(10):
    tk.Label(register_frame, text=f"Register {i+1} Addr=").grid(row=i, column=0, sticky="w")
    entry = tk.Entry(register_frame, width=13)
    entry.grid(row=i, column=1)
    register_entries.append(entry)
    tk.Label(register_frame, text="Value=").grid(row=i, column=2)
    value_entry = tk.Entry(register_frame, width=13)
    value_entry.grid(row=i, column=3)
    register_values.append(value_entry)
tk.Button(register_frame, text="Write All Register", command=lambda: rf_write_registers(), width=30).grid(row=11, column=0, columnspan=4, pady=7)

def rf_write_registers():
    if not check_rf_initialized():
        return
    if not engine.ActiveChipset or not engine.ActiveFamily:
        rf_log_message("Please set Chipset and Family before writing registers.")
        return
    for entry, value_entry in zip(register_entries, register_values):
        register_name = entry.get().strip()
        register_value = value_entry.get().strip()
        if not register_name or not register_value: continue
        if not register_name.lower().startswith("reg"):
            register_name = f"Reg{register_name}"
        try:
            register_value = int(register_value) if register_value.isdigit() else 0
            result = engine.WriteRegister(register_name, register_value)
            rf_log_message(f"{register_name}: Write {'Success' if result else 'Fail'}")
            read_value = engine.ReadRegister2(register_name)
            rf_log_message(f"Register Read: {register_name} -> 0x{read_value:X}")
        except Exception as e:
            rf_log_message(f"Error writing {register_name}: {e}")

# ------------------ PNA Settings 탭 ------------------
pna_log = tk.Text(pna_frame, height=6, width=95)
pna_log.pack(pady=5, padx=5)
def pna_log_message(msg):
    pna_log.insert(tk.END, msg + "\n")
    pna_log.see(tk.END)

pna_setting_frame = tk.LabelFrame(pna_frame, text="PNA Settings", padx=10, pady=10)
pna_setting_frame.pack(fill="x", padx=10, pady=10)

tk.Label(pna_setting_frame, text="GPIB Address:").grid(row=0, column=0)
gpib_entry = tk.Entry(pna_setting_frame)
gpib_entry.grid(row=0, column=1)
tk.Label(pna_setting_frame, text="Channel:").grid(row=1, column=0)
channel_entry = tk.Entry(pna_setting_frame)
channel_entry.grid(row=1, column=1)
tk.Label(pna_setting_frame, text="S-Parameter:").grid(row=2, column=0)
sparam_entry = tk.Entry(pna_setting_frame)
sparam_entry.grid(row=2, column=1)
tk.Label(pna_setting_frame, text="Trace:").grid(row=3, column=0)
trace_entry = tk.Entry(pna_setting_frame)
trace_entry.grid(row=3, column=1)
tk.Label(pna_setting_frame, text="Format:").grid(row=4, column=0)
format_combo = ttk.Combobox(pna_setting_frame, values=["MLOG", "SWR", "PHAS", "REAL", "IMAG"])
format_combo.grid(row=4, column=1)
format_combo.current(0)
tk.Label(pna_setting_frame, text="Marker Freq (MHz):").grid(row=5, column=0)
marker_entry = tk.Entry(pna_setting_frame)
marker_entry.grid(row=5, column=1)
result_label = tk.Label(pna_setting_frame, text="Result will appear here")
result_label.grid(row=7, column=0, columnspan=2)

def apply_pna_settings():
    gpib_addr = gpib_entry.get()
    channel = channel_entry.get()
    sparam = sparam_entry.get()
    trace = trace_entry.get()
    meas_name = f"CH{channel}_S{sparam}_{trace}"
    format_selected = format_combo.get()
    try:
        marker_freq_mhz = float(marker_entry.get())
        marker_freq_hz = marker_freq_mhz * 1e6
    except Exception:
        pna_log_message("Invalid marker frequency input")
        return
    try:
        rm = pyvisa.ResourceManager()
        instrument = rm.open_resource(f'GPIB0::{gpib_addr}::INSTR')
        instrument.write('*RST')
        instrument.write('CALC1:PAR:DEL:ALL')
        instrument.write(f'CALC1:PAR:DEF:EXT "{meas_name}","S{sparam}"')
        instrument.write(f'DISP:WIND:TRAC1:FEED "{meas_name}"')
        instrument.write(f'CALC1:PAR:SEL {meas_name}')
        instrument.write(f'CALC1:FORM {format_selected}')
        instrument.write('CALC1:MARK1:STAT ON')
        instrument.write(f'CALC1:MARK1:X {marker_freq_hz}')
        instrument.query('*OPC?')
        marker_data = instrument.query('CALC1:MARK1:Y?')
        result_label.config(text=f"Marker Data: {marker_data}")
        pna_log_message(f"PNA set! Marker: {marker_data}")
    except Exception as e:
        result_label.config(text=f"Error: {e}")
        pna_log_message(f"Error: {e}")

apply_btn = tk.Button(pna_setting_frame, text="Apply Settings", command=apply_pna_settings)
apply_btn.grid(row=6, column=0, columnspan=2, pady=10)

# ---------------- DC Power Settings 탭 ----------------
ps_log = tk.Text(dc_frame, height=5, width=95)
ps_log.pack(pady=5, padx=5)
def ps_log_message(msg):
    ps_log.insert(tk.END, msg + "\n")
    ps_log.see(tk.END)

device_options = ["E3631A", "E3634A"]
source_options = {
    "E3631A": ["P6V", "P25V"],
    "E3634A": ["P25V", "P50V"]
}
selected_devices = [tk.StringVar() for _ in range(5)]
selected_sources = [tk.StringVar() for _ in range(5)]
gpib_addresses = [tk.StringVar() for _ in range(5)]
current_limits = [tk.StringVar() for _ in range(5)]
default_voltages = [tk.StringVar() for _ in range(5)]
connection_labels = []
source_dropdowns = []
def validate_numeric_input(value):
    return value.isdigit() or value == ""
ps_validate_cmd = root.register(validate_numeric_input)

def update_source_options(event, ch):
    selected_device = selected_devices[ch].get()
    if selected_device in source_options:
        source_dropdowns[ch]["values"] = source_options[selected_device]

def apply_ps_settings():
    rm = pyvisa.ResourceManager()
    for ch in range(5):  
        gpib_address = gpib_addresses[ch].get()
        device_name = selected_devices[ch].get()
        source_name = selected_sources[ch].get()
        current_limit = current_limits[ch].get()
        default_voltage = default_voltages[ch].get()
        if not gpib_address or not device_name or not source_name or not default_voltage or not current_limit:
            ps_log_message(f"⚠ CH{ch+1}: 설정값이 입력되지 않아 건너뜀")
            continue
        try:
            instrument = rm.open_resource(f"GPIB::{gpib_address}::INSTR")
            if device_name == "E3631A":
                instrument.write(f"APPL {source_name}, {default_voltage}, {current_limit}")
            elif device_name == "E3634A":
                instrument.write(f"VOLT:RANG {source_name}")
                instrument.write(f"VOLT {default_voltage}")
                instrument.write(f"CURR {current_limit}")
            connection_labels[ch].config(text=f"CH{ch+1} 설정 완료!", fg="green")
            ps_log_message(f"CH{ch+1} 설정 완료!")
        except Exception as e:
            connection_labels[ch].config(text=f"연결 실패: {e}", fg="red")
            ps_log_message(f"CH{ch+1} 연결 실패: {e}")

for ch in range(5):
    frame = tk.Frame(dc_frame)
    frame.pack(pady=2)
    tk.Label(frame, text=f"Ch{ch+1}").pack(side="left")
    tk.Label(frame, text="장비 선택:").pack(side="left")
    device_dropdown = ttk.Combobox(frame, textvariable=selected_devices[ch], values=device_options, width=8)
    device_dropdown.pack(side="left", padx=2)
    device_dropdown.bind("<<ComboboxSelected>>", lambda event, ch=ch: update_source_options(event, ch))
    tk.Label(frame, text="소스 선택:").pack(side="left")
    source_dropdown = ttk.Combobox(frame, textvariable=selected_sources[ch], width=8)
    source_dropdown.pack(side="left", padx=2)
    source_dropdowns.append(source_dropdown)
    tk.Label(frame, text="GPIB 주소:").pack(side="left")
    gpib_entry = tk.Entry(frame, textvariable=gpib_addresses[ch], width=5, validate="key", validatecommand=(ps_validate_cmd, "%P"))
    gpib_entry.pack(side="left", padx=2)
    gpib_entry.insert(0, str(ch + 1))
    tk.Label(frame, text="전류(A):").pack(side="left")
    current_entry = tk.Entry(frame, textvariable=current_limits[ch], width=5)
    current_entry.pack(side="left", padx=2)
    current_entry.insert(0, "1.0")
    tk.Label(frame, text="전압(V):").pack(side="left")
    default_voltage_entry = tk.Entry(frame, textvariable=default_voltages[ch], width=5)
    default_voltage_entry.pack(side="left", padx=2)
    default_voltage_entry.insert(0, "3.8")
    conn_status = tk.Label(frame, text="설정 대기 중", fg="gray")
    conn_status.pack(side="left", padx=10)
    connection_labels.append(conn_status)
btn_apply_settings = tk.Button(dc_frame, text="계측기 설정", command=apply_ps_settings)
btn_apply_settings.pack(side="top", pady=10)

# --- Program 탭 UI ---
program_listbox = tk.Listbox(program_frame, height=10, width=40)
program_listbox.grid(row=1, column=0, rowspan=7, padx=10, pady=10, sticky="ns")
program_name_var = tk.StringVar()
tk.Label(program_frame, text="Program Name:").grid(row=1, column=1, sticky="e")
tk.Entry(program_frame, textvariable=program_name_var, width=20).grid(row=1, column=2, sticky="w")
btn_save = tk.Button(program_frame, text="SAVE (설정저장)", width=18)
btn_save.grid(row=2, column=2, sticky="w", pady=2)
btn_load = tk.Button(program_frame, text="LOAD (불러오기)", width=18)
btn_load.grid(row=3, column=2, sticky="w", pady=2)
btn_rename = tk.Button(program_frame, text="이름 변경 (수정)", width=18)
btn_rename.grid(row=4, column=2, sticky="w", pady=2)
btn_delete = tk.Button(program_frame, text="삭제", fg="red", width=18)
btn_delete.grid(row=5, column=2, sticky="w", pady=2)

PROGRAM_FILE = "rf_programs.json"

def load_all_programs():
    if not os.path.exists(PROGRAM_FILE):
        return []
    with open(PROGRAM_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_all_programs(programs):
    with open(PROGRAM_FILE, "w", encoding="utf-8") as f:
        json.dump(programs, f, ensure_ascii=False, indent=2)

def refresh_program_list():
    program_listbox.delete(0, tk.END)
    for prog in load_all_programs():
        program_listbox.insert(tk.END, prog['name'])

refresh_program_list()

# 저장 (SAVE) 기능
def save_current_program():
    name = program_name_var.get().strip()
    if not name:
        messagebox.showwarning("경고", "프로그램 이름을 입력하세요.")
        return
    # RF Optimizer
    rf_data = dict(
        interface=rf_interface_var.get(),
        clock=rf_frequency_var.get(),
        chipset=rf_chipset_var.get(),
        family=rf_family_var.get(),
        usid=rf_usid_var.get(),
        registers=[(e.get(), v.get()) for e, v in zip(register_entries,register_values)]
    )
    # PNA
    pna_data = dict(
        gpib=gpib_entry.get(),
        channel=channel_entry.get(),
        sparam=sparam_entry.get(),
        trace=trace_entry.get(),
        format=format_combo.get(),
        marker=marker_entry.get()
    )
    # Power
    power_data = []
    for ch in range(5):
        power_data.append({
            "device": selected_devices[ch].get(),
            "source": selected_sources[ch].get(),
            "gpib": gpib_addresses[ch].get(),
            "current": current_limits[ch].get(),
            "voltage": default_voltages[ch].get()
        })
    # 이름 중복 제거 후 추가
    sets = load_all_programs()
    sets = [s for s in sets if s["name"] != name]
    sets.append({"name":name, "rf":rf_data, "pna":pna_data, "power":power_data})
    save_all_programs(sets)
    refresh_program_list()
    messagebox.showinfo("완료", f'"{name}" 저장됨!')

btn_save.config(command=save_current_program)

# 불러오기 (LOAD)
def load_selected_program():
    sel = program_listbox.curselection()
    if not sel:
        messagebox.showwarning("경고", "불러올 설정을 선택하세요.")
        return
    name = program_listbox.get(sel[0])
    sets = load_all_programs()
    found = next((s for s in sets if s["name"]==name), None)
    if not found:
        messagebox.showerror("오류", "설정 정보를 찾을 수 없습니다.")
        return
    # 1. RF Optimizer 설정 복구
    rf = found['rf']
    rf_interface_var.set(rf.get('interface',""))
    rf_frequency_var.set(rf.get('clock',""))
    rf_chipset_var.set(rf.get('chipset',""))
    rf_family_var.set(rf.get('family',""))
    rf_usid_var.set(rf.get('usid',""))
    regdata = rf.get("registers", [])
    for i, (e,v) in enumerate(regdata):
        try:
            register_entries[i].delete(0, tk.END)
            register_entries[i].insert(0, e)
            register_values[i].delete(0, tk.END)
            register_values[i].insert(0, v)
        except Exception: pass
    # 2. PNA 설정 복구
    pna = found['pna']
    gpib_entry.delete(0, tk.END); gpib_entry.insert(0, pna.get("gpib",""))
    channel_entry.delete(0, tk.END); channel_entry.insert(0, pna.get("channel",""))
    sparam_entry.delete(0, tk.END); sparam_entry.insert(0, pna.get("sparam",""))
    trace_entry.delete(0, tk.END); trace_entry.insert(0, pna.get("trace",""))
    format_combo.set(pna.get("format","MLOG"))
    marker_entry.delete(0, tk.END); marker_entry.insert(0, pna.get("marker",""))
    # 3. Power 설정 복구
    power = found['power']
    for i,ch in enumerate(power):
        selected_devices[i].set(ch.get("device",""))
        selected_sources[i].set(ch.get("source",""))
        gpib_addresses[i].set(ch.get("gpib",""))
        current_limits[i].set(ch.get("current",""))
        default_voltages[i].set(ch.get("voltage",""))
    # 이름 필드도 세팅
    program_name_var.set(found["name"])
    messagebox.showinfo("완료", f'"{name}" 불러오기 완료!')

btn_load.config(command=load_selected_program)

# 이름 변경 (RENAME)
def rename_selected_program():
    sel = program_listbox.curselection()
    new_name = program_name_var.get().strip()
    if not sel:
        messagebox.showwarning("경고", "이름을 변경할 설정을 선택하세요.")
        return
    if not new_name:
        messagebox.showwarning("경고", "새 이름을 입력하세요.")
        return
    old_name = program_listbox.get(sel[0])
    sets = load_all_programs()
    if any(s["name"]==new_name for s in sets):
        messagebox.showwarning("경고", "이미 존재하는 이름입니다.")
        return
    found = next((s for s in sets if s["name"]==old_name), None)
    if not found:
        messagebox.showerror("오류", "설정 정보를 찾을 수 없습니다.")
        return
    found["name"] = new_name
    save_all_programs(sets)
    refresh_program_list()
    program_name_var.set(new_name)
    messagebox.showinfo("완료", f'"{old_name}" → "{new_name}" 이름 변경!')

btn_rename.config(command=rename_selected_program)

# 삭제 (DELETE)
def delete_selected_program():
    sel = program_listbox.curselection()
    if not sel:
        messagebox.showwarning("경고", "삭제할 설정을 선택하세요.")
        return
    name = program_listbox.get(sel[0])
    if not messagebox.askyesno("삭제확인", f'"{name}" 프로그램을 삭제하시겠습니까?'):
        return
    sets = load_all_programs()
    sets = [s for s in sets if s["name"] != name]
    save_all_programs(sets)
    refresh_program_list()
    program_name_var.set("")
    messagebox.showinfo("삭제", f'"{name}" 삭제 완료!')

btn_delete.config(command=delete_selected_program)

# 목록 클릭 시 현재 이름 필드 자동 입력
def on_program_select(evt):
    sel = program_listbox.curselection()
    if sel:
        name = program_listbox.get(sel[0])
        program_name_var.set(name)
program_listbox.bind("<<ListboxSelect>>", on_program_select)

refresh_program_list()

root.mainloop()
