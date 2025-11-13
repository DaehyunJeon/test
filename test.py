import tkinter as tk
from tkinter import ttk, Label
from PIL import Image, ImageTk
import win32com.client
import pyvisa
import time

# COM 객체 생성
engine = win32com.client.Dispatch("LoaderEngine.LoaderEngine")

root = tk.Tk()
root.title("RFOptimizer for QORVO")
root.geometry("420x750")

notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill="both")

# 첫 번째 탭: RF Optimizer
rf_frame = ttk.Frame(notebook)
notebook.add(rf_frame, text="RF Optimizer")

# 두 번째 탭: PNA Settings
pna_frame = ttk.Frame(notebook)
notebook.add(pna_frame, text="PNA Settings")

# 세 번째 탭: Power Supply Settings
pna_frame = ttk.Frame(notebook)
notebook.add(pna_frame, text="DC Power Settings")

# 상단 로고 이미지 로드
logo_image = Image.open(r"C:\Users\daehy\OneDrive\Desktop\Python\Python study\Project\logo.png")  # 로고 파일 경로
logo_image = logo_image.resize((120, 40))  # 크기 조정
logo_photo = ImageTk.PhotoImage(logo_image)

# Label에 이미지 추가 (상단 중앙)
logo_label = Label(root, image=logo_photo)
logo_label.pack(side="top", pady=0)  # 상단에 배치, 여백 10px

# 출력 로그
button_frame = tk.Frame(pna_frame)
button_frame.pack(fill="x", padx=10, pady=5)

# 엔진 초기화 버튼
tk.Button(pna_frame, text="Start", command=lambda: initialize_engine(), bg="lightgreen", width=16).grid(row=0, column=0, padx=5)
tk.Button(pna_frame, text="Dispose", command=lambda: dispose_engine(), bg="lightyellow", width=16).grid(row=0, column=1, padx=5)
tk.Button(pna_frame, text="Exit", command=lambda: root.quit(), bg="red", width=16).grid(row=0, column=2, padx=5)

# 출력 로그
output_frame = tk.Frame(root)
output_frame.pack(fill="x", padx=10, pady=5)

output_text = tk.Text(output_frame, height=6, width=60)
output_text.pack(side="left")
scrollbar = tk.Scrollbar(output_frame, command=output_text.yview)
scrollbar.pack(side="right", fill="y")
output_text.config(yscrollcommand=scrollbar.set)

def log_message(msg):
    output_text.insert(tk.END, msg + "\n")
    output_text.see(tk.END)

engine_initialized = False

def check_initialized():
    if not engine_initialized:
        log_message("Please RUN Program first.")
        return False
    return True

def initialize_engine():
    global engine_initialized
    try:
        result = engine.InitializeEngine("")
        engine_initialized = True
        log_message(f"Engine Initialized: {result}")
    except Exception as e:
        log_message(f"Error initializing engine: {e}")

def dispose_engine():
    global engine_initialized
    try:
        engine.Dispose()
        engine_initialized = False
        log_message("Engine Disposed")
    except Exception as e:
        log_message(f"Error disposing engine: {e}")

# MIPI 설정 UI
before_main_frame = tk.Frame(root)
before_main_frame.pack(fill="x", padx=10, pady=10)
tk.Label(before_main_frame, text="---------------------------RF_Optimizer_setting---------------------------").grid(row=0, column=0, padx=5, pady=5, sticky="w")

# 메인 설정 UI
main_frame = tk.Frame(root, bg="Lightyellow")
main_frame.pack(fill="both", padx=10, pady=10)

# Comm Interface
tk.Label(main_frame, text="Interface", bg="lightyellow").grid(row=1, column=0, padx=5, pady=5, sticky="w")
interface_var = tk.StringVar(value="RFMDComm2")
interface_dropdown = ttk.Combobox(main_frame, textvariable=interface_var, values=["RFMDComm", "RFMDComm2"], width=15)
interface_dropdown.grid(row=1, column=1, padx=5)
tk.Button(main_frame, text="Set", width=23, command=lambda: set_interface()).grid(row=1, column=2,columnspan=2, padx=5)

def set_interface():
    if not check_initialized(): return
    try:
        result = engine.SetInterface(interface_var.get())
        log_message(f"Set Interface: {interface_var.get()} -> {result}")
    except Exception as e:
        log_message(f"Error setting interface: {e}")

# Clock rate
tk.Label(main_frame, text="Clock rate", bg="lightyellow").grid(row=2, column=0, padx=5, pady=5, sticky="w")
frequency_var = tk.StringVar(value="10MHz")
frequency_dropdown = ttk.Combobox(main_frame, textvariable=frequency_var, values=["1MHz", "2MHz", "5MHz", "10MHz"], width=15)
frequency_dropdown.grid(row=2, column=1, padx=5)
tk.Button(main_frame, text="Set", width=23, command=lambda: set_frequency()).grid(row=2, column=2,columnspan=2, padx=5)

def set_frequency():
    if not check_initialized(): return
    freq_map = {"1MHz": 1000000, "2MHz": 2000000, "5MHz": 5000000, "10MHz": 10000000}
    try:
        engine.Frequency_Hz = freq_map.get(frequency_var.get(), 10000000)
        log_message(f"Set Frequency: {engine.Frequency_Hz} Hz")
    except Exception as e:
        log_message(f"Error setting frequency: {e}")

# Chipset
tk.Label(main_frame, text="Chipset", bg="lightyellow").grid(row=3, column=0, padx=5, pady=5, sticky="w")
chipset_var = tk.StringVar()
chipset_dropdown = ttk.Combobox(main_frame, textvariable=chipset_var, width=15)
chipset_dropdown.grid(row=3, column=1, padx=5)
tk.Button(main_frame, text="Load", width=10, command=lambda: load_chipsets()).grid(row=3, column=2, padx=5)
tk.Button(main_frame, text="Set", width=10, command=lambda: set_chipset()).grid(row=3, column=3, padx=5)

def load_chipsets():
    if not check_initialized(): return
    try:
        chipsets = engine.GetChipsetNamesAvailable()
        chipset_dropdown["values"] = list(chipsets)
        log_message("Chipset list loaded.")
    except Exception as e:
        log_message(f"Error loading chipsets: {e}")

def set_chipset():
    if not check_initialized(): return
    try:
        engine.ActiveChipset = chipset_var.get()
        log_message(f"Set Chipset: {chipset_var.get()}")
    except Exception as e:
        log_message(f"Error setting chipset: {e}")

# Family
tk.Label(main_frame, text="Family", bg="lightyellow").grid(row=4, column=0, padx=5, pady=5, sticky="w")
family_var = tk.StringVar()
family_dropdown = ttk.Combobox(main_frame, textvariable=family_var, width=15)
family_dropdown.grid(row=4, column=1, padx=5)
tk.Button(main_frame, text="Load", width=10, command=lambda: load_families()).grid(row=4, column=2, padx=5)
tk.Button(main_frame, text="Set", width=10, command=lambda: set_family()).grid(row=4, column=3, padx=5)

def load_families():
    if not check_initialized(): return
    try:
        families = engine.GetFamilyNamesOfCurrentChipset()
        family_dropdown["values"] = list(families)
        log_message("Family list loaded.")
    except Exception as e:
        log_message(f"Error loading families: {e}")

def set_family():
    if not check_initialized(): return
    try:
        engine.ActiveFamily = family_var.get()
        log_message(f"Set Family: {family_var.get()}")
    except Exception as e:
        log_message(f"Error setting family: {e}")

# USID
tk.Label(main_frame, text="USID", bg="lightyellow").grid(row=5, column=0, padx=5, pady=5, sticky="w")
usid_var = tk.StringVar(value="default")
usid_dropdown = ttk.Combobox(main_frame, textvariable=usid_var, values=["default"] + [str(i) for i in range(1, 21)], width=15)
usid_dropdown.grid(row=5, column=1, padx=5)
tk.Button(main_frame, text="Set", width=23, command=lambda: set_usid()).grid(row=5, column=2,columnspan=2, padx=5)

def set_usid():
    if not check_initialized(): return
    try:
        if usid_var.get() == "default":
            log_message(f"Current Slave Address: {engine.SlaveAddress}")
        else:
            engine.SetSlaveAddress(int(usid_var.get()))
            log_message(f"Set Slave Address: {usid_var.get()}")
    except Exception as e:
        log_message(f"Error setting USID: {e}")

# 레지스터 설정 UI
register_frame = tk.Frame(root, bg="lightyellow")
register_frame.pack(fill="both", padx=10, pady=10)

tk.Label(register_frame, text="-----------------------Register write(Value = Dec.)-----------------------",bg="lightyellow").grid(row=0, column=0, columnspan=4, padx=5, pady=5, sticky="w")

register_entries = []
register_values = []

for i in range(10):
    tk.Label(register_frame, text=f"Register {i+1}||Addr=", bg="lightyellow").grid(row=i+1, column=0, padx=5, pady=2, sticky="w")
    entry = tk.Entry(register_frame, width=13)
    entry.grid(row=i+1, column=1, padx=5, pady=2)
    register_entries.append(entry)

    tk.Label(register_frame, text="Value=",bg="lightyellow").grid(row=i+1, column=2, padx=5, pady=2, sticky="w")
    value_entry = tk.Entry(register_frame, width=13)
    value_entry.grid(row=i+1, column=3, padx=5, pady=2)
    register_values.append(value_entry)

tk.Button(register_frame, text="Write All Register", command=lambda: write_registers(), width=50).grid(row=12, column=0, columnspan=4, pady=10)

def write_registers():
    if not check_initialized():
        return
    if not engine.ActiveChipset or not engine.ActiveFamily:
        log_message("Please set Chipset and Family before writing registers.")
        return

    for entry, value_entry in zip(register_entries, register_values):
        register_name = entry.get().strip()
        register_value = value_entry.get().strip()

        if not register_name or not register_value:
            continue

        if not register_name.lower().startswith("reg"):
            register_name = f"Reg{register_name}"

        try:
            register_value = int(register_value) if register_value.isdigit() else 0
            result = engine.WriteRegister(register_name, register_value)
            log_message(f"{register_name}: Write {'Success' if result else 'Fail'}")

            read_value = engine.ReadRegister2(register_name)
            log_message(f"Register Read: {register_name} -> 0x{read_value}")
        except Exception as e:
            log_message(f"Error writing {register_name}: {e}")


# PNA 설정 UI
def apply_settings():
    gpib_addr = gpib_entry.get()
    channel = channel_entry.get()
    sparam = sparam_entry.get()
    trace = trace_entry.get()
    meas_name = f"CH{channel}_S{sparam}_{trace}"
    format_selected = format_combo.get()
    marker_freq_mhz = float(marker_entry.get())
    marker_freq_hz = marker_freq_mhz * 1e6

    try:
        rm = pyvisa.ResourceManager()
        instrument = rm.open_resource(f'GPIB0::{gpib_addr}::INSTR')

        # PNA 초기화 및 설정
        instrument.write('*RST')
        instrument.write('CALC1:PAR:DEL:ALL')
        instrument.write(f'CALC1:PAR:DEF:EXT "{meas_name}","S{sparam}"')
        instrument.write(f'DISP:WIND:TRAC1:FEED "{meas_name}"')
        instrument.write(f'CALC1:PAR:SEL {meas_name}')
        instrument.write(f'CALC1:FORM {format_selected}')

        # Marker 설정
        instrument.write('CALC1:MARK1:STAT ON')
        instrument.write(f'CALC1:MARK1:X {marker_freq_hz}')
        instrument.query('*OPC?')  # 안정적 대기
        marker_data = instrument.query('CALC1:MARK1:Y?')
        result_label.config(text=f"Marker Data: {marker_data}")
    except Exception as e:
        result_label.config(text=f"Error: {e}")

# GUI 생성
frame = tk.LabelFrame(root, text="PNA Settings", padx=10, pady=10)
frame.pack(padx=10, pady=10, fill="x")

# 입력 필드
tk.Label(frame, text="GPIB Address:").grid(row=0, column=0)
gpib_entry = tk.Entry(frame)
gpib_entry.grid(row=0, column=1)

tk.Label(frame, text="Channel:").grid(row=1, column=0)
channel_entry = tk.Entry(frame)
channel_entry.grid(row=1, column=1)

tk.Label(frame, text="S-Parameter:").grid(row=2, column=0)
sparam_entry = tk.Entry(frame)
sparam_entry.grid(row=2, column=1)

tk.Label(frame, text="Trace:").grid(row=3, column=0)
trace_entry = tk.Entry(frame)
trace_entry.grid(row=3, column=1)

tk.Label(frame, text="Format:").grid(row=4, column=0)
format_combo = ttk.Combobox(frame, values=["MLOG", "SWR", "PHAS", "REAL", "IMAG"])
format_combo.grid(row=4, column=1)
format_combo.current(0)

tk.Label(frame, text="Marker Freq (MHz):").grid(row=5, column=0)
marker_entry = tk.Entry(frame)
marker_entry.grid(row=5, column=1)

# 버튼
apply_btn = tk.Button(frame, text="Apply Settings", command=apply_settings)
apply_btn.grid(row=6, column=0, columnspan=2, pady=10)

# 결과 표시
result_label = tk.Label(frame, text="Result will appear here")
result_label.grid(row=7, column=0, columnspan=2)

import pyvisa
import tkinter as tk
from tkinter import ttk

# Power supply GUI 생성
root = tk.Tk()
root.title("Power Supply Control")
root.geometry("1000x200")

# 장비 및 소스 목록
device_options = ["E3631A", "E3634A"]
source_options = {
    "E3631A": ["P6V", "P25V"],
    "E3634A": ["P25V", "P50V"]
}

# 채널별 선택값 저장
selected_devices = [tk.StringVar() for _ in range(5)]
selected_sources = [tk.StringVar() for _ in range(5)]
gpib_addresses = [tk.StringVar() for _ in range(5)]
current_limits = [tk.StringVar() for _ in range(5)]
default_voltages = [tk.StringVar() for _ in range(5)]
connection_labels = []

# GPIB 주소 숫자 입력 제한
def validate_numeric_input(value):
    return value.isdigit() or value == ""

validate_cmd = root.register(validate_numeric_input)

def update_source_options(event, ch):
    """ 사용자가 장비를 선택하면 해당 장비의 소스 옵션을 업데이트 """
    selected_device = selected_devices[ch].get()
    if selected_device in source_options:
        source_dropdowns[ch]["values"] = source_options[selected_device]

def apply_settings():
    """ 사용자가 지정한 GPIB 주소로 계측기 설정 적용 (E3631A와 E3634A 구분) """
    for ch in range(5):  
        gpib_address = gpib_addresses[ch].get()
        device_name = selected_devices[ch].get()
        source_name = selected_sources[ch].get()
        current_limit = current_limits[ch].get()
        default_voltage = default_voltages[ch].get()

        # 값이 입력되지 않은 채널은 설정하지 않음
        if not gpib_address or not device_name or not source_name or not default_voltage or not current_limit:
            print(f"⚠ CH{ch+1}: 설정값이 입력되지 않아 건너뜀")
            continue  # 다음 채널로 이동

        try:
            instrument = rm.open_resource(f"GPIB::{gpib_address}::INSTR")

            # E3631A 설정 (채널 포함)
            if device_name == "E3631A":
                instrument.write(f"APPL {source_name}, {default_voltage}, {current_limit}")
                
            # E3634A 설정 (출력 범위 설정 후 전압/전류 입력)
            elif device_name == "E3634A":
                instrument.write(f"VOLT:RANG {source_name}")  # P25V 또는 P50V 설정
                instrument.write(f"VOLT {default_voltage}")  # 전압 설정
                instrument.write(f"CURR {current_limit}")    # 전류 설정

            connection_labels[ch].config(text=f"CH{ch+1} 설정 완료!", fg="green")

        except Exception as e:
            connection_labels[ch].config(text=f"연결 실패: {e}", fg="red")
                                     
# 채널별 GUI 요소 생성
source_dropdowns = []
for ch in range(5):
    frame = tk.Frame(root)
    frame.pack(pady=5)

    # 채널 표시
    label = tk.Label(frame, text=f"Ch{ch+1}")
    label.pack(side="left")

    # 장비 선택 레이블 및 입력
    tk.Label(frame, text="장비 선택:").pack(side="left")
    device_dropdown = ttk.Combobox(frame, textvariable=selected_devices[ch], values=device_options)
    device_dropdown.pack(side="left", padx=5)
    device_dropdown.bind("<<ComboboxSelected>>", lambda event, ch=ch: update_source_options(event, ch))

    # 소스 선택 레이블 및 입력
    tk.Label(frame, text="소스 선택:").pack(side="left")
    source_dropdown = ttk.Combobox(frame, textvariable=selected_sources[ch])
    source_dropdown.pack(side="left", padx=5)
    source_dropdowns.append(source_dropdown)

    # GPIB 주소 입력 레이블 및 필드
    tk.Label(frame, text="GPIB 주소:").pack(side="left")
    gpib_entry = tk.Entry(frame, textvariable=gpib_addresses[ch], width=5, validate="key", validatecommand=(validate_cmd, "%P"))
    gpib_entry.pack(side="left", padx=5)
    gpib_entry.insert(0, str(ch + 1))

    # 전류 제한 입력 레이블 및 필드
    tk.Label(frame, text="전류 제한(A):").pack(side="left")
    current_entry = tk.Entry(frame, textvariable=current_limits[ch], width=5)
    current_entry.pack(side="left", padx=5)
    current_entry.insert(0, "1.0")

    # 기본 전압 입력 레이블 및 필드
    tk.Label(frame, text="기본 전압(V):").pack(side="left")
    default_voltage_entry = tk.Entry(frame, textvariable=default_voltages[ch], width=5)
    default_voltage_entry.pack(side="left", padx=5)
    default_voltage_entry.insert(0, "3.8")

    conn_status = tk.Label(frame, text="설정 대기 중", fg="gray")
    conn_status.pack(side="left", padx=10)
    connection_labels.append(conn_status)

# 계측기 설정 버튼
btn_apply_settings = tk.Button(root, text="계측기 설정", command=apply_settings)
btn_apply_settings.pack(side="top", pady=10)

root.mainloop()
