import tkinter as tk
import socket
import threading
import mysql.connector
from mysql.connector import Error
import subprocess


def run_powershell_command(command):
    try:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE

        process = subprocess.Popen(
            ["powershell", "-Command", command],
            startupinfo=startupinfo,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            universal_newlines=True,
        )
        output, error = process.communicate()

        if error:
            return f"Error: {error}"
        return output.strip()
    except subprocess.CalledProcessError as e:
        return f"Error: {e}"


def get_device_name():
    command = "(Get-WmiObject -Class Win32_ComputerSystem).Name"
    return run_powershell_command(command)


def get_monitor_info():
    command = "Get-WmiObject WmiMonitorID -Namespace root\\wmi | ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_.UserFriendlyName -ne 0) }"
    output = run_powershell_command(command)
    monitors = output.strip().split("\n")
    monitor_info = f"{len(monitors)} monitors/"
    for i, monitor in enumerate(monitors, 1):
        monitor_info += f"Monitor {i}: {monitor.strip()}/"
    return monitor_info.strip()


def get_mac_address():
    command = "Get-NetAdapter | Select-Object -First 1 -ExpandProperty MacAddress"
    return run_powershell_command(command)


def get_cpu_info():
    command = (
        "Get-WmiObject -Class Win32_Processor | Select-Object -ExpandProperty Name"
    )
    return run_powershell_command(command)


def get_ram_info():
    command = """
    $ram = Get-WmiObject -Class Win32_PhysicalMemory
    $totalCapacity = ($ram | Measure-Object -Property Capacity -Sum).Sum / 1GB
    $modules = $ram | ForEach-Object {
        $capacity = $_.Capacity / 1GB
        $speed = $_.Speed
        "$($capacity.ToString('F2')) GB Bus $speed MHz"
    }
    "$($ram.Count) modules, Total: $($totalCapacity.ToString('F2')) GB/$($modules -join '/')"
    """
    return run_powershell_command(command)


def get_disk_info():
    command = """
    $disks = Get-PhysicalDisk | Select-Object FriendlyName, MediaType, Size
    $totalCapacity = ($disks | Measure-Object -Property Size -Sum).Sum / 1GB
    $diskDetails = $disks | ForEach-Object {
        $size = $_.Size / 1GB
        $type = if ($_.MediaType -eq 'SSD') { 'SSD' } else { 'HDD' }
        "$($_.FriendlyName) ($type, $($size.ToString('F2')) GB)"
    }
    "$($disks.Count) disks, Total: $($totalCapacity.ToString('F2')) GB/$($diskDetails -join '/')"
    """
    return run_powershell_command(command)


def get_windows_version():
    command = "(Get-WmiObject -Class Win32_OperatingSystem).Caption"
    return run_powershell_command(command)


def get_system_info():
    cpu_name = get_cpu_info()
    ram_info = get_ram_info()
    disk_info = get_disk_info()
    monitor_info = get_monitor_info()

    # GPU info
    gpu_command = "Get-WmiObject -Class Win32_VideoController | Select-Object -ExpandProperty Name"
    gpu_info = run_powershell_command(gpu_command).split("\n")

    # IP Address
    hostname = socket.gethostname()
    ip_address = socket.gethostbyname(hostname)

    # MAC Address
    mac_address = get_mac_address()

    # Windows username
    username = get_device_name()

    # Windows version
    windows_version = get_windows_version()

    # Motherboard info
    motherboard_command = (
        "Get-WmiObject -Class Win32_BaseBoard | Select-Object -ExpandProperty Product"
    )
    motherboard_info = run_powershell_command(motherboard_command)

    return (
        cpu_name,
        ram_info,
        disk_info,
        gpu_info,
        ip_address,
        username,
        windows_version,
        motherboard_info,
        mac_address,
        monitor_info,
    )


def create_connection():
    """Create a database connection to MySQL"""
    try:
        conn = mysql.connector.connect(
            host="198.168.0.54",
            user="sa",
            password="passwd",
            database="system_info_db",
        )
        if conn.is_connected():
            return conn
    except Error as e:
        print(f"Error: {e}")
    return None


def insert_system_info(
    conn,
    cpu,
    ram,
    storage,
    gpu,
    ip_address,
    username,
    windows_version,
    motherboard_info,
    mac_address,
    monitor_info,
):
    """Insert system information into the database"""
    try:
        cursor = conn.cursor()

        insert_sql = """
        INSERT INTO system_info(cpu, ram, storage, gpu, ip_address, username, windows_version, motherboard_info, mac_address, monitor_info)
        VALUES(%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        cursor.execute(
            insert_sql,
            (
                cpu,
                ram,
                storage,
                "/".join(gpu),
                ip_address,
                username,
                windows_version,
                motherboard_info,
                mac_address,
                monitor_info,
            ),
        )

        conn.commit()

    except Error as e:
        print(f"Error: {e}")


def show_system_info():
    loading_label.config(text="กำลังโหลดข้อมูล...")
    loading_label.pack(expand=True)

    def fetch_and_display():
        (
            cpu_name,
            ram_info,
            disk_info,
            gpu_info,
            ip_address,
            username,
            windows_version,
            motherboard_info,
            mac_address,
            monitor_info,
        ) = get_system_info()

        info = (
            f"CPU: {cpu_name}\n"
            f"RAM: {ram_info}\n"
            f"Disks: {disk_info}\n"
            f"GPU: {'/'.join(gpu_info)}\n"
            f"IP Address: {ip_address}\n"
            f"Windows Account: {username}\n"
            f"Windows Version: {windows_version}\n"
            f"Motherboard: {motherboard_info}\n"
            f"MAC Address: {mac_address}\n"
            f"Monitors: {monitor_info}"
        )

        label.config(text=info)
        loading_label.pack_forget()

        conn = create_connection()
        if conn:
            insert_system_info(
                conn,
                cpu_name,
                ram_info,
                disk_info,
                gpu_info,
                ip_address,
                username,
                windows_version,
                motherboard_info,
                mac_address,
                monitor_info,
            )
            conn.close()

            # ปิดโปรแกรมหลังจาก insert ข้อมูลเสร็จ
            root.after(500, root.quit)  # รอ 0.5 วินาทีแล้วปิดโปรแกรม

    threading.Thread(target=fetch_and_display, daemon=True).start()


# Create the main window
root = tk.Tk()
root.title("hossystem")
root.geometry("800x400")

# ลบปุ่มปิดโปรแกรม
root.overrideredirect(True)

# Create and place the label
label = tk.Label(root, text="", font=("Helvetica", 12), anchor="w", justify="left")
label.pack(padx=20, pady=20, fill="both", expand=True)

# Create loading label
loading_label = tk.Label(root, text="", font=("Helvetica", 12))
loading_label.place(relx=0.5, rely=0.5, anchor="center")

# Display system information on startup
show_system_info()

# Start the GUI event loop
root.mainloop()
