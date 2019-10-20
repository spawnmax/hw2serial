from tkinter import ttk, messagebox, Tk, N, W, E, S, Label, StringVar, Button, Frame, BooleanVar, Canvas
from clr import AddReference
from serial import Serial, SerialException
from os import path
from ctypes import windll
import time
from sys import executable
from json import dump, load

config_file = 'hw2serial.conf'
ico_path = 'hw2serial.ico'
time_sensor = '#Time'
cpus_freq = '#CPUs Clock'
cpus_temp = '#CPUs Temp'
cpus_load = '#CPUs Load'
hwTypes = ['Mainboard','SuperIO','CPU','RAM','GpuNvidia','GpuAti','TBalancer','Heatmaster','HDD']
sensorTypes = ['Voltage','Clock','Temperature','Load','Fan','Flow','Control','Level','Factor','Power','Data','SmallData','Data']
sensorUnits = [' V', 'MHz', '°C', '%', ' RPM', ' L/h', '%', '%', '', 'W', '', '', ''] # \u00B0C
run_key = 'Software\Microsoft\Windows\CurrentVersion\Run'
title = 'HW2Serial'

refresh_periods = [2, 1, 0.5, 0.2, 0.1]      # period in seconds
port_rates = ['9600', '19200', '38400', '57600', '115200']
# Default config
defaults = {
    'serial_port': '',
    'baudrate': 9600,
    'refresh': 1,
    'minimize_to_tray': True,
    'launch_at_startup': False,
    'admin_only': True,
    'sensors': [time_sensor,'CPU Package Temperature','CPU Total Load', '', 'GPU Core Clock', '']
}

class HW2Serial:
    def __init__(self, parent):
        self.root = parent
        self.root.title(title)

        self.conf = self.load_config()

        self.sensorValues = {}
        self.sensorUnits = {}
        self.init_ohm()
        self.fetch_stats()
        self.available_ports = self.check_serial_ports()
        # self.root.bind('<Return>', self.save_config())

        self.sensors_frame = None
        self.draw_GUI()

        self.update_all()

    def update_all(self):
        """Update Hardware info, show it on GUI, send to serial. Launched every {} seconds""".format(self.conf['refresh'])
        self.fetch_stats()
        for i, sensor in enumerate(self.conf['sensors']):
            self.format_sensor_value(i)
        self.transfer_data()
        self.root.after(int(self.conf['refresh']*1000), self.update_all)

    def transfer_data(self):
        """Sending data to serial port"""
        try:
            with Serial(self.conf['serial_port'], self.conf['baudrate'], timeout=1) as ser:
                ser.write(' '.join(self.data2transfer).encode())
                #self.canvas.itemconfig(self.connection, text='Connected')
                self.canvas.itemconfig(self.circle, fill='green2')
        except (OSError, SerialException):
            #self.canvas.itemconfig(self.connection, text='Not connected')
            self.canvas.itemconfig(self.circle, fill='red2')

    def init_ohm(self):
        """Initialize sensors object from OpenHardwareMonitorLib.dll"""
        AddReference('OpenHardwareMonitorLib')
        from OpenHardwareMonitor import Hardware

        self.handle = Hardware.Computer()
        self.handle.MainboardEnabled = True
        self.handle.CPUEnabled = True
        self.handle.RAMEnabled = True
        self.handle.GPUEnabled = True
        self.handle.HDDEnabled = True
        self.handle.Open()

    def fetch_stats(self):
        """Read hardware info"""
        self.sensorValues[time_sensor] = time.strftime("%H:%M:%S")
        self.cores_freq, self.cores_temp, self.cores_load = [], [], []
        for i in self.handle.Hardware:
            i.Update()
            for sensor in i.Sensors:
                self.parse_sensor(sensor)
            for j in i.SubHardware:
                j.Update()
                for subsensor in j.Sensors:
                    self.parse_sensor(subsensor)
        self.sensorValues[cpus_freq] = ":".join(self.cores_freq)
        self.sensorValues[cpus_load] = ":".join(self.cores_load)
        self.sensorValues[cpus_temp] = ":".join(self.cores_temp)
        self.sensorNames = sorted(list(self.sensorValues.keys()))
        self.update_data2transfer()

    def parse_sensor(self, sensor):
        """Parsing sensor data and save it to class variables"""
        value = sensor.Value if sensor.Value else '0'
        unit = sensorUnits[sensor.SensorType] if sensor.Value else ''
        unit = unit if sensorTypes[sensor.SensorType] not in ['Data', 'SmallData', 'Factor'] else ''
        if sensorTypes[sensor.SensorType] in ['Data', 'SmallData', 'Factor']:
            sensorname = '{}'.format(sensor.Name)
        else:
            sensorname = '{} {}'.format(sensor.Name, sensorTypes[sensor.SensorType]) # sensor.Hardware.Name
        if sensorname in ['Used Memory', 'Available Memory']:
            sensorvalue = '{:.3f}'.format(float(value))
        elif sensorTypes[sensor.SensorType] in ['Voltage']:
            sensorvalue = '{:.1f}'.format(float(value))
        else:
            sensorvalue = '{:.0f}'.format(float(value))
        sensorunit = unit
        if 'CPU Core ' in sensor.Name:
            if sensorTypes[sensor.SensorType] == 'Clock':
                self.cores_freq.append(sensorvalue)
            if sensorTypes[sensor.SensorType] == 'Load':
                self.cores_load.append(sensorvalue)
            if sensorTypes[sensor.SensorType] == 'Temperature':
                self.cores_temp.append(sensorvalue)
        self.sensorValues[sensorname] = sensorvalue
        self.sensorUnits[sensorname] = sensorunit

    def check_serial_ports(self):
        """ Lists serial port names """
        ports = ['COM%s' % (i + 1) for i in range(256)]
        result = ['']
        for port in ports:
            try:
                if port == self.conf['serial_port']: continue
                s = Serial(port)
                s.close()
                result.append(port)
            except (OSError, SerialException):
                pass
        return result

    def update_data2transfer(self):
        """Get sensors from combobox that should be transfered to serial port"""
        self.data2transfer = [None] * len(self.conf['sensors'])
        for i, sensor_name in enumerate(self.conf['sensors']):
            self.data2transfer[i] = self.sensorValues.get(sensor_name, '0')

    def minimize(self):
        if self.conf['minimize_to_tray']:
            self.root.withdraw()
        else:
            self.root.iconify()

    def quit(self, sender=None, args=None):
        self.tray_icon.Visible = False
        self.components.Dispose()
        self.tray_icon.Dispose()
        self.root.destroy()

    def draw_GUI(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.root.resizable(False, False)
        self.sensors_frame = self.draw_sensors_frame(self.root)
        self.configs_frame = self.draw_frame_configs(self.root)
        self.buttons_frame = self.draw_control_buttons(self.root)
        self.update_config()
        self.tray_icon = self.icon()

    def icon(self):
        """Creating icon in system tray"""
        AddReference('System.ComponentModel')
        AddReference('System.Windows.Forms')
        AddReference('System.Drawing')
        from System.ComponentModel import Container
        from System.Windows.Forms import NotifyIcon, MenuItem, ContextMenu
        from System.Drawing import Icon

        self.components = Container()
        context_menu = ContextMenu()
        menu_item = MenuItem('Show')
        menu_item.Click += self.open_from_tray
        context_menu.MenuItems.Add(menu_item)
        menu_item = MenuItem('Quit')
        menu_item.Click += self.quit
        context_menu.MenuItems.Add(menu_item)
        notifyIcon = NotifyIcon(self.components)
        notifyIcon.Icon = Icon(ico_path)
        notifyIcon.Text = title
        notifyIcon.Visible = True
        notifyIcon.ContextMenu = context_menu
        notifyIcon.DoubleClick += self.open_from_tray

        return notifyIcon

    def draw_sensors_frame(self, parent):
        if self.sensors_frame:
            self.sensors_frame.grid_forget()
        sensors_frame = ttk.LabelFrame(parent, text='Hardware Data')
        sensors_frame.grid(column=0, row=0, padx=5, pady=5, ipady=3, sticky=(N, W, E)) #.pack(side="top", fill="both", expand=True, padx=5, pady=5, ipady=3)
        sensors_frame.columnconfigure(3, weight=1)

        self.number_labels, self.sensor_combos, self.value_labels = [], [], []
        Label(sensors_frame, text="#").grid(column=0, row=0, sticky=W)
        Label(sensors_frame, text="Data").grid(column=1, row=0, sticky=W)
        Label(sensors_frame, text="Value").grid(column=2, row=0, sticky=W)
        for i, sensor in enumerate(self.conf['sensors']):
            self.number_labels.append(StringVar())
            Label(sensors_frame, textvariable=self.number_labels[i]).grid(column=0, row=i+1, sticky=W)
            self.sensor_combos.append(StringVar())
            cmb = ttk.Combobox(sensors_frame, textvariable=self.sensor_combos[i], values=self.sensorNames, width=25, state='readonly')
            cmb.grid(column=1, row=i+1, sticky=W)
            cmb.bind('<<ComboboxSelected>>', self.change_sensor)
            self.value_labels.append(StringVar())
            Label(sensors_frame, textvariable=self.value_labels[i]).grid(column=2, row=i+1, sticky=W)
            if i>0:
                Button(sensors_frame, text='-', height=1, borderwidth=1,
                       command = lambda i=i: self.remove_sensor(i)).grid(column=3, row=i+1, sticky=E)

        self.update_sensors_show()
        if i<19:
            Button(sensors_frame, text='+', command=self.add_sensor, height=1).grid(column=0, row=i+2, sticky=W)
        for child in sensors_frame.winfo_children(): child.grid_configure(padx=3)

        return sensors_frame

    def add_sensor(self):
        self.conf['sensors'].append('')
        self.sensors_frame = self.draw_sensors_frame(self.root)

    def remove_sensor(self, i):
        if i>0:
            del self.conf['sensors'][i]
            self.sensors_frame = self.draw_sensors_frame(self.root)

    def draw_frame_configs(self, parent):
        configs_frame = Frame(parent)
        configs_frame.grid(column=0, row=1, sticky=(W, E))
        #configs_frame.columnconfigure(0, weight=1)
        combos_width = 15

        self.serial_port = StringVar()
#        Label(configs_frame, text="Serial port").grid(column=0, row=1, sticky=W)
        self.canvas = Canvas(configs_frame, width=80, height=16)
        r, x, y = 5, 75, 11
        self.circle = self.canvas.create_oval(x-r, y-r, x+r, y+r, fill="red", outline="")
        self.connection = self.canvas.create_text(4, 11, anchor=W, text='Serial port')
        self.canvas.grid(column=0, row=1, sticky=W)
        self.port_cmb = ttk.Combobox(configs_frame, textvariable=self.serial_port, values=self.check_serial_ports(),
                                     postcommand=self.update_ports, state='readonly', width=combos_width)
        self.port_cmb.grid(column=2, row=1, sticky=W)
        self.port_cmb.bind('<<ComboboxSelected>>', self.change_port)

        self.baudrate = StringVar()
        Label(configs_frame, text="Baudrate").grid(column=0, row=2, sticky=W)
        cmb = ttk.Combobox(configs_frame, textvariable=self.baudrate, values=port_rates, state='readonly', width=combos_width)
        cmb.grid(column=2, row=2, sticky=W)
        cmb.bind('<<ComboboxSelected>>', self.change_baudrate)

        self.refresh_period = StringVar()
        Label(configs_frame, text="Refresh period").grid(column=0, row=3, sticky=W)
        cmb = ttk.Combobox(configs_frame, textvariable=self.refresh_period, values=refresh_periods, state='readonly', width=combos_width)
        cmb.grid(column=2, row=3, sticky=W)
        cmb.bind('<<ComboboxSelected>>', self.change_refresh)

        self.launch_at_start = BooleanVar()
        chb = ttk.Checkbutton(configs_frame, text='Launch at startup', variable=self.launch_at_start, command=self.change_launch_at_start)
        chb.grid(column=3, row=1, sticky=(W, E))

        self.minimize_to_tray = BooleanVar()
        chb = ttk.Checkbutton(configs_frame, text='Minimize to tray', variable=self.minimize_to_tray, command=self.change_minimize_to_tray)
        chb.grid(column=3, row=2, sticky=(W, E))

        self.admin_only = BooleanVar()
        chb = ttk.Checkbutton(configs_frame, text='Run as admin only', variable=self.admin_only, command=self.change_admin_only)
        chb.grid(column=3, row=3, sticky=(W, E))

        for child in configs_frame.winfo_children(): child.grid_configure(padx=2)
        return configs_frame

    def draw_control_buttons(self, parent):
        buttons_frame = Frame(parent)
        buttons_frame.grid(column=0, row=2, padx=5, pady=3, sticky=(E, S))
        #buttons_frame.columnconfigure(0, weight=1)

        Button(buttons_frame, text='Default Config', command=self.restore_defaults).grid(column=1, row=0, sticky=(W, E))
        Button(buttons_frame, text='Load Saved Config', command=self.load_config_button).grid(column=2, row=0, sticky=(W, E))
        Button(buttons_frame, text='Save Config', command=self.save_config).grid(column=3, row=0, sticky=(W, E))

        Button(buttons_frame, text='Minimize', command=self.minimize).grid(column=2, row=5, sticky=(W, E))
        Button(buttons_frame, text='Quit', command=self.quit).grid(column=3, row=5, sticky=(W, E))

        for child in buttons_frame.winfo_children(): child.grid_configure(padx=2, pady=1)

        return buttons_frame

    def update_ports(self):
        self.port_cmb['values'] = self.check_serial_ports()

    def format_sensor_value(self, i):
        sensor = self.conf['sensors'][i]
        value = self.sensorValues.get(sensor, 0)
        if value != 'n/a' and sensor not in [time_sensor, cpus_freq, cpus_load, cpus_temp]:
            value = '{}{}'.format(value, self.sensorUnits.get(sensor, ''))
        self.value_labels[i].set(value)

    def change_sensor(self, event=None):
        """Change which sensor to send to serial port"""
        if event:       # check if it's called from binding
            for i, sensor in enumerate(self.conf['sensors']):
                if self.sensor_combos[i].get() != sensor:
                    self.conf['sensors'][i] = self.sensor_combos[i].get()
                    self.format_sensor_value(i)

    def update_sensors_show(self):
        for i, sensor in enumerate(self.conf['sensors']):
            self.number_labels[i].set(i)
            self.sensor_combos[i].set(sensor)
            self.format_sensor_value(i)

    def update_config(self):
        """Redraw config parameters in GUI from dict 'conf'"""
        self.update_sensors_show()
        self.serial_port.set(self.conf.get('serial_port', ''))
        self.baudrate.set(self.conf.get('baudrate', defaults.get('baudrate', 9600)))
        self.refresh_period.set(self.conf.get('refresh', defaults.get('refresh', 1)))
        self.launch_at_start.set(self.conf.get('launch_at_startup', defaults.get('launch_at_startup', 1)))
        self.minimize_to_tray.set(self.conf.get('minimize_to_tray', defaults.get('minimize_to_tray', False)))
        self.admin_only.set(self.conf.get('admin_only', defaults.get('admin_only', False)))
        self.change_minimize_to_tray()
        self.change_launch_at_start()

    def load_config_button(self):
        if messagebox.askyesno(message='Are you sure you want to reload config?', icon = 'question', title = 'Loading saved config'):
            self.conf = self.load_config()
            self.update_config()

    def save_config(self):
        if messagebox.askyesno(message='Are you sure you want to save current config?', icon='question', title='Saving config'):
            with open(config_file, 'w') as conf_file:
                dump(self.conf, conf_file)

    def load_config(self):
        """Loading config from {}""".format(config_file)
        try:
            with open(config_file) as conf_file:
                conf = load(conf_file)
        except IOError:
            conf = {}
        if 'sensors' not in conf:
            conf = defaults.copy()
        self.savedConf = conf
        return conf

    def restore_defaults(self):
        if messagebox.askyesno(message='Are you sure you want to set default config?', icon = 'question', title = 'Loading default config'):
            self.conf = defaults.copy()
            self.update_config()

    def change_admin_only(self):
        self.conf['admin_only'] = self.admin_only.get()

    def change_minimize_to_tray(self):
        self.conf['minimize_to_tray'] = self.minimize_to_tray.get()
        if self.conf['minimize_to_tray']:
            self.root.protocol("WM_DELETE_WINDOW", self.root.withdraw)
        else:
            self.root.protocol("WM_DELETE_WINDOW", self.quit)

    def change_launch_at_start(self):
        self.conf['launch_at_startup'] = self.launch_at_start.get()
        if self.conf['launch_at_startup']:
            self.set_launch_at_boot()
        else:
            self.remove_launch_at_boot()

    def change_port(self, event=None):
        if event:
            value = self.serial_port.get()
            self.conf['serial_port'] = value

    def change_baudrate(self, event=None):
        if event:
            value = self.baudrate.get()
            self.conf['baudrate'] = int(value)

    def change_refresh(self, event=None):
        if event:
            value = self.refresh_period.get()
            self.conf['refresh'] = int(value) if value in ['1', '2'] else float(value)

    def open_from_tray(self, sender, args):
        self.root.deiconify()

    def set_launch_at_boot(self):
        """Setting start on boot in windows registry"""
        from Microsoft.Win32 import Registry

        if 'python.exe' in executable:
            command_run = executable + ' ' + path.dirname(path.realpath(__file__))
        else:
            command_run = executable
        reg_key = Registry.CurrentUser.OpenSubKey(run_key, True)
        if reg_key.GetValue(title, 'no_key') != 'no_key':
            return None
        reg_key.SetValue(title, command_run)

    def remove_launch_at_boot(self):
        """Removing start on boot in windows registry"""
        from Microsoft.Win32 import Registry

        reg_key = Registry.CurrentUser.OpenSubKey(run_key, True)
        if reg_key.GetValue(title, 'no_key') == 'no_key':
            return None
        reg_key.DeleteValue(title)
        reg_key.Close()

def is_admin():
    """Check if program runs with admin rights. Needed for some sensors"""
    try:
        return windll.shell32.IsUserAnAdmin()
    except:
        return False

def icon_path(ico_path):
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = path.abspath(".")
    return path.join(base_path, ico_path)

if __name__ == '__main__':
    root = Tk()
    ico_path = icon_path('hw2serial.ico')
    h2s = HW2Serial(root)
    if not is_admin() and h2s.conf.get('admin_only', False):
        windll.shell32.ShellExecuteW(None, "runas", executable, __file__, None, 1)
    else:
        root.iconbitmap(ico_path)
        root.mainloop()
