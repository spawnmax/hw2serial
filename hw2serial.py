from tkinter import *
from tkinter import ttk, messagebox
import time
import json
import clr
import os
import ctypes, sys

config_file = 'hw2serial.conf'
hwTypes = ['Mainboard','SuperIO','CPU','RAM','GpuNvidia','GpuAti','TBalancer','Heatmaster','HDD']
sensorTypes = ['Voltage','Clock','Temperature','Load','Fan','Flow','Control','Level','Factor','Power','Data','SmallData','Data']
sensorUnits = [' V', 'MHz', '°C', '%', ' RPM', ' L/h', '%', '%', '', 'W', '', '', ''] # \u00B0C
run_key = 'Software\Microsoft\Windows\CurrentVersion\Run'
prog_name = 'hw2serial.py'
title = 'HW2Serial'

# Default config
refreshPeriods = [2, 1, 0.5, 0.2, 0.1]
defaults = {
    'refresh': 1,
    'minimize_to_tray': True,
    'launch_at_startup': False,
    'sensors': ['Time','CPU Package Temperature','CPU Total Load', '', 'GPU Core Clock', '']
}

class HW2Serial:
    def __init__(self, parent):
        self.root = parent
        self.root.title(title)

        self.savedConf = None
        self.conf = self.load_config()

        self.sensorValues = {}
        self.sensorUnits = {}
        self.data2transfer = [None] * len(self.conf['sensors'])
        self.init_ohm()
        self.fetch_stats()
        # self.root.bind('<Return>', self.save_config())

        self.draw_GUI()

        self.update_all()

    def load_config(self):
        try:
            with open(config_file) as conf_file:
                conf = json.load(conf_file)
        except IOError:
            conf = {}
        if 'sensors' not in conf:
            conf = defaults
        self.savedConf = conf
        return conf

    def update_config(self):
        for i, sensor in enumerate(self.conf['sensors']):
            self.number_labels[i].set(i)
            self.sensor_combos[i].set(sensor)
            value = self.sensorValues.get(sensor, 0)
            if value != 'n/a' and sensor not in ['Time']:
                value = '{}{}'.format(round(float(value), 2), self.sensorUnits.get(sensor, ''))
            self.value_labels[i].set(value)
        self.refresh_period.set(self.conf['refresh'])
        self.launch_at_start.set(self.conf['launch_at_startup'])
        self.minimize_to_tray.set(self.conf['minimize_to_tray'])
        self.change_minimize_to_tray()
        self.change_launch_at_start()

    def load_config_button(self):
        if messagebox.askyesno(message='Are you sure you want to reload config?', icon = 'question', title = 'Loading saved config'):
            self.conf = self.load_config()
            self.update_config()

    def save_config(self):
        if messagebox.askyesno(message='Are you sure you want to save current config?', icon='question', title='Saving config'):
            with open(config_file, 'w') as conf_file:
                json.dump(self.conf, conf_file)

    def init_ohm(self):
        clr.AddReference('OpenHardwareMonitorLib')
        from OpenHardwareMonitor import Hardware

        self.handle = Hardware.Computer()
        self.handle.MainboardEnabled = True
        self.handle.CPUEnabled = True
        self.handle.RAMEnabled = True
        self.handle.GPUEnabled = True
        self.handle.HDDEnabled = True
        self.handle.Open()

    def fetch_stats(self):
        self.sensorValues['Time'] = time.strftime("%H:%M:%S")
        for i in self.handle.Hardware:
            i.Update()
            for sensor in i.Sensors:
                self.parse_sensor(sensor)
            for j in i.SubHardware:
                j.Update()
                for subsensor in j.Sensors:
                    self.parse_sensor(subsensor)
        self.sensorNames = sorted(list(self.sensorValues.keys()))
        self.update_data2transfer()

    def parse_sensor(self, sensor):
        value = sensor.Value if sensor.Value else 'n/a'
        unit = sensorUnits[sensor.SensorType] if value != 'n/a' else ''
        unit = unit if sensorTypes[sensor.SensorType] not in ['Data', 'SmallData', 'Factor'] else ''
        if sensorTypes[sensor.SensorType] in ['Data', 'SmallData', 'Factor']:
            sensorname = '{}'.format(sensor.Name)
        else:
            sensorname = '{} {}'.format(sensor.Name, sensorTypes[sensor.SensorType]) # sensor.Hardware.Name
        sensorvalue = str(value)
        sensorunit = unit
        self.sensorValues[sensorname] = sensorvalue
        self.sensorUnits[sensorname] = sensorunit

    def update_data2transfer(self):
        for i, sensor_name in enumerate(self.conf['sensors']):
            self.data2transfer[i] = self.sensorValues.get(sensor_name, 0)

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
        self.update_config()
        self.tray_icon = self.icon()

    def draw_sensors_frame(self, parent):
        sensors_frame = ttk.LabelFrame(parent, text='Hardware Data')
        sensors_frame.pack(side="top", fill="both", expand=True, padx=5, pady=5, ipady=3) #.grid(column=0, row=0, padx=5, pady=5, sticky=(N, W, E))
        sensors_frame.columnconfigure(1, weight=1)

        self.number_labels = [None] * len(self.conf['sensors'])
        self.sensor_combos = [None] * len(self.conf['sensors'])
        self.value_labels = [None] * len(self.conf['sensors'])
        Label(sensors_frame, text="#").grid(column=0, row=0, sticky=W)
        Label(sensors_frame, text="Data").grid(column=1, row=0, sticky=W)
        Label(sensors_frame, text="Value").grid(column=2, row=0, sticky=W)
        for i, sensor in enumerate(self.conf['sensors']):
            self.number_labels[i] = StringVar()
            lbl = Label(sensors_frame, textvariable=self.number_labels[i])
            self.sensor_combos[i] = StringVar()
            cmb = ttk.Combobox(sensors_frame, textvariable=self.sensor_combos[i], values=self.sensorNames, width=25, state='readonly')
            self.value_labels[i] = StringVar()
            vlbl = Label(sensors_frame, textvariable=self.value_labels[i])

            lbl.grid(column=0, row=i+1, sticky=W)
            cmb.grid(column=1, row=i+1, sticky=(W, E))
            cmb.bind('<<ComboboxSelected>>', self.change_sensor)
            vlbl.grid(column=2, row=i+1, sticky=E)
        for child in sensors_frame.winfo_children(): child.grid_configure(padx=3)

        return sensors_frame

    def update_all(self):
        self.fetch_stats()
        for i, sensor in enumerate(self.conf['sensors']):
            value = self.sensorValues.get(sensor, 0)
            if value != 'n/a' and sensor not in ['Time']:
                value = '{}{}'.format(round(float(value), 2), self.sensorUnits.get(sensor, ''))
            self.value_labels[i].set(value)

        self.root.after(int(self.conf['refresh']*1000), self.update_all)

    def draw_frame_configs(self, parent):
        configs_frame = Frame(parent)
        configs_frame.pack(side="bottom", fill="both", expand=True, padx=5, pady=5) #.grid(column=0, row=0, padx=5, pady=5, sticky=(W, E, S))
        configs_frame.columnconfigure(0, weight=1)

        self.refresh_period = StringVar()
        Label(configs_frame, text="Data refresh period").grid(column=2, row=1, sticky=W)
        cmb = ttk.Combobox(configs_frame, textvariable=self.refresh_period, values=refreshPeriods, state='readonly')
        cmb.grid(column=2, row=2, sticky=(W, E))
        cmb.bind('<<ComboboxSelected>>', self.change_refresh)

        self.launch_at_start = BooleanVar()
        chb = ttk.Checkbutton(configs_frame, text='Launch at startup', variable=self.launch_at_start, command=self.change_launch_at_start)
        chb.grid(column=3, row=1, sticky=(W, E))

        self.minimize_to_tray = BooleanVar()
        chb = ttk.Checkbutton(configs_frame, text='Minimize to tray', variable=self.minimize_to_tray, command=self.change_minimize_to_tray)
        chb.grid(column=3, row=2, sticky=(W, E))

        Button(configs_frame, text='Restore Default Config', command=self.restore_defaults).grid(column=1, row=3, sticky=(W, E))
        Button(configs_frame, text='Load Saved Config', command=self.load_config_button).grid(column=2, row=3, sticky=(W, E))
        Button(configs_frame, text='Save Config', command=self.save_config).grid(column=3, row=3, sticky=(W, E))

        Button(configs_frame, text='Minimize', command=self.minimize).grid(column=2, row=5, sticky=(W, E))
        Button(configs_frame, text='Quit', command=self.quit).grid(column=3, row=5, sticky=(W, E))

        for child in configs_frame.winfo_children(): child.grid_configure(padx=2, pady=2)

        return configs_frame

    def restore_defaults(self):
        if messagebox.askyesno(message='Are you sure you want to set default config?', icon = 'question', title = 'Loading default config'):
            self.conf = defaults
            self.update_config()

    def change_sensor(self, event=None):
        if event:
            for i, sensor in enumerate(self.conf['sensors']):
                if self.sensor_combos[i].get() != sensor:
                    self.conf['sensors'][i] = self.sensor_combos[i].get()
                    value = self.sensorValues.get(self.sensor_combos[i].get(), 0)
                    if value != 'n/a' and self.sensor_combos[i].get() not in ['Time']:
                        value = '{}{}'.format(round(float(value), 2), self.sensorUnits.get(self.sensor_combos[i].get(), ''))
                    self.value_labels[i].set(value)

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

    def change_refresh(self, event=None):
        if event:
            value = self.refresh_period.get()
            self.conf['refresh'] = int(value) if value in ['1', '2'] else float(value)

    def open_from_tray(self, sender, args):
        self.root.deiconify()

    def icon(self):
        clr.AddReference('System.ComponentModel')
        clr.AddReference('System.Windows.Forms')
        clr.AddReference('System.Drawing')
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
        notifyIcon.Icon = Icon('hw2serial.ico')
        notifyIcon.Text = title
        notifyIcon.Visible = True
        notifyIcon.ContextMenu = context_menu
        notifyIcon.DoubleClick += self.open_from_tray

        return notifyIcon

    def set_launch_at_boot(self):
        from Microsoft.Win32 import Registry

        path = os.path.dirname(os.path.realpath(__file__))
        reg_key = Registry.CurrentUser.OpenSubKey(run_key, True)
        if reg_key.GetValue(title, 'no_key') != 'no_key':
            return None
        reg_key.SetValue(title, os.path.join(path,prog_name))

    def remove_launch_at_boot(self):
        from Microsoft.Win32 import Registry

        reg_key = Registry.CurrentUser.OpenSubKey(run_key, True)
        if reg_key.GetValue(title, 'no_key') == 'no_key':
            return None
        reg_key.DeleteValue(title)
        reg_key.Close()

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if __name__ == '__main__':
    if is_admin():
        root = Tk()
        HW2Serial(root)
        root.mainloop()
    else:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)