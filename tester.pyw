#!/usr/bin/env python3
"""Resistance Measuring software"""

import csv
import os
import json

from datetime import datetime
from functools import partial
from glob import glob
from queue import Queue
from random import randint
from threading import Thread
from time import sleep
from tkinter import ttk, Tk, Menu, filedialog, Label, Button, \
        StringVar, messagebox as mBox, BOTH

import serial

from openpyxl import Workbook
from serial.tools.list_ports import comports

VERSION = '0.1'
TITLE = 'Res Measuring V{}'.format(VERSION)

LOOP_TIME = 150 # milliseconds

class Config:
    """Save a configuration"""

    path = os.path.join(os.path.dirname(os.path.realpath(__file__)),
            'config.json')

    comport = None
    test_wo_sensor = False

    def __init__(self):

        if not os.path.exists(self.path):
            return

        with open(self.path) as readfile:
            obj = json.load(readfile)

            if 'comport' in obj:
                self.comport = obj['comport']
            if 'test_wo_sensor' in obj:
                self.test_wo_sensor = obj['test_wo_sensor']

    def save(self):
        """Save to json file"""

        with open(self.path, 'w') as outfile:
            out = {
                'comport': self.comport,
                'test_wo_sensor': self.test_wo_sensor,
            }
            json.dump(out, outfile, indent=2)


class Recorder:
    """Record resistance value into csv log file
    """

    def __init__(self):
        month = datetime.now().strftime("%Y-%m")
        self.path = os.path.join(os.path.dirname(os.path.realpath(__file__)),
                '{}.csv'.format(month))


    def log(self, val):
        """Record value

        Args:
            val (float): resistance value

        Returns:
            string: filename
        """

        time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        month = time[:7]
        if month not in self.path:
            self.path = os.path.join(
                    os.path.dirname(os.path.realpath(__file__)),
                    '{}.csv'.format(month))

        with open(self.path, 'a+') as outfile:
            msg = '{},{}\n'.format(time, val)
            print('recording: {}'.format(msg), end='')

            outfile.write(msg)

        return month


class MainApp(Tk):
    """Main Window"""

    cfg = Config()
    recorder = Recorder()
    queue = Queue()
    serial = None

    def __init__(self, win):
        """Big frame"""

        self.master = win

        menu_bar = Menu(win)
        win.config(menu=menu_bar)

        self.export_menu = Menu()
        self._create_csv_log_menu()

        file_menu = Menu()
        file_menu.add_cascade(label='Export', menu=self.export_menu)
        file_menu.add_separator()
        file_menu.add_command(label='Exit', command=self.quit)
        menu_bar.add_cascade(label='File', menu=file_menu)

        self.resistance_label = Label(win, text='XXX.XX',
                font=('times', 192, 'bold'),
                height=2, width=8
                )
        self.resistance_label.config(background='green', foreground='black')
        self.resistance_label.grid(column=0, row=0,
                sticky='new', padx=12, pady=12)

        self.check_button = Button(win, text='Check',
                command=self.check,
                font=('times', 20),
                )
        self.check_button.config(width=20)
        self.check_button.grid(column=0, row=1, pady=10)
        self.check_button.focus()

        ports = self._portlist()
        self.selected_port = StringVar()
        self.port_combobox = ttk.Combobox(win,
                values=ports, textvariable=self.selected_port)
        self.port_combobox.grid(column=0, row=3)

        if self.cfg.comport and self.cfg.comport in ports:
            self.selected_port.set(self.cfg.comport)


    def quit(self):
        """Exit app"""

        self.master.quit()
        self.master.destroy()
        exit()


    def check(self):
        """Check button pressed"""

        #
        # Manage ui
        #
        self.resistance_label.config(bg='gray', text='?')
        self.check_button.config(state='disabled')
        self.check_button.focus()

        self.queue.queue.clear()
        self._thread = Thread(target=self._get_resistance,
                name='resistance', daemon=True)
        self._thread.start()
        self.master.after(LOOP_TIME, self._listen_for_resistance)

        ##
        ## Save selected comport
        ##
        #selected_port = self.selected_port.get()
        #if selected_port:
            #self.cfg.comport = selected_port
            #self.cfg.save()


    def _get_resistance(self):
        """Get resistance from serial port"""

        if self.cfg.test_wo_sensor:
            sleep(1)
            val = randint(25000, 45000) / 100
        else:
            try:
                port = self.selected_port.get()
                if not self.serial:
                    self.serial =  serial.Serial(port, 9600, timeout=3)
                    sleep(2)
                elif port != self.serial.name:
                    self.serial.close()
                    self.serial =  serial.Serial(port, 9600, timeout=3)
                    sleep(2)

                self.serial.write(b'M\n')
                data = self.serial.readline()
                if data:
                    print('{}: {}'.format(self.serial.name, data))
                    val = float(data.strip())
                else:
                    val = 0.0

                if port != self.cfg.comport:
                    self.cfg.comport = port
                    self.cfg.save()

            except Exception as e:
                print(e)
                self.serial = None
                val = -1
        self.queue.put(val)


    def _listen_for_resistance(self):
        """Get returned value from sensor reading"""

        if self._thread.is_alive():
            self.master.after(LOOP_TIME, self._listen_for_resistance)
            return
        #if self.queue.empty():
            #self.master.after(LOOP_TIME, self._listen_for_resistance)
            #return

        val = self.queue.get()
        if self.recorder.log(val) != self.export_menu.entrycget(0, 'label'):
            self._create_csv_log_menu()

        if val >= 300 and val <= 400:
            self.resistance_label.config(bg='green', text='{}'.format(val))
        else:
            self.resistance_label.config(bg='red', text='{}'.format(val))

        self.check_button.config(state='active')


    def _portlist(self):
        """Scan avaiable ports"""

        ports = []
        _ports = comports()

        if len(_ports) < 1:
            return ports

        for port in _ports:
            #print("{}\n\t{}, {}".format(port, port.device, dir(port)))
            ports.append(port.device)
        return ports


    def _create_csv_log_menu(self):
        """Create log file list"""

        last = self.export_menu.index("end")
        if last:
            for i in reversed(range(last+1)):
                self.export_menu.delete(i)
        months = sorted([ filename[:-4] for filename in glob('*.csv') ])
        for month in reversed(months):
            action_with_arg = partial(self._export_data, month)
            self.export_menu.add_command(label=month, command=action_with_arg)
        print('_create_csv_log_menu: {}'.format(months))


    def _export_data(self, filename):
        """Export month log as xlsx

        Args:
            filename (string): month name
        """

        #mBox.showinfo(TITLE, filename)

        folder = filedialog.askdirectory()
        if not folder:
            return

        wb = Workbook()
        ws = wb.active
        with open('{}.csv'.format(filename), 'r') as f:
            ws.append(('timestamp','value'))
            for row in csv.reader(f):
                ws.append(row)
        path = os.path.join(folder, 'resistance-{}.xlsx'.format(filename))
        wb.save(path)


if __name__ == '__main__':

    win = Tk()
    app = MainApp(win)

    win.title(TITLE)
    win.geometry('%dx%d+%d+%d' % (800, 560, 64, 32))
    win.resizable(0, 0)

    win.mainloop()
