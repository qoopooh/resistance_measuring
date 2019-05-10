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
from tkinter import Tk, Menu, Entry, LabelFrame, Label, \
        StringVar, messagebox as mBox, filedialog
from tkinter.ttk import Style, Combobox, Button

import serial

from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill
from serial.tools.list_ports import comports

VERSION = '0.4'
TITLE = 'Resistance Measuring V{}'.format(VERSION)

LOOP_TIME = 150 # milliseconds

class Config:
    """Save a configuration"""

    path = os.path.join(os.path.dirname(os.path.realpath(__file__)),
            'config.json')

    comport = None
    test_wo_sensor = False
    lot_no = ''
    cable_no = 0

    def __init__(self):

        if not os.path.exists(self.path):
            return

        with open(self.path) as readfile:
            obj = json.load(readfile)

            if 'comport' in obj:
                self.comport = obj['comport']
            if 'test_wo_sensor' in obj:
                self.test_wo_sensor = obj['test_wo_sensor']
            if 'lot_no' in obj:
                self.lot_no = obj['lot_no']
            if 'cable_no' in obj:
                self.cable_no = obj['cable_no']

    def save(self):
        """Save to json file"""

        with open(self.path, 'w') as outfile:
            out = {
                'comport': self.comport,
                'test_wo_sensor': self.test_wo_sensor,
                'lot_no': self.lot_no,
                'cable_no': self.cable_no,
            }
            json.dump(out, outfile, indent=2)


class Recorder:
    """Record resistance value into csv log file
    """

    def __init__(self):
        month = datetime.now().strftime("%Y-%m")
        self.path = os.path.join(os.path.dirname(os.path.realpath(__file__)),
                '{}.csv'.format(month))


    def record(self, val, result='', lot='', cable_no=None):
        """Record value

        Args:
            val (float): resistance value
            result (string): resistance value
            lot (string): lot number
            cable_no (number): cable number

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
            if cable_no:
                msg = '{},{},{},{},{}\n'.format(time, lot, cable_no, val, result)
            else:
                msg = '{},{},,{},{}\n'.format(time, lot, val, result)
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

        self.resistance_label = Label(win, text='XXX.XX')
        self.resistance_label.config(background='green', foreground='black')
        self.resistance_label.grid(
                column=0, row=0,
                sticky='nsew',
                padx=12, pady=10)
        win.rowconfigure(0, weight=1)
        win.columnconfigure(0, weight=1)

        s = Style()
        s.configure('my.TButton', font=(18))
        self.check_button = Button(win, text='\nCheck\n',
                command=self.check,
                style='my.TButton',
                )
        self.check_button.config(width=20)
        self.check_button.grid(column=0, row=1)
        self.check_button.focus()

        setting_frame = LabelFrame(win, text='Setting')
        setting_frame.grid(row=2, column=0, pady=10)

        self.selected_port_var = StringVar()
        self.port_combobox = Combobox(setting_frame, state='readonly',
                textvariable=self.selected_port_var)
        self.port_combobox.grid(row=0, column=0, padx=8)
        self._update_portlist()

        refresh_button = Button(setting_frame,
                text='Refresh',
                command=self._update_portlist,
                )
        refresh_button.grid(row=1, column=0)

        lot_label = Label(setting_frame, text='Lot no.')
        lot_label.grid(row=0, column=1)
        self.lot_var = StringVar()
        lot_entry = Entry(setting_frame, width=20, textvariable=self.lot_var)
        lot_entry.grid(row=0, column=2, padx=8, pady=4)

        cable_label = Label(setting_frame, text='Cable no.')
        cable_label.grid(row=1, column=1)
        self.cable_var = StringVar()
        cable_entry = Entry(setting_frame, width=20, textvariable=self.cable_var, justify='center', state='readonly')
        cable_entry.grid(row=1, column=2, padx=8, pady=4)

        #
        # init cable info
        #
        self.lot_var.set(self.cfg.lot_no)
        self.cable_var.set(self.cfg.cable_no)

        win.bind("<Configure>", self.on_resize)


    def on_resize(self, event):
        #print('on_resize: {} {}'.format(event.width, event.height))
        self.resistance_label.config(
                font=('times', int(self.master.winfo_width() / 5), 'bold'))


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
            val = randint(28000, 42000) / 100
        else:
            try:
                port = self.selected_port_var.get()
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
                    val = -2

                if port != self.cfg.comport:
                    self.cfg.comport = port
                    self.cfg.save()

            except Exception as e:
                print(e)
                self.serial = None
                val = -1
                #self._update_portlist()
        self.queue.put(val)


    def _listen_for_resistance(self):
        """Get returned value from sensor reading"""

        if self._thread.is_alive():
            self.master.after(LOOP_TIME, self._listen_for_resistance)
            return

        val = self.queue.get()
        lot_no, cable_no = self._get_cable_info()

        result = 'Pass'
        if val >= 300 and val <= 400:
            self.resistance_label.config(bg='green', text='{}'.format(val))
        else:
            result = 'Fail'
            self.resistance_label.config(bg='red', text='{}'.format(val))

        if val > 0:
            month = self.recorder.record(val, result, lot_no, cable_no)
            if month != self.last_month:
                self._create_csv_log_menu()
            self.cfg.lot_no = lot_no
            self.cfg.cable_no = cable_no
            self.cfg.save()
            self.cable_var.set(str(cable_no))

        self.check_button.config(state='active')


    def _get_cable_info(self):

        lot_no = self.lot_var.get().strip()
        cable_no = 1
        if lot_no == self.cfg.lot_no:
            try:
                cable_no = int(self.cable_var.get()) + 1
            except Exception as e:
                print(e)

        return lot_no, cable_no


    def _update_portlist(self):
        """Re-assign port list"""

        ports = self._portlist()
        self.port_combobox['values'] = ports
        if self.cfg.comport and self.cfg.comport in ports:
            self.selected_port_var.set(self.cfg.comport)


    def _portlist(self):
        """Scan avaiable ports"""

        ports = []

        ##
        ## dummy ports
        ##
        #val = randint(1, 10)
        #for i in range(val):
            #ports.append('com{}'.format(i+1))
        #return ports

        _ports = comports()

        if len(_ports) < 1:
            return ports

        #for port in _ports:
            #ports.append(port.device)
        #return ports
        return [ port.device for port in _ports ]


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
        if len(months) > 0:
            self.last_month = months[-1]
        else:
            self.last_month = ''


    def _export_data(self, filename):
        """Export month log as xlsx

        Args:
            filename (string): month name
        """

        folder = filedialog.askdirectory()
        if not folder:
            return

        wb = Workbook()
        ws = wb.active
        headers = ('Time', 'Lot No.', 'Cable No.', 'Value', 'Result')
        no_cols = []
        for name in headers:
            if name in ('Cable No.', 'Value'):
                no_cols.append(headers.index(name))
        with open('{}.csv'.format(filename), 'r') as f:
            ws.append(headers)
            for row in csv.reader(f):
                for i in no_cols:
                    try:
                        row[i] = int(row[i])
                    except Exception as e:
                        row[i] = float(row[i])

                ws.append(row)

        fill_format = PatternFill(start_color="e0efd4", end_color="e0efd4", fill_type = "solid")
        self._adjust_column_width(ws)
        for i in range(len(headers)):
            ws['{}1'.format(chr(ord('A')+i))].fill = fill_format

        #
        # Freeze first row / col
        #
        c = ws['B2']
        ws.freeze_panes = c

        path = os.path.join(folder, 'resistance-{}.xlsx'.format(filename))
        wb.save(path)


    def _adjust_column_width(self, worksheet):
        """Manage columns of exported file

        Args:
            worksheet (object): current worksheet
        """

        adjusted_width = 0
        for col in worksheet.columns:
            max_length = 0
            col_idx = col[0].column # Get the column index
            for cell in col:
                try: # Necessary to avoid error on empty cells
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
                cell.alignment = Alignment(horizontal='center')
            adjusted_width = (max_length + 2) * 1.2
            column_name = chr(ord('A') + col_idx - 1)
            worksheet.column_dimensions[column_name].width = adjusted_width


if __name__ == '__main__':

    win = Tk()
    app = MainApp(win)

    win.title(TITLE)
    win.geometry('%dx%d+%d+%d' % (800, 550, 64, 32))
    # win.resizable(0, 0)

    win.mainloop()
