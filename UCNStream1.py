import telnetlib
import threading
import time
import os
import sys
import getpass
from telnetlib import Telnet
import datetime
import subprocess
import re
from tkinter import *
from threading import Thread
import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
import multiprocessing
from multiprocessing import Process


def function_MES(i, wb, user, password):
    uplink_IP = 'F' + str(i)
    uplink_Name = 'E' + str(i)
    port_Name = 'G' + str(i)
    ucn_IP = 'B' + str(i)
    ucn_Name = 'C' + str(i)

    password1 = "admin25"

    uplink_IP_arg = (wb.active[uplink_IP]).value.encode('ascii')
    uplink_Name_arg = (wb.active[uplink_Name]).value.encode('ascii')
    port_Name_arg1 = str((wb.active[port_Name]).value)
    port_Name_arg = port_Name_arg1.encode('ascii')
    ucn_IP_arg = (wb.active[ucn_IP]).value.encode('ascii')
    ucn_Name_arg = (wb.active[ucn_Name]).value.encode('ascii')

    try:
        tn = telnetlib.Telnet(uplink_IP_arg)
        tn.read_until(b"ser")
        tn.write(user.encode('ascii') + b"\n")
        time.sleep(1)
        tn.read_until(b"asswor")
        tn.write(password.encode('ascii') + b"\n")
        time.sleep(1)
        t = tn.read_very_eager().decode('ascii')
        if ">" in t:
            tn.write(b"enable" + b"\n")
            time.sleep(1)
            tn.read_until(b"sswor")
            tn.write(password1.encode('ascii') + b"\n")
            time.sleep(1)
            tn.write(b"show interface description " + port_Name_arg + b"\n")
            time.sleep(2)
            t = tn.read_very_eager().decode('ascii')
            print(t)
            if ucn_IP_arg.decode('utf-8') in t:
                tn.write(b"configure terminal" + b"\n")
                time.sleep(1)
                tn.write(b"interface " + port_Name_arg + b"\n")
                time.sleep(1)
                tn.write(b"shutdown" + b"\n")
                time.sleep(5)
                tn.write(b"no shutdown" + b"\n")
                print(ucn_Name_arg.decode('utf-8') + " " + uplink_Name_arg.decode('utf-8') + " " + port_Name_arg.decode('utf-8') + " SUCCESS")
            else:
                print(uplink_Name_arg.decode('utf-8') + " " + uplink_IP_arg.decode('utf-8') + " " + port_Name_arg.decode('utf-8') + " description was CHANGED " + ucn_Name_arg)
                tn.write(b"exit" + b"\n")
                time.sleep(1)
                tn.write(b"exit" + b"\n")
        else:
            tn.write(b"show interface description " + port_Name_arg + b"\n")
            time.sleep(2)
            t = tn.read_very_eager().decode('ascii')
            if ucn_IP_arg.decode('utf-8') in t:
                tn.write(b"configure terminal" + b"\n")
                time.sleep(1)
                tn.write(b"interface " + port_Name_arg + b"\n")
                time.sleep(1)
                tn.write(b"shutdown" + b"\n")
                time.sleep(5)
                tn.write(b"no shutdown" + b"\n")
                print(ucn_Name_arg.decode('utf-8') + " " + uplink_Name_arg.decode('utf-8') + " " + port_Name_arg.decode('utf-8') + " SUCCESS")
                time.sleep(1)
            else:
                print(uplink_Name_arg.decode('utf-8') + " " + uplink_IP_arg.decode('utf-8') + " " + port_Name_arg.decode('utf-8') + " description was CHANGED " + ucn_Name_arg.decode('utf-8'))
                tn.write(b"exit" + b"\n")
                time.sleep(1)
    except:
        print("ERROR RESET port to " + ucn_Name_arg.decode('utf-8'))

def function_SBA_MES(i, wb, user, password):
    userSBA = "qwerty"
    passSBA = "qwerty"
    password1 = "qwerty"

    uplink_IP = 'F' + str(i)
    uplink_Name = 'E' + str(i)
    port_Name = 'G' + str(i)
    ucn_IP = 'B' + str(i)
    ucn_Name = 'C' + str(i)

    uplink_IP_arg = (wb.active[uplink_IP]).value.encode('ascii')
    uplink_Name_arg = (wb.active[uplink_Name]).value.encode('ascii')
    port_Name_arg = (wb.active[port_Name]).value.encode('ascii')
    ucn_IP_arg = (wb.active[ucn_IP]).value.encode('ascii')
    ucn_Name_arg = (wb.active[ucn_Name]).value.encode('ascii')
    try:
        HOST2 = "212.48.195.228"
        tn = telnetlib.Telnet(HOST2)
        tn.read_until(b"login: ")
        tn.write(userSBA.encode('ascii') + b"\n")
        time.sleep(1)
        tn.read_until(b"Password: ")
        tn.write(passSBA.encode('ascii') + b"\n")
        time.sleep(1)
        tn.write(b"telnet routing-instance ri-is-l3vpn-dslam-ctrl " + uplink_IP_arg + b"\n")
        time.sleep(2)
        tn.read_until(b"ser")
        tn.write(user.encode('ascii') + b"\n")
        time.sleep(1)
        tn.read_until(b"sswor")
        tn.write(password.encode('ascii') + b"\n")
        time.sleep(1)
        k = tn.read_very_eager().decode('ascii')
        if ">" in k:
            tn.write(b"enable" + b"\n")
            time.sleep(1)
            tn.read_until(b"sswor")
            tn.write(password1.encode('ascii') + b"\n")
            time.sleep(1)
            tn.write(b"show interface description " + port_Name_arg + b"\n")
            time.sleep(3)
            t = tn.read_very_eager().decode('ascii')
            if ucn_IP_arg.decode('utf-8') in t:
                tn.write(b"configure terminal" + b"\n")
                time.sleep(1)
                tn.write(b"interface " + port_Name_arg + b"\n")
                time.sleep(2)
                tn.write(b"shutdown" + b"\n")
                time.sleep(7)
                tn.write(b"no shutdown" + b"\n")
                time.sleep(2)
                tn.write(b"exit" + b"\n")
                time.sleep(1)
                tn.write(b"exit" + b"\n")
                time.sleep(1)
                tn.write(b"exit" + b"\n")
                time.sleep(1)
                print(ucn_Name_arg.decode('utf-8') + " " + uplink_Name_arg.decode('utf-8') + " " + port_Name_arg.decode('utf-8') + " SUCCESS")
                tn.write(b"exit" + b"\n")
            else:
                print(uplink_Name_arg.decode('utf-8') + " " + uplink_IP_arg.decode('utf-8') + " " + port_Name_arg.decode('utf-8') + " description was CHANGED " + ucn_Name_arg.decode('utf-8'))
                tn.write(b"exit" + b"\n")
                time.sleep(1)
                tn.write(b"exit" + b"\n")
        else:
            tn.write(b"show interface description " + port_Name_arg + b"\n")
            time.sleep(3)
            t = tn.read_very_eager().decode('ascii')
            if ucn_IP_arg.decode('utf-8') in t:
                tn.write(b"configure terminal" + b"\n")
                time.sleep(1)
                tn.write(b"interface " + port_Name_arg + b"\n")
                time.sleep(2)
                tn.write(b"shutdown" + b"\n")
                time.sleep(7)
                tn.write(b"no shutdown" + b"\n")
                time.sleep(2)
                tn.write(b"exit" + b"\n")
                time.sleep(1)
                tn.write(b"exit" + b"\n")
                time.sleep(1)
                tn.write(b"exit" + b"\n")
                time.sleep(1)
                print(ucn_Name_arg.decode('utf-8') + " " + uplink_Name_arg.decode('utf-8') + " " + port_Name_arg.decode('utf-8') + " SUCCESS")
                tn.write(b"exit" + b"\n")
            else:
                print(uplink_Name_arg.decode('utf-8') + " " + uplink_IP_arg.decode('utf-8') + " " + port_Name_arg.decode('utf-8') + " description was CHANGED " + ucn_Name_arg.decode('utf-8'))
                tn.write(b"exit" + b"\n")
                time.sleep(1)
                tn.write(b"exit" + b"\n")
    except:
        print("ERROR RESET port to " + ucn_Name_arg.decode('utf-8'))



def function_CISCO(i, wb, user, password):
    password1 = "qwerty"

    uplink_IP = 'F' + str(i)
    uplink_Name = 'E' + str(i)
    port_Name = 'G' + str(i)
    ucn_IP = 'B' + str(i)
    ucn_Name = 'C' + str(i)

    uplink_IP_arg = (wb.active[uplink_IP]).value.encode('ascii')
    uplink_Name_arg = (wb.active[uplink_Name]).value.encode('ascii')
    port_Name_arg = (wb.active[port_Name]).value.encode('ascii')
    ucn_IP_arg = (wb.active[ucn_IP]).value.encode('ascii')
    ucn_Name_arg = (wb.active[ucn_Name]).value.encode('ascii')
    try:
        tn = telnetlib.Telnet(uplink_IP_arg)
        tn.read_until(b"sernam")
        tn.write(user.encode('ascii') + b"\n")
        time.sleep(1)
        tn.read_until(b"asswor")
        time.sleep(1)
        tn.write(password.encode('ascii') + b"\n")
        time.sleep(1)
        k = tn.read_very_eager().decode('ascii')
        if ">" in k:
            tn.write(b"enable" + b"\n")
            time.sleep(1)
            tn.read_until(b"sswor")
            tn.write(password1.encode('ascii') + b"\n")
            time.sleep(1)
            tn.write(b"show interface " + port_Name_arg + b" description" + b"\n")
            time.sleep(3)
            t = tn.read_very_eager().decode('utf-8')
            if ucn_IP_arg.decode('utf-8') in t:
                tn.write(b"configure terminal\n" + b"\n")
                time.sleep(1)
                tn.write(b"interface " + port_Name_arg + b"\n")
                time.sleep(2)
                tn.write(b"shutdown" + b"\n")
                time.sleep(7)
                tn.write(b"no shutdown\n" + b"\n")
                time.sleep(2)
                print(ucn_Name_arg.decode('utf-8') + " " + uplink_Name_arg.decode('utf-8') + " " + port_Name_arg.decode('utf-8') + " SUCCESS")
            else:
                print(uplink_Name_arg.decode('utf-8') + " " + uplink_IP_arg.decode('utf-8') + " " + port_Name_arg.decode('utf-8') + " description was CHANGED " + ucn_Name_arg.decode('utf-8'))
                tn.write(b"exit\n" + b"\n")
                time.sleep(1)
                tn.write(b"exit\n" + b"\n")
                time.sleep(1)
        else:
            tn.write(b"show interface " + port_Name_arg + b" description" + b"\n")
            time.sleep(3)
            t = tn.read_very_eager().decode('ascii')
            if ucn_IP_arg.decode('utf-8') in t:
                tn.write(b"configure terminal\n" + b"\n")
                time.sleep(1)
                tn.write(b"interface " + port_Name_arg + b"\n")
                time.sleep(2)
                tn.write(b"shutdown" + b"\n")
                time.sleep(7)
                tn.write(b"no shutdown\n" + b"\n")
                time.sleep(2)
                print(ucn_Name_arg.decode('utf-8') + " " + uplink_Name_arg.decode('utf-8') + " " + port_Name_arg.decode('utf-8') + " SUCCESS")
            else:
                print(uplink_Name_arg.decode('utf-8') + " " + uplink_IP_arg.decode('utf-8') + " " + port_Name_arg.decode('utf-8') + " description was CHANGED " + ucn_Name_arg.decode('utf-8'))
                tn.write(b"exit\n" + b"\n")
                time.sleep(1)
                tn.write(b"exit\n" + b"\n")
                time.sleep(1)
    except:
        print("ERROR RESET port to " + ucn_Name_arg.decode('utf-8'))

def function_RC(i, wb, user, password):
    uplink_IP = 'F' + str(i)
    uplink_Name = 'E' + str(i)
    port_Name = 'J' + str(i)
    ucn_IP = 'B' + str(i)
    ucn_Name = 'C' + str(i)

    uplink_IP_arg = (wb.active[uplink_IP]).value.encode('ascii')
    uplink_Name_arg = (wb.active[uplink_Name]).value.encode('ascii')
    port_Name_arg1 = str((wb.active[port_Name]).value)
    port_Name_arg = port_Name_arg1.encode('ascii')
    ucn_IP_arg = (wb.active[ucn_IP]).value.encode('ascii')
    ucn_Name_arg = (wb.active[ucn_Name]).value.encode('ascii')
    try:
        tn = telnetlib.Telnet(uplink_IP_arg)
        time.sleep(1)
        tn.read_until(b"ogi")
        time.sleep(1)
        tn.write(user.encode('ascii') + b"\n")
        time.sleep(1)
        tn.read_until(b"asswo")
        time.sleep(1)
        tn.write(password.encode('ascii') + b"\n")
        time.sleep(2)
        tn.write(b"show interface port-list " + port_Name_arg + b"\n")
        time.sleep(3)
        t = tn.read_very_eager().decode('ascii')
        if ucn_IP_arg.decode('utf-8') in t:
            tn.write(b"config" + b"\n")
            tn.write(b"interface " + port_Name_arg + b"\n")
            time.sleep(1)
            tn.write(b"shutdown" + b"\n")
            time.sleep(5)
            tn.write(b"no shutdown" + b"\n")
            time.sleep(2)
            tn.write(b"exit" + b"\n")
            time.sleep(1)
            tn.write(b"exit" + b"\n")
            time.sleep(2)
            tn.write(b"exit" + b"\n")
            time.sleep(2)
            print(ucn_Name_arg.decode('utf-8') + " " + uplink_Name_arg.decode('utf-8') + " " + port_Name_arg.decode('utf-8') + " SUCCESS")
        else:
            print(uplink_Name_arg.decode('utf-8') + " " + uplink_IP_arg.decode('utf-8') + " " + port_Name_arg.decode('utf-8') + " description was CHANGED " + ucn_Name_arg.decode('utf-8'))
            tn.write(b"exit" + b"\n")
            time.sleep(2)
    except:
        print("ERROR RESET port to " + ucn_Name_arg.decode('utf-8'))


def function_JUN(i, wb, user, password):
    uplink_IP = 'F' + str(i)
    uplink_Name = 'E' + str(i)
    port_Name = 'G' + str(i)
    ucn_IP = 'B' + str(i)
    ucn_Name = 'C' + str(i)

    uplink_IP_arg = (wb.active[uplink_IP]).value.encode('ascii')
    uplink_Name_arg = (wb.active[uplink_Name]).value.encode('ascii')
    port_Name_arg = (wb.active[port_Name]).value.encode('ascii')
    ucn_IP_arg = (wb.active[ucn_IP]).value.encode('ascii')
    ucn_Name_arg = (wb.active[ucn_Name]).value.encode('ascii')

    tn = telnetlib.Telnet(uplink_IP_arg)
    tn.read_until(b"ogi")
    tn.write(user.encode('ascii') + b"\n")
    time.sleep(1)
    tn.read_until(b"asswor")
    time.sleep(1)
    tn.write(password.encode('ascii') + b"\n")
    time.sleep(4)
    try:
        tn.write(b"show interfaces " + port_Name_arg + b" descriptions" + b"\n")
        time.sleep(3)
        t = tn.read_very_eager().decode('ascii')
        # print(t)
        if ucn_IP_arg.decode('utf-8') in t:
            tn.write(b"configure" + b"\n")
            time.sleep(1)
            tn.write(b"set interfaces " + port_Name_arg + b" disable" + b"\n")
            time.sleep(3)
            tn.write(b"commit" + b"\n")
            time.sleep(5)
            tn.write(b"set interfaces " + port_Name_arg + b" enable" + b"\n")
            time.sleep(3)
            tn.write(b"commit" + b"\n")
            time.sleep(5)
            tn.write(b"exit" + b"\n")
            time.sleep(3)
            print(ucn_Name_arg.decode('utf-8') + " " + uplink_Name_arg.decode('utf-8') + " " + port_Name_arg.decode('utf-8') + " SUCCESS")
        else:
            print(uplink_Name_arg.decode('utf-8') + " " + uplink_IP_arg.decode('utf-8') + " " + port_Name_arg.decode('utf-8') + " description was CHANGED " + ucn_Name_arg.decode('utf-8'))
    except:
        print("ERROR RESET port to " + ucn_Name_arg.decode('utf-8'))




if __name__ == "__main__":
    # i = 2
    user = input("Login: ")
    # password = input("Password: ")
    password = getpass.getpass("Password: ")

    wb = load_workbook(filename='UCN Rotc.xlsx', data_only=True)

    arrayCountUcn = []

    i = 2
    k = 2
    wr_cell = 'A' + str(i)
    rd_cell = (wb.active[wr_cell]).value
    while True:
        wr_cell = 'A' + str(i)
        rd_cell = (wb.active[wr_cell]).value
        if rd_cell is None:
            break
        else:
            i = i + 1
            k = k + 1
    for i in range(2, k):
        wr_cell = 'A' + str(i)
        rd_cell = (wb.active[wr_cell]).value
        uplink_IP = 'F' + str(i)
        uplink_Name = 'E' + str(i)
        port_Name = 'G' + str(i)
        ucn_IP = 'B' + str(i)
        ucn_Name = 'C' + str(i)
        typeEquip = 'I' + str(i)

        typeEquip_arg = (wb.active[typeEquip]).value


        if b"JUN" in typeEquip_arg.encode('ascii'):
            # Process(target=function_JUN, args=(i, wb, user, password)).start()
            # time.sleep(1)

            # p = Process(target=function_JUN, args=(i, wb, user, password))
            # arrayCountUcn.append(p)

            thread = Thread(target=function_JUN, args=(i, wb, user, password))
            thread.start()
            arrayCountUcn.append(thread)
            time.sleep(1)

            if k == len(arrayCountUcn):
                print("Finish!")
                input('Press ENTER to exit')
        elif b"Cisco" in typeEquip_arg.encode('ascii'):
            # Process(target=function_CISCO, args=(i, wb, user, password)).start()
            # time.sleep(1)
            #
            # p = Process(target=function_CISCO, args=(i, wb, user, password))
            # arrayCountUcn.append(p)

            thread = Thread(target=function_CISCO, args=(i, wb, user, password))
            thread.start()
            arrayCountUcn.append(thread)
            time.sleep(1)

            if k == len(arrayCountUcn):
                print("Finish!")
                input('Press ENTER to exit')
        elif b"Mes" in typeEquip_arg.encode('ascii'):
            # function_MES(i, wb, user, password)
            # Process(target=function_MES, args=(i, wb, user, password)).start()
            # time.sleep(1)
            #
            # p = Process(target=function_MES, args=(i, wb, user, password))
            # arrayCountUcn.append(p)

            thread = Thread(target=function_MES, args=(i, wb, user, password))
            thread.start()
            arrayCountUcn.append(thread)
            time.sleep(1)

            if k == len(arrayCountUcn):
                print("Finish!")
                input('Press ENTER to exit')
        elif b"RC" in typeEquip_arg.encode('ascii'):
            # Process(target=function_RC, args=(i, wb, user, password)).start()
            # time.sleep(1)

            # p = Process(target=function_RC, args=(i, wb, user, password))
            # arrayCountUcn.append(p)

            thread = Thread(target=function_RC, args=(i, wb, user, password))
            thread.start()
            arrayCountUcn.append(thread)
            time.sleep(1)

            if k == len(arrayCountUcn):
                print("Finish!")
                input('Press ENTER to exit')
        elif b"SBA" in typeEquip_arg.encode('ascii'):
            # Process(target=function_SBA_MES, args=(i, wb, user, password)).start()
            # time.sleep(1)
            #
            # p = Process(target=function_SBA_MES, args=(i, wb, user, password))
            # arrayCountUcn.append(p)

            thread = Thread(target=function_SBA_MES, args=(i, wb, user, password))
            thread.start()
            arrayCountUcn.append(thread)
            time.sleep(1)

            if k == len(arrayCountUcn):
                print("Finish!")
                input('Press ENTER to exit')
        else:
            print("Неизвестный тип оборудования")
    time.sleep(30)
    input('Press ENTER to exit')
