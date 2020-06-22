from input import inputA
from output import WriteArrangement
from distribution import distribution
from distribution import worker
import topwindow


import os
import tkinter as tk
from tkinter.filedialog import askopenfilenames,askopenfilename
from tkinter import messagebox
import glob
from enum import Enum
import win32com.client
from tkinter import *
import pandas as pd
from tabulate import tabulate
from pandas import ExcelWriter
from datetime import datetime
from calendar import Calendar
import xlsxwriter
import nexmo
from shutil import copyfile



class Types(Enum):
    ERR =1
    INFO = 2
    YESORNOT = 3
    QUESTION = 4
    OKCANCEL = 5


class local:

    def __init__(self, list_worker, regions, days):
        self.list_worker = list_worker
        self.regions = regions
        self.days = days

global res
global presed
global filenames
global used
global lockfile
global lockKB
global w
global top

res = None
presed = None
filenames = []
used = False
lockKB = False
lockfile = False
w=[]
top = None




class press:
    # for item in list_worker:
    #    print(item.name, item.regions, item.day)
    def print_msg(title, str, type,bg="blue"):
        global top
        if type is Types.INFO:
            return messagebox.showinfo(title, str)
        if type is Types.QUESTION:
            return messagebox.askquestion(title, str)
        if type is Types.OKCANCEL:
            return messagebox.askokcancel(title, str)
        if type is Types.ERR:
            return messagebox.showwarning(title, str)
        if type is Types.YESORNOT:
            return messagebox.askyesno(title, str)
        if type is None:
            global top
            if top is not None and top.winfo_exists() is not 0:
                return
            top = tk.Toplevel(bg = bg)
            top.title(title)
            icon = PhotoImage(file="logo/logo2.png")
            top.iconphoto(top._w, icon)
            l = tk.Label(top, text=str)
            l.pack(ipadx=50, ipady=10, fill='both', expand=True)
            b = tk.Button(top, text="OK", command=press.AfterMsg)
            b.pack(pady=10, padx=10, ipadx=20, side='right')

    def AfterMsg():
        global presed
        global top
        global used
        if not presed or presed.top.winfo_exists() is 0:
            top.destroy()
        else:
            used = False
            top.destroy()
            presed.TopME()


    def About():
        press.print_msg("About", "Employee Management Software\n\n Programmed by Etiel\n\n Version 0.1", None, "blue")



    def do_it_auto(): #print massage, to do function that print
        res = inputA(glob.glob('files/*worker.xlsx'))
        res.ReadData()
        res.TakeRegionOfWorker()
        d = distribution(res.list_worker, res.regions, res.days).findmatch()
        if d == None:
            res = inputA(glob.glob('files/*worker.xlsx'),False)
            res.ReadData()
            res.TakeRegionOfWorker()
            d = distribution(res.list_worker, res.regions, res.days).findmatch()
        if d == None:
            press.print_msg("Ops", "There is no legal arrangement, you need more worker", Types.ERR)
            return
        WriteArrangement(d).write()
        press.print_msg("Well done!", "Your workers are ready to work:)", None)


    def PressLoad():
        global presed
        if not presed or presed.top.winfo_exists() is 0:
            button_for_RB1 = ["print result to file", "print result to screen"]
            button_for_frame1 = [("Load from File", press.load_files), ("Load from Keyboard", press.LoadFromKeyBoard)]
            last_buttons = [("START", press.PresStart)]
            presed = topwindow.LoadData(button_for_frame1, button_for_RB1, None,last_buttons)
            presed.WindowDesign()
        if presed and presed.top.winfo_exists() is not 0:
            presed.TopME()


    def load_files():  #if load twice need to be deleted, i can chouse initial path (the last one), title, type of file (with describe)
        global filenames
        global presed
        global res
        global used
        global lockfile
        global lockKB
        filenames = askopenfilenames() # To do - need to check if files are match
        if filenames != [] and filenames != "" and not used and not lockfile:
            #if len(filenames) < 6:
             #   press.print_msg("ERROR", "You don't have enough employees", Types.ERR)
              #  return
            lockKB = True
            res = inputA(filenames, True)
            res.ReadData()
            if res.TakeRegionOfWorker() is -1:
                press.print_msg("oh", "One of the employees has no domain. \n"
                                      "The names in the files may not match the names in the system", Types.ERR)
                return
            used = True
            press.print_msg("Well done!", "Files loaded successfully!   \n\n Note: You can load preferences for your"
                                          " workers", None) #To do adding checkbox with dont show me agen
            lockKB = False
        presed.TopME()


    def print_to_file():
        global res
        global filenames
        if not res:
            return
        d = distribution(res.list_worker, res.regions, res.days).findmatch()
        if d == None and lockKB: # if there is no answer with pref
            res = inputA(filenames, pref=False)
            res.ReadData()
            if res.TakeRegionOfWorker() is -1:
                press.print_msg("oh", "One of the employees has no domain. \n"
                                      "The names in the files may not match the names in the system", Types.ERR)
                return
            d = distribution(res.list_worker, res.regions, res.days).findmatch()
        if d == None:
            press.print_msg("Ops", "There is no legal arrangement, you need more worker", Types.ERR)
            return -1
        WriteArrangement(d).write()

    def PresStart():
        global presed
        global filenames
        global res
        if filenames == [] and not res:
            press.print_msg(title="ERORR!",str= "please load files before pressing start",type = Types.ERR)
            presed.TopME()
            return
        if press.print_to_file() == -1:
            return
        if presed.valofrb.get() == 1:
            press.print_msg("Well done!", "Your workers are ready to work:)", None) #To do adding checkbox with dont show me agen
            presed.Destroy()
        if presed.valofrb.get() == 2:
            press.Print_to_scream()


    #works only if all files in mail inbox and there is no other worker files nounths in inbox
    def download_worker_files():
        path = os.path.realpath(r'files')
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items
        i = 0
        for message in messages:
            # To iterate through email items using message.Attachments object.
            for attachment in message.Attachments:
                if ("קובץ זמינות" in str(attachment)):
                    i += 1
                    # To save the perticular attachment at the desired location in your hard disk.
                    attachment.SaveAsFile(os.path.join(path, str(i) + "worker.xlsx"))
                break
        press.print_msg("Note", str(i)+" employee files downloaded", None)

    def Delete_worker_files():
        for file in glob.glob('files/*worker.xlsx'):
            os.remove(file)
        press.print_msg("Great!", "File Removed!", None)

    def _PrintFileToScream(path):
        rows = []
        hedear = []
        path = os.path.realpath(path)
        data = pd.read_excel(path, header=0)
        for i in data.head(0):
            hedear += [i]
        for row in data.values:
            rows += [row]

        str1 = tabulate(rows, headers=hedear, tablefmt="grid", numalign="decimal",
                        stralign="left", showindex="default", disable_numparse=False,
                        colalign=None, )
        press.print_msg("Result", str1, None, bg = "blue")

    def Print_to_scream():
        press._PrintFileToScream(r'res/res.xlsx')

    def _GetSizeOfOverlapArea():
        data = pd.read_excel("system/regions.xlsx", header=0)
        regions = press._GetRegion()
        temp = press._ConvertToList(data.values)
        temp = press._TransposeList(temp)
        msg = ""
        for i in range(len(regions)):
            size1 = 0
            size2 = 0
            for char in temp[i]:
                if char == "X" or char == "-":
                    size1+=1
                if char == "-":
                    size2+=1
            msg += regions[i]+" has "+str(size1)+" workers with "+str(size2)+" Preferences\n"
        msg += "\n***Note that at least 3 workers recommended for region***\n"
        return msg


    def Allert():
        size = len(press._GetWorker())
        msg = ""
        if size < 9 and size > 7:
            msg = "You have "+ str(size)+" workers.\n at least 9 workers are recomended\n"
        if size <= 7:
            msg = "Be careful you are missing employees. You have only "+str(size)+" workers\n"
        msg += "\n"+press._GetSizeOfOverlapArea()
        press.print_msg("Allert",msg,None)

    # return the header of "regions.xlsx"
    def _GetHedear():
        temp = []
        data = pd.read_excel("system/regions.xlsx", header=0)
        for i in data.head(0):
            temp += [i]
        hedear = [temp[i] for i in range(len(temp))]
        return hedear

    # return the all existing regions
    def _GetRegion():
        temp = []
        data = pd.read_excel("system/regions.xlsx", header=0)
        for i in data.head(0):
            temp += [i]
        regions = [temp[i] for i in range(len(temp)-3)]
        return regions

    # return the all existing regions
    def _GetPhoneNumber():
        nums = []
        data = pd.read_excel("system/regions.xlsx", header=0)
        for row in data.values:
            nums += [row[len(row) - 3]]
            continue
        return nums

    def _GetWorker():
        worker = []
        data = pd.read_excel("system/regions.xlsx", header=0)
        for row in data.values:
            worker += [row[len(row)-1]]
            continue
        return worker

    def _GetMailOfWorker():
        mail = []
        data = pd.read_excel("system/regions.xlsx", header=0)
        for row in data.values:
            mail += [row[len(row) - 2]]
            continue
        return mail

    # getting an iterable object in iterable object and convert it into a list
    def _ConvertToList(data):
        res_convert = []
        for i in data:
            intern_convert = list(i)
            res_convert += [intern_convert]
        return res_convert

    # get list of a list and transpose column and rows
    def _TransposeList(data):
        res_list = []
        intern_list = []
        size = len(data[0])
        for i in range(size):
            for line in data:
                intern_list += [line.pop()]
            res_list += [intern_list]
            intern_list = []
        res_list.reverse()
        return res_list

    # add list to the last row of existing file
    def _AddListToExistFile(name, mail, phone, region):
        data = pd.read_excel("system/regions.xlsx", header=0)
        hedear = press._GetHedear()
        regions = press._GetRegion()
        new_row = []
        for i in range(len(regions)):
            if regions[i] == region:
                new_row += ["X"]
                continue
            new_row += [float("nan")]
        new_row += [phone, mail, name]
        temp_list = press._ConvertToList(data.values) + [new_row]
        res_list = press._TransposeList(temp_list)
        # wrote it to Excel
        df = pd.DataFrame({hedear[i]: res_list[i] for i in range(len(hedear))})
        writer = ExcelWriter("system/regions.xlsx")
        df.to_excel(writer, index=False, inf_rep=8)
        writer.save()


    def Add_worker():

        regions = press._GetRegion()
        # this function handle when press Add button
        def Add_Worker_helper():
            if variable.get() == 'Select an area':
                press.print_msg("ERROR", "No region was selected", Types.ERR)
                top.lift()
                return
            if name.get() == "":
                press.print_msg("ERROR", "No name was writen", Types.ERR)
                top.lift()
                return
            if mail.get() == "":
                press.print_msg("ERROR", "No mail was writen", Types.ERR)
                top.lift()
                return
            if Phone.get() == "":
                press.print_msg("ERROR", "No Phone was writen", Types.ERR)
                top.lift()
                return
            press._AddListToExistFile(name.get(),mail.get(), Phone.get(), variable.get())
            press.print_msg("Yes", name.get()+" added successfully with region "+variable.get(), None)
            top.destroy()

        global top
        if top is not None and top.winfo_exists() is not 0:
            return
        top = tk.Toplevel()
        top.geometry("+30+30")
        top.title("Add worker")
        icon = PhotoImage(file="logo/logo2.png")
        top.iconphoto(top._w, icon)
        #add name
        name = Entry(top, width=30)
        name.grid(row=0, column=1, padx=20)
        name_label = Label(top, text="Name")
        name_label.grid(row=0, column=0, padx=20)
        # add mail
        mail = Entry(top, width=30)
        mail.grid(row=1, column=1, padx=20)
        mail_label = Label(top, text="Mail")
        mail_label.grid(row=1, column=0, padx=20)
        # add Phone
        Phone = Entry(top, width=30)
        Phone.grid(row=2, column=1, padx=20)
        Phone_label = Label(top, text="Phone number +972")
        Phone_label.grid(row=2, column=0, padx=20)
        # add is regions
        variable = StringVar(top)
        variable.set('Select an area')

        w = OptionMenu(top, variable, *regions)
        w.grid(padx=10, pady=10, row=3, column=0, columnspan=2, ipadx=60)
        # button Add
        sub_b = Button(top, text="Add", command=Add_Worker_helper)
        sub_b.grid(row=4, column=0, columnspan=2, pady=10, padx=10, ipadx=60)

    def _RemoveRowFromExcel(name):
        data = pd.read_excel("system/regions.xlsx", header=0)
        hedear = press._GetHedear()
        temp_list = press._ConvertToList(data.values)
        for item in temp_list:
            if item[len(item)-1] == name:
                temp_list.remove(item)
                break
        res_list = press._TransposeList(temp_list)
        # wrote it to Excel
        df = pd.DataFrame({hedear[i]: res_list[i] for i in range(len(hedear))})
        writer = ExcelWriter("system/regions.xlsx")
        df.to_excel(writer, index=False, inf_rep=8)
        writer.save()


    def DeleteWorker():
        global top
        if top is not None and top.winfo_exists() is not 0:
            return
        worker = press._GetWorker()
        if len(worker) == 0:
            press.print_msg("ERROR","There is no worker",Types.ERR)
        top = tk.Toplevel()
        top.geometry("+30+30")
        top.title("Delete worker")
        icon = PhotoImage(file="logo/logo2.png")
        top.iconphoto(top._w, icon)
        # add name of worker
        variable = StringVar(top)
        variable.set('Select a worker')

        def Delete_Worker_helper():
            if variable.get() == 'Select a worker':
                press.print_msg("ERROR", "No name was selected", Types.ERR)
                top.lift()
                return
            press._RemoveRowFromExcel(variable.get())
            press.print_msg("Yes", variable.get()+ " removed successfully", None)
            top.destroy()

        w = OptionMenu(top, variable, *worker)
        w.grid(padx=10, pady=10, row=0, column=0, columnspan=2, ipadx=60)
        # button Add
        sub_b = Button(top, text="Delete", command=Delete_Worker_helper)
        sub_b.grid(row=1, column=0, columnspan=2, pady=10, padx=10, ipadx=60)

    def _AddBlockToExcel(name, region, char):
        data = pd.read_excel("system/regions.xlsx", header=0)
        header = press._GetHedear()
        regions = press._GetRegion()
        index = 0
        for i in range(len(regions)):
            if regions[i] == region:
                index = i
                break
        temp_list = press._ConvertToList(data.values)
        for item in temp_list:
            if item[len(item) - 1] == name:
                item[index] = char
                break
        res_list = press._TransposeList(temp_list)
        df = pd.DataFrame({header[i]: res_list[i] for i in range(len(header))})
        writer = ExcelWriter("system/regions.xlsx")
        df.to_excel(writer, index=False, inf_rep=8)
        writer.save()

    def AddRegionToWorker():
        global top
        if top is not None and top.winfo_exists() is not 0:
            return
        worker = press._GetWorker()
        if len(worker)==0:
            press.print_msg("ERROR", "There is no worker", Types.ERR)
            return
        regions = press._GetRegion()
        top = tk.Toplevel()
        top.geometry("+30+30")
        top.title("Add region to workerr")
        icon = PhotoImage(file="logo/logo2.png")
        top.iconphoto(top._w, icon)
        # add name of worker
        variable1 = StringVar(top)
        variable1.set('Select a worker')
        variable2 = StringVar(top)
        variable2.set('Select a region')

        def Add_region_to_worker_helper():
            if variable1.get() == 'Select a worker':
                press.print_msg("ERROR", 'Select a worker',Types.ERR)
                top.lift()
                return
            if variable2.get() == 'Select a region':
                press.print_msg("ERROR", 'Select a region',Types.ERR)
                top.lift()
                return
            press._AddBlockToExcel(variable1.get(), variable2.get(),"X")
            press.print_msg("Yes!", "the region "+ variable2.get()+" added to "+variable1.get()+" successfully", None)
            top.destroy()

        w1 = OptionMenu(top, variable1, *worker)
        w1.grid(padx=10, pady=10, row=0, column=0, columnspan=2, ipadx=60)
        w2 = OptionMenu(top, variable2, *regions)
        w2.grid(padx=10, pady=10, row=1, column=0, columnspan=2, ipadx=60)
        # button Add
        sub_b = Button(top, text="Add", command=Add_region_to_worker_helper)
        sub_b.grid(row=2, column=0, columnspan=2, pady=10, padx=10, ipadx=60)

    # this function get list of lists and adding nan to the beginning of every list
    def _AddNanToList(list):
        new_list = []
        intern_list = []
        for item in list:
            intern_list = [float("nan")] + item
            new_list += [intern_list]
        return new_list

    def _Add_column_to_Excel(name):
        data = pd.read_excel("system/regions.xlsx", header=0)
        header = press._GetHedear()
        header = [name] + header
        res_list = press._ConvertToList(data.values)
        res_list = press._AddNanToList(res_list)
        res_list = press._TransposeList(res_list)
        df = pd.DataFrame({header[i]: res_list[i] for i in range(len(header))})
        writer = ExcelWriter("system/regions.xlsx")
        df.to_excel(writer, index=False, inf_rep=8)
        writer.save()

    def AddRegion():
        global top
        if top is not None and top.winfo_exists() is not 0:
            return
        top = tk.Toplevel()
        top.geometry("+30+30")
        top.title("Add region")
        icon = PhotoImage(file="logo/logo2.png")
        top.iconphoto(top._w, icon)

        def Add_region_helper():
            if name.get() == "":
                press.print_msg("Error", "Enter an area name", Types.ERR)
                top.lift()
                return
            regions = press._GetRegion()
            if regions.__contains__(name.get()):
                press.print_msg("Error", "Name already exists", Types.ERR)
                top.lift()
                return
            press._Add_column_to_Excel(name.get())
            press.print_msg("Yes", name.get()+" added successfully", None)
            top.destroy()

        # add name
        name = Entry(top, width=30)
        name.grid(row=0, column=1, padx=20)
        name_label = Label(top, text="name of region")
        name_label.grid(row=0, column=0, padx=20)
        # button Add
        sub_b = Button(top, text="Add", command=Add_region_helper)
        sub_b.grid(row=2, column=0, columnspan=2, pady=10, padx=10, ipadx=60)

    # get a list of lists and delete the item in index from all of them
    def _remove_index_from_list(list, index):
        new_list =[]
        intern_list = []
        size = len(list[0])
        for item in list:
            for j in range(size):
                if j == index:
                    continue
                intern_list += [item[j]]
            new_list += [intern_list]
            intern_list = []
        return new_list

    def _Delete_column_to_Excel(name):
        data = pd.read_excel("system/regions.xlsx", header=0)
        header = press._GetHedear()
        index = 0
        regions = press._GetRegion()
        for i in range(len(regions)):
            if regions[i] == name:
                index = i
                break
        header.remove(name)
        res_list = press._ConvertToList(data.values)
        res_list = press._remove_index_from_list(res_list,index)
        res_list = press._TransposeList(res_list)
        df = pd.DataFrame({header[i]: res_list[i] for i in range(len(header))})
        writer = ExcelWriter("system/regions.xlsx")
        df.to_excel(writer, index=False, inf_rep=8)
        writer.save()

    def DeleteRegion():
        global top
        if top is not None and top.winfo_exists() is not 0:
            return
        top = tk.Toplevel()
        top.geometry("+30+30")
        top.title("Delete region")
        icon = PhotoImage(file="logo/logo2.png")
        top.iconphoto(top._w, icon)
        regions = press._GetRegion()
        variable = StringVar(top)
        variable.set('Select a region')

        def Delete_region_helper():
            if variable.get() == 'Select a region':
                press.print_msg("Error", 'Select a region',Types.ERR)
                return
            press._Delete_column_to_Excel(variable.get())
            press.print_msg("Yes", variable.get() + " deleted successfully", None)
            top.destroy()

        w = OptionMenu(top, variable, *regions)
        w.grid(padx=10, pady=10, row=0, column=0, columnspan=2, ipadx=60)
        # button Add
        sub_b = Button(top, text="Delete", command=Delete_region_helper)
        sub_b.grid(row=1, column=0, columnspan=2, pady=10, padx=10, ipadx=60)


    def AddPreferencesToWorker():
        worker = press._GetWorker()
        regions = press._GetRegion()
        global top
        if top is not None and top.winfo_exists() is not 0:
            return
        top = tk.Toplevel()
        top.geometry("+30+30")
        top.title("Add preferences to worker")
        icon = PhotoImage(file="logo/logo2.png")
        top.iconphoto(top._w, icon)
        # add name of worker
        variable1 = StringVar(top)
        variable1.set('Select a worker')
        variable2 = StringVar(top)
        variable2.set('Select a region')

        def Add_pref_to_worker_helper():
            if variable1.get() == 'Select a worker':
                press.print_msg("ERROR", 'Select a worker', Types.ERR)
                top.lift()
                return
            if variable2.get() == 'Select a region':
                press.print_msg("ERROR", 'Select a region', Types.ERR)
                top.lift()
                return
            press._AddBlockToExcel(variable1.get(), variable2.get(),"-")
            press.print_msg("Yes!", "Preference for not working in area " + variable2.get() +
                            " added to " + variable1.get() + " successfully",
                            None)
            top.destroy()

        pref_lable = Label(top, text="Choose an area where your employee prefers not to work",
                           foreground="blue", font ="david")
        pref_lable.grid(row=0, column=0, padx=20, pady =5)
        w1 = OptionMenu(top, variable1, *worker)
        w1.grid(padx=10, pady=10, row=1, column=0, columnspan=2, ipadx=60)
        w2 = OptionMenu(top, variable2, *regions)
        w2.grid(padx=10, pady=10, row=2, column=0, columnspan=2, ipadx=60)
        # button Add
        sub_b = Button(top, text="Add", command=Add_pref_to_worker_helper)
        sub_b.grid(row=3, column=0, columnspan=2, pady=10, padx=10, ipadx=60)

    def _GetDateOfMonth(day, month):
        n = datetime.now()
        cal = Calendar()  # week starts Monday
        days = []
        weeks = cal.monthdayscalendar(n.year, month)
        for w in weeks:
            if w[(day-2)%7] == 0: # day 4 in week
                continue
            date = str(w[(day-2)%7]) + "/" + str(month) + "/" + str(n.year)
            days += [date]
        return days

    def _GetRegionsOfWorker(worker, pref = False):
        data = pd.read_excel("system/regions.xlsx", header=0)
        # Extracting the regions of worker
        for item in data.values:
            if item[len(item)-1] != worker:
                continue
            is_region = []
            for j in range(len(item)):
                # if there is solution with prefenses add the X value
                if pref and item[j] is "X":  # 7 is max day. can be upgrade
                    is_region += [j]
                # if there is not solution with prefenses add also - value
                if not pref and (item[j] is "X" or item[j] is "-"):
                    is_region += [j]
            return is_region


    def _AddDaysToWorker(worker_name, days, pref=None):
        is_region = press._GetRegionsOfWorker(worker_name,pref)
        return worker(is_region,days,worker_name)

    def Next(day, month, pref):

        days = press._GetDateOfMonth(int(day), int(month))
        var_for_checkbox = [IntVar(0) for i in range(len(days))]

        def Add_Helper(worker1):
            global w
            global top
            is_days = []
            for i in range(len(var_for_checkbox)):
                if var_for_checkbox[i].get() == 1:
                    var_for_checkbox[i].set(0)
                    is_days += [i+1]
            w += [press._AddDaysToWorker(worker1, is_days, pref)]
            top.destroy()
            Next_helper(workers, day,month,pref)

        def Next_helper(workers, day, month, pref):
            global w
            global top
            global res
            if len(workers) == 0:
                res = local(w,press._GetRegion(),days)
                top.destroy()
                return
            worker = workers.pop()
            top = tk.Toplevel()
            top.geometry("+30+30")
            top.title("chose day to "+worker)
            icon = PhotoImage(file="logo/logo2.png")
            top.iconphoto(top._w, icon)
            lable = Label(top, text="Choose days for " +worker,
                          foreground="blue", font="david")
            lable.grid(row=0, column=0, padx=20, pady=5)
            # add name of worker
            checkboxes = []
            for i in range(len(days)):
                checkboxes += [Checkbutton(top, text=days[i], variable=var_for_checkbox[i])]
                checkboxes[i].grid(row=i+1, sticky=W)
            # button Add
            sub_b = Button(top, text="Add", command=lambda: Add_Helper(worker))
            sub_b.grid(row=len(days)+1, column=0, columnspan=2, pady=10, padx=10, ipadx=60)

        workers = press._GetWorker()
        Next_helper(workers,day,month,pref)


    def LoadFromKeyBoard():
        days = [i+1 for i in range(7)]
        month = [i+1 for i in range(12)]
        top = tk.Toplevel()
        top.geometry("+30+30")
        top.title("chose day and month")
        icon = PhotoImage(file="logo/logo2.png")
        top.iconphoto(top._w, icon)
        # add name of worker
        variable1 = StringVar(top)
        variable1.set('Select a day at week')
        variable2 = StringVar(top)
        variable2.set('Select a month')
        var = IntVar(0)

        def load_from_keyboard_helper():
            if variable1.get() == 'Select a day at week':
                press.print_msg("Error", 'Select a day at week', Types.ERR)
                top.lift()
                return
            if variable2.get() == 'Select a month':
                press.print_msg("Error", 'Select a month', Types.ERR)
                top.lift()
                return
            press.Next(variable1.get(), variable2.get(), var.get())
            top.destroy()

        lable = Label(top, text="Choose a day of the week and month of work",
                           foreground="blue", font="david")
        lable.grid(row=0, column=0, padx=20, pady=5)
        w1 = OptionMenu(top, variable1, *days)
        w1.grid(padx=10, pady=2, row=1, column=0, columnspan=2, ipadx=60)
        w2 = OptionMenu(top, variable2, *month)
        w2.grid(padx=10, pady=2, row=2, column=0, columnspan=2, ipadx=60)
        # button Add
        sub_b = Button(top, text="Next", command=load_from_keyboard_helper)
        sub_b.grid(row=3, column=0, columnspan=2, pady=2, padx=10, ipadx=60)
        checkboxes = Checkbutton(top, text="Preferences", variable=var)
        checkboxes.grid(row=4, sticky=W)

    # sending mail to a list of mails
    def _SendMailTo(list_of_mails, Subject, Body, list_path_to_file):
        for m in list_of_mails:
            outlook = win32com.client.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = m
            mail.Subject = Subject
            mail.Body = Body
            # To attach a file to the email (optional):
            for path in list_path_to_file:
                mail.Attachments.Add(path)
            mail.Send()


    def _GetHebrewNameOfMonth(indek):
        list_of_month = ["","ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני", "יולי", "אוגוסט", "ספטמבר", "אוקטובר",
                         "נובמבר", "דצמבר"]
        return list_of_month[int(indek)]

    def _CreateFileAvailability(day, month):
        n = datetime.now()
        sub = "קובץ זמינות חודש " + press._GetHebrewNameOfMonth(month) + " " + str(n.year)
        dates = press._GetDateOfMonth(int(day),int(month))
        header = ["תאריך", "יום", "שם"]
        day_heb = ["ד" for i in range(len(dates))]
        # note = "לכתוב את השם שלכם בתאריכים שאתם יכולים.\n תזכרו שזה לאו דווקא אותו היום, גם היום שאחריו.\n יש למלא לפחות 3 תאריכים."
        workbook = xlsxwriter.Workbook('system/' + sub + ".xlsx")
        worksheet = workbook.add_worksheet(sub)
        worksheet.write_row(0, 0, header)
        worksheet.write_column(1, 0, dates)
        worksheet.write_column(1, 1, day_heb)
        worksheet.set_default_row(hide_unused_rows=True)
        worksheet.set_column('D:XFD', None, None, {'hidden': True})
        workbook.close()


    # this function is sending file availability to all workers
    def sendFileAvailability():
        days = [i + 1 for i in range(7)]
        month = [i + 1 for i in range(12)]
        global top
        if top is not None and top.winfo_exists() is not 0:
            return
        top = tk.Toplevel()
        top.geometry("+30+30")
        top.title("chose day and month")
        icon = PhotoImage(file="logo/logo2.png")
        top.iconphoto(top._w, icon)
        # add name of worker
        variable1 = StringVar(top)
        variable1.set('Select a day at week')
        variable2 = StringVar(top)
        variable2.set('Select a month')

        def Send_file_availability_helper():
            if variable1.get() == 'Select a day at week':
                press.print_msg("Error", 'Select a day at week', Types.ERR)
                top.lift()
                return
            if variable2.get() == 'Select a month':
                press.print_msg("Error", 'Select a month', Types.ERR)
                top.lift()
                return
            press._CreateFileAvailability(variable1.get(),variable2.get())
            n = datetime.now()
            list_of_worker = ["thtkxu@walla.com", "leah293546@gmail.com"]
            path = "\\" + "קובץ זמינות חודש " + press._GetHebrewNameOfMonth(variable2.get()) + " " + str(n.year)
            sub = "קובץ זמינות חודש " + press._GetHebrewNameOfMonth(variable2.get()) + " " + str(n.year) + " מוכן"
            body = "\nנא לכתוב את השם שלכם בתאריכים שאתם יכולים. תזכרו שזה לאו דווקא אותו היום, אלא גם היום שאחריו.\n יש למלא לפחות 3 תאריכים.\n"
            body2 = "שלום לכולם.\n מצ\"ב קובץ זמינות לחודש הקרוב.\n אבקש למלא בהקדם ולהחזיר לי במייל." + body + "יום נפלא,\n איתיאל"
            press._SendMailTo(list_of_worker, sub, body2, [r"C:\Users\ETL\PycharmProjects\boardworker\system" + path + ".xlsx"])
            press.print_msg("Nice","File Availability has been sent successfully",None)
            top.destroy()

        lable = Label(top, text="Choose a day of the week and month of work",
                      foreground="blue", font="david")
        lable.grid(row=0, column=0, padx=20, pady=5)
        w1 = OptionMenu(top, variable1, *days)
        w1.grid(padx=10, pady=2, row=1, column=0, columnspan=2, ipadx=60)
        w2 = OptionMenu(top, variable2, *month)
        w2.grid(padx=10, pady=2, row=2, column=0, columnspan=2, ipadx=60)
        # button Add
        sub_b = Button(top, text="Send", command=Send_file_availability_helper)
        sub_b.grid(row=3, column=0, columnspan=2, pady=2, padx=10, ipadx=60)

    # this function is sending month arrangment to all workers phones
    def SmsMonthArrangment(phone_of_worker):
        client = nexmo.Client(key='2aed3cd8', secret='d9Jiv7aQar3u7XRP')
        for num in phone_of_worker:
            client.send_message({
                'from': 'Etiel Sofer',
                'to': num,
                'text': 'Work order is ready and sent to your email.\n It can be viewed at the following link -'
                        + " https://technionmail-my.sharepoint.com/:f:/g/personal/ads_asat_technion_ac_il/EuG96T3FPg9NsizGwqnmAlABfQEmlISGhAEO-clpf6C_Yw?e=tb8wcp \n.",
            })


    # this function is sending month arrangment to all workers mail
    def sendMonthArrangment():
        n = datetime.now()
        list_of_worker = ["thtkxu@walla.com", "leah293546@gmail.com"]
        list_of_phone = [972587618955]
        sub = "סידור עבודה חודש " + press._GetHebrewNameOfMonth(4) + " " + str(n.year) + " מוכן"
        body = "שלום לכולם מצ\"ב סידור העבודה לחודש הקרוב.\n יום נפלא, איתיאל"
        press._SendMailTo(list_of_worker, sub, body, [r"C:\Users\ETL\PycharmProjects\boardworker\res\res.xlsx"])
        press.SmsMonthArrangment(list_of_phone)
        press.print_msg("Well done", "The monthly arrangement is sent", None)

    def _ChangeStatus(row, status):
        new_row = []
        statuses = ["בדפוס","הגיע ללקוח","נשלח להדפסה","טרם נשלח"]
        for item in row:
            if statuses.__contains__(item):
                new_row += [status]
                continue
            new_row += [item]
        return new_row

    # convertting all item in list to str
    def _ConvertToString(list):
        new_list = []
        for item in list:
            if str(item) == "nan":
                new_list += [""]
                continue
            new_list += [str(item)]
        return new_list



    # this function can delete or add row. if the row is a list then add it otherwise delete it
    def _ChangeOrder(new_row, path, param="add", status=None):
        data = pd.read_excel(path, header=0)
        workbook = xlsxwriter.Workbook(path)
        worksheet = workbook.add_worksheet("sheet")
        cell_format = workbook.add_format({'bold': True, 'italic': True, "color": "green"})
        worksheet.write_row(0, 0, data.head(0), cell_format)
        i = 1
        for row in data.values:
            cur_row = press._ConvertToString(row)
            if param == "del" and str(row[1]) == str(new_row):
                continue
            if param == "change" and str(row[1]) == str(new_row):
                cur_row = press._ChangeStatus(cur_row, status)
            worksheet.write_row(i, 0, cur_row)
            i += 1
        if param == "add":
            worksheet.write_row(i, 0, new_row)
        worksheet.set_default_row(hide_unused_rows=True)
        worksheet.set_column(chr(65+len(list(data.head(0)))) +":XFD", None, None, {'hidden': True})
        workbook.close()

    def PressAddOrder():
        global top
        if top is not None and top.winfo_exists() is not 0:
            return
        top = tk.Toplevel()
        top.geometry("+30+30")
        top.title("Add orders")
        icon = PhotoImage(file="logo/logo2.png")
        top.iconphoto(top._w, icon)

        def add_order_helper():
            if Order_date.get() == "" or Invitation_name.get() == "" or importer.get() == "" or\
                Order_size.get() == "" or Due_Date.get() == "" or Loading_on.get() == "" or Amount.get() == "":
                press.print_msg("ERROR", "One of the marked fields is empty", Types.ERR)
                top.lift()
                return
            press._ChangeOrder([Order_date.get(),Invitation_name.get(),variable.get(), importer.get(),Order_size.get(),
                                  Due_Date.get(),Amount.get(),Loading_on.get(),price.get(),Remarks.get()],path = "system/orders.xlsx")
            press.print_msg("Well done!","Invitation successfully added", None)
            top.destroy()

        # add Order date
        Order_date = Entry(top, width=30)
        Order_date.grid(row=0, column=1, padx=20)
        Order_date_label = Label(top, text="Order date*")
        Order_date_label.grid(row=0, column=0, padx=20)
        # add Invitation name
        Invitation_name = Entry(top, width=30)
        Invitation_name.grid(row=1, column=1, padx=20)
        Invitation_name_label = Label(top, text="Order name*")
        Invitation_name_label.grid(row=1, column=0, padx=20)
        # add importer
        importer = Entry(top, width=30)
        importer.grid(row=2, column=1, padx=20)
        importer_label = Label(top, text="Importer*")
        importer_label.grid(row=2, column=0, padx=20)
        # add Order size
        Order_size = Entry(top, width=30)
        Order_size.grid(row=3, column=1, padx=20)
        Order_size_label = Label(top, text="Order size*")
        Order_size_label.grid(row=3, column=0, padx=20)
        # add Due Date
        Due_Date = Entry(top, width=30)
        Due_Date.grid(row=4, column=1, padx=20)
        Due_Date_label = Label(top, text="Due Date*")
        Due_Date_label.grid(row=4, column=0, padx=20)
        # add Amount
        Amount = Entry(top, width=30)
        Amount.grid(row=5, column=1, padx=20)
        Amount_label = Label(top, text="Amount*")
        Amount_label.grid(row=5, column=0, padx=20)
        # add Loading on
        Loading_on = Entry(top, width=30)
        Loading_on.grid(row=6, column=1, padx=20)
        Loading_on_label = Label(top, text="Loading on*")
        Loading_on_label.grid(row=6, column=0, padx=20)
        # add price
        price = Entry(top, width=30)
        price.grid(row=7, column=1, padx=20)
        price_label = Label(top, text="Price")
        price_label.grid(row=7, column=0, padx=20)
        # add Remarks
        Remarks = Entry(top, width=30)
        Remarks.grid(row=8, column=1, padx=20)
        Remarks_label = Label(top, text="Remarks")
        Remarks_label.grid(row=8, column=0, padx=20)
        # add is status
        variable = StringVar(top)
        variable.set('טרם נשלח')
        status = ["בדפוס","הגיע ללקוח","נשלח להדפסה","טרם נשלח"]
        w = OptionMenu(top, variable, *status)
        w.grid(padx=10, pady=10, row=9, column=0, columnspan=2, ipadx=60)
        # button Add
        sub_b = Button(top, text="Add", command=add_order_helper)
        sub_b.grid(row=10, column=0, columnspan=2, pady=10, padx=10, ipadx=60)

    # this function is getting all orders by names
    def _GetOders():
        orders = []
        data = pd.read_excel("system/orders.xlsx", header=0)
        for row in data.values:
            orders += [row[1]]
            continue
        return orders

    def DeleteOrder():
        orders = press._GetOders()
        if len(orders) == 0:
            press.print_msg("ERROR", "There is no orders", Types.ERR)
            return
        global top
        if top is not None and top.winfo_exists() is not 0:
            return
        top = tk.Toplevel()
        top.geometry("+30+30")
        top.title("Delete order")
        icon = PhotoImage(file="logo/logo2.png")
        top.iconphoto(top._w, icon)

        def delete_order_helper():
            if variable.get() == 'Order name':
                press.print_msg("ERROR", "pick order", Types.ERR)
                return
            press._ChangeOrder(variable.get(),"system/orders.xlsx","del")
            press.print_msg("Well done!","Order successfully removed", None)
            top.destroy()

        variable = StringVar(top)
        variable.set('Order name')
        w = OptionMenu(top, variable, *orders)
        w.grid(padx=10, pady=10, row=0, column=0, columnspan=2, ipadx=60)
        # button Add
        sub_b = Button(top, text="Delete", command=delete_order_helper)
        sub_b.grid(row=1, column=0, columnspan=2, pady=10, padx=10, ipadx=60)

    def ChangeStatus():
        orders = press._GetOders()
        if len(orders) == 0:
            press.print_msg("ERROR", "There is no orders", Types.ERR)
            return
        global top
        if top is not None and top.winfo_exists() is not 0:
            return
        top = tk.Toplevel()
        top.geometry("+30+30")
        top.title("Change order status")
        icon = PhotoImage(file="logo/logo2.png")
        top.iconphoto(top._w, icon)

        def change_order_helper():
            if variable.get() == 'Order name':
                press.print_msg("ERROR", "pick order", Types.ERR)
                return
            press._ChangeOrder(variable.get(), "system/orders.xlsx", "change",variable1.get())
            press.print_msg("Well done!","Status Successfully changed", None)
            top.destroy()

        variable = StringVar(top)
        variable.set('Order name')
        w = OptionMenu(top, variable, *orders)
        w.grid(padx=10, pady=10, row=0, column=0, columnspan=2, ipadx=60)
        variable1 = StringVar(top)
        variable1.set('טרם נשלח')
        status = ["בדפוס", "הגיע ללקוח", "נשלח להדפסה", "טרם נשלח"]
        w1 = OptionMenu(top, variable1, *status)
        w1.grid(padx=10, pady=10, row=1, column=0, columnspan=2, ipadx=60)
        # button Add
        sub_b = Button(top, text="change", command=change_order_helper)
        sub_b.grid(row=2, column=0, columnspan=2, pady=10, padx=10, ipadx=60)

    # this function taking file archive it in folder and starting new one
    # param - namefile = the name of file in system folder without suffix. msg = define if to print msg or not.
    def _StartNew(namefile,msg = True):
        n = datetime.now()
        if n.month == 1:
            name = namefile + " " + press._GetHebrewNameOfMonth(12) + " " + str(n.year-1) + ".xlsx"
        else:
            name = namefile + " " + press._GetHebrewNameOfMonth(n.month-1) + " " + str(n.year) + ".xlsx"
        copyfile("system/"+namefile+".xlsx","system/archive/"+name)
        data = pd.read_excel("system/"+namefile+".xlsx", header=0)
        workbook = xlsxwriter.Workbook("system/"+namefile+".xlsx")
        worksheet = workbook.add_worksheet(namefile)
        cell_format = workbook.add_format({'bold': True, 'italic': True, "color": "green"})
        worksheet.write_row(0, 0, data.head(0), cell_format)
        worksheet.set_default_row(hide_unused_rows=True)
        worksheet.set_column('K:XFD', None, None, {'hidden': True})
        workbook.close()
        if msg:
            press.print_msg("GOOD","Now you can start your " + press._GetHebrewNameOfMonth(n.month) + " bookings. Previous "+ namefile+ " have been archived", None)

    def StartNewOrders():
        press._StartNew("orders")

    # unused
    # this function get list and list of sizes and feet the word to middle of size
    def _ListToString(list,sizes):
        cur_list = press._ConvertToString(list)
        #cur_list.reverse()
        str1 = ""
        for i in range(len(cur_list)):
            temp_size = int((sizes[i]) - len(cur_list[i])/2)
            if temp_size < 0:
                temp_size = 0
            for j in range(temp_size):
                str1 += " "
            str1 += str(cur_list[i])
            for j in range(temp_size):
                str1 += " "
            str1 += " | "
        return str1

    # unused
    def _PrintFile(path):
        sizes = []
        str1 = ""
        data = pd.read_excel(path, header=0)
        for i in data.head(0):
            sizes += [len(i)]
        str1 += press._ListToString(data.head(0),sizes)+"\n"
        #sizes.reverse()
        for row in data.values:
            str1 += press._ListToString(row,sizes) + "\n"
        press.print_msg("Result",str1,None)

    def PrintOrders():
        press._PrintFileToScream("system/orders.xlsx")

    def PrintHours():
        press._PrintFileToScream("system/hours.xlsx")

    def _Hours(op):
        global top
        if top is not None and top.winfo_exists() is not 0:
            return
        n = datetime.now()
        top = tk.Toplevel()
        top.geometry("+30+30")
        top.title("Add orders")
        icon = PhotoImage(file="logo/logo2.png")
        top.iconphoto(top._w, icon)

        def enter_hours_helper():
            if Worker_name.get() == "" or hours.get() =="" or Loading_on.get() == "" or reason.get() == "":
                press.print_msg("ERROR", "One of the marked fields is empty", Types.ERR)
                top.lift()
                return
            date = str(n.day) + "/" + str(n.month) + "/" + str(n.year)
            hour = hours.get() if op == "plus" else "-"+hours.get()
            press._ChangeOrder([date, Worker_name.get(),hour, Loading_on.get(),reason.get(),note.get()],"system/hours.xlsx")
            press.print_msg("Good!","The hours were successfully added to the employee", None)
            top.destroy()

        # add Worker name
        Worker_name = Entry(top, width=30)
        Worker_name.grid(row=0, column=1, padx=20)
        Worker_name_label = Label(top, text="Worker full name*")
        Worker_name_label.grid(row=0, column=0, padx=20)
        # add num of hours
        hours = Entry(top, width=30)
        hours.grid(row=1, column=1, padx=20)
        hours_label = Label(top, text="Num of hours*")
        hours_label.grid(row=1, column=0, padx=20)
        # add Loading on
        Loading_on = Entry(top, width=30)
        Loading_on.grid(row=2, column=1, padx=20)
        Loading_on_lable = Label(top, text="Loading on*")
        Loading_on_lable.grid(row=2, column=0, padx=20)
        # add reason
        reason = Entry(top, width=30)
        reason.grid(row=3, column=1, padx=20)
        reason_label = Label(top, text="Reason*")
        reason_label.grid(row=3, column=0, padx=20)
        # add note
        note = Entry(top, width=30)
        note.grid(row=4, column=1, padx=20)
        note_label = Label(top, text="Note")
        note_label.grid(row=4, column=0, padx=20)
        # button enter
        sub_b = Button(top, text="Enter", command=enter_hours_helper)
        sub_b.grid(row=5, column=0, columnspan=2, pady=10, padx=10, ipadx=60)

    def EnterHours():
        press._Hours("plus")

    def StartNewHours():
        press._StartNew("hours")

    def DeleteHours():
        press._Hours("minus")

    def restart():
        if press.print_msg("Are you sure?", "The action sends an email to the office manager and a new monthly concentration begins. to approve?", Types.YESORNOT) == 1:
            n = datetime.now()
            sub = "ריכוז חודש " + press._GetHebrewNameOfMonth(n.month-1) +" " +str(n.year) + " מוכן"
            mail_Office_Manager = "leah293546@gmail.com"
            press._SendMailTo([mail_Office_Manager],sub, "", [r"C:\Users\ETL\PycharmProjects\boardworker\system\orders.xlsx",r"C:\Users\ETL\PycharmProjects\boardworker\system\hours.xlsx"])
            press._StartNew("hours",False)
            press._StartNew("orders",False)
            press.print_msg("Wooo", "The email was sent to the office manager. Have a successful month :)",None)