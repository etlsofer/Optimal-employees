from tkinter import *
from PIL import ImageTk, Image






class TopWindow:
    def __init__(self):
        self.top = Toplevel()
        self.top.geometry("+30+30")
        self.button = []
        self.frames = []
        self.buttonforframe = []
        self.rediobottun = []
        self.valofrb = IntVar()

    def AddTitle(self,title):
        self.top.title(title)

    def AddLogo(self, path):
        icon = PhotoImage(file=path)
        self.top.iconphoto(self.top._w, icon)
        return

    # updatting the top window to the front
    def TopME(self):
        self.top.lift()

    # adding a buck ground
    def AddBg(self, path_to_image,columnspan):
        self.image1 = ImageTk.PhotoImage(Image.open(path_to_image))
        self.my_label = Label(self.top, image=self.image1)
        self.my_label.grid(row=0, column=0, columnspan=columnspan)

    def AddFrame(self, padx, pady, row, column, columnspan):
        self.frames += [Frame(self.top)]
        self.frames[len(self.frames)-1].grid(padx = padx, pady = pady, row=row,column=column,columnspan=columnspan)

    def AddButton(self, name, func, padx, pady, row, column, columnspan):  # need to update to add button into frame
        self.button += [Button(self.top, text=name, padx=padx, pady=pady, command=func)]
        self.button[len(self.button) - 1].grid(row=row, column=column, columnspan=columnspan)

    def AddButtonToLastFrame(self, button_list):
        for i in range(len(button_list)):
            self.buttonforframe += [Button(self.frames[len(self.frames)-1], text=button_list[i][0], padx=2, pady=2, command=button_list[i][1])]
            self.buttonforframe[len(self.buttonforframe)-1].pack()

    def AddRedioButton(self, text, row, column):
        self.rediobottun += [Radiobutton(self.top, text=text, variable=self.valofrb, value=len(self.rediobottun)+1)]
        self.rediobottun[len(self.rediobottun)-1].grid(row = row, column = column)


class LoadData(TopWindow):

    def __init__(self,buuton, Rb, path,lastbuttons):
        super().__init__()
        self.buttons = buuton
        self.RB = Rb
        self.path=path
        self.lastbuttons = lastbuttons

    def TopME(self):
        self.top.lift()

    def WindowDesign(self,turnon = True):
        self.AddTitle("Load Data")
        #self.AddBg(self.path,0)
        self.AddLogo("logo/logo2.png")
        self.AddFrame(2,2,0,0,3)
        self.AddButtonToLastFrame(self.buttons)
        for r in range(len(self.RB)):
            self.AddRedioButton(self.RB[r], r+1, 0)

        self.AddFrame(2, 2, 3, 0, 3)
        self.AddButtonToLastFrame(self.lastbuttons)
        if not turnon:
            self.button[len(self.button) - 1]['state'] = DISABLED

    def Destroy(self):
        self.top.destroy()
        self.button = []
        self.frames = []
        self.buttonforframe = []
        self.rediobottun = []


class LoadFromKB(TopWindow):

    def __init__(self, list_of_checkbox_names, button):
        super().__init__()
        self.list_of_names = list_of_checkbox_names
        self.var_for_checkbox =[]
        for i in range(len(self.list_of_names)):
            self.var_for_checkbox += [IntVar()]
        self.checkboxes =[]
        self.check_button = button


    def AddCheckBoxes(self):
        for i in range(len(self.list_of_names)):
            self.checkboxes += [Checkbutton(self.top, text=self.list_of_names[i], variable=self.var_for_checkbox[i])]
            self.checkboxes[i].grid(row=i, sticky=W)

    def LoadKBDesign(self):
        self.AddCheckBoxes()
        self.AddButton(self.check_button[0], self.check_button[1], 20, 10, 3,0,1)


