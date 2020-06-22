#!/usr/bin/env python

from tkinter import *
from PIL import ImageTk, Image
from press import press





class GraphicBoard:

    def __init__(self, button_list, size=None):
        self.button = []
        self.frames = []
        self.buttonforframe = []
        self.rediobottun = []
        self.valofrb = 0
        self.root = Tk()
        self.root.geometry("+30+30")
        self.root.geometry(size)
        self.button_list = button_list
        self.color = "blue"
        self.font = None

    def AddTitle(self, title):
        self.root.title(title)

    def AddLogo(self, path):
        icon = PhotoImage(file=path)
        self.root.call('wm', 'iconphoto', self.root._w, icon)

    def AddBg(self, path_to_image, columnspan):
        self.image1 = ImageTk.PhotoImage(Image.open(path_to_image))
        self.my_label = Label(image=self.image1, bg=self.color)
        self.my_label.grid(row=0, column=0, padx = 10, pady =10, columnspan=columnspan)

    def AddFrame(self,padx, pady, row, column, columnspan):
        self.frames += [Frame(self.root, bg=self.color)]
        self.frames[len(self.frames)-1].grid(padx=padx, pady=pady, row=row, column=column, columnspan=columnspan)

    def DestroyFrame(self, frame):
        frame.destroy()
        
    def AddMenu(self):
        self.menubar = Menu(self.root,bg=self.color)

        # create a pulldown menu, and add it to the menu bar
        self.filemenu1 = Menu(self.menubar, tearoff=0,bg=self.color)
        self.filemenu1 .add_command(label="Open", command=press.PressLoad,foreground="white", font=self.font)
        self.filemenu1 .add_command(label="About", command=press.About,foreground="white", font=self.font)
        self.filemenu1.add_command(label="Pricelist", command=None,foreground="white", font=self.font)
        self.filemenu1.add_command(label="Current boards", command=None,foreground="white", font=self.font)
        self.filemenu1 .add_separator()
        self.filemenu1 .add_command(label="Instructions", command=None,foreground="white", font=self.font)
        self.filemenu1.add_command(label="Automatic result", command=press.do_it_auto,foreground="white", font=self.font)
        self.filemenu1.add_command(label="Restart", command=None,foreground="white", font=self.font) #delete workers name and workers file and regions-need to ask if you sure

        self.filemenu2 = Menu(self.menubar, tearoff=0,bg=self.color,foreground="white", font=self.font)
        self.filemenu2.add_command(label="Add worker", command=press.Add_worker)
        self.filemenu2.add_command(label="Delete worker", command=press.DeleteWorker)
        self.filemenu2.add_command(label="Add region to worker", command=press.AddRegionToWorker,foreground="white", font=self.font)
        self.filemenu2.add_command(label="Add Preferences to worker", command=press.AddPreferencesToWorker,foreground="white", font=self.font)
        self.filemenu2.add_separator()
        self.filemenu2.add_command(label="Download workers file from outlook", command=press.download_worker_files,foreground="white", font=self.font)
        self.filemenu2.add_command(label="Delete old workers file", command=press.Delete_worker_files,foreground="white", font=self.font)
        self.filemenu2.add_command(label="Send file availability", command=press.sendFileAvailability ,foreground="white", font=self.font)
        self.filemenu2.add_command(label="Send month arrangement", command=press.sendMonthArrangment,foreground="white", font=self.font)
        #self.filemenu2.add_command(label="Send sms month arrangement", command=press.SmsMonthArrangment, foreground="white", font=self.font)
        self.filemenu2.add_command(label="Send week arrangement", command=None,foreground="white", font=self.font)
        self.filemenu2.add_separator()
        self.filemenu2.add_command(label="Alerts", command=press.Allert,foreground="white", font=self.font)

        self.filemenu3 = Menu(self.menubar, tearoff=0, bg=self.color)
        self.filemenu3.add_command(label="Add region", command=press.AddRegion,foreground="white", font=self.font)
        self.filemenu3.add_command(label="Delete region", command=press.DeleteRegion, foreground="white", font=self.font)

        self.filemenu4 = Menu(self.menubar, tearoff=0, bg=self.color)
        self.filemenu4.add_command(label="Add order", command=press.PressAddOrder, foreground="white", font=self.font)
        self.filemenu4.add_command(label="Delete order", command=press.DeleteOrder, foreground="white", font=self.font)
        self.filemenu4.add_command(label="Change status", command=press.ChangeStatus, foreground="white", font=self.font)
        self.filemenu4.add_command(label="View orders", command=press.PrintOrders, foreground="white", font=self.font)
        self.filemenu4.add_separator()
        self.filemenu4.add_command(label="Start new monthly orders", command=press.StartNewOrders, foreground="white",font=self.font)

        self.filemenu5 = Menu(self.menubar, tearoff=0, bg=self.color)
        self.filemenu5.add_command(label="Enter hours", command=press.EnterHours, foreground="white", font=self.font)
        self.filemenu5.add_command(label="Delete hours", command=press.DeleteHours, foreground="white", font=self.font)
        self.filemenu5.add_command(label="View hours", command=press.PrintHours, foreground="white", font=self.font)
        self.filemenu5.add_separator()
        self.filemenu5.add_command(label="Start new monthly Hours", command=press.StartNewHours, foreground="white", font=self.font)


        self.menubar .add_cascade(label="File", menu=self.filemenu1, font=self.font)
        self.menubar.add_cascade(label="Worker", menu=self.filemenu2, font=self.font)
        self.menubar.add_cascade(label="Region", menu=self.filemenu3, font=self.font)
        self.menubar.add_cascade(label="Orders", menu=self.filemenu4, font=self.font)
        self.menubar.add_cascade(label="Hours", menu=self.filemenu5, font=self.font)
        self.root.config(menu=self.menubar)



    def AddButton(self, name, func, padx, pady, row, column, columnspan):#need to update to add button into frame
        self.button += [Button(self.root, text=name, padx=padx, pady=pady, command=func,bg=self.color)]
        self.button[len(self.button)-1].grid(row = row, column = column, columnspan = columnspan)

    def AddButtonToLastFrame(self, button_list):
        for i in range(len(button_list)):
            self.buttonforframe += [Button(self.frames[len(self.frames) - 1], text=button_list[i][0], padx=20, pady=10,
                                           command=button_list[i][1], bg=self.color,font=self.font ,highlightcolor="white",foreground="white")]
            self.buttonforframe[len(self.buttonforframe) - 1].grid(row = 0, column = i,ipadx=40, ipady=10)

    def AddRedioButton(self, text, row, column):
        self.rediobottun += [Radiobutton(self.root, text=text, variable=self.valofrb, value=len(self.rediobottun)+1, bg=self.color)]
        self.rediobottun[len(self.rediobottun)-1].grid(row = row, column = column)

    def StartWindow(self):
        #define title and logo
        self.AddTitle("Welcome to Distributor")
        self.AddLogo("logo/logo2.png")
        #define images and bg
        self.AddBg("files/image1.jpeg",4)
        #define bottom
        self.AddFrame(10,10, 1,0,4)
        self.AddButtonToLastFrame(self.button_list)
        #adding menu
        self.AddMenu()
        self.root.mainloop()



# here we define the window function --------------------------------------------------------------------------------

# ---------------------------------------------------------------------------------------------------------------------

# the list of main buttons
button_list = [("Load", press.PressLoad), ("Enter hours",press.EnterHours), ("Add order",press.PressAddOrder), ("Start new month", press.restart)]


#main
if __name__ == "__main__":
    GraphicBoard(button_list).StartWindow()

