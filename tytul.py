#-*- coding: utf-8 -*-

import Tkinter, Tkconstants, tkFileDialog

import datetime

format = "%Y%m%d_%H%M%S"




class Window(Tkinter.Frame):
    counter = 0


    def __init__(self, master=None):
        Tkinter.Frame.__init__(self, master)
        self.master = master
        self.init_window()

    #Creation of init_window
    def init_window(self):

        # changing the title of our master widget
        self.master.title("Wróbelek Iksemelek 0.1")

        # allowing the widget to take the full space of the root window
        self.pack(fill=Tkinter.BOTH, expand=1)

        # creating a button instance
       # quitButton = Tkinter.Button(self, text="Quit")

        # placing the button on my window
       # quitButton.place(x=0, y=0)

        button_opt = {'fill': Tkconstants.BOTH, 'padx': 5, 'pady': 5}

        # define buttons
        Tkinter.Button(self, text='Otwórz plik Excel (*.xlsx)...', command=self.askopenfile).pack(**button_opt)
        #Tkinter.Button(self, text='Zapisz XML', command=self.askopenfilename).pack(**button_opt)
        #Tkinter.Button(self, text='asksaveasfile', command=self.asksaveasfile).pack(**button_opt)
        Tkinter.Button(self, text='Zapisz XML', command=self.asksaveasfilename).pack(**button_opt)
        Tkinter.Button(self, text='Okno', command=self.create_window).pack(**button_opt)
        Tkinter.Button(self, text="Quit", command=self.koniec).pack(**button_opt)

        today = datetime.datetime.today()

        s = today.strftime(format)

        # define options for opening or saving a file
        self.file_opt = options = {}
        options['defaultextension'] = '.xlsx'
        options['filetypes'] = [('all files', '.*'), ('Pliki Excel 2003+', '.xlsx')]
        options['initialdir'] = 'C:\\'
        options['initialfile'] = s+'.xml'
        options['parent'] = root
        options['title'] = 'This is a title'

        # defining options for opening a directory
        self.dir_opt = options = {}
        options['initialdir'] = 'C:\\'
        options['mustexist'] = False
        options['parent'] = root
        options['title'] = 'This is a title'

    def askopenfile(self):

        """Returns an opened file in read mode."""

        return tkFileDialog.askopenfile(mode='r', **self.file_opt)

    def askopenfilename(self):

        """Returns an opened file in read mode.
        This time the dialog just returns a filename and the file is opened by your own code.
        """

        # get filename
        filename = tkFileDialog.askopenfilename(**self.file_opt)

        # open file on your own
        if filename:
            return open(filename, 'r')

    def asksaveasfile(self):

        """Returns an opened file in write mode."""

        return tkFileDialog.asksaveasfile(mode='w', **self.file_opt)

    def asksaveasfilename(self):

        """Returns an opened file in write mode.
        This time the dialog just returns a filename and the file is opened by your own code.
        """

        today = datetime.datetime.today()

        s = today.strftime(format)
        self.file_opt['initialfile'] = s + '.xml'

        #options['initialfile'] = s + '.xml'

        # get filename
        filename = tkFileDialog.asksaveasfilename(**self.file_opt)

        # open file on your own
        if filename:
            return open(filename, 'w')

    def askdirectory(self):

        """Returns a selected directoryname."""

        return tkFileDialog.askdirectory(**self.dir_opt)

    def create_window(self):
        self.counter += 1
        t = Tkinter.Toplevel(self)
        t.wm_title("Window #%s" % self.counter)
        l = Tkinter.Text(t, state='normal')
        l.insert('end',"This is window #%s" % self.counter)
        l.pack(side="top", fill="both", expand=True, padx=10, pady=10)

    def koniec(self):
        root.quit()


root = Tkinter.Tk()

#size of the window
root.geometry("350x250")

app = Window(root)
root.mainloop()