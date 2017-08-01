import Tkinter as tk
import ttk
import tkFont, tkFileDialog, tkMessageBox

import csv

import webbrowser

import platform
import ctypes

APP_TITLE = "Cintra-DB"
APP_VERSION = "v0.05"
APP_AUTHOR = "Joao Borrego"
AUTHOR_URL = "https://github.com/jsbruglie"

# Generic Utilities

def center(toplevel):
    """ Place a frame on the center of the screen
    :param toplevel: The container to be centered
    """
    toplevel.update_idletasks()
    w = toplevel.winfo_screenwidth()
    h = toplevel.winfo_screenheight()
    size = tuple(int(_) for _ in toplevel.geometry().split('+')[0].split('x'))
    x = w/2 - size[0]/2
    y = h/2 - size[1]/2
    toplevel.geometry("%dx%d+%d+%d" % (size + (x, y)))

def raiseAboveAll(window):
    """ Raises a window above all others
    :param window: The window to be raised
    """
    window.attributes('-topmost', 1)
    window.attributes('-topmost', 0)

def openLink(link):
    """ Open url on web browser
    :param link: The desired url
    """
    webbrowser.open_new(link)

class MainApp(tk.Frame):
    """ Main App class """

    def __init__(self, parent):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        self.parent.title(APP_TITLE)
        try:
            self.parent.iconbitmap(default="logo.ico")
        except tk.TclError:
            pass
        self.parent.minsize(width=300, height=200)
        self.parent.maxsize(width=parent.winfo_screenwidth(), 
                            height=parent.winfo_screenheight())

        self.header = []
        self.data= []

        self.setupWidgets()

    def setupWidgets(self):

        # Allow frame to be extended
        self.pack(fill="both", expand=True)

        self.createMenuBar()
        ttk.Style().configure("Treeview", font= ('', 10), background="#383838", 
                                foreground="white", fieldbackground="yellow")

    def createMenuBar(self):
        """ Create top menu bar """
        menu_bar = tk.Menu(self.parent)
        self.parent.config(menu=menu_bar)

        file_menu = tk.Menu(menu_bar, tearoff=False)
        file_menu.add_command(label="Open", command=self.onOpen)
        file_menu.add_separator()
        file_menu.add_command(label="New Window", command=self.onNewWindow)
        menu_bar.add_cascade(label="File", menu=file_menu)
        
        menu_bar.add_command(label="Find", command=None)

        view_menu = tk.Menu(menu_bar, tearoff=False)
        view_menu.add_command(label="Show/Hide Columns", command=None)
        menu_bar.add_cascade(label="View", menu=view_menu)

        menu_bar.add_command(label="About", command=self.onAbout)

    def createPopupMenu(self, parent):
        """ Create a popup context menu when right click """
        self.popup_menu = tk.Menu(parent, tearoff=False)
        self.popup_menu.add_command(label="New Window", command=self.onNewWindow)
        self.popup_menu.add_separator()
        self.popup_menu.add_command(label="Select All", command=self.onSelectAll)
        self.popup_menu.add_command(label="Copy", command=self.onCopy)
        parent.bind("<Button-3>", self.popup)

    def popup(self, event):
        # Show popup menu on event
        self.popup_menu.post(event.x_root, event.y_root)

    def createTree(self):
        """ Create treeview  for table representation """
        self.tree = ttk.Treeview(self.parent, columns=self.header, show="headings")
        y_scrollbar = ttk.Scrollbar(self.parent, orient="vertical", command=self.tree.yview)
        x_scrollbar = ttk.Scrollbar(self.parent, orient="horizontal", command=self.tree.xview)
        
        self.tree.configure(yscrollcommand=y_scrollbar.set,
                            xscrollcommand=x_scrollbar.set)

        self.tree.grid(column=0, row=0, sticky='nsew', in_=self)
        y_scrollbar.grid(column=1, row=0, sticky='ns', in_=self)
        x_scrollbar.grid(column=0, row=1, sticky='ew', in_=self)

        # Bind events
        self.tree.bind('<Control-c>', self.onCopy)
        self.tree.bind('<Control-a>', self.onSelectAll)
        y_scrollbar.bind_all("<MouseWheel>", self.onMousewheel)

        self.createPopupMenu(self.tree)

        # Allow treeview to be resized
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.buildTree()

    def buildTree(self):
        """ Update the contents of the treeview table """
        for col in self.header:
            self.tree.heading(col, text=col.title(),
                command=lambda c=col: self.sortBy(self.tree, c, 0))
            self.tree.column(col, width=tkFont.Font().measure(col.title()))

        for row in self.data:
            self.tree.insert("","end", values=row)

    def onSelectAll(self, event=None):
        selected = [child for child in self.tree.get_children('')]
        for item in selected:
            self.tree.selection_add(item)

    def onCopy(self, event=None):
        """ Handle Ctrl+C copy selection to clipboard action """
        self.clipboard_clear()
        sel_id = self.tree.selection()
        rows = [self.tree.item(row_id)["values"] for row_id in sel_id]
        text = '\n'.join(','.join(map(unicode, row)) for row in rows)
        self.clipboard_append(text)
    
    def onNewWindow(self, event=None):
        """ Creates a duplicate of the main window """
        new_window = tk.Toplevel()
        new_instance = MainApp(new_window)

    def onMousewheel(self, event):
        self.tree.yview_scroll(-1*(event.delta/120), "units")

    def sortBy(self, tree, col, descending):
        
        data = [(tree.set(child, col), child) for child in tree.get_children('')]
        try:
            data.sort(key=lambda tup: int(tup[0]), reverse=descending)
        except ValueError:
            data.sort(key=lambda tup: tup[0], reverse=descending)
        
        for idx, item in enumerate(data):
            tree.move(item[1], '', idx)
        
        # switch the heading so that it will sort in the opposite direction
        tree.heading(col,
            command=lambda col=col: self.sortBy(tree, col, int(not descending)))

    # Menu commands
    def onOpen(self):
        global tree_header, tree_data

        file_types = [('CSV files', '*.csv'), ('All files', '*')]
        dialogue = tkFileDialog.Open(self, filetypes = file_types)
        file_path = dialogue.show()

        if file_path != '':
            with open(file_path, "r") as csv_file:
                csv_file.seek(3,1) # Skip initial 3 garbage bytes; Don't ask
                reader = csv.reader(csv_file)
                self.header = list(reader.next())
                self.data = map(list, reader)
                self.createTree()
                raiseAboveAll(self.parent)
                #self.tree["displaycolumns"]=("num_processo")

    def onAbout(self):
        """ Detached toplevel about frame """
        toplevel = tk.Toplevel()
        toplevel.geometry("300x80")
        center(toplevel)
        toplevel.title("About")
        app_title = tk.Label(toplevel, text=APP_TITLE, font=("Helvetica", 16))
        app_version = tk.Label(toplevel, text=APP_VERSION)
        app_author = tk.Label(toplevel, text=APP_AUTHOR, fg="blue", cursor="hand2")
        app_author.bind("<Button-1>", lambda event, link=AUTHOR_URL: openLink(link))
        app_title.pack()
        app_version.pack()
        app_author.pack()



def main():

    # Windows only
    if platform.system() == "Windows":
        my_app_id = 'borrego.cintra-db.v0.05'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(my_app_id)

    root = tk.Tk()

    def onClosing():
        """ Root close confirmation prompt """
        if tkMessageBox.askokcancel("Quit", "Are you sure you want to quit?"):
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", onClosing)
    app = MainApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()