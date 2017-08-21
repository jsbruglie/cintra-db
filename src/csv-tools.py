# GUI
import Tkinter as tk
import ttk
import tkFont, tkFileDialog, tkMessageBox
# Open CSV files
import csv
# Open Developer URL
import webbrowser
# OS specific commands
import os
# Windows specific behavior
import platform
import ctypes
# Operator as parameter
import operator

# Debug
import pprint

# Application properties

APP_TITLE = "CSV Tools"
APP_VERSION = "v1.0"
APP_AUTHOR = "Joao Borrego"
AUTHOR_URL = "https://github.com/jsbruglie"

# Config files

PREPROC_FILE = "preproc.csv"

# TODO - Debug
pp = pprint.PrettyPrinter(indent=4)

# Generic Utilities

def get_truth(inp, relate, cut):
    ops = {">": operator.gt,
           "<": operator.lt,
           ">=": operator.ge,
           "<=": operator.le,
           "=": operator.eq}
    return ops[relate](inp, cut)

def filterData(data, col, operator, value):
    """ Filter table and return matching entries in new table
        :param table: 
    """
    filtered_data = []
    
    # Try float conversion and then comparison
    try:
        for i in range(len(data)):
            if get_truth(float(data[i][col]), operator, float(value)):
                filtered_data.append(data[i])    
    
    except ValueError:
        for i in range(len(data)):
            if get_truth(data[i][col], operator, value):
                filtered_data.append(data[i])
    
    return filtered_data

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

def readPreProcOpt():
    """ Opens preprocessing file returns required changes """

    # Whether the file is to be preprocessed
    modify = True
    # The amount of characters to skip at the begining of the file
    skip = 0
    # Other desired data modifications
    changes = []

    if os.path.exists(PREPROC_FILE):    
        with open(PREPROC_FILE, "r") as preproc:
            reader = csv.reader(preproc)
            changes = map(list, reader)

            preproc_msg = 'Detected pre-processing file with changes:' + '\n'
            for change in changes:
                preproc_msg = preproc_msg + '-' + change[0] + '\n'

            # Prompt user to confirm changes
            if not tkMessageBox.askokcancel("Apply Changes", preproc_msg):
                modify = False

            for change in changes:
                if (change[1] == "skipInitialNChars"):
                    skip = int(change[2])

    return [modify, skip, changes]


def preProcDb(data, changes):
    """ Applies preprecessing to the database """

    new_data = data

    for change in changes:
        if (change[1] == "removeNCharsFromCol"):
            [num_chars, col, start] = [int(change[2]), int(change[3]), (change[4]=='True')]
            new_data = removeNCharsFromCol(new_data,
                n=num_chars, col=col, start=start)

    return new_data

def removeNCharsFromCol(data, n, col, start):
    """ Removes n characters from the value of a given column 
        for every row either from the start or the end of
        the string
        :param data: The data to process
        :param n: The number of characters
        :param col: The index of the column to alter
        :param start: Remove from start (True) or end (False)
    """
    for i in range(len(data)):
        try:
            data[i][col] =  data[i][col][n:] if start else data[i][col][:-n]
        except IndexError:
            pass # Empty field
    return data

class MainApp(tk.Frame):
    """ Main App class """

    def __init__(self, parent, header=None, data=None):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        self.parent.title(APP_TITLE)
        
        try:
            self.parent.iconbitmap(default="logo.ico")
        except tk.TclError:
            pass # Icon image not present, load default
        
        self.parent.minsize(width=300, height=200)
        self.parent.maxsize(width=parent.winfo_screenwidth(), 
                            height=parent.winfo_screenheight())

        self.header = header
        self.data= data

        self.setupWidgets()

    def setupWidgets(self):
        """ Declare and initialize widgets """

        # Allow frame to be extended
        self.pack(fill="both", expand=True)

        self.createMenuBar()
        
        ttk.Style().configure("Treeview", font=('', 10), background="#383838", 
                                foreground="white", fieldbackground="#383838")

        if (self.header is not None) and (self.data is not None):
            self.createTree()

    def createMenuBar(self):
        """ Create top menu bar """
        menu_bar = tk.Menu(self.parent)
        self.parent.config(menu=menu_bar)

        file_menu = tk.Menu(menu_bar, tearoff=False)
        file_menu.add_command(label="Open", command=self.onOpen)
        file_menu.add_separator()
        file_menu.add_command(label="Save", command=self.onSave)
        file_menu.add_command(label="New Window", command=self.onNewWindow)
        menu_bar.add_cascade(label="File", menu=file_menu)
        
        menu_bar.add_command(label="Find", command=self.onFind)

        menu_bar.add_command(label="About", command=self.onAbout)

    def createPopupMenu(self, parent):
        """ Create a popup context menu when right click """
        
        select_all_label = "Select All (" + str(self.num_entries) +") - Ctrl+A"

        self.popup_menu = tk.Menu(parent, tearoff=False)
        self.popup_menu.add_command(label="New Window", command=self.onNewWindow)
        self.popup_menu.add_separator()
        self.popup_menu.add_command(label=select_all_label, command=self.onSelectAll)
        self.popup_menu.add_command(label="Copy - Ctrl+C", command=self.onCopy)
        self.popup_menu.add_command(label="Copy Header", command=self.onCopyHeader)

        parent.bind("<Button-3>", self.popup)

    def popup(self, event):
        """ Show popup menu on event """
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

        # Allow treeview to be resized
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.buildTree()
        self.num_entries = len(self.data)

        self.createPopupMenu(self.tree)

    def buildTree(self):
        """ Update the contents of the treeview table """
        for col in self.header:
            self.tree.heading(col, text=col.title(),
                command=lambda c=col: self.sortBy(self.tree, c, 0))
            self.tree.column(col, width=tkFont.Font().measure(col.title()))

        for row in self.data:
            self.tree.insert("","end", values=row)

    def onSelectAll(self, event=None):
        """ Select all the entries in the tree view table"""
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

    def onCopyHeader(self, event=None):
        """ Handle copy header line to clipboard action """
        self.clipboard_clear()
        text = ''.join(','.join(map(unicode, self.header))) + '\n,'
        self.clipboard_append(text)

    def onNewWindow(self, event=None, header=None, data=None):
        """ Creates a duplicate of the main window """
        new_window = tk.Toplevel()
        new_instance = MainApp(new_window, header, data)

    def onMousewheel(self, event):
        """ Scroll treeview table vertically """
        self.tree.yview_scroll(-1*(event.delta/120), "units")

    def sortBy(self, tree, col, descending):
        """ Sort the treeview rows according to the value of a given column """ 
        data = [(tree.set(child, col), child) for child in tree.get_children('')]
        try:
            data.sort(key=lambda tup: float(tup[0]), reverse=descending)
        except ValueError:
            data.sort(key=lambda tup: tup[0], reverse=descending)
        
        for idx, item in enumerate(data):
            tree.move(item[1], '', idx)
        
        # switch the heading so that it will sort in the opposite direction
        tree.heading(col,
            command=lambda col=col: self.sortBy(tree, col, int(not descending)))

    # Menu commands

    def onOpen(self):
        """ Open CSV file to memory """
        file_types = [('CSV files', '*.csv'), ('All files', '*')]
        dialogue = tkFileDialog.Open(self, filetypes=file_types)
        file_path = dialogue.show()
        if file_path is '':
            return

        with open(file_path, "r") as csv_file:
            
            [preproc, skip, changes] = readPreProcOpt()

            csv_file.seek(skip, 1)
            reader = csv.reader(csv_file)
            self.header = list(reader.next())
            self.data = map(list, reader)

            if (preproc):
                self.data = preProcDb(self.data, changes)

            self.createTree()
            raiseAboveAll(self.parent)

    def onSave(self):
        """ Save current table to CSV file """
        
        if (self.header is None or self.data is None):
            return

        file_types = [('CSV files', '*.csv'), ('All files', '*')]
        file_path = tkFileDialog.asksaveasfilename(filetypes=file_types, defaultextension=".csv")
        if file_path is '':
            return
        
        with open(file_path, "wb") as out_file:
            writer = csv.writer(out_file)
            writer.writerow(self.header)
            writer.writerows(self.data)

    def onFind(self):
        """ Detached toplevel find frame """
        
        if (self.header is None or self.data is None):
            return

        toplevel = tk.Toplevel()
        toplevel.title("Find")

        # Column selection
        column_label = tk.Label(toplevel, text="Column: ")
        col_var = tk.StringVar()
        col_var.set(self.header[0])
        column_options = tk.OptionMenu(toplevel, col_var, *self.header)

        # Operation selection
        valid_ops = ["=","<", ">","<=",">="]
        ops_var = tk.StringVar()
        ops_var.set(valid_ops[0])
        ops_options = tk.OptionMenu(toplevel, ops_var, *valid_ops)

        # Entry
        txt_entry = tk.Entry(toplevel)

        # Submit button
        submit_btn = tk.Button(toplevel, text="Filter",
            command=lambda: self.onFilter(toplevel, col_var.get(), ops_var.get(), txt_entry.get()))

        # Grid
        column_label.grid(row=0, column=0, sticky="W", padx=5, pady=5)
        column_options.grid(row=0, column=1, sticky="W", padx=5, pady=5)
        ops_options.grid(row=0, column=2, sticky="W", padx=10, pady=5)
        txt_entry.grid(row=0, column=3, sticky="W", padx=10, pady=5)
        submit_btn.grid(row=1, column=3, sticky="E", padx=10, pady=5)
    
    def onFilter(self, parent, column, op, value):
        """ Filter table data according to query """        
        col_idx = self.header.index(column)
        filtered_data = filterData(self.data, col_idx, op, value)
        self.onNewWindow(header=self.header, data=filtered_data)
        parent.destroy()

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
        my_app_id = 'borrego.cintra-db.v0.1'
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