"""
Ballot Manifest Data Entry Software

The goal of this software is to create an easy to use interface for creating, updating, and removing entries
in a ballot manifest that will be used for a Risk Limiting Audit (RLA).

This software uses the tkinter module for rendering a graphical user interface (GUI) and uses the pandas module
for editing and sorting data frames of information.

All GUI variables are declared in the __init__ method of the BallotManifestGui class, all "views" are generated in the
views methods, and all functions and data frame manipulation occurs in the remaining methods.

For feedback, questions, or support:
Contact Jonathan Layman at jlayman@arapahoegov.com
"""


import pandas as pd
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText


class BallotManifestGui:
    def __init__(self, master):
        # Establish base GUI settings.
        # Determines the master tkinter window, the title of the window, the resizable status of the window, and
        # the size and render location of the window.
        self.master = master
        self.master.title('Ballot Manifest')
        self.master.resizable(False, False)
        self.master.geometry('525x600+50+25')

        # Establish frame header settings
        # This sets the display for the top part of the window. It adds the Arapahoe County Logo and the Title Text
        self.frame_header = ttk.Frame(master)
        self.frame_header.pack()
        self.logo = PhotoImage(file='Data/ArapahoeLogo.png').subsample(20, 20)
        self.arapahoe_logo = ttk.Label(self.frame_header, image=self.logo).grid(row=0, column=0, rowspan=2,
                                                                                pady=(5, 25))
        self.title_text = ttk.Label(self.frame_header, font=('Arial', 18, 'bold'), wraplength=400,
                                    text='Ballot Manifest Data Entry').grid(row=1, column=1, padx=25)

        # Establish Load File Frame Settings
        # This is the landing page of the program that prompts for a filename stored under the self.filename variable.
        # The self.content_textbox variable is the text input where a filename can be loaded.
        self.frame_load = ttk.Frame(master)
        self.filename = None
        self.content_text = ttk.Label(self.frame_load, text='Select File to Begin').grid(row=0, column=0, columnspan=2,
                                                                                         sticky='w')
        self.content_textbox = ttk.Entry(self.frame_load, width=60)
        self.select_button = ttk.Button(self.frame_load, text='Browse Computer', command=lambda: self.locate_file())
        self.load_button = ttk.Button(self.frame_load, text='Load File',
                                      command=lambda: self.load_to_form())

        # Establish Form Settings
        # This page is the form that the user will enter the information on the container label.
        # The self.labels and self.batch_entries will be determined after a label is scanned.
        self.frame_form = ttk.Frame(master)
        self.scan_text = ttk.Label(self.frame_form, text='Scan Label: ')
        self.scan_label = ttk.Entry(self.frame_form, width=25, show='*')
        self.scan_submit = ttk.Button(self.frame_form, text='Submit', command=lambda: self.load_label())
        self.scanner = ttk.Label(self.frame_form, text='Scanner: ')
        self.container = ttk.Label(self.frame_form, text='Container: ')
        self.seal_text_1 = ttk.Label(self.frame_form, text='Seal 1: ')
        self.seal_text_2 = ttk.Label(self.frame_form, text='Seal 2: ')
        self.seal_entry_1 = ttk.Entry(self.frame_form, width=25)
        self.seal_entry_2 = ttk.Entry(self.frame_form, width=25)
        self.labels = []
        self.batch_entries = []
        self.form_save = ttk.Button(self.frame_form, text='Save', command=lambda: self.submit_form())
        self.form_clear = ttk.Button(self.frame_form, text='Clear', command=lambda: self.form_view())

        # Establish Menu Bar Settings
        # These are the options available under the "File" and "Help" drop down menus.
        self.master.option_add('*tearOff', False)
        self.menu_bar = Menu(master)
        self.master.config(menu=self.menu_bar)
        self.file = Menu(self.menu_bar)
        self.help_ = Menu(self.menu_bar)
        self.menu_bar.add_cascade(menu=self.file, label='File')
        self.menu_bar.add_cascade(menu=self.help_, label='Help')
        self.file.add_command(label='Create New From Template', command=lambda: self.load_from_template_view())
        self.file.add_command(label='Load Manifest', command=lambda: self.load_view())
        self.file.add_command(label='Edit or Remove Previous Entry', command=lambda: self.edit_or_remove_view())
        self.file.entryconfig('Edit or Remove Previous Entry', state='disabled')
        self.file.add_command(label='Save and Quit', command=lambda: self.save_and_quit())
        self.file.entryconfig('Save and Quit', state='disabled')
        self.help_.add_command(label='How to use this software', command=lambda: self.how_to_use_view())
        self.help_.add_separator()
        self.help_.add_command(label='About', command=lambda: self.about_view())

        # Establish Edit or Remove Settings
        # This is the for the edit or remove page. The self.scanner_entry and self.starting_batch_entry variables
        # are the text input locations that will be used for looking up an entry.
        self.frame_edit = ttk.Frame(master)
        self.instruction_text = ttk.Label(self.frame_edit, text='To Edit or Remove an entry: enter the scanner and'
                                                                'starting batch number, then select "Edit" or "Remove"',
                                          wraplength=250)
        self.scanner_prompt = ttk.Label(self.frame_edit, text='Scanner: ')
        self.starting_batch_prompt = ttk.Label(self.frame_edit, text='First Batch in Container: ')
        self.scanner_entry = ttk.Entry(self.frame_edit, width=25)
        self.starting_batch_entry = ttk.Entry(self.frame_edit, width=25)
        self.edit_button = ttk.Button(self.frame_edit, text='Edit', command=lambda: self.edit_entry())
        self.remove_button = ttk.Button(self.frame_edit, text='Remove', command=lambda: self.remove_entry())

        # Establish Load From New Template Settings
        # These are the settings for the load from a new template screen.
        self.frame_load_template = ttk.Frame(master)
        self.load_template_text = ttk.Label(self.frame_load_template, text='To create a new ballot manifest from the '
                                                                           'manifest template: click the "Create From '
                                                                           'Template" button below', wraplength=350)
        self.load_template_button = ttk.Button(self.frame_load_template, text='Create From Template',
                                               command=lambda: self.create_from_template())

        # Establish Pandas DataFrame Variables
        # These variables can be changed to change the county using the software and the column names of the ballot
        # manifest.
        self.county = 'Arapahoe'
        self.df = None
        self.column_names = ['County', 'Scanner', 'ICC Batch', 'Ballot Count', 'Container Number', 'Seal 1', 'Seal 2']
        self.printed_label = None

        # Render load View
        self.load_view()

    # Methods to render the Various views.

    # This method will clear the view on the app, except for the header view which is static.
    def clear_view(self):
        self.frame_load.pack_forget()
        self.frame_form.pack_forget()
        self.frame_edit.pack_forget()
        self.frame_load_template.pack_forget()

    # This view is the screen prompting to load a ballot manifest into the software.
    # The self.select_button will call the self.locate_file() method, and the self.load_button will call the
    # self.load_to_form() method.
    def load_view(self):
        self.clear_view()
        self.df = None
        self.filename = None
        self.update_menu_view()
        self.frame_load.pack()
        self.content_textbox.grid(row=1, column=0, columnspan=4, sticky='w', pady=5)
        self.content_textbox.config(state='normal')
        self.content_textbox.delete(0, END)
        self.content_textbox.config(state='disabled')
        self.select_button.grid(row=2, column=0, sticky='w', pady=10)
        self.load_button.grid(row=2, column=1)

    # This view loads the form to input data from completed labels. It calls the self.generate_batch_widgets() method
    # to load the correct ICC Station, Batch Numbers, and Container Number according to the label. The self.form_save
    # button will call the self.submit_form() method. The self.form_clear button will call the self.form_view() method,
    # effectively looping back to the base form screen without saving the information.
    def form_view(self):
        self.update_menu_view()
        self.frame_form.pack(fill=BOTH, expand=1)
        self.scan_text.grid(row=1, column=0, sticky='w', pady=15, padx=10)
        self.scan_label.grid(row=1, column=1, sticky='w', padx=10)
        self.scan_label.config(state='normal')
        self.scan_label.delete(0, END)
        self.scan_submit.grid(row=1, column=2)
        self.scan_submit.config(state='normal')
        self.scanner.grid(row=2, column=0, sticky='w', padx=10, columnspan=2)
        self.scanner.config(text='Scanner: ')
        self.container.grid(row=3, column=0, sticky='w', padx=10, columnspan=2)
        self.container.config(text='Container: ')
        self.seal_text_1.grid(row=4, column=2)
        self.seal_entry_1.grid(row=4, column=3)
        self.seal_entry_1.delete(0, END)
        self.seal_entry_1.config(state='disabled')
        self.seal_text_2.grid(row=5, column=2)
        self.seal_entry_2.grid(row=5, column=3)
        self.seal_entry_2.delete(0, END)
        self.seal_entry_2.config(state='disabled')
        self.form_save.grid(row=7, column=3)
        self.form_save.config(state='disabled')
        self.form_clear.grid(row=15, column=3)
        self.form_clear.config(state='disabled')
        self.generate_batch_widgets()

    # This view is for the edit or remove screen selected from the file drop down menu. It is only available once a
    # ballot manifest has been loaded. The self.edit_button will call the self.edit_entry() method and the
    # self.remove_button will call the self.remove_entry() method.
    def edit_or_remove_view(self):
        self.clear_view()
        self.frame_edit.pack(fill=BOTH, expand=1)
        self.instruction_text.grid(row=0, column=1, columnspan=2, pady=20)
        self.scanner_prompt.grid(row=1, column=1, sticky='e')
        self.scanner_entry.grid(row=1, column=2)
        self.starting_batch_prompt.grid(row=2, column=1, sticky='e')
        self.starting_batch_entry.grid(row=2, column=2)
        self.edit_button.grid(row=3, column=1)
        self.remove_button.grid(row=3, column=2)

    # This view is for creating a new ballot manifest from a a template. The self.load_template_button will call the
    # self.create_from_template() method.
    def load_from_template_view(self):
        self.clear_view()
        self.frame_load_template.pack()
        self.load_template_text.grid(row=1, column=0, pady=25)
        self.load_template_button.grid(row=2, column=0)

    # This is used to dynamically verify which option in the file drop down menu are active. It will only activate them
    # if there is an active data frame.
    def update_menu_view(self):
        if self.df is not None:
            self.file.entryconfig('Edit or Remove Previous Entry', state='normal')
            self.file.entryconfig('Save and Quit', state='normal')
        else:
            self.file.entryconfig('Edit or Remove Previous Entry', state='disabled')
            self.file.entryconfig('Save and Quit', state='disabled')

    # This view is an informational pop-up with basic information about the software.
    def about_view(self):
        messagebox.showinfo('About', 'Arapahoe County Ballot Manifest Data Entry Software.\nDeveloped by Jonathan '
                                     'Layman')

    # This view is a pop-up containing the data from a text file describing how to use the software.
    def how_to_use_view(self):
        with open('Data/How_to_use.txt', 'r') as txt_file:
            use_doc = txt_file.read()
        use_window = Toplevel(self.master)
        use_window.geometry('500x400+75+40')
        use_window.title('How to Use')
        use_text = ScrolledText(use_window, width=135, height=50, wrap='word')
        use_text.insert(1.0, use_doc)
        use_text.pack()
        use_text.config(state='disabled')

    # The following methods are used to read and manipulate the data entered into this software. Many pandas
    # data frame methods are called and used to sort and store data.

    # This method is called from the self.form_view() and the self.generate_form() methods. It
    # is used to dynamically label the batch numbers in the form according to their starting batch number. It starts by
    # removing any existing data from the relevant variables. It then creates a list of the 15 batches from the start
    # point it is given. The self.form_view() method does not provide a starting point, so a default starting point of 1
    # is used. The self.generate_form() method does provide a starting point which is determined from the label.
    def generate_batch_widgets(self, start=1):
        for label in self.labels:
            label.destroy()
        del self.labels[:]
        del self.batch_entries[:]
        for i in range(start, start + 15):
            self.labels.append(ttk.Label(self.frame_form, text='Batch ' + str(i) + ':'))
            self.batch_entries.append(ttk.Entry(self.frame_form, width=25))
        for index, label in enumerate(self.labels, start=4):
            label.grid(row=index, column=0, sticky='w', padx=10, pady=2)
        for index, entry in enumerate(self.batch_entries, start=4):
            entry.grid(row=index, column=1, sticky='w')
            entry.config(state='disabled')

    # This method is called by the self.scan_submit() method and the edit view screen. It takes data from a scanned
    # barcode (3 of 9 format) and splits the data. The raw data looks like: 01/L1/LICC 01-1
    # The split point is "/L". The first part of the data is the Scanner Number, The second part is the first batch in
    # the container, and the third part is the name of the container.
    #
    # After splitting the data, this method checks the data frame to make sure the length of the split data is 3 because
    # each barode has 3 parts. This will reduce the risk of scanning a wrong barcode. This method then checks to see if
    # #there is a matching entry, and if there is, it prompts the user to confirm that they want to overwrite the
    # exiting data. If the user answers "Yes" or if there is no matching data, it passes back to the
    # self.generate_form() or self.form_view() methods as appropriate.
    def load_label(self):
        self.printed_label = self.scan_label.get().split('/L')
        self.printed_label[0] = int(self.printed_label[0])
        self.printed_label[1] = int(self.printed_label[1])
        if len(self.printed_label) != 3:
            self.form_view()
        else:
            if (self.county, self.printed_label[0], self.printed_label[1]) in self.df.index:
                result = messagebox.askyesno('Entry Found in Manifest', 'This entry already exists in the manifest, '
                                                                        'would you like to overwrite it?')
                if result:
                    for i in range(15):
                        self.df.drop((self.county, self.printed_label[0], self.printed_label[1] + i), inplace=True)
                    self.generate_form()
                else:
                    self.form_view()
            else:
                self.generate_form()

    # This method is called from the self.load_label() method and is used to set the appropriate labels according to the
    # batch that is provided from the scanned label.
    def generate_form(self):
        self.scanner.config(text='Scanner:  ICC STATION ' + str(self.printed_label[0]))
        self.container.config(text='Container:  ' + self.printed_label[2])
        self.generate_batch_widgets(start=self.printed_label[1])
        self.scan_label.delete(0, END)
        self.scan_label.config(state='disabled')
        self.scan_submit.config(state='disabled')
        self.seal_entry_1.config(state='normal')
        self.seal_entry_2.config(state='normal')
        self.form_clear.config(state='normal')
        self.form_save.config(state='normal')
        for label in self.labels:
            label.config(state='normal')
        for entry in self.batch_entries:
            entry.config(state='normal')

    # This method is called from the self.select_button. It opens up a dialog to locate a file on the computer, and then
    # checks to see if that file is a .csv. If the File is a .csv it will save the filename under the self.filename
    # variable and load the contents of that file as the main data frame under the self.df variable.
    def locate_file(self):
        self.filename = filedialog.askopenfile()
        if self.filename is not None and self.filename.name[-3:] == 'csv':
            self.filename = self.filename.name
            self.df = pd.read_csv(self.filename, index_col=['County', 'Scanner', 'ICC Batch'])
            self.df.sort_index(inplace=True)
            self.content_textbox.config(state='normal')
            self.content_textbox.insert(0, self.filename)
            self.content_textbox.config(state='disabled')
        else:
            messagebox.showerror(title='Ballot Manifest - Error',
                                 message='Please select a CSV file')

    # This method is called by the self.create_from_template() method and the self.load_button. It checks that there is
    # content in the self.filename variable, and if there is, it clears the screen and displays the form.
    def load_to_form(self):
        if self.filename is not None:
            self.clear_view()
            self.form_view()

    # This method is called by the self.form_save button. It creates a new data frame in a local variable called
    # new_data. It uses list comprehension to populate the form based off of the information gathered from the form.
    # The method then appends the data to the main data frame, sorts the data, and then saves it back to the .csv. If
    # the .csv is open, than this method will be denied access to the file and not be able to save. However, the data
    # frame will remain intact.
    def submit_form(self):
        new_data = pd.DataFrame(columns=self.column_names)
        new_data['County'] = [self.county for i in range(15)]
        new_data['Scanner'] = [self.printed_label[0] for i in range(15)]
        new_data['ICC Batch'] = [self.printed_label[1] + i for i in range(15)]
        new_data['Ballot Count'] = [entry.get() for entry in self.batch_entries]
        new_data['Container Number'] = [self.printed_label[2] for i in range(15)]
        new_data['Seal 1'] = [self.seal_entry_1.get() for i in range(15)]
        new_data['Seal 2'] = [self.seal_entry_2.get() for i in range(15)]
        new_data.set_index(keys=['County', 'Scanner', 'ICC Batch'], inplace=True)
        self.df = self.df.append(new_data)
        self.df.sort_index(inplace=True)
        self.df.to_csv(self.filename)
        messagebox.showinfo('Ballot Manifest', 'Form successfully submitted')
        self.form_view()

    # This method is called from the file menu. It saves a copy of the main data frame to the .csv file. A copy is saved
    # every time the submit form button is pressed, so saving here is redundant. This method then kills the master
    # window.
    def save_and_quit(self):
        if self.df is not None:
            self.df.to_csv(self.filename)
        self.master.destroy()
        quit()

    # This method is called from the self.edit_button. It prompts the user to enter the scanner number and the starting
    # batch number. If the data entered does not match a possible entry, the software will show an error explaining how
    # to load a proper batch. If a data is a possible match, it will generate a fake bar code that is fed into the
    # self.load_label() method.
    def edit_entry(self):
        list_of_scanners = ['01', '02', '03', '04', '05', '06', '07', '08', '09' '10']
        if self.scanner_entry.get() in list_of_scanners and int(self.starting_batch_entry.get()) % 15 == 1:
            scanner = self.scanner_entry.get()
            batch = self.starting_batch_entry.get()
            container = 'ICC ' + scanner + '-' + str(int((int(batch) + 14) / 15))
            self.clear_view()
            self.form_view()
            self.scan_label.insert(0, scanner + '/L' + batch + '/L' + container)
            self.scanner_entry.delete(0, END)
            self.starting_batch_entry.delete(0, END)
            self.load_label()
        else:
            messagebox.showerror('Data Mismatch', 'The data entered does not match a valid entry. Make sure that the '
                                                  'Scanner is entered with two numbers (ie. 05 or 10) and the batch is '
                                                  'a valid starting point (is the start of a batch of 15 ie. 31)')

    # This method is called by the self.remove_button. It matches the data similar to the self.edit_entry() method.
    # If a match is found it will prompt the user if they want to remove the entry, or it will inform the user that
    # There is no such entry (even if the search criteria is valid)
    def remove_entry(self):
        list_of_scanners = ['01', '02', '03', '04', '05', '06', '07', '08', '09' '10']
        print(self.df)
        print('Check 1')
        if self.scanner_entry.get() in list_of_scanners and int(self.starting_batch_entry.get()) % 15 == 1:
            if (self.county, int(self.scanner_entry.get()), int(self.starting_batch_entry.get())) in self.df.index:
                result = messagebox.askyesno('Entry Found in Manifest', 'Are you sure you want to remove this entry?')
                if result:
                    scanner = int(self.scanner_entry.get())
                    batch = int(self.starting_batch_entry.get())
                    for i in range(15):
                        self.df.drop((self.county, scanner, batch + i), inplace=True)
                    print(self.df)
                    print(self.filename)
                    self.df.to_csv(self.filename)
                    self.scanner_entry.delete(0, END)
                    self.starting_batch_entry.delete(0, END)
                    self.clear_view()
                    self.form_view()
            else:
                messagebox.showerror('No Entry Found',
                                     'Your submission does not match any existing entry in the manifest')

    # This method is called from the create new from template option in the file menu. It prompts the user to select
    # a location and provides a default name of "Ballot Manifest.csv". Once a file location and name is selected it will
    # create a new data frame from the column list outlined in self.column_names and exports that to a csv. Finally,
    # this method loads the newly created template into the entry form.
    def create_from_template(self):
        directory = filedialog.asksaveasfilename(title='Select a Location and Name for the Ballot Manifest',
                                                 filetypes=[('csv files', '*.csv')],
                                                 defaultextension='.csv',
                                                 initialfile='Ballot Manifest')
        new_from_template = pd.DataFrame(columns=self.column_names)
        new_from_template.to_csv(directory, index=False)
        self.df = new_from_template.set_index(keys=['County', 'Scanner', 'ICC Batch'])
        self.filename = directory
        self.load_to_form()


# This function launches the program by creating a tkinter root, and then calling the BallotManifestGui class.
def main():
    root = Tk()
    manifest_gui = BallotManifestGui(root)
    root.mainloop()


if __name__ == '__main__': main()
