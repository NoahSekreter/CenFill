import smartsheet
import functions
from tkinter import *
import tkinter.filedialog as tk


class Company:
    def __init__(self, sheet_id, name):
        self.id = sheet_id
        self.name = name.split(" Enrollment", 1)[0]
        self.carrier = None
        self.market = None
        self.month = None
        self.year = None

    # Creates all the required spreadsheets
    def create_spreadsheets(self, client, output):
        base_loc = tk.askdirectory()
        if base_loc is '':
            return
        # status.set("Generating census fillers...")
        output.insert(tk.END, "Generating census filler...\n")
        new_loc = functions.generate_spreadsheet(base_loc, functions.generate_client_matrix(client, self.id), self.name)
        if self.carrier and self.market is not None or "N/A":
            if "100+" in self.market:
                if "Cigna" not in self.carrier:
                    functions.generate_carrier_spreadsheets(new_loc, '4')
                    output.insert(tk.END, "Generating Cigna...\n")
                if "Aetna" not in self.carrier:
                    functions.generate_carrier_spreadsheets(new_loc, '7')
                    output.insert(tk.END, "Generating Aetna CAT 100+...\n")
                    functions.generate_carrier_spreadsheets(new_loc, '8')
                    output.insert(tk.END, "Generating Aetna Tier 100+...\n")
            elif "51-99" in self.market:
                if "UHC" not in self.carrier:
                    functions.generate_carrier_spreadsheets(new_loc, '3')
                    output.insert(tk.END, "Generating UHC...\n")
                if "Cigna" not in self.carrier:
                    functions.generate_carrier_spreadsheets(new_loc, '4')
                    output.insert(tk.END, "Generating Cigna...\n")
                if "Aetna" not in self.carrier:
                    functions.generate_carrier_spreadsheets(new_loc, '6')
                    output.insert(tk.END, "Generating Aetna 51-99...\n")
            elif "-50" in self.market:
                if "Anthem" not in self.carrier:
                    functions.generate_carrier_spreadsheets(new_loc, '1')
                    output.insert(tk.END, "Generating Anthem MEWA...\n")
                if "MMO" not in self.carrier:
                    functions.generate_carrier_spreadsheets(new_loc, '2')
                    output.insert(tk.END, "Generating MMO...\n")
                if "UHC" not in self.carrier:
                    functions.generate_carrier_spreadsheets(new_loc, '3')
                    output.insert(tk.END, "Generating UHC...\n")
                if "SummaCare" not in self.carrier:
                    functions.generate_carrier_spreadsheets(new_loc, '5')
                    output.insert(tk.END, "Generating SummaCare...\n")
            output.insert(tk.END, "All censuses generated!\n")
            # status.set("Censuses generated!")

    # Assigns the carrier and month variable with the company's carrier and renewal month
    def find_carrier_and_renewal(self, client):
        search_results = client.Search.search(self.name + " - Medical")

        if 'row' in search_results.results[0].object_type:
            med_sheet = client.Sheets.get_sheet(search_results.results[0].parent_object_id)
        else:
            med_sheet = client.Sheets.get_sheet(search_results.results[0].object_id)

        self.month = str(med_sheet.rows[1].cells[2].value) if med_sheet.rows[1].cells[2].value is not None else "N/A"
        self.year = str(med_sheet.rows[1].cells[3].value) if med_sheet.rows[1].cells[3].value is not None else "N/A"
        self.carrier = str(med_sheet.rows[1].cells[5].value) if med_sheet.rows[1].cells[5].value is not None else "N/A"
        self.market = str(med_sheet.rows[1].cells[6].value) if med_sheet.rows[1].cells[6].value is not None else "N/A"


class Interface:
    Excel_File = ""
    Autofill_Button = None
    Template_Button = None
    TNF_Button = None
    ICHRA_Button = None
    Excel_Name = None
    Comp_Name = None
    Comp_Carrier = None
    Comp_Market = None
    Comp_Month = None
    Comp_Year = None
    Smartsheet_Census_Button = None
    Output = None

    def __init__(self, master):
        home_check = open('C:/Program Files/CenFill/version.txt', 'r')
        current_version = str(home_check.readline())
        master.title('BBG CenFill - Ver ' + current_version)
        master.iconbitmap(default='icon.ico')

        # master.title('BBG CenFill - Ver 2.0')
        # Prepare company variables
        self.Comp_Name = StringVar()
        self.Comp_Carrier = StringVar()
        self.Comp_Market = StringVar()
        self.Comp_Month = StringVar()
        self.Comp_Year = StringVar()

        e_frame = Frame(master, padx=4)
        e_frame.grid(row=0, column=0)
        s_frame = Frame(master, padx=4, bg='#131369', pady=6)
        s_frame.grid(row=0, column=1)

        Label(e_frame, text='BBG CenFill', font=('Helvetica', 20, "bold")).grid(row=0, column=0, columnspan=2)
        description = Label(e_frame, text='Welcome to CenFill, BBG\'s automation multi-tool! The functionality of this '
                                          'program includes auto-completing census fillers, creating carrier templates '
                                          'from complete census fillers, generating FT Censuses from 1095\'s, and '
                                          'auto-completing ICHRA censuses. In addition, this program comes with '
                                          'Smartsheet integration, allowing you to create every census needed for '
                                          'groups with an enrollment/term sheet on Smartsheets.', wraplength=380)
        description.grid(row=1, column=0, columnspan=2)
        self.Excel_Name = Label(e_frame, text='Current File: None', width=54, bg='#dadada')
        self.Excel_Name.grid(row=2, column=0, columnspan=2, pady=2)
        Button(e_frame, text="Open Excel Spreadsheet", width=54,
               command=lambda: self.select_excel_file()).grid(row=3, column=0, columnspan=2)
        self.Autofill_Button = Button(e_frame, text="Autofill Remaining Info", width=26, state=DISABLED,
                                      command=lambda: self.autofill_window())
        self.Template_Button = Button(e_frame, text="Copy onto Carrier Template", width=26, state=DISABLED,
                                      command=lambda: self.template_window())
        self.TNF_Button = Button(e_frame, text="Create 1095", width=26, state=DISABLED,
                                 command=lambda: self.tnf_window())
        self.ICHRA_Button = Button(e_frame, text="Create ICHRA", width=26, state=DISABLED,
                                   command=lambda: self.ichra_window())
        self.Output = Text(e_frame, width=47, pady=4, height=12, bg='white')
        self.Autofill_Button.grid(row=4, column=0, pady=1)
        self.Template_Button.grid(row=4, column=1, pady=1)
        self.TNF_Button.grid(row=5, column=0, pady=1)
        self.ICHRA_Button.grid(row=5, column=1, pady=1)
        self.Output.grid(row=6, column=0, columnspan=2, pady=(2, 8))

        Label(s_frame, text='Smartsheets', font=('Helvetica', 20, 'bold'), fg='white', bg='#131369').grid(
            row=0, column=0)
        Label(s_frame, textvariable=self.Comp_Name, font=('Helvetica', 12, 'bold'), fg='white', bg='#131369').grid(
            row=1, column=0)
        Label(s_frame, textvariable=self.Comp_Carrier, font=('Helvetica', 9, 'bold'), bg='#131369', fg='white').grid(
            row=2, column=0)
        Label(s_frame, textvariable=self.Comp_Market, font=('Helvetica', 9, 'bold'), bg='#131369', fg='white').grid(
            row=3, column=0)
        Label(s_frame, textvariable=self.Comp_Month, font=('Helvetica', 9, 'bold'), bg='#131369', fg='white').grid(
            row=4, column=0)
        Label(s_frame, textvariable=self.Comp_Year, font=('Helvetica', 9, 'bold'), bg='#131369', fg='white').grid(
            row=5, column=0)

        list_frame = Frame(s_frame, relief=GROOVE, pady=3, width=330, height=262, bd=1)
        list_frame.grid(row=7, column=0, pady=(4, 2))

        self.Smartsheet_Census_Button = Button(
            s_frame, text='Retrieve Smartsheets Data', command=lambda: self.populate_enrollments(list_frame),
            font=('Helvetica', 9, 'bold'), width=46)
        self.Smartsheet_Census_Button.grid(row=6, column=0, padx=2)

        self.Comp_Name.set("Company: N/A")
        self.Comp_Carrier.set("Carrier: N/A")
        self.Comp_Market.set("Market Size: N/A")
        self.Comp_Month.set("Renewal Month: N/A")
        self.Comp_Year.set("Renewal Year: N/A")

        self.Output.insert(tk.END, 'Program started successfully!\n')

    # Generate a list of companies to populate a scrollable frame
    def populate_enrollments(self, sub_frame):
        print("\nPlease wait a moment while the program retrieves all available enrollment sheets")
        client = smartsheet.Smartsheet('wwau865vx80lma53o0gubkh3ht')  # API-TEST Access Token
        client.errors_as_exceptions(True)
        workspace = client.Workspaces.get_workspace('4570508340553604')  # Enrollments Workspace

        # Create subframe for scroll area
        canvas = Canvas(sub_frame)
        frame_content = Frame(canvas)
        scrollbar = Scrollbar(sub_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        def canvas_event(event):
            canvas.configure(scrollregion=canvas.bbox("all"), width=309, height=262)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.create_window((0, 0), window=frame_content, anchor='nw')
        frame_content.bind("<Configure>", canvas_event)

        # Create a list of all companies with enrollment sheets
        enrollments = list()
        i = 1
        for sheet in workspace.sheets:
            if "advantage diagnostics" not in str(client.Sheets.get_sheet(sheet.id).name).lower():
                print("Retrieving clients... [ " + str(i) + " / " + str(len(workspace.sheets)) + " ]")
                enrollments.append(Company(sheet.id, client.Sheets.get_sheet(sheet.id).name))
            else:
                print("Advantage Diagnostics sheet found. Skipping...")
            # self.Output.insert(tk.END, "Retrieving clients [ " + str(i) + " / " + str(len(workspace.sheets)) + " ]\n")
            i += 1
        self.Output.insert(tk.END, "All clients have been retrieved\n")
        enrollments.sort(key=lambda x: x.name)
        for e in range(0, len(enrollments)):
            Button(frame_content, text=enrollments[e].name, width=43,
                   command=lambda e=e: self.select_company(enrollments[e], client)).pack()
        self.Smartsheet_Census_Button.configure(text="Create Censuses from Smartsheets")
        self.Smartsheet_Census_Button.configure(state=DISABLED)

    # Display company name and allow for user to create censuses
    def select_company(self, company, client):
        company.find_carrier_and_renewal(client)
        self.Comp_Name.set(company.name)
        self.Comp_Carrier.set("Carrier: " + company.carrier)
        self.Comp_Market.set("Market Size: " + company.market)
        self.Comp_Month.set("Renewal Month: " + company.month)
        self.Comp_Year.set("Renewal Year: " + company.year)
        self.Smartsheet_Census_Button['state'] = 'normal'
        self.Smartsheet_Census_Button['command'] = lambda: company.create_spreadsheets(client, self.Output)

    # Select an excel spreadsheet
    def select_excel_file(self):
        self.Excel_File = tk.askopenfilename()
        if ".xlsx" not in self.Excel_File:
            if self.Excel_File is not "":
                self.Output.insert(tk.END, "Error: Select an '.xlsx' file\n")
            self.Excel_File = ""
            self.Autofill_Button['state'] = 'disabled'
            self.Template_Button['state'] = 'disabled'
            self.TNF_Button['state'] = 'disabled'
            self.ICHRA_Button['state'] = 'disabled'
            self.Excel_Name['text'] = 'Current File: None'
        elif self.Excel_File is not "":
            file_type = functions.check_input_file(self.Excel_File)
            if file_type is 1:
                self.Output.insert(tk.END, "Census filler detected\n")
                self.Autofill_Button['state'] = 'normal'
                self.Template_Button['state'] = 'normal'
                self.TNF_Button['state'] = 'disabled'
                self.ICHRA_Button['state'] = 'disabled'
            elif file_type is 2:
                self.Output.insert(tk.END, "1095 detected\n")
                self.TNF_Button['state'] = 'normal'
                self.ICHRA_Button['state'] = 'disabled'
                self.Autofill_Button['state'] = 'disabled'
                self.Template_Button['state'] = 'disabled'
            elif file_type is 3:
                self.Output.insert(tk.END, "ICHRA detected\n")
                self.ICHRA_Button['state'] = 'normal'
                self.TNF_Button['state'] = 'disabled'
                self.Autofill_Button['state'] = 'disabled'
                self.Template_Button['state'] = 'disabled'
            else:
                self.TNF_Button['state'] = 'disabled'
                self.ICHRA_Button['state'] = 'disabled'
                self.Autofill_Button['state'] = 'disabled'
                self.Template_Button['state'] = 'disabled'
            self.Excel_Name['text'] = 'Current File: ' + functions.basename(self.Excel_File)

    # Open up the window for the autofill option
    def autofill_window(self):
        win = Toplevel()
        win.resizable(width=False, height=False)
        win.wm_title('')
        Label(win, text='Include child count?').grid(row=0, column=0, columnspan=2)
        Button(win, text='Yes', width=12, command=lambda: fill_and_destroy(True)).grid(row=1, column=0)
        Button(win, text='No', width=12, command=lambda: fill_and_destroy(False)).grid(row=1, column=1)

        # Fill out census based on button pressed and remove the pop up window
        def fill_and_destroy(answer):
            win.destroy()
            functions.auto_fill(self.Excel_File, answer)
            self.Output.insert(tk.END, 'Census has been completed!\n')

    def template_window(self):
        win = Toplevel()
        win.resizable(width=False, height=False)
        win.wm_title('')
        Label(win, text='Which template would you like to copy info onto?').grid(row=0, column=0, columnspan=2)
        Button(win, text='Anthem MEWA', width=20, command=lambda: template_fill('1')).grid(row=1, column=0)
        Button(win, text='Anthem ACA', width=20, command=lambda: template_fill('9')).grid(row=1, column=1)
        Button(win, text='MMO', width=20, command=lambda: template_fill('2')).grid(row=2, column=0)
        Button(win, text='UHC GRX', width=20, command=lambda: template_fill('3')).grid(row=2, column=1)
        Button(win, text='Cigna GRX', width=20, command=lambda: template_fill('4')).grid(row=3, column=0)
        Button(win, text='SummaCare', width=20, command=lambda: template_fill('5')).grid(row=3, column=1)
        Button(win, text='Aetna 51-99', width=20, command=lambda: template_fill('6')).grid(row=4, column=0)
        Button(win, text='Aetna CAT 100+', width=20, command=lambda: template_fill('7')).grid(row=4, column=1)
        Button(win, text='Aetna Tier 100+', width=20, command=lambda: template_fill('8')).grid(row=5, column=0)

        # Fill out census based on button pressed and remove the pop up window
        def template_fill(answer):
            win.destroy()
            functions.generate_carrier_spreadsheets(self.Excel_File, answer)
            carrier_answer = {'1': 'Anthem MEWA',
                              '2': 'MMO',
                              '3': 'UHC GRX',
                              '4': 'Cigna GRX',
                              '5': 'SummaCare',
                              '6': 'Aetna 51-99',
                              '7': 'Aetna CAT 100+',
                              '8': 'Aetna Tier 100+',
                              '9': 'Anthem ACA'}
            self.Output.insert(tk.END, carrier_answer[answer] + ' spreadsheet has been generated!\n')

    def tnf_window(self):
        win = Toplevel()
        win.resizable(width=False, height=False)
        win.wm_title('')
        entry = Entry(win, width=6)
        Label(win, text='Enter the wait period in months: ').grid(row=0, column=0)
        entry.grid(row=0, column=1)
        Button(win, text='Submit', width=29, command=lambda: fill_and_destroy(entry.get())).grid(row=1, columnspan=2)

        def fill_and_destroy(wait):
            win.destroy()
            try:
                if wait is None:
                    wait = 0
                int(wait)
                self.Output.insert(tk.END, "Creating 1095...\n")
                functions.create_ft_census(self.Excel_File, int(wait))
                self.Output.insert(tk.END, "1095 has been generated!\n")
            except(PermissionError, Exception):
                self.Output.insert(tk.END, "Error: Please close the selected file and try again\n")
            except(ValueError, Exception):
                self.Output.insert(tk.END, "An error has occurred\n")

    def ichra_window(self):
        win = Toplevel()
        win.resizable(width=False, height=False)
        win.wm_title('')
        entry = Entry(win, width=7)
        m1_frame = Frame(win)
        m2_frame = Frame(win)
        Label(win, text='Enter the year you want to retrieve quotes for: ').grid(row=0, column=0, pady=2)
        Label(win, text='Select the metal levels of plans you want to retrieve').grid(row=1, column=0, columnspan=2)
        entry.grid(row=0, column=1, pady=2)
        m1_frame.grid(row=2, columnspan=2)
        m2_frame.grid(row=3, columnspan=2)
        Button(win, text='Submit', width=40, command=lambda: fill_and_destroy(entry.get())).grid(
            row=4, columnspan=2, pady=2)

        # Note to self: Find a more efficient way of doing this
        var1 = tk.IntVar()
        var2 = tk.IntVar()
        var3 = tk.IntVar()
        var4 = tk.IntVar()
        var5 = tk.IntVar()
        check_bronze = tk.Checkbutton(m1_frame, text='Bronze', variable=var1, onvalue=1, offvalue=0)
        check_silver = tk.Checkbutton(m1_frame, text='Silver', variable=var2, onvalue=1, offvalue=0)
        check_gold = tk.Checkbutton(m1_frame, text='Gold', variable=var3, onvalue=1, offvalue=0)
        check_platinum = tk.Checkbutton(m2_frame, text='Platinum', variable=var4, onvalue=1, offvalue=0)
        check_cat = tk.Checkbutton(m2_frame, text='Catastrophic', variable=var5, onvalue=1, offvalue=0)
        check_bronze.grid(row=0, column=0, padx=2)
        check_silver.grid(row=0, column=1, padx=2)
        check_gold.grid(row=0, column=2, padx=2)
        check_platinum.grid(row=0, column=0, padx=2)
        check_cat.grid(row=0, column=1, padx=2)

        def fill_and_destroy(year):
            win.destroy()
            # try:
            if year is "":
                year = functions.get_current_year()
            # Find a more efficient way of doing getting the metal levels
            metal_levels = []
            if var1.get():
                metal_levels.append("Bronze")
            if var2.get():
                metal_levels.append("Silver")
            if var3.get():
                metal_levels.append("Gold")
            if var4.get():
                metal_levels.append("Platinum")
            if var5.get():
                metal_levels.append("Catastrophic")
            self.Output.insert(tk.END, "Creating ICHRA...\n")
            # Create the plans and input all the data
            # We won't need to specify which plans would be needed or not. Just input all of them
            functions.create_ichra_plans(self.Excel_File, year, metal_levels)
            self.Output.insert(tk.END, "ICHRA has been generated!\n")
            # except(PermissionError, Exception):
            #     self.Output.insert(tk.END, "Error: Please close the selected file and try again\n")
            # except(ValueError, Exception):
            #     self.Output.insert(tk.END, "An error has occurred\n")


root = Tk()
root.resizable(width=False, height=False)
app = Interface(root)
root.mainloop()
sys.exit()
