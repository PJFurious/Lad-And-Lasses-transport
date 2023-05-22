import requests
from bs4 import BeautifulSoup
from tkinter import messagebox, filedialog, simpledialog
import win32com.client as win32
import os

# Declaring paths to set files that should exist, if not it will be created
vehicles_path = "C:/ProgramData/LaddAndLasses/vehicle.txt"
spreadsheet_path = "C:/ProgramData/LaddAndLasses/spreadsheet.txt"

def ScrapeData(year, mon, reg):
    # Set arrays
    from_mine_a = []
    loading_for_a = []
    arrival_date_a = []
    empty_weight_a = []
    load_weight_a = []
    load_gross_weight_a = []
    km_at_mine_a = []
    dep_date_a = []
    offload_at_a = []
    offload_weight_a = []
    offload_gross_weight_a = []
    rand_per_ton_a = []
    trip_monetry_value_a = []
    km_at_destination_a = []
    trip_hours_a = []
    rand_value_per_hour_a = []
    driver_name_a = []
    invoice_no_a = []

    # URL of the website to scrape
    url = "https://www.steelhorse.co.za/0430Report_trip_TruckMonthly.php"

    if mon < 10:
        mon = "0" + str(mon)
    else: mon = str(mon)

    # Form data to submit with the POST request
    form_data = {
        'var_repcompany': reg,
        'var_month': mon,
        'var_year': year
    }

    # Send POST request to the URL with the form data
    response = requests.post(url, data=form_data)

    # Check if the request was successful
    if response.status_code == 200:
        # Parse the HTML content using BeautifulSoup
        soup = BeautifulSoup(response.content, 'html.parser')

        # Find the table headers
        table_headers = soup.find_all('th')
        table_headers = table_headers[:len(table_headers)//2]

        # Extract the data for the truck registration numbers and months
        rows = soup.find_all('tr')
        no_data = True
        for row in rows:
            cells = row.find_all('td')
            if len(cells) > 1:  
                # If there is data save it into the an array
                from_mine = cells[3].text.strip()
                loading_for = cells[4].text.strip()
                arrival_date = cells[6].text.strip()
                try:    empty_weight = float(cells[7].text.strip())
                except: empty_weight = cells[7].text.strip()
                try:    load_weight = float(cells[8].text.strip())
                except: load_weight = cells[8].text.strip()
                try:    load_gross_weight = float(cells[9].text.strip())
                except: load_gross_weight = cells[9].text.strip()
                try:    km_at_mine = float(cells[10].text.strip())
                except: km_at_mine = cells[10].text.strip()
                dep_date = cells[12].text.strip()
                offload_at = cells[13].text.strip()
                try:    offload_weight = float(cells[14].text.strip())
                except: offload_weight = cells[14].text.strip()
                try:    offload_gross_weight = float(cells[15].text.strip())
                except: offload_gross_weight = cells[15].text.strip()
                try:    rand_per_ton = float(cells[16].text.strip())
                except: rand_per_ton = cells[16].text.strip()
                try:    trip_monetry_value = float(cells[18].text.strip())
                except: trip_monetry_value = cells[18].text.strip()
                try:    km_at_destination = float(cells[19].text.strip())
                except: km_at_destination = cells[19].text.strip()
                try:    trip_hours = float(cells[20].text.strip())
                except: trip_hours = cells[20].text.strip()
                try:    rand_value_per_hour = float(cells[21].text.strip())
                except: rand_value_per_hour = cells[21].text.strip()
                driver_name = cells[2].text.strip()
                try:    invoice_no = int(cells[23].text.strip())
                except: invoice_no = cells[23].text.strip()
                
                from_mine_a.insert(0, from_mine)
                loading_for_a.insert(0, loading_for)
                arrival_date_a.insert(0, arrival_date)
                empty_weight_a.insert(0, empty_weight)
                load_weight_a.insert(0, load_weight)
                load_gross_weight_a.insert(0, load_gross_weight)
                km_at_mine_a.insert(0, km_at_mine)
                dep_date_a.insert(0, dep_date)
                offload_at_a.insert(0, offload_at)
                offload_weight_a.insert(0, offload_weight)
                offload_gross_weight_a.insert(0, offload_gross_weight)
                rand_per_ton_a.insert(0, rand_per_ton)
                trip_monetry_value_a.insert(0, trip_monetry_value)
                km_at_destination_a.insert(0, km_at_destination)
                trip_hours_a.insert(0, trip_hours)
                rand_value_per_hour_a.insert(0, rand_value_per_hour)
                driver_name_a.insert(0, driver_name)
                invoice_no_a.insert(0, invoice_no)

                no_data = False
        
        if no_data:
            # messagebox.showerror("Error", f"Registration ({reg}) not found")
            if messagebox.askquestion("Warning", f"Registration ({reg}) not found.\nDo you want to add the registration?") == "yes":
                load_SpreadSheet(reg, from_mine_a, loading_for_a, arrival_date_a, empty_weight_a, load_weight_a, load_gross_weight_a, km_at_mine_a, 
                         dep_date_a, offload_at_a, offload_weight_a, offload_gross_weight_a, rand_per_ton_a, trip_monetry_value_a,
                         km_at_destination_a, trip_hours_a, rand_value_per_hour_a, driver_name_a, invoice_no_a)
            return
        else:
            load_SpreadSheet(reg, from_mine_a, loading_for_a, arrival_date_a, empty_weight_a, load_weight_a, load_gross_weight_a, km_at_mine_a, 
                         dep_date_a, offload_at_a, offload_weight_a, offload_gross_weight_a, rand_per_ton_a, trip_monetry_value_a,
                         km_at_destination_a, trip_hours_a, rand_value_per_hour_a, driver_name_a, invoice_no_a)
    else:
        print("Failed to retrieve data. Status Code:", response.status_code)


def get_Path_sheet():
    # get spreadsheet path that has been saved before
    try:
        with open(spreadsheet_path, "r") as f:
            path = f.readline().strip()
            sheet = f.readline().strip()
    except:
        messagebox.showerror("Error", "Please select a excel Spreadsheet")

    return path, sheet


def load_SpreadSheet(reg, a, b, c, d, e, f, g, h, j, k, l, m, n, o, p , q, r, s):
    win32.pythoncom.CoInitialize()
    # Save data into spreadsheet
    try:
        # Get path and sheet
        path, sheet = get_Path_sheet()
        
        # Open excel excel
        xlApp = win32.Dispatch("Excel.Application")
        xlApp.DisplayAlerts = False
        xlApp.Visible = True
        # Specify which spreadsheet and sheet to open
        workbook = xlApp.Workbooks.Open(path, False, False) # the third parameter indicates whether to try to open the file as read-only
        worksheet = workbook.Worksheets(sheet)

        # Insert data into spreadsheet
        worksheet.Rows(4).Insert()
        for x in range(len(a)):
            i = len(a)-(x+1)
            worksheet.Rows(4).Insert()
            new_row_data = [(i+1), a[i], b[i], c[i], d[i], e[i], f[i], g[i], "", h[i], j[i], k[i], l[i], m[i], n[i], o[i], p[i] , q[i], r[i], s[i]]

            for count, data in enumerate(new_row_data):
                worksheet.Cells(5, count+2).Value = data
                cell_to_color = worksheet.Cells(5, 2)
                cell_to_color.Interior.Color = 0xCEC7ff

        # Basic formatting to what was wanted
        worksheet.Rows(4).Insert()
        worksheet.Cells(5, 3).Value = reg

        cell_to_color = worksheet.Cells(5, 3)
        cell_to_color.Interior.Color = 0xC6EFCE
        cell_to_color.Font.Color = 0x006100

        workbook.Save()
    except Exception as e:
        messagebox.showerror("Error: ", e)
        print(e)
    win32.pythoncom.CoUninitialize()


def add_Heading(year, mon):
    win32.pythoncom.CoInitialize()
    # Allows to save the heading of the year and month when all vehicle data have been pulled
    path, sheet = get_Path_sheet()

    xlApp = win32.Dispatch("Excel.Application")
    xlApp.DisplayAlerts = False
    workbook = xlApp.Workbooks.Open(path, False, False) # the third parameter indicates whether to try to open the file as read-only

    worksheet = workbook.Worksheets(sheet)

    worksheet.Rows(4).Insert()

    worksheet.Cells(5, 2).Value = int(year)
    worksheet.Cells(5, 3).Value = mon

    cell_to_color = worksheet.Cells(5, 3)
    cell_to_color.Interior.Color = 0xA5A5A5
    cell_to_color.Font.Bold = True
    cell_to_color.Font.Color = 0xFFFFFF
    cell_to_color.Borders.LineStyle = 9

    workbook.Save()
    win32.pythoncom.CoUninitialize()


def get_name():
    # Get the sheet name where data should be saved
    name = simpledialog.askstring("Name Input", "Enter the sheet that should be used:")
    if name:
        return name
    else:
        messagebox.showerror("Error", "Please enter a name.")


def set_path():
    # Save path to the spreadsheet into a txt file
    file_path = filedialog.askopenfilename()
    sheet = get_name()
    if not os.path.exists("C:/ProgramData/LaddAndLasses/"):
        os.makedirs("C:/ProgramData/LaddAndLasses/")
    with open(spreadsheet_path, "w") as f:
        f.write(file_path + "\n" + sheet)
        

def Open_existing_Sheet():
    # Opens spreadsheet after it has been saved into a txt
    try:
        try:
            path, sheet = get_Path_sheet()
        except:
            return

        xlApp = win32.Dispatch("Excel.Application")
        xlApp.DisplayAlerts = False
        workbook = xlApp.Workbooks.Open(path,  False, False)
        xlApp.Visible = True 

        workbook.Worksheets(sheet)
    except:
        messagebox.showerror("Error", "Please select excel spreadsheet or ensure the password was correct.")


def add_vehicle(reg):
    # Save vehicle registration for future use
    if not os.path.exists("C:/ProgramData/LaddAndLasses/"):
        os.makedirs("C:/ProgramData/LaddAndLasses/")    
    with open(vehicles_path, "a") as f:
        f.write(reg + "\n")


def read_vehicles():
    # Extract saved vehicle registrations from txt
    try:
        reg = []
        with open(vehicles_path, 'r') as file:
            lines = file.readlines()

        for line in lines:
            reg.append(line.strip())
    except:
        pass

    return reg



def delete_vehicle(reg):
    # Remove saved vehicle registration
    with open(vehicles_path, "r") as file:
        lines = file.readlines()

    with open(vehicles_path, "w") as file:
        for line in lines:
            if reg not in line:
                file.write(line)


def submit_all(reg, month, year, months):
    for reg in read_vehicles():
        ScrapeData(year, months, reg.strip())
    add_Heading(year, month)