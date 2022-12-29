# Importing required library
import pygsheets
  
def openfile(client, filename, sheetname):
    # opens a spreadsheet by its name/title
    spreadsht = client.open(filename)
    
    # opens a worksheet by its name/title
    worksht = spreadsht.worksheet("title", sheetname)
    return worksht

def lookup(worksht, name):
    cell = worksht.find(name, cols = (1,3))
    if len(cell) == 0:
        return None
    result = ''.join([i for i in str(cell[0]) if i.isnumeric()])
    return worksht.get_row(int(result) + 1)


if __name__ == "__main__":
    # Create the Client
    client = pygsheets.authorize(service_account_file="timesheet-search-and-fill-446d34b4d7c0.json")

    mainsheet = openfile(client, "[VN] Timesheet_Dec2022", "HCM Office")
    subsheet = openfile(client, "01-16.12.2022", "FTA Report")

    # print(lookup(subsheet, "Nguyen Thi Ha My"))
    allIndexs = mainsheet.get_col(1)
    for i in allIndexs:
        if i.isnumeric():
            rowData = mainsheet.get_row(int(i) + 2)
            data = lookup(subsheet, rowData[1])
            if data != None:
                alltime = data[4:]
                mainsheet.update_row(int(i) + 2, alltime, 5)
            else:
                print(f"not found ${rowData[1]}")