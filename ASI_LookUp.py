import json
import pandas as pd
import argparse
import sys
import win32com.client as win32
import xlwings as xw

def lstConvert(lst):

    #l=len(lst)
    col=lst.keys()[0]
    mylst=[]
    i = 0
    while i<len(lst):
        id=df['id'][i]
        #print(id)
        j = 0
        while j<len(lst[col][i]):
            mydist = {}
            #print(lst[col][i][j] +" - "+ id)
            mydist['id']=id
            mydist[col]=lst[col][i][j]
            mylst.append(mydist)
            j+=1
        i+=1
    return(mylst)

def xlstbl(data_frame,XLsheet_name,TableName):
    if data_frame.empty:
        #print(data_frame)
        data_frame = pd.DataFrame({'col1': ['NA'], 'col2': ['NA']})

    data_frame.to_excel(writer, sheet_name=XLsheet_name, index=False)
    workbook = writer.book
    worksheet = writer.sheets[XLsheet_name]
    # Get the dimensions of the dataframe.
    (max_row, max_col) = data_frame.shape

    # Create a list of column headers, to use in add_table().

    column_settings = []
    for header in data_frame.columns:
        column_settings.append({'header': header})

    # Add the table.
    worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings, "name": TableName})
    # Make the columns wider for clarity.
    worksheet.set_column(0, max_col - 1, 12)


def Index(workbook_path):
    print("[*] Indexing excel sheets")
    try:
        # Open the existing Excel file
        # Replace with the path to your file
        workbook = xw.Book(workbook_path)
        sheet_names_list = [sheet.name for sheet in workbook.sheets]

        # print(a)
        # workbook.sheets['Index'].delete()

        # Check if 'Index' sheet exists and delete it
        if 'Index' in sheet_names_list:
            workbook.sheets['Index'].delete()

        # Add a new sheet named "Index"
        index_sheet = workbook.sheets.add('Index')

        # Get the list of existing sheet names
        sheet_names = [sheet.name for sheet in workbook.sheets]

        # Write the header
        index_sheet.range('A1').value = ['Sheet Name', 'Hyperlink']

        # Create hyperlinks for each sheet name
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        for row, sheet_name in enumerate(sheet_names, start=2):
            index_sheet.range(f'A{row}').value = sheet_name
            hyperlink = f"'{sheet_name}'!A1"
            index_sheet.range(f'B{row}').value = "Go to Sheet"

            # Set hyperlink using win32com
            hyperlink_range = index_sheet.range(f'B{row}')
            hyperlink_range.api.Hyperlinks.Add(Anchor=hyperlink_range.api, Address="", SubAddress=hyperlink,
                                               TextToDisplay="Go to Sheet")

        # Save the workbook (optional)
        workbook.save()

        # Close the workbook
        workbook.close()

        # Quit Excel
        excel.Quit()
    except:
        exception_type, exception_value, traceback = sys.exc_info()
        print(f"An error of type {exception_type} occurred: {exception_value}")



if __name__ == '__main__':
    print("""      
                         
        d8888  .d8888b. 8888888      888                       888      888     888 8888888b.  
      d88888 d88P  Y88b  888        888                       888      888     888 888   Y88b 
     d88P888 Y88b.       888        888                       888      888     888 888    888 
    d88P 888  "Y888b.    888        888      .d88b.   .d88b.  888  888 888     888 888   d88P 
   d88P  888     "Y88b.  888        888     d88""88b d88""88b 888 .88P 888     888 8888888P"  
  d88P   888       "888  888        888     888  888 888  888 888888K  888     888 888        
 d8888888888 Y88b  d88P  888        888     Y88..88P Y88..88P 888 "88b Y88b. .d88P 888        
d88P     888  "Y8888P" 8888888      88888888 "Y88P"   "Y88P"  888  888  "Y88888P"  888        
                                                                                              
Version - 2.0                                                         
Washim Rabbani
Information Security Analyst
SecurityScoreCard
                                                           
                                                          """)
    parser = argparse.ArgumentParser(description="Command Line Options Example")

    parser.add_argument("--json", help="Path to JSON file", type=str)
    parser.add_argument("--index", help="Index Excel sheets", type=str)

    args = parser.parse_args()

    if args.json:
        print("[*] Loading JSON files")
        data = json.load(open(args.json, encoding="utf8"))
        print("[*] Creating Dataframes")
        df = pd.json_normalize(data)

        print("[*] Converting Data Frames into excel table")
        dfcves = pd.json_normalize(lstConvert(pd.DataFrame(df['cves'])))
        dfports = pd.json_normalize(lstConvert(pd.DataFrame(df['ports'])))
        dfservices = pd.json_normalize(lstConvert(pd.DataFrame(df['services'])))
        dfhostnames = pd.json_normalize(lstConvert(pd.DataFrame(df['hostnames'])))
        dfproducts = pd.json_normalize(lstConvert(pd.DataFrame(df['products'])))
        dfdomains = pd.json_normalize(lstConvert(pd.DataFrame(df['domains'])))

        dfdetectedLibraries = pd.json_normalize(lstConvert(pd.DataFrame(df['detectedLibraries'])))
        dfmaliciousReputation = pd.json_normalize(lstConvert(pd.DataFrame(df['maliciousReputation'])))

        dfcvss = pd.json_normalize(lstConvert(pd.DataFrame(df['cvss'])))
        dfdeviceType = pd.json_normalize(lstConvert(pd.DataFrame(df['deviceType'])))
        dfransomwareVictims = pd.json_normalize(lstConvert(pd.DataFrame(df['ransomwareVictims'])))
        dfransomwareGroups = pd.json_normalize(lstConvert(pd.DataFrame(df['ransomwareGroups'])))

        dfsensorObservationCategory = pd.json_normalize(lstConvert(pd.DataFrame(df['sensorObservationCategory'])))
        dfsensorObservationSource = pd.json_normalize(lstConvert(pd.DataFrame(df['sensorObservationSource'])))
        if 'httpFaviconHash' in df:
            dfhttpFaviconHash = pd.json_normalize(lstConvert(pd.DataFrame(df['httpFaviconHash'])))
        else:
            dfhttpFaviconHash = pd.json_normalize(lstConvert(pd.DataFrame(df['hxxpFaviconHash'])))

        if 'httpTitle' in df:
            dfhttpTitle = pd.json_normalize(lstConvert(pd.DataFrame(df['httpTitle'])))
        else:
            dfhttpTitle = pd.json_normalize(lstConvert(pd.DataFrame(df['hxxpTitle'])))

        if 'httpStatus' in df:
            dfhttpStatus = pd.json_normalize(lstConvert(pd.DataFrame(df['httpStatus'])))
        else:
            dfhttpStatus = pd.json_normalize(lstConvert(pd.DataFrame(df['hxxpStatus'])))

        dfsslVersion = pd.json_normalize(lstConvert(pd.DataFrame(df['sslVersion'])))
        dfsslCipherName = pd.json_normalize(lstConvert(pd.DataFrame(df['sslCipherName'])))
        dfsslCertAlpn = pd.json_normalize(lstConvert(pd.DataFrame(df['sslCertAlpn'])))
        dfsslCertExtension = pd.json_normalize(lstConvert(pd.DataFrame(df['sslCertExtension'])))
        dfmitreTactics = pd.json_normalize(lstConvert(pd.DataFrame(df['mitreTactics'])))
        dfmitreSoftware = pd.json_normalize(lstConvert(pd.DataFrame(df['mitreSoftware'])))
        dfmitreMitigations = pd.json_normalize(lstConvert(pd.DataFrame(df['mitreMitigations'])))
        dforganizations = pd.json_normalize(lstConvert(pd.DataFrame(df['organizations'])))
        # dfips = pd.json_normalize(lstConvert(pd.DataFrame(df['ips'])))
        dfthreatActors = pd.json_normalize(lstConvert(pd.DataFrame(df['threatActors'])))
        dfindustries = pd.json_normalize(lstConvert(pd.DataFrame(df['industries'])))
        dfdnsRecords = pd.json_normalize(lstConvert(pd.DataFrame(df['dnsRecords'])))
        dfmainAttribution = pd.json_normalize(lstConvert(pd.DataFrame(df['mainAttribution'])))

        dfmitreTechniques = pd.json_normalize(lstConvert(pd.DataFrame(df['mitreTechniques'])))
        dfsensorObservationAlerts = pd.json_normalize(lstConvert(pd.DataFrame(df['sensorObservationAlerts'])))
        dfbreachSourceUrl = pd.json_normalize(lstConvert(pd.DataFrame(df['breachSourceUrl'])))
        dfbreachSourceType = pd.json_normalize(lstConvert(pd.DataFrame(df['breachSourceType'])))
        dfbreachSourceDomain = pd.json_normalize(lstConvert(pd.DataFrame(df['breachSourceDomain'])))
        dfosTypes = pd.json_normalize(lstConvert(pd.DataFrame(df['osTypes'])))

        print("[*] Writing Excel file")
        with pd.ExcelWriter("Output.xlsx") as writer:
            xlstbl(df, "RawHit", "tblRawHit")
            xlstbl(dfports, "ports", "Table9")
            xlstbl(dfservices, "services", "Table10")
            xlstbl(dfhostnames, "hostnames", "Table11")
            xlstbl(dfproducts, "products", "Table12")
            xlstbl(dfdomains, "domains", "Table30")

            xlstbl(dfdetectedLibraries, "dfdetectedLibraries", "Table2")
            xlstbl(dfmaliciousReputation, "maliciousReputation", "Table3")
            xlstbl(dfcves, "cves", "Table4")
            xlstbl(dfcvss, "cvss", "Table5")
            xlstbl(dfdeviceType, "deviceType", "Table6")
            xlstbl(dfransomwareVictims, "ransomwareVictims", "Table7")
            xlstbl(dfransomwareGroups, "ransomwareGroups", "Table8")

            xlstbl(dfsensorObservationCategory, "sensorObservationCategory", "Table13")
            xlstbl(dfsensorObservationSource, "sensorObservationSource", "Table14")
            xlstbl(dfhttpFaviconHash, "httpFaviconHash", "Table15")
            xlstbl(dfhttpTitle, "httpTitle", "Table16")
            xlstbl(dfhttpStatus, "httpStatus", "Table17")
            xlstbl(dfsslVersion, "sslVersion", "Table18")
            xlstbl(dfsslCipherName, "sslCipherName", "Table19")
            xlstbl(dfsslCertAlpn, "sslCertAlpn", "Table20")
            xlstbl(dfsslCertExtension, "sslCertExtension", "Table21")
            xlstbl(dfmitreTactics, "mitreTactics", "Table22")
            xlstbl(dfmitreSoftware, "mitreSoftware", "Table23")
            xlstbl(dfmitreMitigations, "mitreMitigations", "Table24")
            xlstbl(dforganizations, "organizations", "Table25")
            xlstbl(dfthreatActors, "threatActors", "Table26")
            xlstbl(dfindustries, "industries", "Table27")
            xlstbl(dfdnsRecords, "dnsRecords", "Table28")
            xlstbl(dfmainAttribution, "mainAttribution", "Table29")

            xlstbl(dfmitreTechniques, "mitreTechniques", "Table31")
            xlstbl(dfsensorObservationAlerts, "sensorObservationAlerts", "Table32")
            xlstbl(dfbreachSourceUrl, "breachSourceUrl", "Table33")
            xlstbl(dfbreachSourceType, "breachSourceType", "Table34")
            xlstbl(dfbreachSourceDomain, "breachSourceDomain", "Table35")
            xlstbl(dfosTypes, "osTypes", "Table36")

        Index("Output.xlsx")

    elif args.index:
        Index(args.index)
        exit()

    else:
        print("No valid option provided.")
        exit()






    print("[*] Successfully completed\n[*] Check at the output folder for the outcome.")

