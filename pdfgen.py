import pandas as pd
import os
from PyPDF2 import PdfFileWriter, PdfFileReader
from PyPDF2.generic import BooleanObject, NameObject, IndirectObject
import datetime
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
import time
import math

def set_need_appearances_writer(writer: PdfFileWriter):
    try:
        catalog = writer._root_object
        if "/AcroForm" not in catalog:
            writer._root_object.update({
                NameObject("/AcroForm"): IndirectObject(len(writer._objects), 0, writer)})
        need_appearances = NameObject("/NeedAppearances")
        writer._root_object["/AcroForm"][need_appearances] = BooleanObject(True)

        return writer

    except Exception as e:
        print('set_need_appearances_writer() catch : ', repr(e))
        return writer

def generate2():
    print("hello world")

def setExcel():
    global exlFile
    exlFile = filedialog.askopenfilename(filetypes=(("Excel File", "*.xlsx"),))
    print(exlFile)
    label2.config(text=exlFile)

def setFolder():
    global savedFolder
    savedFolder = filedialog.askdirectory()
    print(savedFolder)
    label31.config(text=savedFolder)


def start():
    isExisting = os.path.exists(os.path.join(os.getcwd(), "PMRFv1-2020.pdf"))
    
    if exlFile == '' or savedFolder == '':
        messagebox.showerror('Error','You must choose an .xlsx file and folder before starting.')
    elif isExisting == False:
        messagebox.showerror('Error','The PMRFv1-2020.pdf file was not found in the root directory.')
        print('path of pdf',os.path.join(os.getcwd(), "PMRFv1-2020.pdf"))
    else:
        generate()
        my_progress.stop()
        my_progress['value'] = 100
        label50.config(text='Completed')
        os.system(f'start {os.path.realpath(savedFolder)}')


def generate():
    pdf_filename = "PMRFv1-2020.pdf"
    
    print('xl',exlFile)
    csvin = exlFile
    pdfin = os.path.join(os.getcwd(), pdf_filename)
    print('PDFIN',pdfin)
    pdfout = savedFolder

    #gui.update()
    #my_progress.start(20)
    label50.config(text='Loading')
    gui.update()

    data = pd.read_excel(csvin)
    pdf = PdfFileReader(open(pdfin, "rb"), strict=False)  
    if "/AcroForm" in pdf.trailer["/Root"]:
        pdf.trailer["/Root"]["/AcroForm"].update(
            {NameObject("/NeedAppearances"): BooleanObject(True)})
    pdf_fields = [str(x) for x in pdf.getFields().keys()] 
    csv_fields = data.columns.tolist()
    
    separator = " " 

    i = 0 
    for j, rows in data.iterrows():
        i += 1
        pdf2 = PdfFileWriter()
        set_need_appearances_writer(pdf2)
        if "/AcroForm" in pdf2._root_object:
            pdf2._root_object["/AcroForm"].update(
                {NameObject("/NeedAppearances"): BooleanObject(True)})
        
        field_dictionary_1 = {"PIN": separator.join(str(math.trunc(rows['PhilHealth Identification Number']))) if len(str(rows['PhilHealth Identification Number'])) == 14 else '', 
                            "PurposeReg": '✓' if (str(rows['Purpose'])) == "Registration" else '',
                            "PurposeUpdate": '✓' if (str(rows['Purpose'])) == "Updating/Amendment" else '',

                            # I. PERSONAL DETAILS
                            "LAST NAMEMEMBER": (str(rows['Last Name']).upper()),
                            "FIRST NAMEMEMBER": (str(rows['First Name']).upper()),
                            "MIDDLE NAMEMEMBER": (str(rows['Middle Name']).upper()) if str(rows['Middle Name']) != 'nan' else '', 
                            "NAME EXT1": (str(rows['Extension Name']).upper()) if str(rows['Extension Name']) != 'nan' else '',  
                            #"nomid": '✓' if (str(rows['Middle Name'])) == '' else '',
                            #"mono": '✓' if (str(rows['Biological Sex'])) == "MALE" else '',

                            "DOB MEMBER": datetime.datetime.strptime(str(rows['Date of Birth'].date()), "%Y-%m-%d").strftime("%m - %d - %Y"),
                            "SEXM": '✓' if (str(rows['Biological Sex'])) == "MALE" else '',
                            "SEXF": '✓' if (str(rows['Biological Sex'])) == "FEMALE" else '',

                            "Single": '✓' if (str(rows['Married'])) == "NO" else '',
                            "Married": '✓' if (str(rows['Married'])) == "YES" else '',
                            #"Separated": '✓' if (str(rows['Biological Sex'])) == "MALE" else '',
                            #"Annulled": '✓' if (str(rows['Biological Sex'])) == "MALE" else '',
                            #"Widow": '✓' if (str(rows['Biological Sex'])) == "MALE" else '',

                            "LAST NAMEMOTHER": (str(rows["Mother's Maiden LAST NAME"]).upper()) if str(rows["Mother's Maiden LAST NAME"]) != 'nan' else '', 
                            "FIRST NAMEMOTHER": (str(rows["Mother's Maiden FIRST NAME"]).upper()) if str(rows["Mother's Maiden LAST NAME"]) != 'nan' else '', 
                            "MIDDLE NAMEMOTHER": (str(rows["Mother's Maiden MIDDLE NAME"]).upper()) if str(rows["Mother's Maiden MIDDLE NAME"]) != 'nan' else '', 
                            "NAME EXT2": (str(rows["Mother's Maiden EXTENSION NAME"]).upper()) if str(rows["Mother's Maiden EXTENSION NAME"]) != 'nan' else '', 
                            #"nomid2": '✓' if (str(rows["Mother's Maiden MIDDLE NAME"])) == '' else '', 

                            "LAST NAMESPOUSE": (str(rows["Spouse's LAST NAME"]).upper()) if str(rows["Spouse's LAST NAME"]) != 'nan' else '', 
                            "FIRST NAMESPOUSE": (str(rows["Spouse's FIRST NAME"]).upper()) if str(rows["Spouse's LAST NAME"]) != 'nan' else '', 
                            "MIDDLE NAMESPOUSE": (str(rows["Spouse's MIDDLE NAME"]).upper()) if str(rows["Spouse's MIDDLE NAME"]) != 'nan' else '', 
                            "NAME EXT3": (str(rows["Spouse's EXTENSION NAME"]).upper()) if str(rows["Spouse's EXTENSION NAME"]) != 'nan' else '', 
                            #"nomid3": '✓' if (str(rows["Spouse's MIDDLE NAME"])) == 'nan' else '',  

                            "CITY POB": (str(rows['Municipality/City']).upper()) if str(rows['Municipality/City']) != 'nan' else '',  
                            "PROVINCE POB": (str(rows['Province/State/Country (if abroad)']).upper()) if str(rows['Province/State/Country (if abroad)']) != 'nan' else '',  
                            #"COUNTRY POB": '(str(rows['Middle Name']).upper())',

                            "Filipino": '✓' if (str(rows['Citizenship'])) == "FILIPINO" else '',
                            "Dual": '✓' if (str(rows['Citizenship'])) == "DUAL CITIZEN" else '',
                            "Foreign": '✓' if (str(rows['Citizenship'])) == "NON-FILIPINO" else '',

                            "PHILSYS": separator.join(str(rows['Philsys ID Number (Optional)'])) if len(str(rows['Philsys ID Number (Optional)'])) == 12 else '', 
                            "TIN": separator.join(str(math.trunc(rows['Tax Payer Identification Number (Optional)']))) if str(rows['Tax Payer Identification Number (Optional)']) != 'nan' else '', 

                            # II. ADDRESS and CONTACT DETAILS
                            "Unit": (str(rows['Unit/Room No./Floor']).upper()) if str(rows['Unit/Room No./Floor']) != 'nan' else '', 
                            "Building": (str(rows['Building Name']).upper()) if str(rows['Building Name']) != 'nan' else '', 
                            "House Number": (str(rows['Lot/Block/Phase/House Number']).upper()) if str(rows['Lot/Block/Phase/House Number']) != 'nan' else '', 
                            "Street": (str(rows['Street Name']).upper()) if str(rows['Street Name']) != 'nan' else '', 
                            "Subdivision": (str(rows['Subdivision']).upper()) if str(rows['Subdivision']) != 'nan' else '', 
                            "Barangay": (str(rows['Barangay']).upper()) if str(rows['Barangay']) != 'nan' else '', 
                            "City": (str(rows['Municipality/City']).upper()) if str(rows['Municipality/City']) != 'nan' else '', 
                            "Province": (str(rows['Province/State/Country (if abroad)']).upper()) if str(rows['Province/State/Country (if abroad)']) != 'nan' else '', 
                            "ZIP": math.trunc(rows['ZIP Code']) if str(rows['ZIP Code']) != 'nan' else '',

                            "mail same": '✓' if (str(rows['Is your Mailing Address same as Permanent Home Address?'])) == "YES" else '',

                            "Unit2": (str(rows['Unit/Room No./Floor (Mailing)']).upper()) if str(rows['Unit/Room No./Floor (Mailing)']) != 'nan' else '', 
                            "Building2": (str(rows['Building Name (Mailing)']).upper()) if str(rows['Building Name (Mailing)']) != 'nan' else '', 
                            "House Number2": (str(rows['Lot/Block/Phase/House Number (Mailing)']).upper()) if str(rows['Lot/Block/Phase/House Number (Mailing)']) != 'nan' else '', 
                            "Street2": (str(rows['Street Name (Mailing)']).upper()) if str(rows['Street Name (Mailing)']) != 'nan' else '', 
                            "Subdivision2": (str(rows['Subdivision (Mailing)']).upper()) if str(rows['Subdivision (Mailing)']) != 'nan' else '', 
                            "Barangay2": (str(rows['Barangay (Mailing)']).upper()) if str(rows['Barangay (Mailing)']) != 'nan' else '', 
                            "City2": (str(rows['Municipality/City (Mailing)']).upper()) if str(rows['Municipality/City (Mailing)']) != 'nan' else '', 
                            "Province2": (str(rows['Province/State/Country (Mailing)']).upper()) if str(rows['Province/State/Country (Mailing)']) != 'nan' else '', 
                            "ZIP2": (str(math.trunc(rows['ZIP Code (Mailing)'])).upper()) if str(rows['ZIP Code (Mailing)']) != 'nan' else '', 

                            "Home Phone Number": math.trunc(rows['Home']) if str(rows['Home']) != 'nan' else '',
                            "Mobile Number Required": math.trunc(rows['Mobile Number (Required)']) if str(rows['Mobile Number (Required)']) != 'nan' else '',
                            "Business Direct Line": math.trunc(rows['Business (Direct Line)']) if str(rows['Business (Direct Line)']) != 'nan' else '',
                            "Email Address": rows['E-mail Address (Required for OFW)'] if str(rows['E-mail Address (Required for OFW)']) != 'nan' else '',
                            
                            # III. DECLARATION OF DEPENDENTS
                            "LAST NAMERow1": (str(rows['<span style="display:none">row_3-LAST NAME</span>']).upper()) if str(rows['<span style="display:none">row_3-LAST NAME</span>']) != 'nan' else '', 
                            "FIRST NAMERow1": (str(rows['<span style="display:none">row_3-FIRST NAME</span>']).upper()) if str(rows['<span style="display:none">row_3-FIRST NAME</span>']) != 'nan' else '', 
                            "NAME EXTD1": (str(rows['<span style="display:none">row_3-EXTENSION NAME</span>']).upper()) if str(rows['<span style="display:none">row_3-EXTENSION NAME</span>']) != 'nan' else '', 
                            "MIDDLE NAMERow1": (str(rows['<span style="display:none">row_3-MIDDLE NAME</span>']).upper()) if str(rows['<span style="display:none">row_3-MIDDLE NAME</span>']) != 'nan' else '', 
                            "REL1": (str(rows['<span style="display:none">row_3-Relationship</span>']).upper()) if str(rows['<span style="display:none">row_3-Relationship</span>']) != 'nan' else '', 
                            "DOBD1": datetime.datetime.strptime(str(rows['''<span style="display:none">row_3-Date of Birth</span>'''].date()), "%Y-%m-%d").strftime("%m - %d - %Y") if type(rows['''<span style="display:none">row_3-Date of Birth</span>''']) == datetime.date else str(rows['''<span style="display:none">row_3-Date of Birth</span>''']) if str(rows['''<span style="display:none">row_3-Date of Birth</span>''']) != 'nan' else '', 
                            "CITIZENSHIPRow1": (str(rows['<span style="display:none">row_3-Citizenship</span>']).upper()) if str(rows['<span style="display:none">row_3-Citizenship</span>']) != 'nan' else '',                                                           
                            "nomidd1": '✓' if (str(rows['<span style="display:none">row_3-No Middle Name?</span>'])) == "YES" else '',
                            "monod1": '✓' if (str(rows['<span style="display:none">row_3-Mononym?</span>'])) == "YES" else '',
                            "pwd1": '✓' if (str(rows['<span style="display:none">row_3-Permanent Disability?</span>'])) == "YES" else '',

                            "LAST NAMERow2": (str(rows['<span style="display:none">row_4-LAST NAME</span>']).upper()) if str(rows['<span style="display:none">row_4-LAST NAME</span>']) != 'nan' else '', 
                            "FIRST NAMERow2": (str(rows['<span style="display:none">row_4-FIRST NAME</span>']).upper()) if str(rows['<span style="display:none">row_4-LAST NAME</span>']) != 'nan' else '', 
                            "NAME EXTD2": (str(rows['<span style="display:none">row_4-EXTENSION NAME</span>']).upper())if str(rows['<span style="display:none">row_4-LAST NAME</span>']) != 'nan' else '', 
                            "MIDDLE NAMERow2": (str(rows['<span style="display:none">row_4-MIDDLE NAME</span>']).upper())if str(rows['<span style="display:none">row_4-LAST NAME</span>']) != 'nan' else '', 
                            "REL2": (str(rows['<span style="display:none">row_4-Relationship</span>']).upper())if str(rows['<span style="display:none">row_4-LAST NAME</span>']) != 'nan' else '', 
                            "DOBD2": datetime.datetime.strptime(str(rows['''<span style="display:none">row_4-Date of Birth</span>'''].date()), "%Y-%m-%d").strftime("%m - %d - %Y") if type(rows['''<span style="display:none">row_4-Date of Birth</span>''']) == datetime.date else str(rows['''<span style="display:none">row_4-Date of Birth</span>''']) if str(rows['<span style="display:none">row_4-LAST NAME</span>']) != 'nan' else '', 
                            "CITIZENSHIPRow2": (str(rows['<span style="display:none">row_4-Citizenship</span>']).upper()) if str(rows['<span style="display:none">row_4-LAST NAME</span>']) != 'nan' else '',                                                          
                            "nomidd2": '✓' if (str(rows['<span style="display:none">row_4-No Middle Name?</span>'])) == "YES" else '',
                            "monod2": '✓' if (str(rows['<span style="display:none">row_4-Mononym?</span>'])) == "YES" else '',
                            "pwd2": '✓' if (str(rows['<span style="display:none">row_4-Permanent Disability?</span>'])) == "YES" else '',

                            "LAST NAMERow3": (str(rows['<span style="display:none">row_2-LAST NAME</span>']).upper()) if str(rows['<span style="display:none">row_2-LAST NAME</span>']) != 'nan' else '', 
                            "FIRST NAMERow3": (str(rows['<span style="display:none">row_2-FIRST NAME</span>']).upper()) if str(rows['<span style="display:none">row_2-LAST NAME</span>']) != 'nan' else '', 
                            "NAME EXTD3": (str(rows['<span style="display:none">row_2-EXTENSION NAME</span>']).upper()) if str(rows['<span style="display:none">row_2-LAST NAME</span>']) != 'nan' else '', 
                            "MIDDLE NAMERow3": (str(rows['<span style="display:none">row_2-MIDDLE NAME</span>']).upper()) if str(rows['<span style="display:none">row_2-LAST NAME</span>']) != 'nan' else '', 
                            "REL3": (str(rows['<span style="display:none">row_2-Relationship</span>']).upper()) if str(rows['<span style="display:none">row_2-LAST NAME</span>']) != 'nan' else '', 
                            "DOBD3": datetime.datetime.strptime(str(rows['''<span style="display:none">row_2-Date of Birth</span>'''].date()), "%Y-%m-%d").strftime("%m - %d - %Y") if type(rows['''<span style="display:none">row_2-Date of Birth</span>''']) == datetime.date else str(rows['''<span style="display:none">row_2-Date of Birth</span>''']) if str(rows['<span style="display:none">row_2-LAST NAME</span>']) != 'nan' else '', 
                            "CITIZENSHIPRow3": (str(rows['<span style="display:none">row_2-Citizenship</span>']).upper()) if str(rows['<span style="display:none">row_2-LAST NAME</span>']) != 'nan' else '',                                                           
                            "nomidd3": '✓' if (str(rows['<span style="display:none">row_2-No Middle Name?</span>'])) == "YES" else '',
                            "monod3": '✓' if (str(rows['<span style="display:none">row_2-Mononym?</span>'])) == "YES" else '',
                            "pwd3": '✓' if (str(rows['<span style="display:none">row_2-Permanent Disability?</span>'])) == "YES" else '',

                            "LAST NAMERow4": (str(rows['<span style="display:none">row-LAST NAME</span>']).upper()) if str(rows['<span style="display:none">row-LAST NAME</span>']) != 'nan' else '', 
                            "FIRST NAMERow4": (str(rows['<span style="display:none">row-FIRST NAME</span>']).upper()) if str(rows['<span style="display:none">row-LAST NAME</span>']) != 'nan' else '', 
                            "NAME EXTD4": (str(rows['<span style="display:none">row-EXTENSION NAME</span>']).upper()) if str(rows['<span style="display:none">row-LAST NAME</span>']) != 'nan' else '', 
                            "MIDDLE NAMERow4": (str(rows['<span style="display:none">row-MIDDLE NAME</span>']).upper()) if str(rows['<span style="display:none">row-LAST NAME</span>']) != 'nan' else '', 
                            "REL4": (str(rows['<span style="display:none">row-Relationship</span>']).upper()) if str(rows['<span style="display:none">row-LAST NAME</span>']) != 'nan' else '', 
                            "DOBD4": datetime.datetime.strptime(str(rows['''<span style="display:none">row-Date of Birth</span>'''].date()), "%Y-%m-%d").strftime("%m - %d - %Y") if type(rows['''<span style="display:none">row-Date of Birth</span>''']) == datetime.date else str(rows['''<span style="display:none">row-Date of Birth</span>''']) if str(rows['<span style="display:none">row-LAST NAME</span>']) != 'nan' else '', 
                            "CITIZENSHIPRow4": (str(rows['<span style="display:none">row-Citizenship</span>']).upper()) if str(rows['<span style="display:none">row-LAST NAME</span>']) != 'nan' else '',                                                           
                            "nomidd4": '✓' if (str(rows['<span style="display:none">row-No Middle Name?</span>'])) == "YES" else '',
                            "monod4": '✓' if (str(rows['<span style="display:none">row-Mononym?</span>'])) == "YES" else '',
                            "pwd4": '✓' if (str(rows['<span style="display:none">row-Permanent Disability?</span>'])) == "YES" else '',

                            # # IV. MEMBER TYPE

                            #### Direct Contributer
                            "EmployedPrivate": '✓' if rows['Direct Contributor/Employed Private'] == 1 else '',
                            "EmployedGovernment": '✓' if rows['Direct Contributor/Employed Government'] == 1 else '',
                            "Professional Practitioner": '✓' if rows['Direct Contributor/Professional Practicioner'] == 1 else '',
                            "SelfEarning": '✓' if rows['Direct Contributor/Self-Earning Individual'] == 1 else '',
                            "SelfIndividual": '✓' if rows['Direct Contributor/Self-Earning Individual'] == 1 else '',
                            #"SelfSole": '✓' if rows['Direct Contributor/Employed Private'] == 1 else '',
                            #"SelfGroup": '✓' if rows['Direct Contributor/Employed Private'] == 1 else '',

                            "Kasambahay": '✓' if rows['Direct Contributor/Kasambahay'] == 1 else '',
                            "Driver": '✓' if rows['Direct Contributor/Family Driver'] == 1 else '',
                            "Migrant": '✓' if rows['Direct Contributor/Migrant Worker'] == 1 else '',
                            "MigrantLand": '✓' if rows['Migrant Worker'] == 'Land-Based' else '',
                            "MigrantSea": '✓' if rows['Migrant Worker'] == 'Sea-Based' else '',
                            "Lifetime": '✓' if rows['Direct Contributor/Lifetime Member'] == 1 else '',
                            "FilipinoDual": '✓' if rows['Direct Contributor/Filipinos with Dual Citizenship/ Living Abroad'] == 1 else '',
                            "ForeignNational": '✓' if rows['Direct Contributor/Foreign National'] == 1 else '',
                            
                            "PRA SRRV No": math.trunc(rows["Foreign National's PRA SRRV No."]) if str(rows["Foreign National's PRA SRRV No."]) != 'nan' else '',
                            "ACR ICard No": math.trunc(rows["Foreign National's ACR I-Card No."]) if str(rows["Foreign National's ACR I-Card No."]) != 'nan' else '',

                            #### Indirect Contributer

                            #"Listahanan": '✓' if rows['Direct Contributor/Employed Private'] == 1 else '',
                            "4P": '✓' if rows['Indirect Contributor/4Ps/MCCT'] == 1 else '',
                            "Senior": '✓' if rows['Indirect Contributor/Senior Citizen'] == 1 else '',
                            #"PAMANA": '✓' if rows['Direct Contributor/Employed Private'] == 1 else '',
                            "KIA": '✓' if rows['Survivorship'] == 'Killed In Action (KIA)' else '',
                            #"BANG": '✓' if rows['Direct Contributor/Employed Private'] == 1 else '',
                            #"LGU": '✓' if rows['Direct Contributor/Employed Private'] == 1 else '',
                            #"NGA": '✓' if rows['Direct Contributor/Employed Private'] == 1 else '',
                            #"Private": '✓' if rows['Direct Contributor/Employed Private'] == 1 else '',
                            "PWD": '✓' if rows['Indirect Contributor/Person with Disability (PWD)'] == 1 else '',
                            #"PWD ID No": rows['PhilHealth Identification Number'],

                            "POS": '✓' if rows['Indirect Contributor/Point of Service/Financially Incapable'] == 1 else '',
                            "FI": '✓' if rows['Indirect Contributor/Point of Service/Financially Incapable'] == 1 else '',

                            "Profession": (str(rows['Profession']).upper()) if str(rows['Profession']) != 'nan' else '',
                            "MONTHLY INCOME": (rows['Monthly Income']),
                            "PROOF OF INCOME": (str(rows['Proof of Income']).upper()) if str(rows['Proof of Income']) != 'nan' else '',

                            #V. UPDATING/AMENDMENT

                            # "ChangeN": '✓' if (str(rows['<span style="display:none">row-FROM</span>'])) != "" else '',
                            # "ChangeDOB": '✓' if (str(rows['<span style="display:none">row_1-FROM</span>'])) != "" else '',
                            # "ChangeS": '✓' if (str(rows['<span style="display:none">row_2-FROM</span>'])) != "" else '',
                            # "ChangeC": '✓' if (str(rows['<span style="display:none">row_3-FROM</span>'])) != "" else '',
                            # "UpdatePI": '✓' if (str(rows['<span style="display:none">row_4-FROM</span>'])) != "" else '',

                            # "change name from": (str(rows['<span style="display:none">row-FROM</span>']).upper()),
                            # "change name to": (str(rows['<span style="display:none">row-TO</span>']).upper()),
                            # "dob from": (str(rows['<span style="display:none">row_1-FROM</span>']).upper()),
                            # "dob to": (str(rows['<span style="display:none">row_1-TO</span>']).upper()),
                            # "sex from": (str(rows['<span style="display:none">row_2-FROM</span>']).upper()),
                            # "sex to": (str(rows['<span style="display:none">row_2-TO</span>']).upper()),
                            # "civil from": (str(rows['<span style="display:none">row_3-FROM</span>']).upper()),
                            # "civil to": (str(rows['<span style="display:none">row_3-TO</span>']).upper()),
                            # "info from": (str(rows['<span style="display:none">row_4-FROM</span>']).upper()),
                            # "info to": (str(rows['<span style="display:none">row_4-TO</span>']).upper()),


                            # "Date": datetime.datetime.strptime(str(rows['Health Screening & Assessment Date'].date()), "%Y-%m-%d").strftime("%m-%d-%Y"),
                            
                            # "FN": (str(rows['First Name']).upper()),
                            # "MN": (str(rows['Middle Name']).upper()),
                            # "LN": (str(rows['Last Name']).upper()),

                            }
        

        if (str(rows['Purpose'])) == 'Registration' or (str(rows['Purpose'])) == 'Updating/Amendment':
            temp_out_dir = os.path.normpath(os.path.join(pdfout,str(i+1) +'_' +str(rows['Last Name']).upper() + '.pdf'))
            #print(i+1,'',len(str(rows['PhilHealth Identification Number'])))
            pdf2.addPage(pdf.getPage(0))
            pdf2.updatePageFormFieldValues(pdf2.getPage(0), field_dictionary_1)
            #pdf2.addPage(pdf.getPage(1))
            #pdf2.addPage(pdf.getPage(2))
            #pdf2.addPage(pdf.getPage(3))
            outputStream = open(temp_out_dir, "wb")
            pdf2.write(outputStream)
            outputStream.close()
            label50.config(text='Writing '+str(rows['Last Name']).upper()+'.pdf')
            my_progress['value'] += 0.3
            gui.update()
            #print(f'Process Complete: {i} PDFs Processed!')
            


if __name__ == '__main__':
    print("Current working directory: {0}".format(os.getcwd()))
    exlFile = ''
    pdfFile = ''
    savedFolder = ''

    gui2 = Tk()
    gui = Frame(gui2)
 
    gui2.configure(background="#ebf3f3")
    gui2.title("AER: Philhealth Form Generator")
    gui2.geometry("600x300")

    label0 = Label(gui, text='Version: PMRF UHC v.1 January 2020', height=2, width=30, font='Helvetica 12 bold')
    label0.grid(row=0, column=0, sticky=W)

    #XL FILE
    label1 = Label(gui, text='Specify the file path of the .xlsx file', height=2, width=30)
    label1.grid(row=1, column=0, sticky=W)
    label2 = Label(gui, text=exlFile, height=1, width=60, bg='white',)
    label2.grid(row=2, column=0)
    button2 = Button(gui, text='Browse', fg='#e4e3e7', bg='#005181', command=lambda: setExcel(), height=1, width=10)
    button2.grid(row=2, column=1, padx=5)

    #DIRECTORY FOR PDF PRINTING
    label30 = Label(gui, text='Specify the directory path to save the generated PDF files', height=2, width=45)
    label30.grid(row=3, column=0, sticky=W)
    label31 = Label(gui, text=savedFolder, height=1, width=60, bg='white')
    label31.grid(row=4, column=0)
    button3 = Button(gui, text='Browse', fg='#e4e3e7', bg='#005181', command=lambda: setFolder(), height=1, width=10)
    button3.grid(row=4, column=1, padx=5)

    #GENERATE PDF
    label40 = Label(gui, text='Click Start to generate the PDF files', height=2, width=30)
    label40.grid(row=5, column=0, sticky=W)
    button4 = Button(gui, text='Start', fg='#e4e3e7', bg='#005181', command=lambda: start(), height=4, width=10)
    button4.grid(row=6, column=1, pady= 0, padx=5, sticky=W)

    my_progress = ttk.Progressbar(gui, orient=HORIZONTAL, length=428, mode='determinate')
    my_progress.grid(row=6, column=0, pady= 0, sticky=N)

    label50 = Label(gui, text='Not started.', height=1, width=30)
    label50.grid(row=6, column=0, pady=25, sticky=W)

    gui.place(relx=0.5, rely=0.5, anchor=CENTER)

    #update()
     # start the GUI
    gui.mainloop()