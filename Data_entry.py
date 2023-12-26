# Necessary imports

import PySimpleGUI as Sg
import pandas as pd
import cv2
from pdf_gen import PdfGenerator
from openpyxl import Workbook

# ------------- Config ----------------- #

Sg.theme('DarkTeal9')  # Setting theme
EXCEL_FILE = 'Data.xlsx'
try:
    dd = pd.read_excel(EXCEL_FILE)
except FileNotFoundError:
    # Sg.popup_error("Excel Sheet Not Found")
    workbook = Workbook()
    sheet = workbook.active
    data = [
        ["ApplicationNumber", "Name", "Class", "Section", "Address1", "Address2", "Address3",
         "From", "To", "Distance", "Fare"]
    ]
    for row in data:
        sheet.append(row)
    # Save the workbook to a file
    excel_file_path = "Data.xlsx"
    workbook.save(excel_file_path)
    dd = pd.read_excel(EXCEL_FILE)
cam = cv2.VideoCapture(0)  # Selecting first Camera (front Camera)
cam_open = False
img_counter = len(dd)

# -------------- Layout design ----------------- #
layout = [
    [Sg.Text(' ', size=10), Sg.Text('TamilNadu Transport Corporation - Kumbakonam')],
    [Sg.Text('Ranees Govt. Hr. Sec. School', size=(40, 1)), Sg.Text('School Code: 199')],
    [Sg.Text('App.No.', size=(15, 1)), Sg.InputText(key='ApplicationNumber')],
    [Sg.Text('Name', size=(15, 1)), Sg.InputText(key='Name')],
    [Sg.Text('Class', size=(15, 1)), Sg.Combo(['6', '7', '8', '9', '10', '11', '12'], key='Class')],
    [Sg.Text('Section', size=(15, 1)), Sg.InputText(key='Section')],
    [Sg.Text('Address 1', size=(15, 1)), Sg.InputText(key='Address1')],
    [Sg.Text('Address 2', size=(15, 1)), Sg.InputText(key='Address2')],
    [Sg.Text('Address 3', size=(15, 1)), Sg.InputText(key='Address3')],
    [Sg.Text('From', size=(15, 1)), Sg.InputText(key='From')],
    [Sg.Text('To', size=(15, 1)), Sg.InputText(key='To')],
    [Sg.Text('Distance(in Kms)', size=(15, 1)), Sg.InputText(key='Distance')],
    [Sg.Text('Fare (₹)', size=(15, 1)), Sg.InputText(key='Fare')],

    [Sg.Submit(), Sg.Button("Camera"), Sg.Button('Clear'), Sg.Button("Generate Report"), Sg.Exit()]
]

window = Sg.Window('Bus-Pass Registration', layout)


# ---------------- functions --------------- #

def clear_input():  # clear function
    for key in values:
        window[key](' ')
    return None


while True:
    event, values = window.read()
    if event == Sg.WIN_CLOSED or event == 'Exit':   # Exit button
        break
    elif event == 'Clear':    # clear button
        clear_input()

    elif event == 'Submit':    # submit button
        dd = dd._append(values, ignore_index=True)
        dd.to_excel(EXCEL_FILE, index=False)
        Sg.popup('Data saved!')
        clear_input()

    elif event == "Generate Report":    # Generate report in pdf
        pdff = PdfGenerator()
        pdff.pdf_generation()

    elif event == 'Camera':    # Camera button
        cv2.namedWindow("Image Capture")
        while True:
            success, frame = cam.read()
            if not success:
                cam = cv2.VideoCapture(0)
            else:
                cv2.imshow("Test", frame)
                k = cv2.waitKey(1)      # waiting for a keyboard interrupt

                # ------- Keyboard interrupts -------- #
                if k % 256 == 27:  # if escape key is pressed
                    Sg.popup("Camera Closed")
                    cam.release()
                    break

                elif k % 256 == 32:  # if space bar is pressed
                    img_name = f"opencv_frame_{img_counter}.png"
                    cv2.imwrite(img_name, frame)
                    Sg.popup("Image Captured")
                    img_counter += 1

        cv2.destroyAllWindows()
        cam.release()

# -----------------------
window.close()
# ------------------------
