import PySimpleGUI as sg
import pandas as pd
import cv2

#------------------------------

sg.theme('DarkTeal9')

#-------------------------------

EXCEL_FILE = 'Data.xlsx'
dd = pd.read_excel(EXCEL_FILE)
cam = cv2.VideoCapture(0)
img_counter = 0

#-------------------------------
layout = [
    [sg.Text(' ', size = (10)), sg.Text('TamilNadu Transport Corporation - Kumbakonam')],
    [sg.Text('Ranees Govt. Hr. Sec. School', size = (40,1)), sg.Text('School Code: 199')],
    [sg.Text('Name', size = (15, 1)), sg.InputText(key = 'Name')],
    [sg.Text('Class', size = (15, 1)), sg.Combo(['6','7','8','9','10','11','12'], key = 'Class')],
    [sg.Text('Section', size = (15,1)), sg.InputText(key = 'Section')],
    [sg.Text('Address 1', size = (15, 1)), sg.InputText(key = 'Address1')],
    [sg.Text('Address 2', size = (15, 1)), sg.InputText(key = 'Address2')],
    [sg.Text('Address 3', size = (15, 1)), sg.InputText(key = 'Address3')],
    [sg.Text('From', size = (15, 1)), sg.InputText(key = 'From')],
    [sg.Text('To', size = (15, 1)), sg.InputText(key = 'To')],
    [sg.Text('Distance(in Kms)', size = (15, 1)), sg.InputText(key = 'Distance')],
    [sg.Text('Fare ($)', size = (15, 1)), sg.InputText(key = 'Fare')],

    [sg.Submit(), sg.Button("Camera"), sg.Button('Clear'), sg.Exit()]
]
#-------------------------------
window = sg.Window('Bus-Pass Registration', layout)

def clear_input():
    for key in values:
        window[key](' ')
    return None

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        dd = dd._append(values, ignore_index = True)
        dd.to_excel(EXCEL_FILE, index = False)
        sg.popup('Data saved!')
        clear_input()
    if event == 'Camera':
        cv2.namedWindow("Image Capture")
        while True:
            ret, frame = cam.read()

            if not ret:
                print("Failed to grab frame")
                break
            cv2.imshow("Test", frame)
            k = cv2.waitKey(1)

            if k%256 == 27:
                print("Escape hit, closing the app")
                break

            elif k%256 == 32:
                img_name = "opencv_frame_{}.png".format(img_counter)
                cv2.imwrite(img_name, frame)
                print("Image Captured")
                img_counter += 1
        cam.release()
        cv2.destroyAllWindows()
#-----------------------
window.close()
#------------------------