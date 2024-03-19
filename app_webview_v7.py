import flask
from flask import Flask, render_template, request
from pdf2image import convert_from_path
import os
import cv2
import pytesseract
import pandas as pd
import re
import glob
from pandas import ExcelWriter
import datetime
import pathlib
import webview


app = Flask(__name__, template_folder= './templates')
window = webview.create_window('value extraction',  app)

# Create a route to the main page
@app.route('/')
def main_page():
    return render_template('login_dashboard.html')

#
# Create a route to handle the file upload
@app.route('/upload', methods=['POST'])
def handle_upload():

    uploaded_file_1 =  request.form['text_input']
    uploaded_file_2 =  request.form['text_input_1']
    uploaded_file_3 =  request.form['text_input_2']

    poppler_path = uploaded_file_1 + r'\Release-22.04.0-0\poppler-22.04.0\Library\bin'
    print("pp", poppler_path)
    folder_path = uploaded_file_1
    #filename = uploaded_file_2
    pdf_path = folder_path + '/' + uploaded_file_2

    head_tail = os.path.split(pdf_path)
    get_folder_name = head_tail[1]
    ret = get_folder_name.split(".")[0]
    p = pathlib.Path(folder_path)
    os.chdir(p)
    ret_solu = ret + str("_files")
    print("folder_name", ret_solu)
    pages = convert_from_path(pdf_path=pdf_path, poppler_path=poppler_path)
    multiple_img = "images"
    saving_folder = os.path.join(ret_solu, multiple_img)
    os.umask(0)
    os.makedirs(saving_folder, mode=0o666)
    os.chdir(saving_folder)

    c = 1

    for page in pages:
        img_name = f"img-{c}.jpeg"
        page.save(img_name)
        c += 1
    print("***************************poppler finished**********************************")

    # Mention the installed location of Tesseract-OCR in your system

    pytesseract.pytesseract.tesseract_cmd = uploaded_file_1 + r'\Tesseract-OCR\tesseract.exe'
    path = glob.glob("*.jpeg")

    result = []
    for i, k in enumerate(path):

        img = cv2.imread(k)

        # Convert the image to gray scale
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        # Performing OTSU threshold
        ret, thresh1 = cv2.threshold(gray, 0, 255, cv2.THRESH_OTSU | cv2.THRESH_BINARY_INV)

        rect_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (18, 18))

        # Applying dilation on the threshold image
        dilation = cv2.dilate(thresh1, rect_kernel, iterations=1)

        # Finding contours
        contours, hierarchy = cv2.findContours(dilation, cv2.RETR_EXTERNAL,
                                               cv2.CHAIN_APPROX_NONE)

        # Creating a copy of image
        im2 = img.copy()

        for j, cnt in enumerate(contours):
            inner = ''
            x, y, w, h = cv2.boundingRect(cnt)
            # Drawing a rectangle on copied image
            rect = cv2.rectangle(im2, (x, y), (x + w, y + h), (0, 255, 0), 2)
            # Cropping the text block for giving input to OCR
            cropped = im2[y:y + h, x:x + w]
            ext_txt = pytesseract.image_to_string(cropped)
            result.append(ext_txt)
        df = pd.DataFrame(result)

        excel_folder_path = folder_path + '\\' + ret_solu
        save_excel_path = excel_folder_path + "\\Excel_files"
        if os.path.exists(save_excel_path):
            pass
        else:
            os.mkdir(excel_folder_path + "\\Excel_files")

        file = "extract_text" + str(i) + ".xlsx"
        writer = ExcelWriter(save_excel_path + '\\' + file)
        df.to_excel(writer, 'Sheet1', index=False)
        # writer.save()
        writer.close()
        result *= 0
    excel_folder_path = folder_path + '\\' + ret_solu
    save_excel_path = excel_folder_path + "\\Excel_files"
    file = "extract_text*.xlsx"
    print("********************Tesseract completed**********************")
    path_excel = glob.glob(save_excel_path + '\\' + file)
    result_value = []
    result_value1 = []
    result_value2 = []
    result_value3 = []
    result_value4 = []
    result_value6 = []
    result_value7 = []
    result_value8 = []
    result_value9 = []
    result_value10 = []
    result_value11 = []
    result_value12 = []
    for i in path_excel:
        df = pd.read_excel(i)
        result = df.to_string()

        pattern1 = 'Name: ([A-Z. A-Za-z]*[a-zA-Z]*[A-ZA-Z]*)'
        pattern2 = 'RN\s*Number\s*:\s*(RNV[0-9]{5}|RNV[0-9]{4})'
        pattern3 = '(?:Total Billable days:)([0-9.]+)'
        pattern4 = '(?:Approved days:\s*|Approved\s*days\s*:|Approved\s*days\s*:\s*)(-?\s*[0-9.0]*)[\),]'
        pattern5 = 'IPN: ([A-Z]?[0-9]{6})'
        pattern7 = '(?: Org.Business days:\s*|Org.Business\s*days\s*:|Org.Business\s*days\s*:\s*)(\s*[0-9]*)[\),]'
        pattern8 = '(?:Working days:\s*|Working\s*days:\s*|days:)([0-9.]+)'
        pattern9 = '(?:Declared Days:\s*|Declared\s*Days:\s*|Days:)([0-9]+)'
        pattern10 = '(?:Undeclared days:\s*|Undeclared\s*days:\s*|days:)([0-9.]+)'
        pattern11 = '(?: Leave Days:\s*|Leave\s*Days:\s*|Days:)([0-9.]+)'
        pattern12 = '(?:Rejected days:\s*|Rejected\s*days:\s*|days:)([0-9.]+)'
        pattern13 = 'Submitted days:[0-9.]+'



        required_data = []
        required_data_rn = []
        required_data_billable = []
        required_data_aprroved = []
        required_data_ipn = []
        required_data_org_bussiness = []
        required_data_Working = []
        required_data_Declared = []
        required_data_Undeclared = []
        required_data_Leave = []
        required_data_Rejected = []
        required_data_Submitted = []

        get_name = re.search(pattern1, result, flags=0)
        get_rn_number = re.search(pattern2, result, flags=0)
        get_total_billable = re.search(pattern3, result, flags=0)
        get_total_approved = re.search(pattern4, result, flags=0)
        get_total_ipn = re.search(pattern5, result, flags=0)
        #print("Declared", get_Declared)

        get_org_busssiness = re.search(pattern7, result, flags=0)
        print("org_busssiness", get_org_busssiness)
        get_Working = re.search(pattern8, result, flags=0)
        print("workig", get_Working)
        get_Declared = re.search(pattern9, result, flags=0)
        print("Declared", get_Declared)
        get_Undeclared= re.search(pattern10, result, flags=0)
        print("undeclared", get_Undeclared)
        get_Leave = re.search(pattern11, result, flags=0)
        print("Leave", get_Leave)
        get_Rejected = re.search(pattern12, result, flags=0)
        print("Rejected", get_Rejected)
        get_Submitted = re.search(pattern13, result, flags=0)
        print("Submitted", get_Submitted)



        try:
            Name = get_name.group(0)
            required_data.append(Name)
            RN_Number = get_rn_number.group(0)
            required_data_rn.append(RN_Number)
            Total_billable = get_total_billable.group(0)
            required_data_billable.append(Total_billable)
            org_busssiness = get_org_busssiness.group(0)
            required_data_org_bussiness.append(org_busssiness)
            Working = get_Working.group(0)
            required_data_Working.append(Working)
            Declared = get_Declared.group(0)
            required_data_Declared.append(Declared)
            Undeclared = get_Undeclared.group(0)
            required_data_Undeclared.append(Undeclared)
            Leave = get_Leave.group(0)
            required_data_Leave.append(Leave)
            Rejected = get_Rejected.group(0)
            required_data_Rejected.append(Rejected)
            Submitted = get_Submitted.group(0)
            required_data_Submitted.append(Submitted)

            if get_total_approved != None:
                Total_approved = get_total_approved.group(0).strip()
                required_data_aprroved.append(Total_approved)
            Total_ipn = get_total_ipn.group(0)
            required_data_ipn.append(Total_ipn)

        except AttributeError:
            continue

        ch = ':'
        for i, h, b, z, a, o, w, d, u, l, r, s in zip(required_data, required_data_rn, required_data_billable, required_data_ipn,
                                 required_data_aprroved, required_data_org_bussiness,required_data_Working, required_data_Declared,required_data_Undeclared, required_data_Leave, required_data_Rejected, required_data_Submitted):
            splitted_name = i.split(':')
            print(splitted_name[1])
            strvalue = splitted_name[1]
            strvalue1 = h.split(ch, 1)[1]
            strvalue2 = b.split(ch, 1)[1]
            strvalue3 = a.split(ch, 1)[1]
            strvalue5_app = strvalue3.split(")", 1)[0]
            strvalue5 = strvalue5_app.split(",")[0]
            strvalue4 = z.split(ch, 1)[1]
            strvalue_bus = o.split(ch, 1)[1]
            strvalue7 = strvalue_bus.split(",")[0]
            strvalue8 = w.split(ch, 1)[1]
            strvalue9 = d.split(ch, 1)[1]
            strvalue10 = u.split(ch, 1)[1]
            strvalue11 = l.split(ch, 1)[1]

            strvalue12 = r.split(ch, 1)[1]
            strvalue13 = s.split(ch, 1)[1]


            result_value.append(strvalue)
            result_value1.append(strvalue1)
            result_value2.append(strvalue2)
            result_value3.append(strvalue5)
            result_value4.append(strvalue4)

            result_value6.append(strvalue7)
            result_value7.append(strvalue8)
            result_value8.append(strvalue9)
            result_value9.append(strvalue10)
            result_value10.append(strvalue11)
            result_value11.append(strvalue12)
            result_value12.append(strvalue13)

        df = pd.DataFrame(
            {Name: result_value, Total_ipn: result_value4, RN_Number: result_value1, Total_billable: result_value2,
              Total_approved: result_value3, org_busssiness : result_value6, Working : result_value7, Declared : result_value8, Undeclared : result_value9, Leave : result_value10, Rejected : result_value11, Submitted : result_value12})
        df.columns = ['NAME', 'IPN', 'RN NUMBER', 'Total Billable Days', 'Total Approved Days', 'Org_bussiness_days', 'Working_days', 'Declared_days', 'Undeclared_days', 'Leave_days', 'Rejected_days', 'Submitted_days']

        now = datetime.datetime.now()
        formatted_date_time = now.strftime("%Y-%m-%d___%H-%M-%S")
        writer = ExcelWriter(uploaded_file_3 + '/' + f'Result_{ret_solu}.xlsx')
        df.to_excel(writer, 'Sheet1', index=False)
        print("********************Extraction Completed**********************")
        writer.close()

    return flask.render_template('result_page.html')

if __name__ == '__main__':
   webview.start()

