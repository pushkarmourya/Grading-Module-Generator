# -*- coding: utf-8 -*-

from flask import Flask, render_template, request
from email.utils import COMMASPACE, formatdate
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from os.path import basename
from hashlib import new
import os
import xlrd
import csv
import pandas as pd
from typing import Counter
import pandas as pd
import xlsxwriter
import openpyxl
from xlsxwriter.workbook import WorksheetMeta
import smtplib
from email.message import EmailMessage
import socket


master_roll = {}
response = {}
concise_marks = {}
correct_answer_marks = 0
incorrect_answer_marks = 0


after_marks = {}

app = Flask(__name__)
app.config['UPLOADED_PATH'] = os.path.join(app.root_path, 'upload')


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        for f in request.files.getlist('file'):
            f.save(os.path.join(app.config['UPLOADED_PATH'], f.filename))
    return render_template('index.html')


@app.route('/generate_individual_marksheet', methods=['GET', "POST"])
def marksheet_generator_function():
    if not os.path.isdir("marksheet"):
        os.mkdir("marksheet")

    if not os.path.isfile('upload/responses.csv') or not os.path.isfile('upload/master_roll.csv'):
        return render_template('upload_file_error.html')
    masterroll_file = pd.read_csv("upload/master_roll.csv")
    response_file = pd.read_csv("upload/responses.csv")

    for index, row in masterroll_file.iterrows():
        master_roll[row["roll"]] = row["name"]

    for index, row in response_file.iterrows():
        response[row["Roll Number"]] = row

    correct_answer_marks = request.form.get("positive_marks")
    incorrect_answer_marks = request.form.get("negative_marks")

    f = open("static/value_store.csv", 'w')
    writer = csv.writer(f)
    writer.writerow([float(correct_answer_marks), float(incorrect_answer_marks)])
    f.close()

    if "ANSWER" not in response:
        return "no roll number with ANSWER is present, Cannot Process!"

    for key in master_roll:
        right_answer = 0
        wrong_answer = 0
        not_attempted = 0

        global no_of_questions
        checked_answers = []
        for i in range(7, len(response["ANSWER"])):
            if pd.isna(response[key][i]):
                response[key][i] = ""
                checked_answers.append(
                    [response[key][i], response["ANSWER"][i], -1, 0])
                not_attempted += 1
                continue

            if response["ANSWER"][i] == response[key][i]:
                checked_answers.append(
                    [response[key][i], response["ANSWER"][i], 1, correct_answer_marks])
                right_answer += 1
            else:
                checked_answers.append(
                    [response[key][i], response["ANSWER"][i], 0, incorrect_answer_marks])
                wrong_answer += 1

        roll_number = str(key[:4]+key[4:6].upper()+key[6:])

        workbook = xlsxwriter.Workbook("marksheet/"+roll_number+".xlsx")
        worksheet = workbook.add_worksheet()

        worksheet.set_column(0, 4, 17)

        merge_format = workbook.add_format({"font_size": 16,
                                            'bold': 1, 'border': 1})
        merge_format.set_align("center")
        worksheet.merge_range('A5:E5', "MARKSHEET", merge_format)

        cell_format_red = workbook.add_format({'border': 1})
        cell_format_red.set_font_color('red')
        cell_format_red.set_align('center')
        cell_format_red.set_font_size(14)

        cell_format_blue = workbook.add_format({'border': 1})
        cell_format_blue.set_font_color('blue')
        cell_format_blue.set_align('center')
        cell_format_blue.set_font_size(14)

        cell_format_green = workbook.add_format({'border': 1})
        cell_format_green.set_font_color('green')
        cell_format_green.set_align('center')
        cell_format_green.set_font_size(14)

        cell_format_bold = workbook.add_format({'border': 1})
        cell_format_bold.set_align('center')
        cell_format_bold.set_bold()
        cell_format_bold.set_font_size(14)

        cell_format_normal = workbook.add_format()
        cell_format_normal.set_align('center')
        cell_format_normal.set_font_size(14)
        worksheet.write("A6", "Name:", cell_format_normal)
        worksheet.write("A7", "Roll Number:", cell_format_normal)

        cell_format_normal = workbook.add_format({'border': 1})
        cell_format_normal.set_align('center')
        cell_format_normal.set_font_size(14)

        worksheet.write("D6", "Exam:", cell_format_normal)
        worksheet.insert_image(0, 0, "static/IITP LOGO.png",
                               {'x_scale': 0.62, 'y_scale': 0.62})

        for i in range(len(checked_answers)):
            if i < 25:
                if checked_answers[i][2] == 1:
                    worksheet.write(
                        'A'+str(i+16), checked_answers[i][0], cell_format_green)
                else:
                    worksheet.write(
                        'A'+str(i+16), checked_answers[i][0], cell_format_red)

                worksheet.write(
                    'B'+str(i+16), checked_answers[i][1], cell_format_blue)
            else:
                if checked_answers[i][2]:
                    worksheet.write(
                        'D'+str(i+16-25), checked_answers[i][0], cell_format_green)
                else:
                    worksheet.write(
                        'D'+str(i+16-25), checked_answers[i][0], cell_format_red)

                worksheet.write(
                    'E'+str(i+16-25), checked_answers[i][1], cell_format_blue)

        worksheet.write("B6", master_roll[key], cell_format_bold)
        worksheet.write("B7", key, cell_format_bold)
        worksheet.write("B11", correct_answer_marks, cell_format_green)
        worksheet.write("C11", incorrect_answer_marks, cell_format_red)
        worksheet.write("D11", "0", cell_format_normal)
        worksheet.write("B10", right_answer, cell_format_green)
        worksheet.write("C10", wrong_answer, cell_format_red)
        worksheet.write("B12", right_answer *
                        float(correct_answer_marks), cell_format_green)
        worksheet.write("C12", float(incorrect_answer_marks)
                        * wrong_answer, cell_format_red)
        worksheet.write("D10", not_attempted, cell_format_normal)
        worksheet.write("E6", "Quiz", cell_format_bold)

        worksheet.write("E10", len(checked_answers), cell_format_normal)
        worksheet.write("E12", str(right_answer *
                        float(correct_answer_marks)+float(incorrect_answer_marks)*wrong_answer) + "/" + str(len(checked_answers)*float(correct_answer_marks)), cell_format_blue)
        worksheet.write("D12", "", cell_format_normal)
        worksheet.write("E11", "", cell_format_normal)
        worksheet.write("A9", "", cell_format_bold)
        worksheet.write("B9", "Right", cell_format_bold)
        worksheet.write("C9", "Wrong", cell_format_bold)
        worksheet.write("D9", "NotAttempt", cell_format_bold)
        worksheet.write("E9", "Max", cell_format_bold)
        worksheet.write("A10", "No.", cell_format_bold)
        worksheet.write("A11", "Marking", cell_format_bold)
        worksheet.write("A12", "Total", cell_format_bold)
        # print(after_marks)
        workbook.close()

    return render_template('generate_marksheet.html')


@app.route('/generate_concise_marksheet', methods=['GET', "POST"])
def concise_marksheet_generator_function():
    if not os.path.isdir("marksheet"):
        os.mkdir("marksheet")

    if not os.path.isfile('upload/responses.csv') or not os.path.isfile('upload/master_roll.csv'):
        return render_template('upload_file_error.html')
    masterroll_file = pd.read_csv("upload/master_roll.csv")
    response_file = pd.read_csv("upload/responses.csv")

    correct_answer_marks = 0
    incorrect_answer_marks = 0
    with open("static/value_store.csv", 'r', newline='') as File:
        reader = list(csv.reader(File))
        for row in reader:
            correct_answer_marks = float(row[0])
            incorrect_answer_marks = float(row[1])
            break
    for index, row in masterroll_file.iterrows():
        master_roll[row["roll"]] = row["name"]

    for index, row in response_file.iterrows():
        response[row["Roll Number"]] = row

    if "ANSWER" not in response:
        return "no roll number with ANSWER is present, Cannot Process!"
    new_response = []
    for key in master_roll:
        right_answer = 0
        wrong_answer = 0
        not_attempted = 0
        no_of_questions = 0

        if key not in response:
            new_response[key]=[key,master_roll[key],"ABSENT"]
            print(key)
            continue

        checked_answers = []
        for i in range(7, len(response["ANSWER"])):
            no_of_questions += 1
            if pd.isna(response[key][i]):
                not_attempted += 1
                continue
            if response["ANSWER"][i] == response[key][i]:
                right_answer += 1
            else:
                wrong_answer += 1

        roll_number = str(key[:4]+key[4:6].upper()+key[6:])

        after_marks[key] = [str(right_answer *
                                correct_answer_marks+incorrect_answer_marks*wrong_answer) + "/" + str(no_of_questions*correct_answer_marks), "["+str(right_answer)+","+str(wrong_answer)+","+str(not_attempted)+"]"]

    for key in master_roll:
        temp = []
        count = 0
        for items in response[key]:
            count += 1
            if count == 6:
                temp.append(after_marks[key][0])
            temp.append(items)
        temp.append(after_marks[key][1])
        new_response.append(temp)

    df = pd.DataFrame(new_response)

    header_list = []
    header_list.extend(["Timestamp", "Email address", "Google_Score", "Name", "IITP webmail", "Score_After_negative", "Phone(10 digit only)",
                       "Roll Number"])
    for _ in range(no_of_questions):
        header_list.append("unnamed")
    header_list.append("StatusAns")
    writer = pd.ExcelWriter(
        'marksheet/concise_marksheet.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='sheet1', index=False, header=header_list)
    writer.save()
    return render_template('generate_concise_marksheet.html')


@app.route('/send_email', methods=['GET', "POST"])
def send_email():

    response_file = pd.read_csv("upload/responses.csv")
    new_response = {}
    for index, row in response_file.iterrows():
        new_response[row["Roll Number"]] = row

    for key in new_response:
        roll_no = key
        socket.getaddrinfo('localhost', 8080)
        msg = EmailMessage()
        msg['subject'] = 'Marksheet'
        msg['From'] = 'Pushkar Mourya'
        msg['To'] = new_response[key][1], new_response[key][4]
        msg.set_content("Please find the masksheet of quiz")
        server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
        server.login("pushkarmourya.iitp@gmail.com", "pushkar250801")
        with open("marksheet/"+roll_no+".xlsx", "rb") as file:
            data = file.read()
            file_name = file.name
            msg.add_attachment(data, maintype="application",
                               subtype="xlsx", filename=file_name)
        server.send_message(msg)
        server.quit()

    return render_template('send_email.html')

if __name__ == '__main__':
    app.run(debug=True)
