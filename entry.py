from flask import Flask, render_template, url_for, flash, redirect, request
from fpdf import FPDF
import os
from datetime import datetime
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json

app = Flask(__name__)
app.static_folder = 'static'


@app.route('/', methods=['GET', 'POST'])
@app.route('/index', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        fname = request.form['fname']
        lname = request.form['lname']
        purpose = request.form['purpose']
        curTime = str(time.ctime())
        id_Gen(fname, lname, curTime)
        # log_visit(fname, lname, curTime, purpose)
        return redirect(url_for('index'))
    return render_template('index.html')


def log_visit(fName, lName, curTime, purpose):
    # Will append Firstname/Lastname/Purpose of visit to a Google Sheet
    googleAuthenticate()
    print('Authentication OK')
    sheet = client.open("SEG_Visitors").sheet1  # open sheet
    row = [curTime, fName, lName, purpose]
    sheet.insert_row(row, 2)


def idleCheck():
    # If the authtication goes stale, reauthenticae
    curTime = curTime[11:19]
    FMT = '%H:%M:%S'
    idleTime = datetime.strptime(curTime, FMT) - \
        datetime.strptime(authTime, FMT)
    idleTime = str(idleTime)[5:7]
    if int(idleTime) > 1:
        googleAuthenticate()


def googleAuthenticate():
    # Authenticate with Google and set authTime to event time
    json_key = json.load(open(os.getcwd()+'/creds.json'))
    scope = ["https://spreadsheets.google.com/feeds",
             'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        'C:/printer/flaskr/creds.json', scope)  # get email and key from creds
    file = gspread.authorize(credentials)  # authenticate with Google
    globals()['client'] = gspread.authorize(credentials)
    authTime = time.ctime()
    globals()['authTime'] = authTime[11:19]
    print('ReAuth')


def id_Gen(fname, lname, curTime):
    # Genearate and print pdf with First/Last name
    pdf = FPDF('L', 'mm', (54, 102))
    pdf.set_margins(left=5, top=2, right=5)
    pdf.set_auto_page_break(False)
    pdf.add_page()

    pdf.image(os.getcwd()+'/logo.png', w=35, h=11)
    pdf.set_font("Arial", size=22)
    pdf.cell(0, 17, txt="VISITOR PASS", ln=1, align="C")
    pdf.set_font("Arial", size=22)
    pdf.cell(0, 3, txt="{} {}".format(fname, lname), ln=1, border=0, align="C")
    pdf.set_font("Arial", size=12)
    pdf.cell(0, 13, txt=curTime, ln=1,  align="C")
    pdf.set_line_width(.5)
    pdf.set_draw_color(255, 0, 0)
    pdf.line(5, 15, 97, 15)
    pdf.output(os.getcwd()+"/entry.pdf")

    os.system('lp entry.pdf')
