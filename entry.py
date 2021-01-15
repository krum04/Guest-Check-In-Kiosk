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

# @app.route("/home")


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
        print('RUN')
        return redirect(url_for('index'))

    return render_template('index.html')


def log_visit(fName, lName, curTime, purpose):
    googleAuthenticate()
    print('Authentication OK')
    sheet = client.open("SEG_Visitors").sheet1  # open sheet
    row = [curTime, fName, lName, purpose]
    sheet.insert_row(row, 2)


def idleCheck():
    #curTime = time.ctime()
    curTime = curTime[11:19]
    FMT = '%H:%M:%S'
    idleTime = datetime.strptime(curTime, FMT) - \
        datetime.strptime(authTime, FMT)
    idleTime = str(idleTime)[5:7]
    if int(idleTime) > 1:
        googleAuthenticate()


def googleAuthenticate():
    # json credentials you downloaded earlier
    json_key = json.load(open('C:/printer/flaskr/creds.json'))
    scope = ["https://spreadsheets.google.com/feeds",
             'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        'C:/printer/flaskr/creds.json', scope)  # get email and key from creds
    file = gspread.authorize(credentials)  # authenticate with Google
    globals()['client'] = gspread.authorize(credentials)
    authTime = time.ctime()
    globals()['authTime'] = authTime[11:19]
    # print(authTime[11:19])
    # print(curTime[11:19])
    print('ReAuth')


def id_Gen(fname, lname, curTime):
    # __location__ = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__)))
    # logo = open(os.path.join(__location__, 'logo.png'));
    pdf = FPDF('L', 'mm', (54, 102))
    pdf.set_margins(left=5, top=2, right=5)
    pdf.set_auto_page_break(False)
    pdf.add_page()

    #pdf.image(R'E:\logo.png', w = 35, h = 11)
    pdf.image(os.getcwd()+'/logo.png', w=35, h=11)
    pdf.set_font("Arial", size=22)
    pdf.cell(0, 17, txt="VISITOR PASS", ln=1, align="C")
    pdf.set_font("Arial", size=22)
    pdf.cell(0, 3, txt="{} {}".format(fname, lname), ln=1, border=0, align="C")
    pdf.set_font("Arial", size=12)
    pdf.cell(0, 13, txt=curTime, ln=1,  align="C")

    #pdf.line(10, 10, 10, 100)
    pdf.set_line_width(.5)
    pdf.set_draw_color(255, 0, 0)
    pdf.line(5, 15, 97, 15)
    pdf.output(os.getcwd()+"/entry.pdf")

    os.system('lp entry.pdf')

#     time.sleep(5)
#     print('Killing Adobe')
#     os.system("TASKKILL /F /IM AcroRd32.exe")
