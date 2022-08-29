from fileinput import filename
import os
from os import walk
import csv
from unicodedata import name
import xlsxwriter
import string
import random
from flask import Flask, flash, request, redirect, render_template, send_file, url_for, make_response, after_this_request
from werkzeug.utils import secure_filename
import time
import zipfile
from io import BytesIO
import requests
import tempfile
import glob


app=Flask(__name__, static_folder='./static', static_url_path='/')
app.secret_key = "secret key"
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024


path = os.getcwd()
# file Upload
UPLOAD_FOLDER = os.path.join(path, 'uploads')
PARSED = os.path.join(path, 'parsed')

if not os.path.isdir(UPLOAD_FOLDER):
    os.mkdir(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PARSED'] = PARSED



ALLOWED_EXTENSIONS = set(['csv','xslx'])


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def upload_form():
    return render_template('upload.html')

@app.route('/list', methods = ['GET', 'POST'])
def list():
    return render_template("list.html")

@app.route('/upload_mms', methods = ['GET', 'POST'])
def upload_mms():
    return render_template("upload_mms.html")

@app.route('/', methods=['GET','POST'])
def upload_file():
    if request.method == 'POST':
        render_template('content.html')
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        utm = request.form['utm']
        if file.filename == '':
            flash('No file selected for uploading')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            file = open(app.config['UPLOAD_FOLDER'] + '/'+filename,"r")
            csv_reader_all = csv.reader(open(app.config['UPLOAD_FOLDER'] + '/'+filename, 'r', encoding='UTF-8'), delimiter=';')
            count = 0
            flag_input = request.form['flag']
            name = request.form['name']
            flag = int(flag_input)
            link = request.form['link']
            link_cutted = link[38:46]

            line_count = 0

            workbook = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'.xlsx'))
            worksheet = workbook.add_worksheet()

            workbook1 = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'-p2.xlsx'))
            worksheet1 = workbook1.add_worksheet()

            workbook2 = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'-p3.xlsx'))
            worksheet2 = workbook2.add_worksheet()

            workbook3 = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'-p4.xlsx'))
            worksheet3 = workbook3.add_worksheet()

            workbook4 = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'-p5.xlsx'))
            worksheet4 = workbook4.add_worksheet()

            workbook5 = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'-p6.xlsx'))
            worksheet5 = workbook5.add_worksheet()


            for line in csv_reader_all:
             if(count == 0):
                first_line = line
                #writer.writerow(line)
                worksheet.write(0,0,first_line[0])
                worksheet.write(0,1, first_line[1])
                worksheet.write(0,2, first_line[2])
                worksheet.write(0,3, first_line[3])
                worksheet.write(0,4, first_line[4])
                worksheet.write(0,5, first_line[5])
                worksheet.write(0,6, first_line[6])
                worksheet.write(0,7, first_line[7])
                worksheet.write(0,8, first_line[8])

             else:
                if line[5] == 'Mme':
                    line[5] = line[5]
                elif line[5] == 'm':
                    line[5] = 'M'
                elif line[5] == 'f':
                    line[5] = 'Mme'
                else:
                    line[5] = 'M'
        
        
                line[3] = str(line[3])
        
                line[0] = line[0].replace("√®",'è')
                line[0] = line[0].replace("√©", 'é')
                line[0] = line[0].replace("√´", 'ë')
                line[0] = line[0].replace("√ß", 'ç')
                line[0] = line[0].replace("√™", 'ê')
                line[0] = line[0].replace("√£¬ß", 'ç')
                line[0] = line[0].replace("√Ø", 'ï')


                line[1] = line[1].replace("√®", 'è')
                line[1] = line[1].replace("√©", 'é')
                line[1] = line[1].replace("√´", 'ë')
                line[1] = line[1].replace("√ß", 'ç')
                line[1] = line[1].replace("√™", 'ê')
                line[1] = line[1].replace("√£¬ß", 'ç')
                line[1] = line[1].replace("√Ø", 'ï')
        
                S = 5  # number of characters in the string.
                # call random.choices() string module to find the string in Uppercase + numeric data.
        
                ran = ''.join(random.choices(string.ascii_uppercase + string.digits, k=S))
                code = ran.replace ("0", "5")
                line[7] = str(code)
                if flag > 1:
                    line[6] = utm+'-p'+str(flag)
                else:
                    line[6] = utm
                line[4] = "https://contact788081.typeform.com/to/"+link_cutted+"?utm_source="+line[6]+"&prenom="+line[1]+"&nom="+line[0]+"&email="+line[2]+"&telephone="+line[3]+"&code="+line[7]+"&civilite="+line[5]+"&code_postal="+line[8]
        
                if count < 50000 :
                    worksheet.write(line_count, 0, line[0])
                    worksheet.write(line_count, 1, line[1])
                    worksheet.write(line_count, 2, line[2])
                    worksheet.write(line_count, 3, line[3])
                    worksheet.write(line_count, 4, line[4])
                    worksheet.write(line_count, 5, line[5])
                    worksheet.write(line_count, 6, line[6])
                    worksheet.write(line_count, 7, line[7])
                    worksheet.write(line_count, 8, line[8])
                elif count > 50000 and count < 100001:
                    worksheet1.write(line_count, 0, line[0])
                    worksheet1.write(line_count, 1, line[1])
                    worksheet1.write(line_count, 2, line[2])
                    worksheet1.write(line_count, 3, line[3])
                    worksheet1.write(line_count, 4, line[4])
                    worksheet1.write(line_count, 5, line[5])
                    worksheet1.write(line_count, 6, line[6])
                    worksheet1.write(line_count, 7, line[7])
                elif count > 100001 and count < 150000:
                    worksheet2.write(line_count, 0, line[0])
                    worksheet2.write(line_count, 1, line[1])
                    worksheet2.write(line_count, 2, line[2])
                    worksheet2.write(line_count, 3, line[3])
                    worksheet2.write(line_count, 4, line[4])
                    worksheet2.write(line_count, 5, line[5])
                    worksheet2.write(line_count, 6, line[6])
                    worksheet2.write(line_count, 7, line[7])
                elif count > 150000 and count < 200000:
                    worksheet3.write(line_count, 0, line[0])
                    worksheet3.write(line_count, 1, line[1])
                    worksheet3.write(line_count, 2, line[2])
                    worksheet3.write(line_count, 3, line[3])
                    worksheet3.write(line_count, 4, line[4])
                    worksheet3.write(line_count, 5, line[5])
                    worksheet3.write(line_count, 6, line[6])
                    worksheet3.write(line_count, 7, line[7])
                elif count > 200000 and count < 250000:
                    worksheet4.write(line_count, 0, line[0])
                    worksheet4.write(line_count, 1, line[1])
                    worksheet4.write(line_count, 2, line[2])
                    worksheet4.write(line_count, 3, line[3])
                    worksheet4.write(line_count, 4, line[4])
                    worksheet4.write(line_count, 5, line[5])
                    worksheet4.write(line_count, 6, line[6])
                    worksheet4.write(line_count, 7, line[7])
                elif count > 250000 and count < 300000:
                    worksheet5.write(line_count, 0, line[0])
                    worksheet5.write(line_count, 1, line[1])
                    worksheet5.write(line_count, 2, line[2])
                    worksheet5.write(line_count, 3, line[3])
                    worksheet5.write(line_count, 4, line[4])
                    worksheet5.write(line_count, 5, line[5])
                    worksheet5.write(line_count, 6, line[6])
                    worksheet5.write(line_count, 7, line[7])

                if count%50000 == 0:
                    flag = flag + 1
                    line_count = 0
        
             count = count + 1
             line_count = line_count +1
             count_str = str(count)
            print(count)
            if count <= 50001 :
                workbook.close()
            elif count > 50001 and count <= 100001:
                workbook.close()
                workbook1.close()
            elif count > 100001 and count <= 150001:
                workbook.close()
                workbook1.close()
                workbook2.close()
            elif count > 150001 and count <= 200001:
                workbook.close()
                workbook1.close()
                workbook2.close()
                workbook3.close()
            elif count > 200001 and count <= 250001:
                workbook.close()
                workbook1.close()
                workbook2.close()
                workbook3.close()
                workbook4.close()
            elif count > 250001 and count <= 300001:
                workbook.close()
                workbook1.close()
                workbook3.close()
                workbook2.close()
                workbook4.close()
                workbook5.close()
            
            filenames = next(walk(os.path.abspath("parsed")), (None, None, []))[2]  # [] if no file
            print(filenames)
            count = 0
            for file in filenames:
                if (file != ".DS_Store"):
                    count = count + 1
                    print("start " + file)
                    sample_file = open("parsed/" + file, "rb")
                    upload_file = {"xlsxFile": sample_file}
                    r = requests.post("https://aud.vc/upload-file", files=upload_file)
        
                    if r.status_code == 200:
                        print("finish parsed_" + file)
                        with open(os.path.abspath("parsed/"+file), "wb") as f:
                            f.write(r.content)
                    else:
                        print(r.status_code)
                        print(r.content)
        
                else:
                    print(file)
            filenames_ = next(walk(os.path.abspath("parsed")), (None, None, []))[2]  # [] if no file
            print("Hello", filenames_)
            print('all file finish')
       

        return render_template("content.html")
        
    else:
        flash('Allowed file types are only csv')
        return redirect(request.url)


@app.route('/mms', methods=['GET','POST'])
def mms():
    if request.method == 'POST':
        render_template('content.html')
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        utm = request.form['utm']
        if file.filename == '':
            flash('No file selected for uploading')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            file = open(app.config['UPLOAD_FOLDER'] + '/'+filename,"r")
            csv_reader_all = csv.reader(open(app.config['UPLOAD_FOLDER'] + '/'+filename, 'r', encoding='UTF-8'), delimiter=';')
            count = 0
            flag_input = request.form['flag']
            name = request.form['name']
            flag = int(flag_input)

            line_count = 0

            workbook = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'-p1.xlsx'))
            worksheet = workbook.add_worksheet()

            workbook1 = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'-p2.xlsx'))
            worksheet1 = workbook1.add_worksheet()

            workbook2 = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'-p3.xlsx'))
            worksheet2 = workbook2.add_worksheet()

          ##  workbook3 = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'-p4.xlsx'))
          ##  worksheet3 = workbook3.add_worksheet()
##
          ##  workbook4 = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'-p3-1.xlsx'))
          ##  worksheet4 = workbook4.add_worksheet()
##
          ##  workbook5 = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'-p3-2.xlsx'))
          ##  worksheet5 = workbook5.add_worksheet()
##
          ##  workbook6 = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'-p4-1.xlsx'))
          ##  worksheet6 = workbook6.add_worksheet()
##
          ##  workbook7 = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'-p4-2.xlsx'))
          ##  worksheet7 = workbook7.add_worksheet()
##
          ##  workbook8 = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'-p5-1.xlsx'))
          ##  worksheet8 = workbook4.add_worksheet()
##
          ##  workbook9 = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'-p5-2.xlsx'))
          ##  worksheet9 = workbook5.add_worksheet()
##
          ##  workbook10 = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'-p6-1.xlsx'))
          ##  worksheet10 = workbook10.add_worksheet()
##
          ##  workbook11 = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'-p6-2.xlsx'))
          ##  worksheet11 = workbook11.add_worksheet()



            for line in csv_reader_all:
             if(count == 0):
                first_line = line
                #writer.writerow(line)
                worksheet.write(0,0,first_line[0])
                worksheet.write(0,1, first_line[1])
                worksheet.write(0,2, first_line[2])
                worksheet.write(0,3, first_line[3])
                worksheet.write(0,4, first_line[4])
                worksheet.write(0,5, first_line[5])
                worksheet.write(0,6, first_line[6])
                worksheet.write(0,7, first_line[7])

             else:
                if line[5] == 'Mme':
                    line[5] = line[5]
                elif line[5] == 'm':
                    line[5] = 'M'
                elif line[5] == 'f':
                    line[5] = 'Mme'
                else:
                    line[5] = 'M'
        
        
                line[3] = str(line[3])
        
                line[0] = line[0].replace("√®",'è')
                line[0] = line[0].replace("√©", 'é')
                line[0] = line[0].replace("√´", 'ë')
                line[0] = line[0].replace("√ß", 'ç')
                line[0] = line[0].replace("√™", 'ê')
                line[0] = line[0].replace("√£¬ß", 'ç')
                line[0] = line[0].replace("√Ø", 'ï')


                line[1] = line[1].replace("√®", 'è')
                line[1] = line[1].replace("√©", 'é')
                line[1] = line[1].replace("√´", 'ë')
                line[1] = line[1].replace("√ß", 'ç')
                line[1] = line[1].replace("√™", 'ê')
                line[1] = line[1].replace("√£¬ß", 'ç')
                line[1] = line[1].replace("√Ø", 'ï')
        
                S = 5  # number of characters in the string.
                # call random.choices() string module to find the string in Uppercase + numeric data.
        
                ran = ''.join(random.choices(string.ascii_uppercase + string.digits, k=S))
                code = ran.replace ("0", "3")
                line[7] = str(code)
                if flag > 0:
                    line[6] = '46.'+utm+'.p'+str(flag)
                else:
                    line[6] = utm
                line[4] = "https://contact788081.typeform.com/to/u4CNV4lF?utm_source="+line[6]+"&name="+line[0]+"&surname="+line[1]+"&email="+line[2]+"&phone="+line[3]+"&code="+line[7]

                line[1] = line[2] = ""

                if count < 50001 :
                    worksheet.write(line_count, 0, line[0])
                    worksheet.write(line_count, 1, line[1])
                    worksheet.write(line_count, 2, line[2])
                    worksheet.write(line_count, 3, line[3])
                    worksheet.write(line_count, 4, line[4])
                    worksheet.write(line_count, 5, line[5])
                    worksheet.write(line_count, 6, line[6])
                    worksheet.write(line_count, 7, line[7])
                elif count > 50001 and count < 100001:
                    worksheet1.write(line_count, 0, line[0])
                    worksheet1.write(line_count, 1, line[1])
                    worksheet1.write(line_count, 2, line[2])
                    worksheet1.write(line_count, 3, line[3])
                    worksheet1.write(line_count, 4, line[4])
                    worksheet1.write(line_count, 5, line[5])
                    worksheet1.write(line_count, 6, line[6])
                    worksheet1.write(line_count, 7, line[7])
                elif count > 100001 and count < 150001:
                    worksheet2.write(line_count, 0, line[0])
                    worksheet2.write(line_count, 1, line[1])
                    worksheet2.write(line_count, 2, line[2])
                    worksheet2.write(line_count, 3, line[3])
                    worksheet2.write(line_count, 4, line[4])
                    worksheet2.write(line_count, 5, line[5])
                    worksheet2.write(line_count, 6, line[6])
                    worksheet2.write(line_count, 7, line[7])
            ## elif count > 75001 and count < 100001:
            ##     worksheet3.write(line_count, 0, line[0])
            ##     worksheet3.write(line_count, 1, line[1])
            ##     worksheet3.write(line_count, 2, line[2])
            ##     worksheet3.write(line_count, 3, line[3])
            ##     worksheet3.write(line_count, 4, line[4])
            ##     worksheet3.write(line_count, 5, line[5])
            ##     worksheet3.write(line_count, 6, line[6])
            ##     worksheet3.write(line_count, 7, line[7])
            ## elif count > 100001 and count < 125001:
            ##     worksheet4.write(line_count, 0, line[0])
            ##     worksheet4.write(line_count, 1, line[1])
            ##     worksheet4.write(line_count, 2, line[2])
            ##     worksheet4.write(line_count, 3, line[3])
            ##     worksheet4.write(line_count, 4, line[4])
            ##     worksheet4.write(line_count, 5, line[5])
            ##     worksheet4.write(line_count, 6, line[6])
            ##     worksheet4.write(line_count, 7, line[7])
            ## elif count > 125001 and count < 150001:
            ##     worksheet5.write(line_count, 0, line[0])
            ##     worksheet5.write(line_count, 1, line[1])
            ##     worksheet5.write(line_count, 2, line[2])
            ##     worksheet5.write(line_count, 3, line[3])
            ##     worksheet5.write(line_count, 4, line[4])
            ##     worksheet5.write(line_count, 5, line[5])
            ##     worksheet5.write(line_count, 6, line[6])
            ##     worksheet5.write(line_count, 7, line[7])
            ## elif count > 150001 and count < 175001:
            ##     worksheet6.write(line_count, 0, line[0])
            ##     worksheet6.write(line_count, 1, line[1])
            ##     worksheet6.write(line_count, 2, line[2])
            ##     worksheet6.write(line_count, 3, line[3])
            ##     worksheet6.write(line_count, 4, line[4])
            ##     worksheet6.write(line_count, 5, line[5])
            ##     worksheet6.write(line_count, 6, line[6])
            ##     worksheet6.write(line_count, 7, line[7])
            ## elif count > 175001 and count < 200001:
            ##     worksheet7.write(line_count, 0, line[0])
            ##     worksheet7.write(line_count, 1, line[1])
            ##     worksheet7.write(line_count, 2, line[2])
            ##     worksheet7.write(line_count, 3, line[3])
            ##     worksheet7.write(line_count, 4, line[4])
            ##     worksheet7.write(line_count, 5, line[5])
            ##     worksheet7.write(line_count, 6, line[6])
            ##     worksheet7.write(line_count, 7, line[7])
            ## elif count > 200001 and count < 225001:
            ##     worksheet8.write(line_count, 0, line[0])
            ##     worksheet8.write(line_count, 1, line[1])
            ##     worksheet8.write(line_count, 2, line[2])
            ##     worksheet8.write(line_count, 3, line[3])
            ##     worksheet8.write(line_count, 4, line[4])
            ##     worksheet8.write(line_count, 5, line[5])
            ##     worksheet8.write(line_count, 6, line[6])
            ##     worksheet8.write(line_count, 7, line[7])
            ## elif count > 225001 and count < 250001:
            ##     worksheet9.write(line_count, 0, line[0])
            ##     worksheet9.write(line_count, 1, line[1])
            ##     worksheet9.write(line_count, 2, line[2])
            ##     worksheet9.write(line_count, 3, line[3])
            ##     worksheet9.write(line_count, 4, line[4])
            ##     worksheet9.write(line_count, 5, line[5])
            ##     worksheet9.write(line_count, 6, line[6])
            ##     worksheet9.write(line_count, 7, line[7])
            ## elif count > 250001 and count < 275001:
            ##     worksheet10.write(line_count, 0, line[0])
            ##     worksheet10.write(line_count, 1, line[1])
            ##     worksheet10.write(line_count, 2, line[2])
            ##     worksheet10.write(line_count, 3, line[3])
            ##     worksheet10.write(line_count, 4, line[4])
            ##     worksheet10.write(line_count, 5, line[5])
            ##     worksheet10.write(line_count, 6, line[6])
            ##     worksheet10.write(line_count, 7, line[7])
            ## elif count > 275001 and count < 300001:
            ##     worksheet11.write(line_count, 0, line[0])
            ##     worksheet11.write(line_count, 1, line[1])
            ##     worksheet11.write(line_count, 2, line[2])
            ##     worksheet11.write(line_count, 3, line[3])
            ##     worksheet11.write(line_count, 4, line[4])
            ##     worksheet11.write(line_count, 5, line[5])
            ##     worksheet11.write(line_count, 6, line[6])
            ##     worksheet11.write(line_count, 7, line[7])

                if count%50000 == 0:
                    flag = flag + 1
                    line_count = 0
        
             count = count + 1
             line_count = line_count +1
             count_str = str(count)
            print(count)
            if count <= 50001 :
                workbook.close()
            elif count > 50001 and count <= 100001:
                workbook.close()
                workbook1.close()
            elif count > 100001 and count <= 150001:
                workbook.close()
                workbook1.close()
                workbook2.close()
            elif count > 150001 and count <= 200001:
                workbook.close()
                workbook1.close()
                workbook2.close()
                workbook3.close()
        ##  elif count > 100001 and count <= 125001:
        ##      workbook.close()
        ##      workbook1.close()
        ##      workbook2.close()
        ##      workbook3.close()
        ##      workbook4.close()
        ##  elif count > 125001 and count <= 150001:
        ##      workbook.close()
        ##      workbook1.close()
        ##      workbook3.close()
        ##      workbook2.close()
        ##      workbook4.close()
        ##      workbook5.close()
        ##  elif count > 150001 and count <= 175001:
        ##      workbook.close()
        ##      workbook1.close()
        ##      workbook3.close()
        ##      workbook2.close()
        ##      workbook4.close()
        ##      workbook5.close()
        ##      workbook6.close()
        ##  elif count > 175001 and count <= 200001:
        ##      workbook.close()
        ##      workbook1.close()
        ##      workbook3.close()
        ##      workbook2.close()
        ##      workbook4.close()
        ##      workbook5.close()
        ##      workbook6.close()
        ##      workbook7.close()
        ##  elif count > 200001 and count <= 225001:
        ##      workbook.close()
        ##      workbook1.close()
        ##      workbook3.close()
        ##      workbook2.close()
        ##      workbook4.close()
        ##      workbook5.close()
        ##      workbook6.close() 
        ##      workbook7.close()
        ##      workbook8.close()
        ##  elif count > 225001 and count <= 250001:
        ##      workbook.close()
        ##      workbook1.close()
        ##      workbook3.close()
        ##      workbook2.close()
        ##      workbook4.close()
        ##      workbook5.close()
        ##      workbook6.close()
        ##      workbook7.close()
        ##      workbook8.close()
        ##      workbook9.close()
        ##  elif count > 250001 and count <= 275001:
        ##      workbook.close()
        ##      workbook1.close()
        ##      workbook3.close()
        ##      workbook2.close()
        ##      workbook4.close()
        ##      workbook5.close()
        ##      workbook6.close()
        ##      workbook7.close()
        ##      workbook8.close()
        ##      workbook9.close()
        ##      workbook10.close()
        ##  elif count > 275001 and count <= 300001:
        ##      workbook.close()
        ##      workbook1.close()
        ##      workbook3.close()
        ##      workbook2.close()
        ##      workbook4.close()
        ##      workbook5.close()
        ##      workbook6.close()
        ##      workbook7.close()
        ##      workbook8.close()
        ##      workbook9.close()
        ##      workbook10.close()
        ##      workbook11.close()

            
            filenames = next(walk(os.path.abspath("parsed")), (None, None, []))[2]  # [] if no file
            print(filenames)
            count = 0
            for file in filenames:
                if (file != ".DS_Store"):
                    count = count + 1
                    print("start " + file)
                    sample_file = open("parsed/" + file, "rb")
                    upload_file = {"xlsxFile": sample_file}
                    r = requests.post("https://sma.vc/upload-file", files=upload_file)
        
                    if r.status_code == 200:
                        print("finish parsed_" + file)
                        with open(os.path.abspath("parsed/"+file), "wb") as f:
                            f.write(r.content)
                    else:
                        print(r.status_code)
                        print(r.content)
        
                else:
                    print(file)
            filenames_ = next(walk(os.path.abspath("parsed")), (None, None, []))[2]  # [] if no file
            print("Hello", filenames_)
            print('all file finish')
       

        return render_template("content.html")
        
    else:
        flash('Allowed file types are only csv')
        return redirect(request.url)


@app.route('/zipped_data')
def zipped_data():
    timestr = time.strftime("%Y%m%d-%H%M%S")
    fileName = "parsed_files{}.zip".format(timestr)
    memory_file = BytesIO()
    file_path = os.path.abspath("parsed")
    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
          for root, dirs, files in os.walk(file_path):
                    for file in files:
                              zipf.write(os.path.join(root, file))
    memory_file.seek(0)
    filenames2 = next(walk(os.path.abspath("parsed")), (None, None, []))[2]  # [] if no file
    for file in filenames2:
                if (file != ".DS_Store"):
                    file_path_del_2 = (os.path.abspath("parsed/"+file))
                    os.remove(file_path_del_2)
    filenames3 = next(walk(os.path.abspath("uploads")), (None, None, []))[2]  # [] if no file
    for file in filenames3:
                if (file != ".DS_Store"):
                    file_path_del_3 = (os.path.abspath("uploads/"+file))
                    os.remove(file_path_del_3)

    print(memory_file)


    return send_file(memory_file,
                     attachment_filename=fileName,
                     as_attachment=True)

if __name__ == "__main__":
    app.run(host = '0.0.0.0',port = 8080, debug = False)
