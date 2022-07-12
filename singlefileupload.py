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


app=Flask(__name__)

app.secret_key = "secret key"
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

path = os.getcwd()
# file Upload
UPLOAD_FOLDER = os.path.join(path, 'uploads')

if not os.path.isdir(UPLOAD_FOLDER):
    os.mkdir(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


ALLOWED_EXTENSIONS = set(['csv','xslx'])


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def upload_form():
    return render_template('upload.html')


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
            csv_reader_all = csv.reader(open(app.config['UPLOAD_FOLDER'] + '/'+filename, 'r', encoding='UTF-8'), delimiter=',')
            count = 0
            flag_input = request.form['flag']
            name = request.form['name']
            flag = int(flag_input)
            line_count = 0

            workbook = xlsxwriter.Workbook('parsed/'+name+'.xlsx')
            worksheet = workbook.add_worksheet()

            workbook1 = xlsxwriter.Workbook('parsed/'+name+'-p2.xlsx')
            worksheet1 = workbook1.add_worksheet()

            workbook2 = xlsxwriter.Workbook('parsed/'+name+'-p3.xlsx')
            worksheet2 = workbook2.add_worksheet()

            workbook3 = xlsxwriter.Workbook('parsed/'+name+'-p4.xlsx')
            worksheet3 = workbook3.add_worksheet()

            workbook4 = xlsxwriter.Workbook('parsed/'+name+'-p5.xlsx')
            worksheet4 = workbook4.add_worksheet()

            workbook5 = xlsxwriter.Workbook('parsed/'+name+'-p6.xlsx')
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
        
                code = ''.join(random.choices(string.ascii_uppercase + string.digits, k=S))
                line[7] = str(code)
        
        
                line[6] = utm+'-p'+str(flag)
                line[4] = "https://contact788081.typeform.com/to/u4CNV4lF?utm_source="+line[6]+"&name="+line[0]+"&surname="+line[1]+"&email="+line[2]+"&phone="+line[3]+"&code="+line[7]
        
                if count < 50000 :
                    worksheet.write(line_count, 0, line[0])
                    worksheet.write(line_count, 1, line[1])
                    worksheet.write(line_count, 2, line[2])
                    worksheet.write(line_count, 3, line[3])
                    worksheet.write(line_count, 4, line[4])
                    worksheet.write(line_count, 5, line[5])
                    worksheet.write(line_count, 6, line[6])
                    worksheet.write(line_count, 7, line[7])
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
            flash('Nombre de lignes détéctées dans le fichier :'+ count_str)
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
            
            filenames = next(walk("parsed"), (None, None, []))[2]  # [] if no file

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
                        with open( "parsed_"+ file, "wb") as f:
                            f.write(r.content)
                    else:
                        print(r.status_code)
                        print(r.content)
        
                else:
                    print(file)
        
            print('all file finish')
                

        flash('File successfully uploaded')
        return render_template("content.html")
        
    else:
        flash('Allowed file types are only csv')
        return redirect(request.url)

@app.route('/zipped_data')
def zipped_data():
    timestr = time.strftime("%Y%m%d-%H%M%S")
    fileName = "parsed_files{}.zip".format(timestr)
    memory_file = BytesIO()
    file_path = 'ready'
    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
          for root, dirs, files in os.walk(file_path):
                    for file in files:
                              zipf.write(os.path.join(root, file))
    memory_file.seek(0)
    return send_file(memory_file,
                     attachment_filename=fileName,
                     as_attachment=True)

if __name__ == "__main__":
    app.run(host = '127.0.0.1',port = 5000, debug = False)
