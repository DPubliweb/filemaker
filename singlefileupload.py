from fileinput import filename
import os
from os import walk
import csv
from unicodedata import name
import xlsxwriter
import string
import random
from flask import Flask, flash, request, redirect, render_template, send_file, url_for, make_response, after_this_request, session
from werkzeug.utils import secure_filename
import time
import zipfile
from io import BytesIO
import requests
import pandas as pd

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

gsm_to_unicode = {
    '@': '\u0000', '£': '\u0001', '$': '\u0002', '¥': '\u0003', 'è': '\u0004',
    'é': '\u00E9', 'ù': '\u0006', 'ì': '\u0007', 'ò': '\u0008', 'Ç': '\u0009',
    '\n': '\u000A', 'Ø': '\u000B', 'ø': '\u000C', '\r': '\u000D', 'Å': '\u000E',
    'å': '\u000F', 'Δ': '\u0010', '_': '\u0011', 'Φ': '\u0012', 'Γ': '\u0013',
    'Λ': '\u0014', 'Ω': '\u0015', 'Π': '\u0016', 'Ψ': '\u0017', 'Σ': '\u0018',
    'Θ': '\u0019', 'Ξ': '\u001A', 'Æ': '\u001C', 'æ': '\u001D', 'ß': '\u001E',
    'É': '\u001F', ' ': '\u0020', '!': '\u0021', '\"': '\u0022', '#': '\u0023',
    '¤': '\u0024', '%': '\u0025', '&': '\u0026', '\'': '\u0027', '(': '\u0028',
    ')': '\u0029', '*': '\u002A', '+': '\u002B', ',': '\u002C', '-': '\u002D',
    '.': '\u002E', '/': '\u002F', '0': '\u0030', '1': '\u0031', '2': '\u0032',
    '3': '\u0033', '4': '\u0034', '5': '\u0035', '6': '\u0036', '7': '\u0037',
    '8': '\u0038', '9': '\u0039', ':': '\u003A', ';': '\u003B', '<': '\u003C',
    '=': '\u003D', '>': '\u003E', '?': '\u003F', '¡': '\u0040', 'A': '\u0041',
    'B': '\u0042', 'C': '\u0043', 'D': '\u0044', 'E': '\u0045', 'F': '\u0046',
    'G': '\u0047', 'H': '\u0048', 'I': '\u0049', 'J': '\u004A', 'K': '\u004B',
    'L': '\u004C', 'M': '\u004D', 'N': '\u004E', 'O': '\u004F', 'P': '\u0050',
    'Q': '\u0051', 'R': '\u0052', 'S': '\u0053', 'T': '\u0054', 'U': '\u0055',
    'V': '\u0056', 'W': '\u0057', 'X': '\u0058', 'Y': '\u0059', 'Z': '\u005A',
    'Ä': '\u005B', 'Ö': '\u005C', 'Ñ': '\u005D', 'Ü': '\u005E', '§': '\u005F',
    '¿': '\u0060', 'a': '\u0061', 'b': '\u0062', 'c': '\u0063', 'd': '\u0064',
    'e': '\u0065', 'f': '\u0066', 'g': '\u0067', 'h': '\u0068', 'i': '\u0069',
    'j': '\u006A', 'k': '\u006B', 'l': '\u006C', 'm': '\u006D', 'n': '\u006E',
    'o': '\u006F', 'p': '\u0070', 'q': '\u0071', 'r': '\u0072', 's': '\u0073',
    't': '\u0074', 'u': '\u0075', 'v': '\u0076', 'w': '\u0077', 'x': '\u0078',
    'y': '\u0079', 'z': '\u007A', 'ä': '\u007B', 'ö': '\u007C', 'ñ': '\u007D',
    'ü': '\u007E', 'à': '\u007F'
}

def convert_to_gsm(text):
    # Convert to GSM 03.38
    return "".join([gsm_to_unicode.get(char, '?') for char in text])

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods = ['GET'])
def home():
  return render_template('login.html')

@app.route('/', methods = ['POST'])
def login():
    session['completed'] = False
    identifiant = request.form['identifiant']
    password = request.form['password']
    if (identifiant == 'contact@publiweb.agency' and password == 'Nanah148148'):
        session['completed'] = True
        return render_template('upload.html')
    else:
        return render_template('login.html')

@app.route('/refused')
def refused():
    return render_template('refused.html')

@app.route('/lea')
def lea_upload_form():
    print(session.get('completed', None))
    if session.get('completed', None) == True :
        return render_template('lea.html')
    else:
        return render_template('refused.html')

@app.route('/b2b')
def b2b():
    return render_template('b2b.html')

@app.route('/upload')
def upload_form():
    print(session.get('completed', None))
    if session.get('completed', None) == True :
        return render_template('upload.html')
    else:
        return render_template('refused.html')

@app.route('/sms', methods = ['GET', 'POST'])
def sms():
    if session.get('completed', None) == True :
        return render_template('sms.html')
    else:
        return render_template('refused.html')

@app.route('/upload', methods=['GET','POST'])
def upload_file():
    if request.method == 'POST':
        render_template('content.html')
    # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        utm = request.form['utm']
        url_shortenner = request.form ['url_shortenner']
        if file.filename == '':
            flash('No file selected for uploading')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            file = open(app.config['UPLOAD_FOLDER'] + '/'+filename,"r")
            csv_reader_all = csv.reader(open(app.config['UPLOAD_FOLDER'] + '/'+filename, 'r', encoding='UTF-8'), delimiter=';')
            count = 0
            #flag_input = request.form['flag']
            name = request.form['name']
            #flag = int(flag_input)
            link = request.form['link']
            link_cutted = link[38:46]

            line_count = 0
            more_than = 0

            workbook = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'.xlsx'))
            worksheet = workbook.add_worksheet()

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
                worksheet.write(0,9, first_line[9])
                worksheet.write(0,10, first_line[10])

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
                code = ran.replace ("0", "1")
                line[6] = str(code)
                #line[6] = utm
                line[4] = "https://contact788081.typeform.com/to/"+link_cutted+"?utm_source="+line[8]+"&prenom="+line[1]+"&nom="+line[0]+"&email="+line[2]+"&telephone="+line[3]+"&code="+line[6]+"&civilite="+line[5]+"&code_postal="+line[7]+"&cohort="+line[9]+"&analytics="+line[10]

                

               
                if count < 50000:
                    worksheet.write(line_count, 0, line[0])
                    worksheet.write(line_count, 1, line[1])
                    worksheet.write(line_count, 2, line[2])
                    worksheet.write(line_count, 3, line[3])
                    worksheet.write(line_count, 4, line[4])
                    worksheet.write(line_count, 5, line[5])
                    worksheet.write(line_count, 6, line[6])
                    worksheet.write(line_count, 7, line[7])
                    worksheet.write(line_count, 8, line[8])
                    worksheet.write(line_count, 9, line[9])
             
             
             count = count + 1
             line_count = line_count +1
             count_str = str(count)
            print(count)
            print(str(url_shortenner))
            workbook.close()
            

            # Récupérer la liste des fichiers
            filenames = next(walk(os.path.abspath("parsed")), (None, None, []))[2]
            
            # Compteur pour le nombre de fichiers traités
            count = 0
            
            # Boucle à travers chaque fichier
            for file in filenames:
                if file != ".DS_Store" and file != "text.html":
                    count += 1
                    print(f"Start {file}")
                    
                    # Utiliser with pour s'assurer que le fichier est fermé après utilisation
                    with open(os.path.join("parsed", file), "rb") as sample_file:
                        upload_file = {"xlsxFile": sample_file}
                        try:
                            # Envoi de la requête avec un délai d'attente de 10 secondes
                            r = requests.post("https://"+str(url_shortenner)+"/upload-file", files=upload_file)
                            if r.status_code == 200:
                                print(f"Finish parsed_{file}")
                                with open(os.path.abspath(os.path.join("parsed", file)), "wb") as f:
                                    f.write(r.content)
                            else:
                                print(r.status_code)
                                print(r.content)
                        except requests.exceptions.RequestException as e:
                            # Gestion des exceptions de requête
                            print(f"Error: {e}")
            
                    # Pause de 1 seconde entre les requêtes pour éviter la surcharge du serveur
                    time.sleep(1)

            
            filenames = next(walk(os.path.abspath("parsed")), (None, None, []))[2]  # [] if no file
            for file in filenames:
                if (file != ".DS_Store" and file != "text.html" ):
                    read_file = pd.read_excel('parsed/'+file)
                    file = file[:-4]
                    read_file.to_csv("ready/"+file+"csv", index=None, header=True)
            print("Salut", filenames)
            print('all file finish')


        return render_template("content.html")
    
    else:
        flash('Allowed file types are only csv')
        return redirect(request.url)

@app.route('/upload_lea', methods=['GET','POST'])
def upload_file_lea():
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
                #flag_input = request.form['flag']
                name = request.form['name']
                #flag = int(flag_input)
                link = request.form['link']
                link_cutted = link[38:46]
    
                line_count = 0
                more_than = 0
    
                workbook = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'.p1.xlsx'))
                worksheet = workbook.add_worksheet()
                workbook1 = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'.p2.xlsx'))
                worksheet1 = workbook1.add_worksheet()
                workbook2 = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'.p3.xlsx'))
                worksheet2 = workbook2.add_worksheet()
                workbook3 = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'.p4.xlsx'))
                worksheet3 = workbook3.add_worksheet()
                workbook4 = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'.p5.xlsx'))
                worksheet4 = workbook4.add_worksheet()
                more_than = 1
                
                    
    
    
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
                    worksheet.write(0,9, first_line[9])
    
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
                    code = ran.replace ("0", "2")
                    line[8] = str(code)
                    if count < 50000 :
                       if more_than == 0:
                           line[6] = utm
                       else:
                           line[6] = utm+".p1"
                    elif count > 50000 and count <= 100001:
                       line[6] = utm+".p2"
                    elif count > 100001 and count <= 150001:
                        line[6] = utm+".p3"
                    elif count > 150001 and count <= 200001:
                        line[6] = utm+".p4"
                    elif count > 200001 and count <= 250001:
                        line[6] = utm+".p5"
    
                    line[4] = "https://contact788081.typeform.com/to/"+link_cutted+"?utm_source="+line[6]+"&prenom="+line[1]+"&nom="+line[0]+"&email="+line[2]+"&telephone="+line[3]+"&code="+line[8]+"&civilite="+line[5]+"&code_postal="+line[7]+"&cohort="+line[9]
                   
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
                        worksheet.write(line_count, 9, line[9])
                    elif count > 50000 and count <= 100001:
                        worksheet1.write(line_count, 0, line[0])
                        worksheet1.write(line_count, 1, line[1])
                        worksheet1.write(line_count, 2, line[2])
                        worksheet1.write(line_count, 3, line[3])
                        worksheet1.write(line_count, 4, line[4])
                        worksheet1.write(line_count, 5, line[5])
                        worksheet1.write(line_count, 6, line[6])
                        worksheet1.write(line_count, 7, line[7])
                        worksheet1.write(line_count, 8, line[8])
                    elif count > 100001 and count <= 150000:
                        worksheet2.write(line_count, 0, line[0])
                        worksheet2.write(line_count, 1, line[1])
                        worksheet2.write(line_count, 2, line[2])
                        worksheet2.write(line_count, 3, line[3])
                        worksheet2.write(line_count, 4, line[4])
                        worksheet2.write(line_count, 5, line[5])
                        worksheet2.write(line_count, 6, line[6])
                        worksheet2.write(line_count, 7, line[7])
                        worksheet2.write(line_count, 8, line[8])
                    elif count > 150000 and count <= 200001:
                        worksheet3.write(line_count, 0, line[0])
                        worksheet3.write(line_count, 1, line[1])
                        worksheet3.write(line_count, 2, line[2])
                        worksheet3.write(line_count, 3, line[3])
                        worksheet3.write(line_count, 4, line[4])
                        worksheet3.write(line_count, 5, line[5])
                        worksheet3.write(line_count, 6, line[6])
                        worksheet3.write(line_count, 7, line[7])
                        worksheet3.write(line_count, 8, line[8])
                    elif count > 200001 and count <= 250001:
                        worksheet4.write(line_count, 0, line[0])
                        worksheet4.write(line_count, 1, line[1])
                        worksheet4.write(line_count, 2, line[2])
                        worksheet4.write(line_count, 3, line[3])
                        worksheet4.write(line_count, 4, line[4])
                        worksheet4.write(line_count, 5, line[5])
                        worksheet4.write(line_count, 6, line[6])
                        worksheet4.write(line_count, 7, line[7])
                        worksheet4.write(line_count, 8, line[8])
    
                 count = count + 1
                 line_count = line_count +1
                 count_str = str(count)
                print(count)
                if count <= 50000 :
                    workbook.close()
                elif count > 50000 and count <= 100001:
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
                
                filenames = next(walk(os.path.abspath("parsed")), (None, None, []))[2]  # [] if no file
                print(filenames)
                count = 0
                for file in filenames:
                    if (file != ".DS_Store" and file != "text.html"):
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
                        
                filenames = next(walk(os.path.abspath("parsed")), (None, None, []))[2]  # [] if no file
                for file in filenames:
                    if (file != ".DS_Store" and file != "text.html" ):
                        read_file = pd.read_excel('parsed/'+file)
                        file = file[:-4]
                        read_file.to_csv("ready/"+file+"csv", index=None, header=True)
                print("Salut", filenames)
                print('all file finish')
            return render_template("content.html")
        
    else:
            flash('Allowed file types are only csv')
            return redirect(request.url)

@app.route('/upload_b2b', methods=['GET','POST'])
def upload_b2b():
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
            #flag_input = request.form['flag']
            name = request.form['name']
            #flag = int(flag_input)
            link = request.form['link']
            link_cutted = link[38:46]

            line_count = 0

            workbook = xlsxwriter.Workbook(os.path.abspath('parsed/'+name+'.xlsx'))
            worksheet = workbook.add_worksheet()

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
                worksheet.write(0,9, first_line[9])

             else:

                line[9] = utm
                line[4] = "https://contact788081.typeform.com/to/"+link_cutted+"?utm_source="+line[9]+"&address="+line[2]+"&company="+line[0]+"&email="+line[1]+"&tel_fixe="+line[6]+"&tel_mobile="+line[3]+"&website="+line[8]+"&city="+line[5]+"&zipcode="+line[7]
               
                line[1] = line[2] = line[6] = line[8] = line[9] = ""
                
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
        
             count = count + 1
             line_count = line_count +1
             count_str = str(count)
            print(count)
            if count <= 50001 :
                workbook.close()
            
            filenames = next(walk(os.path.abspath("parsed")), (None, None, []))[2]  # [] if no file
            print(filenames)
            count = 0
            for file in filenames:
                if (file != ".DS_Store" and file != "text.html"):
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
            filenames = next(walk(os.path.abspath("parsed")), (None, None, []))[2]  # [] if no file
            for file in filenames:
                if (file != ".DS_Store" and file != "text.html"):
                    read_file = pd.read_excel('parsed/'+file)
                    file = file[:-4]
                    read_file.to_csv("ready/"+file+"csv", index=None, header=True)
                    print("Salut",file)
            print("Salut", filenames)
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

    file_path = os.path.abspath("ready")

    # Création de l'archive ZIP en mémoire
    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(file_path):
            for file in files:
                if file not in [".DS_Store", "text.html"]:
                    zipf.write(os.path.join(root, file))

    # Réinitialiser le curseur au début du fichier en mémoire
    memory_file.seek(0)

    # Suppression des fichiers
    directories = ["parsed", "uploads", "ready"]
    for directory in directories:
        dir_path = os.path.abspath(directory)
        for file in os.listdir(dir_path):
            if file not in [".DS_Store", "text.html"]:
                file_path = os.path.join(dir_path, file)
                os.remove(file_path)

    # Envoi de l'archive ZIP au client
    return send_file(memory_file,
                     attachment_filename=fileName,
                     as_attachment=True)

@app.route('/zipped_data_2')
def zipped_data_2():
    timestr = time.strftime("%Y%m%d-%H%M%S")
    fileName = "ready_files{}.zip".format(timestr)
    memory_file = BytesIO()
    file_path = os.path.abspath("final")
    
    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
          for root, dirs, files in os.walk(file_path):
                    for file in files:
                        if (file != ".DS_Store" and file != "text.html"):
                            zipf.write(os.path.join(root, file))
                            #os.remove(file_path)
    memory_file.seek(0)
    filenames2 = next(walk(os.path.abspath("final")), (None, None, []))[2]  # [] if no file
    for file in filenames2:
                if (file != ".DS_Store" and file != "text.html"):
                    file_path_del_2 = (os.path.abspath("final/"+file))
                    os.remove(file_path_del_2)
    filenames3 = next(walk(os.path.abspath("uploads")), (None, None, []))[2]  # [] if no file
    for file in filenames3:
                if (file != ".DS_Store" and file != "text.html"):
                    file_path_del_3 = (os.path.abspath("uploads/"+file))
                    os.remove(file_path_del_3)

    print(memory_file)


    return send_file(memory_file,
                     attachment_filename=fileName,
                     as_attachment=True)

@app.route('/sms_write', methods=['GET','POST'])
def sms_write():
    if request.method == 'POST':
        render_template('sms.html')
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        sms_content = request.form['sms_content']
        if file.filename == '':
            flash('No file selected for uploading')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            file = open(app.config['UPLOAD_FOLDER'] + '/'+filename,"r")
            csv_reader_all = csv.reader(open(app.config['UPLOAD_FOLDER'] + '/'+filename, 'r', encoding='UTF-8'), delimiter=',')
            name = request.form['name']
            short_url = request.form['short_url']
            counter = request.form['counter']
            count = 0
            line_count = 0
            civilite = "{civilite}"
            nom = "{nom}"
            lien = "{lien}"
            print(short_url)
            

            workbook = xlsxwriter.Workbook(os.path.abspath('final/'+name+'.xlsx'))
            worksheet = workbook.add_worksheet()

            for line in csv_reader_all:
             if(count == 0):
                first_line = line
                #writer.writerow(line)
                worksheet.write(0,0,first_line[0])
                worksheet.write(0,1,first_line[1])
                worksheet.write(0,2,first_line[2])
                worksheet.write(0,3,first_line[3])
                worksheet.write(0,4,first_line[4])
                worksheet.write(0,5,first_line[5])
                worksheet.write(0,6,first_line[6])
                worksheet.write(0,7,first_line[7])
                worksheet.write(0,8,first_line[8])
             else:
                if(short_url == 'aud' or short_url == 'inf'):  
                    int_counter = int(counter)
                    cut = (int_counter - 20 + 23 - 160)
                else:
                    int_counter = int(counter)
                    cut = (int_counter - 20 + 24 - 160)

                abs_cut = abs(cut)
                line[0] = line[0][0:abs_cut]
                line[4] = line[4].replace('aud', short_url)
                line[4] = sms_content.replace(lien, line[4]).replace(civilite, line[5]).replace(nom, line[0]).replace('\r\n','\n')
                #line[4] = convert_to_gsm(line[4])
               
                if count < 200001 :
                    worksheet.write(line_count, 0, line[3])
                    worksheet.write(line_count, 1, line[4])
        
             count = count + 1
             line_count = line_count +1
             count_str = str(count)
            print(count)
            workbook.close()

            full_path = os.path.abspath('final/'+name+'.xlsx')
            df = pd.read_excel(full_path)
            df[:-3].to_excel(full_path, index=False)  # overwrite the original file
            df[-3:].to_excel(full_path.replace('.xlsx', '_last3.xlsx'), index=False)  # create new file with last 3 lines

        return render_template("content_sms.html")
        
    else:
        flash('Allowed file types are only csv!')
        return redirect(request.url)
    

@app.route('/dedouble', methods=['GET','POST'])
def dedouble_file():
    if request.method == 'POST':
        render_template('content_dedouble.html')
    # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        column = request.form['column']
        keep = request.form['keep']
        keep_2 = ""
        if file.filename == '':
            flash('No file selected for uploading')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            file = open(app.config['UPLOAD_FOLDER'] + '/'+filename,"r")
            csv_reader_all = csv.reader(open(app.config['UPLOAD_FOLDER'] + '/'+filename, 'r', encoding='UTF-8'), delimiter=';')
            if keep == 'Oui':
                keep_2 == 'first'
            else:
                keep_2 == False
            df = csv_reader_all[(~csv_reader_all[column].duplicated(keep = keep_2 )) | csv_reader_all[column].isna()]
            df.to_csv("ready/"+file+"csv", index=None, header=True)

        return render_template("content_dedouble.html")
    
    else:
        flash('Allowed file types are only csv')
        return redirect(request.url)


if __name__ == "__main__":
    app.run(host = '0.0.0.0',port = 8080, debug = False)