from flask import Flask, flash, redirect, render_template, request, send_file
from txtFileCreator import *
from dataExtractor import *
import os, sys, io
from werkzeug.utils import secure_filename

DOWNLOAD_FOLDER = './output_files/'
ALLOWED_EXTENSIONS = {'xlsx'}

app = Flask(__name__)
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
app.secret_key = b'\xaf\xcd\x0f\xb2`\x16qz\xe8\xed\xea\xeb3b\xfa\xcd'

def allowed_file(filename):
	return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/", methods=['GET', 'POST'])
def home():
	if request.method == 'POST':
		if 'file' not in request.files:
			flash('No file part')
			return redirect(request.url)
		file = request.files['file']
		# if user does not select file, browser also
		# submit an empty part without filename
		if file.filename == '':
			flash('No selected file')
			return redirect(request.url)
		if file and allowed_file(file.filename):
			filename = secure_filename(file.filename)
			errorsExist, result = processInputFile(file)
			if errorsExist:
				#show the errors on the page (in result var) and allow user to reupload spreadsheet
				for error in result:
					flash(error)
				return redirect(request.url)
			else:
				#send output file to the page
				return send_file('upload_file.txt', as_attachment=True)
		flash('Incorrect file type. Only .xlsx files allowed. Please try again.')
		return redirect(request.url)
	else:
		return render_template("index.html")
		

def processInputFile(file):
	resultFile = io.open('upload_file.txt', 'w+')
	dataObj = DataPrepper(file)
	koteretData = dataObj.extractAndValidateKoteretData()
	tenuotData = dataObj.extractAndValidateTenuotData()
	errorList = dataObj.getErrorList()
	if len(errorList) > 0:  # if there are errors
		return True, errorList
	else:
		writerObj = TxtCreator(koteretData, tenuotData)
		resultFile.write(writerObj.ReshumatKoteret_Creator())
		resultFile.write(writerObj.ReshumatTenua_Creator())
		resultFile.write(writerObj.ReshumatTotal_Creator())
		resultFile.close()
		return False, resultFile

	
if __name__=="__main__":
	app.run()
