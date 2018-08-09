# -*- coding: utf-8 -*-
import argparse
import os
import re
import time
import sys
from datetime import datetime
from pprint import pprint

class EmailToCSVConverter: 
	field_delimeter = ';'
	string_delimeter = '"'
	output_filename = 'result.csv'
	input_filename = 'email.txt'
	line_delimeter = "\n"
	input_file_handle = ""
	output_file_handle = ""
	file_headers = ""
	file_contents = ""
	source_file_handle = ""
	row_content = ""
	empty_value = str(0)
	work_mode = False
	source_dir = "maildir"
	debug = False
	progress = False
	source_headers = ['Message-ID: ', 'Date: ', 'From: ', 'To: ', 'Subject: ', 'Cc: ', 'Mime-Version: ', 'Content-Type: ', 'Content-Transfer-Encoding: ', 'Bcc: ', 'X-From: ', 'X-To: ', 'X-cc: ', 'X-bcc: ', 'X-Folder: ', 'X-Origin: ', 'X-FileName: ']
	target_headers = ['Message-ID: ', 'Date: ', 'From: ', 'To: ', 'Subject: ', 'Cc: ', 'Mime-Version: ', 'Content-Type: ', 'Content-Transfer-Encoding: ', 'Bcc: ', 'X-from: ', 'X-to: ', 'X-cc: ', 'X-bcc: ', 'X-folder: ', 'X-origin: ', 'X-fileName: ']
	desc_headers = ['Message-ID', 'Date', 'From', 'To', 'Subject', 'Cc', 'Mime-Version', 'Content-Type', 'Content-Transfer-Encoding', 'Bcc', 'X-from', 'X-to', 'X-cc', 'X-bcc', 'X-folder', 'X-origin', 'X-fileName', 'Message']
	time_start = 0
	time_end = 0
	messageid_found = False
	date_found = False
	from_found = False
	to_found = False
	subject_found = False
	cc_found = False
	mimeversion_found = False
	contenttype_found = False
	contenttransferencoding_found = False
	bcc_found = False
	xfrom_found = False
	xto_found = False
	xcc_found = False
	xbcc_found = False
	xfolder_found = False
	xorigin_found = False
	xfilename_found = False

	def __init__(self):
		parser = argparse.ArgumentParser(description='Convert en e-mail or an e-mail directory to an CSV formatted output.')
		
		parser.add_argument('--input_filename', metavar='I', type=str, nargs='+', required=True, help='name of the input_filename / directory to process. Example: email.txt or maildir')
		parser.add_argument('--output_filename', metavar='O', type=str, nargs='+', required=True, help='name of the output filename. Example: result.csv')
		parser.add_argument('--progress', metavar='P', type=int, nargs='+', required=False, help='should progress bar be enabled ')
		
		args = parser.parse_args()

		# if (self.debug == True) :
		self.time_start = self.microtime(False)

		self.source_filename = args.input_filename[0]
		self.output_filename = args.output_filename[0]

		try:		
			if args.progress[0] == 1 :
				self.progress = True
			else :
				self.progress = False
		except: 
			self.progress = False

		if os.path.exists(self.source_filename) :

			if os.path.isfile(args.input_filename[0]) :
				self.work_mode = "file"
				self.input_file_handle = open(args.input_filename[0])

			if os.path.isdir(args.input_filename[0]) :
				self.work_mode = "dirs"

			self.output_file_handle = open(args.output_filename[0], "w")
			
			if (self.debug == True) :
				print "File " + self.output_filename + " created."

			self.processDescriptionHeaders()

			if (self.work_mode == "file") :
				email = self.input_file_handle.read()
				
				if (self.debug == True) :
					print "Processing file: " + self.input_filename
				
				self.processEmail(email)
				self.input_file_handle.close()
			if (self.work_mode == "dirs") :
				self.processDirectory(args.input_filename[0])

			self.output_file_handle.close()
			print "Output file: " + self.output_filename + " written"

		else :
			print 'Requested directory or file does not exist'

		# if (self.debug == True) :
		self.time_end = self.microtime(False);
		self.execution_time = self.time_end - self.time_start;

		# if (self.debug == True) :
		# print 'Max memory usage: ' + str(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss) + 'MB';
		print "Script Execution time: " + '%.5f' % self.execution_time + " seconds, " + '%.5f' % (self.execution_time/60) + " minutes";

	def update_progress(self, position, total):
		barLength = 50 # Modify this to change the length of the progress bar
		status = ""
		
		progress = position/total

		if isinstance(progress, int):
			progress = float(progress)
		if not isinstance(progress, float):
			progress = 0
			status = "error: progress var must be float\r\n"
		if progress < 0:
			progress = 0
			status = "Halt...\r\n"
		if progress >= 1:
			progress = 1
			status = "Done...\r\n"
		block = int(round(barLength*progress))
		text = "\rPercent: [{0}] {1}% {2}".format( "#"*block + "-"*(barLength-block), progress*100, status)
		
		sys.stdout.write(text)
		sys.stdout.flush()

	def processDescriptionHeaders(self):
		for key,val in enumerate(self.desc_headers) :
			if (key != len(self.desc_headers)-1) :
				self.row_content += self.string_delimeter + self.desc_headers[key] + self.string_delimeter + self.field_delimeter
			else :
				self.row_content += self.string_delimeter + self.desc_headers[key] + self.string_delimeter + self.line_delimeter
		self.writeToFile()

	def processDirectory(self, dirname):
		file_sum = 0
		file_counter = 0

		if self.progress == True :
			print "Calculating amount of files"
			for dirpath, dirs, files in os.walk(dirname):
				for file in files:
					if os.path.isfile(dirpath + "\\" + file) :
						file_sum += 1

		print "Converting..."
		
		if self.progress == True :
			self.update_progress(file_counter, file_sum)

		for dirpath, dirs, files in os.walk(dirname):
			for file in files:
				email_handle = open(dirpath + "\\" + file, "r")
				if (self.debug == True) :
					print "Processing file: " + dirpath + "\\" + file
				email = email_handle.read()
				self.processEmail(email)
				file_counter += 1
				if self.progress == True :
					self.update_progress(float(file_counter), file_sum)
				email_handle.close()

	def processEmail(self, email):
		email_contents = self.separateContent(email)
		email_headers = email_contents[0]
		processed_headers = self.processHeaders(email_headers)

		self.searchForHeaders(processed_headers)
		self.createCSV(processed_headers)

		email_body = ""
		

		if len(email_contents) > 1 :
			if len(email_contents) > 2 :
				for k,v in enumerate(email_contents):
					if (k != 0) :
						email_body += email_contents[k]
				self.processBody(email_body)
			else:
				email_body = email_contents[1]
				self.processBody(email_body)
		else :
			self.row_content += self.field_delimeter + self.string_delimeter + self.line_delimeter + self.string_delimeter;

		self.writeToFile()

	def searchForHeaders(self, processed_headers):
		self.resetFoundValues()
		# processed_headers_value = "\n".join(str(process_headers_value) for process_headers_value in processed_headers)
		for processed_headers_key, processed_headers_value in enumerate(processed_headers) :
			if self.searchForHeader('Message-ID', processed_headers_value):
				self.messageid_found = self.searchForHeader('Message-ID', processed_headers_value)
			if self.searchForHeader('Date', processed_headers_value):
				self.date_found = self.searchForHeader('Date', processed_headers_value)
			if self.searchForHeader('From', processed_headers_value):
				self.from_found = self.searchForHeader('From', processed_headers_value)
			if self.searchForHeader('To', processed_headers_value):
				self.to_found = self.searchForHeader('To', processed_headers_value)
			if self.searchForHeader('Subject', processed_headers_value):
				self.subject_found = self.searchForHeader('Subject', processed_headers_value)
			if self.searchForHeader('Cc', processed_headers_value):
				self.cc_found = self.searchForHeader('Cc', processed_headers_value)
			if self.searchForHeader('Mime-Version', processed_headers_value):
				self.mimeversion_found = self.searchForHeader('Mime-Version', processed_headers_value)
			if self.searchForHeader('Content-Type', processed_headers_value):
				self.contenttype_found = self.searchForHeader('Content-Type', processed_headers_value)
			if self.searchForHeader('Content-Transfer-Encoding', processed_headers_value):
				self.contenttransferencoding_found = self.searchForHeader('Content-Transfer-Encoding', processed_headers_value)
			if self.searchForHeader('Bcc', processed_headers_value):
				self.bcc_found = self.searchForHeader('Bcc', processed_headers_value)
			if self.searchForHeader('X-from', processed_headers_value):
				self.xfrom_found = self.searchForHeader('X-from', processed_headers_value)
			if self.searchForHeader('X-to', processed_headers_value):	
				self.xto_found = self.searchForHeader('X-to', processed_headers_value)
			if self.searchForHeader('X-cc', processed_headers_value):
				self.xcc_found = self.searchForHeader('X-cc', processed_headers_value)
			if self.searchForHeader('X-bcc', processed_headers_value):
				self.xbcc_found = self.searchForHeader('X-bcc', processed_headers_value)
			if self.searchForHeader('X-folder', processed_headers_value):
				self.xfolder_found = self.searchForHeader('X-folder', processed_headers_value)
			if self.searchForHeader('X-origin', processed_headers_value):
				self.xorigin_found = self.searchForHeader('X-origin', processed_headers_value)
			if self.searchForHeader('X-fileName', processed_headers_value):
				self.xfilename_found = self.searchForHeader('X-fileName', processed_headers_value)

	def searchForHeader(self, header, headers):
		# print re.search(header + ': ' , headers)
		# return re.search(header + ': ' , headers)
		# return headers.find(header + ': ')
		return headers.startswith(header + ":")

	def processHeaders(self, email_headers):
		cleaned_headers = self.replaceTabs(email_headers)
		cleaned_headers = self.replaceExtraNewLines(cleaned_headers)

		cleaned_headers = self.replaceSemicolons(cleaned_headers)
		cleaned_headers = self.replaceQuotes(cleaned_headers)
		cleaned_headers = self.replaceApos(cleaned_headers)
		cleaned_headers = self.replaceBackslashes(cleaned_headers)

		normalized_headers = self.normalizeHeaders(cleaned_headers)
		
		separated_headers = self.separateHeaders(normalized_headers)

		return separated_headers

	def processHeaderValue(self, header_name, header_value, field_delimeter=True):
		cleaned_header_value = self.removeHeader(header_name, header_value)
		
		if cleaned_header_value == "" or cleaned_header_value == " " :
			cleaned_header_value = self.empty_value
		
		cleaned_header_value = str(cleaned_header_value).strip()
		
		if field_delimeter:
			self.row_content += self.field_delimeter

		self.row_content += self.string_delimeter + cleaned_header_value

	def processBody(self, email_body):
		email_body = self.replaceSemicolons(email_body)
		email_body = self.replaceQuotes(email_body)
		email_body = self.replaceApos(email_body)
		email_body = self.replaceBackslashes(email_body)
		email_body = self.replaceExtraNewLines(email_body)
		# email_body = " ";
		self.row_content += self.field_delimeter + self.string_delimeter + self.removeNewLines(email_body) + self.string_delimeter + self.line_delimeter;

	def createCSV(self, processed_headers):
		for process_headers_key, processed_headers_value in enumerate(processed_headers) :
			matches = 0;

			if processed_headers_value.startswith('Message-ID:') != False :
				self.processHeaderValue('Message-ID', processed_headers_value, False)

				matches += 1

				if self.date_found == False :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + self.empty_value

			if processed_headers_value.startswith('Date:') != False :
				self.processHeaderValue('Date', processed_headers_value)

				matches += 1

				if self.from_found == False :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + self.empty_value

			if processed_headers_value.startswith('From:') != False :
				self.processHeaderValue('From', processed_headers_value)

				matches += 1

				if self.to_found == False :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + self.empty_value

			if processed_headers_value.startswith('To:') != False :
				self.processHeaderValue('To', processed_headers_value)

				matches += 1

				if self.subject_found == False :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + self.empty_value

			if processed_headers_value.startswith('Subject:') != False :
				self.processHeaderValue('Subject', processed_headers_value)

				matches += 1

				if self.cc_found == False :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + self.empty_value

			if processed_headers_value.startswith('Cc:') != False :
				self.processHeaderValue('Cc', processed_headers_value)

				matches += 1

				if self.mimeversion_found == False :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + self.empty_value

			if processed_headers_value.startswith('Mime-Version:') != False :
				self.processHeaderValue('Mime-Version', processed_headers_value)

				matches += 1

				if self.contenttype_found == False :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + self.empty_value

			if processed_headers_value.startswith('Content-Type:') != False :
				self.processHeaderValue('Content-Type', processed_headers_value)

				matches += 1

				if self.contenttransferencoding_found == -1 :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + self.empty_value

			if processed_headers_value.startswith('Content-Transfer-Encoding:') != False :
				self.processHeaderValue('Content-Transfer-Encoding', processed_headers_value)

				matches += 1

				if self.bcc_found == False :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + self.empty_value

			if processed_headers_value.startswith('Bcc:') != False :
				self.processHeaderValue('Bcc', processed_headers_value)

				matches += 1

				if self.xfrom_found == False :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + self.empty_value

			if processed_headers_value.startswith('X-from:') != False :
				self.processHeaderValue('X-from', processed_headers_value)

				matches += 1

				if self.xto_found == False :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + self.empty_value

			if processed_headers_value.startswith('X-to:') != False :
				self.processHeaderValue('X-to', processed_headers_value)

				matches += 1

				if self.xcc_found == False :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + self.empty_value

			if processed_headers_value.startswith('X-cc:') != False :
				self.processHeaderValue('X-cc', processed_headers_value)

				matches += 1

				if self.xbcc_found == False :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + self.empty_value

			if processed_headers_value.startswith('X-bcc:') != False :
				self.processHeaderValue('X-bcc', processed_headers_value)

				matches += 1

				if self.xfolder_found == False :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + self.empty_value

			if processed_headers_value.startswith('X-folder:') != False :
				self.processHeaderValue('X-folder', processed_headers_value)

				matches += 1

				if self.xorigin_found == False :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + self.empty_value

			if processed_headers_value.startswith('X-origin:') != False :
				self.processHeaderValue('X-origin', processed_headers_value)

				matches += 1

				if self.xfilename_found == False :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + self.empty_value

			if processed_headers_value.startswith('X-fileName:') != False :
				self.processHeaderValue('X-fileName', processed_headers_value)
				
				matches += 1;

			if matches == 0 :
				if self.row_content[-5:] == '";"' + self.empty_value + '"':
					self.row_content = self.row_content[:-5]
					self.row_content += " " + processed_headers_value + ' ";"' + self.empty_value + '"'
				elif self.row_content[-1:] == '"':
					self.row_content = self.row_content[:-1]
					self.row_content += " " + processed_headers_value + ' "'
			else :
				self.row_content += self.string_delimeter

		self.row_content = self.replaceDoubleSpaces(self.row_content)

	def replaceSemicolons(self, content):
		return content.replace(";", "&#59;")

	def replaceQuotes(self, content):
		return content.replace('"', "&quot;")

	def replaceApos(self, content):
		return content.replace("'", "&#39;")

	def replaceBackslashes(self, content):
		return content.replace("\\", "&bsol;")

	def replaceTabs(self, content):
		return content.replace("\t", " ")

	def replaceExtraNewLines(self, content):
		#return re.sub("/\r\n\s+/m", " ", email_headers)
		content_cleaned = re.sub(r"\n\s+", " ", content, flags=re.M)
		return content_cleaned.replace("  ", " ")

	def replaceDoubleSpaces(self, content):
		return content.replace("  ", " ")

	def replaceExtraNewLines(self, email_body):
		return email_body.replace("\n\n", " ")

	def separateHeaders(self, normalized_headers):
		return normalized_headers.split("\n")

	def separateContent(self, email):
		return email.split("\n\n")

	def normalizeHeaders(self, separated_headers):
		normalized_headers = separated_headers
		for k,v in enumerate(self.source_headers):
			normalized_headers = normalized_headers.replace(v, self.target_headers[k])

		return normalized_headers

	def removeHeader(self, header_name, normalized_header_value):
		return normalized_header_value.replace(header_name + ":", "")

	def trimHeader(self, header_value) :
		header_value.strip()

	def removeNewLines(self, value):
		return value.replace("\n", " ")

	def writeToFile(self):
		self.output_file_handle.write(self.row_content)
		self.row_content = ""

	def resetFoundValues(self):
		self.messageid_found = False
		self.date_found = False
		self.from_found = False
		self.to_found = False
		self.subject_found = False
		self.cc_found = False
		self.mimeversion_found = False
		self.contenttype_found = False
		self.contenttransferencoding_found = False
		self.bcc_found = False
		self.xfrom_found = False
		self.xto_found = False
		self.xcc_found = False
		self.xbcc_found = False
		self.xfolder_found = False
		self.xorigin_found = False
		self.xfilename_found = False

	def microtime(self, get_as_float = False) :
		return time.time()

EmailToCSVConverter()