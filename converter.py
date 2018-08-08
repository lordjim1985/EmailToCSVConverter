# -*- coding: utf-8 -*-
import argparse
import os
import re
import time
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
	row_content_text = ""
	work_mode = "dirs"
	source_dir = "maildir"
	debug = True
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
		
		args = parser.parse_args()

		if (self.debug == True) :
			self.time_start = self.microtime(True)

		self.source_filename = args.input_filename[0]
		self.output_filename = args.output_filename[0]

		if (os.path.isfile(args.input_filename[0])) :
			self.work_mode = "file"
			self.input_file_handle = open(args.input_filename[0])

		if (os.path.isdir(args.input_filename[0])) :
			self.work_mode = "dirs"

		self.output_file_handle = open(args.output_filename[0], "w")
		self.processDescriptionHeaders()

		if (self.work_mode == "file") :
			email = self.input_file_handle.read()
			self.processEmail(email)
			self.input_file_handle.close()
		if (self.work_mode == "dirs") :
			self.readDirectory(args.input_filename[0])

		if (self.debug == True) :
			self.time_end = self.microtime(True);
			self.execution_time = (self.time_end - self.time_start) / 60;

		print "File " + self.output_filename + " created."

		if (self.debug == True) :
			# print 'Max memory usage: ' + str(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss) + 'MB';
			print "Script Execution time: " + '%.5f' % self.execution_time + " seconds, " + '%.5f' % (self.execution_time/60) + " minutes";

		self.output_file_handle.close()

	def processDescriptionHeaders(self):
		for key,val in enumerate(self.desc_headers) :
			if (key != len(self.desc_headers)-1) :
				self.row_content += self.string_delimeter + self.desc_headers[key] + self.string_delimeter + self.field_delimeter
			else :
				self.row_content += self.string_delimeter + self.desc_headers[key] + self.string_delimeter + self.line_delimeter
		self.writeToFile()

	def readDirectory(self, dirname):
		for dirpath, dirs, files in os.walk(dirname):

			for file in files:
				email_handle = open(dirpath + "\\" + file, "r")
				email = email_handle.read()
				self.processEmail(email)
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
		processed_headers = "\n".join(str(process_headers_value) for process_headers_value in processed_headers)

		self.messageid_found = self.searchForHeader('Message-ID', processed_headers)
		self.date_found = self.searchForHeader('Date', processed_headers)
		self.from_found = self.searchForHeader('From', processed_headers)
		self.to_found = self.searchForHeader('To', processed_headers)
		self.subject_found = self.searchForHeader('Subject', processed_headers)
		self.cc_found = self.searchForHeader('Cc', processed_headers)
		self.mimeversion_found = self.searchForHeader('Mime-Version', processed_headers)
		self.contenttype_found = self.searchForHeader('Content-Type', processed_headers)
		self.contenttransferencoding_found = self.searchForHeader('Content-Transfer-Encoding', processed_headers)
		self.bcc_found = self.searchForHeader('Bcc', processed_headers)
		self.xfrom_found = self.searchForHeader('X-from', processed_headers)
		self.xto_found = self.searchForHeader('X-to', processed_headers)
		self.xcc_found = self.searchForHeader('X-cc', processed_headers)
		self.xbcc_found = self.searchForHeader('X-bcc', processed_headers)
		self.xfolder_found = self.searchForHeader('X-folder', processed_headers)
		self.xorigin_found = self.searchForHeader('X-origin', processed_headers)
		self.xfilename_found = self.searchForHeader('X-fileName', processed_headers)

	def searchForHeader(self, header, headers):
		return headers.find(header + ':')

	def processHeaders(self, email_headers):
		cleaned_headers = self.replaceTabs(email_headers)
		cleaned_headers = self.replaceDoubleNewLines(cleaned_headers)

		cleaned_headers = self.replaceSemicolons(cleaned_headers)
		cleaned_headers = self.replaceQuotes(cleaned_headers)
		cleaned_headers = self.replaceApos(cleaned_headers)
		cleaned_headers = self.replaceBackslashes(cleaned_headers)

		normalized_headers = self.normalizeHeaders(cleaned_headers)
		
		separated_headers = self.separateHeaders(normalized_headers)
		
		return separated_headers

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

			if processed_headers_value.find('Message-ID:') != -1 :
				cleaned_header_value = self.removeHeader('Message-ID', processed_headers_value)
				if cleaned_header_value == "" or cleaned_header_value == " " :
					cleaned_header_value = 0

				self.row_content += self.string_delimeter + str(cleaned_header_value).strip()

				matches += 1

				if self.date_found == -1 :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + str(0)

			if processed_headers_value.find('Date:') != -1 :
				cleaned_header_value = self.removeHeader('Date', processed_headers_value);
				if cleaned_header_value == "" or cleaned_header_value == " " :
					cleaned_header_value = 0

				self.row_content += self.field_delimeter + self.string_delimeter + str(cleaned_header_value).strip()

				matches += 1

				if self.from_found == -1 :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + str(0)

			if processed_headers_value.find('From:') != -1 :
				cleaned_header_value = self.removeHeader('From', processed_headers_value);
				if cleaned_header_value == "" or cleaned_header_value == " " :
					cleaned_header_value = 0

				self.row_content += self.field_delimeter + self.string_delimeter + str(cleaned_header_value).strip()

				matches += 1

				if self.to_found == -1 :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + str(0)

			if processed_headers_value.find('To:') != -1 :
				cleaned_header_value = self.removeHeader('To', processed_headers_value);
				if cleaned_header_value == "" or cleaned_header_value == " " :
					cleaned_header_value = 0

				self.row_content += self.field_delimeter + self.string_delimeter + str(cleaned_header_value).strip()

				matches += 1

				if self.subject_found == -1 :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + str(0)

			if processed_headers_value.find('Subject:') != -1 :
				cleaned_header_value = self.removeHeader('Subject', processed_headers_value);
				if cleaned_header_value == "" or cleaned_header_value == " " :
					cleaned_header_value = 0

				self.row_content += self.field_delimeter + self.string_delimeter + str(cleaned_header_value).strip()

				matches += 1

				if self.cc_found == -1 :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + str(0)

			if processed_headers_value.find('Cc:') != -1 :
				cleaned_header_value = self.removeHeader('Cc', processed_headers_value);
				if cleaned_header_value == "" or cleaned_header_value == " " :
					cleaned_header_value = 0

				self.row_content += self.field_delimeter + self.string_delimeter + str(cleaned_header_value).strip()

				matches += 1

				if self.mimeversion_found == -1 :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + str(0)

			if processed_headers_value.find('Mime-Version:') != -1 :
				cleaned_header_value = self.removeHeader('Mime-Version', processed_headers_value);
				if cleaned_header_value == "" or cleaned_header_value == " " :
					cleaned_header_value = 0

				self.row_content += self.field_delimeter + self.string_delimeter + str(cleaned_header_value).strip()

				matches += 1

				if self.contenttype_found == -1 :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + str(0)

			if processed_headers_value.find('Content-Type:') != -1 :
				cleaned_header_value = self.removeHeader('Content-Type', processed_headers_value);
				if cleaned_header_value == "" or cleaned_header_value == " " :
					cleaned_header_value = 0

				self.row_content += self.field_delimeter + self.string_delimeter + str(cleaned_header_value).strip()

				matches += 1

				if self.contenttransferencoding_found == -1 :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + str(0)

			if processed_headers_value.find('Content-Transfer-Encoding:') != -1 :
				cleaned_header_value = self.removeHeader('Content-Transfer-Encoding', processed_headers_value);
				if cleaned_header_value == "" or cleaned_header_value == " " :
					cleaned_header_value = 0

				self.row_content += self.field_delimeter + self.string_delimeter + str(cleaned_header_value).strip()

				matches += 1

				if self.bcc_found == -1 :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + str(0)

			if processed_headers_value.find('Bcc:') != -1 :
				cleaned_header_value = self.removeHeader('Bcc', processed_headers_value);
				if cleaned_header_value == "" or cleaned_header_value == " " :
					cleaned_header_value = 0

				self.row_content += self.field_delimeter + self.string_delimeter + str(cleaned_header_value).strip()

				matches += 1

				if self.xfrom_found == -1 :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + str(0)

			if processed_headers_value.find('X-from:') != -1 :
				cleaned_header_value = self.removeHeader('X-from', processed_headers_value);
				if cleaned_header_value == "" or cleaned_header_value == " " :
					cleaned_header_value = 0

				self.row_content += self.field_delimeter + self.string_delimeter + str(cleaned_header_value).strip()

				matches += 1

				if self.xto_found == -1 :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + str(0)

			if processed_headers_value.find('X-to:') != -1 :
				cleaned_header_value = self.removeHeader('X-to', processed_headers_value);
				if cleaned_header_value == "" or cleaned_header_value == " " :
					cleaned_header_value = 0

				self.row_content += self.field_delimeter + self.string_delimeter + str(cleaned_header_value).strip()

				matches += 1

				if self.xcc_found == -1 :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + str(0)

			if processed_headers_value.find('X-cc:') != -1 :
				cleaned_header_value = self.removeHeader('X-cc', processed_headers_value);
				if cleaned_header_value == "" or cleaned_header_value == " " :
					cleaned_header_value = 0

				self.row_content += self.field_delimeter + self.string_delimeter + str(cleaned_header_value).strip()

				matches += 1

				if self.xbcc_found == -1 :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + str(0)

			if processed_headers_value.find('X-bcc:') != -1 :
				cleaned_header_value = self.removeHeader('X-bcc', processed_headers_value);
				if cleaned_header_value == "" or cleaned_header_value == " " :
					cleaned_header_value = 0

				self.row_content += self.field_delimeter + self.string_delimeter + str(cleaned_header_value).strip()

				matches += 1

				if self.xfolder_found == -1 :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + str(0)

			if processed_headers_value.find('X-folder:') != -1 :
				cleaned_header_value = self.removeHeader('X-folder', processed_headers_value);
				if cleaned_header_value == "" or cleaned_header_value == " " :
					cleaned_header_value = 0

				self.row_content += self.field_delimeter + self.string_delimeter + str(cleaned_header_value).strip()

				matches += 1

				if self.xorigin_found == -1 :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + str(0)

			if processed_headers_value.find('X-origin:') != -1 :
				cleaned_header_value = self.removeHeader('X-origin', processed_headers_value);
				if cleaned_header_value == "" or cleaned_header_value == " " :
					cleaned_header_value = 0

				self.row_content += self.field_delimeter + self.string_delimeter + str(cleaned_header_value).strip()

				matches += 1

				if self.xfilename_found == -1 :
					self.row_content += self.string_delimeter + self.field_delimeter + self.string_delimeter + str(0)

			if processed_headers_value.find('X-fileName:') != -1 :
				cleaned_header_value = self.removeHeader('X-fileName', processed_headers_value);
				if cleaned_header_value == "" or cleaned_header_value == " " :
					cleaned_header_value = 0

				self.row_content += self.field_delimeter + self.string_delimeter + str(cleaned_header_value).strip()
				
				matches += 1;

			if matches == 0 :
				self.row_content = self.row_content[:-2];
				self.row_content += self.removeNewLines(processed_headers_value);
			else :
				self.row_content += self.string_delimeter;

	def replaceSemicolons(self, unified_headers):
		return unified_headers.replace(";", "&#59;")

	def replaceQuotes(self, unified_headers):
		return unified_headers.replace('"', "&quot;")

	def replaceApos(self, unified_headers):
		return unified_headers.replace("'", "&#39;")

	def replaceBackslashes(self, unified_headers):
		return unified_headers.replace("\\", "&bsol;")

	def replaceTabs(self, separated_headers):
		return separated_headers.replace("\t", " ")

	def replaceDoubleNewLines(self, email_headers):
		#return re.sub("/\r\n\s+/m", " ", email_headers)
		email_headers_cleaned = re.sub(r"\n\s+", " ", email_headers, flags=re.M)
		return email_headers_cleaned.replace("  ", " ")

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
		return normalized_header_value.replace(header_name + ": ", "")

	def trimHeader(self, header_value) :
		header_value.strip()

	def removeNewLines(self, value):
		return value.replace("\n", " ")

	def writeToFile(self):
		self.output_file_handle.write(self.row_content)
		self.row_content = ""

	def getScriptPath(self):
		print "test"


	def microtime(self, get_as_float = False) :
		d = datetime.now()
		t = time.mktime(d.timetuple())
		if get_as_float:
			return t
		else:
			ms = d.microsecond / 1000000.
		return '%f %d' % (ms, t)

EmailToCSVConverter()