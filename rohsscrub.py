import os
import re
import sys
import string
import argparse  # TODO: Use this
import requests
import xlsxwriter
from datetime import datetime
from tabulate import tabulate

#tkinter handles the file dialogs for loading BOMs
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

# The API key that actually allows requests
API_KEY = ''

#Log File
log = []

#Lists for holding the relevant values
#Lifecycle Status
lcs = []
#RoHS Compliance
rohsc = []
#Compliance Documentation downloaded
cdo = []

#Creates prompt for user to select BOM file, returns path to file
def open_file():
	
	#Set the root window and hide it
	#(only show the file dialog)
	root = tk.Tk()
	root.withdraw()

	#Ask for a file until the user selects a proper format
	while True:
		#Create global variable for where the text file is
		global save_dir
		#Set the file path to a user-selected file
		file_path = filedialog.askopenfilename(title='Load BOM', filetypes=[('Parts List','*.txt'),('All Files','*.*')])
		#Checks for excel file
		if file_path.endswith('.txt'):
			save_dir = os.path.dirname(file_path)
			return parse_list(file_path)
			break
		#Exit if the user presses cancel or closes the window
		elif not file_path:
			break
		#Throw an error if the user selects a non-excel file and re-open the window
		else:
			messagebox.showinfo('Error','Please select a text file ending with .txt')
			continue
	return

#Parses the BOM file and returns a list of part number strings
def parse_list(loc):
		
	with open(loc, 'r') as f:
		#Strip all whitespace from line
		lines = (line.rstrip() for line in f)
		#Remove all blank lines
		p = list(line for line in lines if line)
	
	#Returns a list containing all the part numbers
	return p

#Formats the URL properly and makes the API request. Returns the JSON
def make_API_call(pn):
	url = 'http://octopart.com/api/v3/parts/match?'
	url += '&queries=[{"mpn":"' + pn + '"}]'
	url += '&include[]=datasheets'
	url += '&include[]=compliance_documents'
	url += '&include[]=specs'
	url += '&include[]=descriptions'
	url += '&pretty_print=true'
	url += '&apikey=' + API_KEY

	partjson = requests.get(url).json()

	return partjson

#Return the lifecycle status for the selected part
#	partjson is the JSON file for the part returned by the API function
#	sel is the user's part selection (automatic if only 1 item returned)
def lifecycle_status(partjson, sel):
	
	try:
		log.append('PN ' + partjson['results'][0]['items'][sel]['mpn'] + ' - Checking lifecycle status...')
		lifecycle_status = partjson['results'][0]['items'][sel]['specs']['lifecycle_status']['value'][0]
	except KeyError:
		log.append('PN ' + partjson['results'][0]['items'][sel]['mpn'] + ' - Error - No Lifecycle Information Given')
		return 'No lifecycle information given'
	
	log.append('PN ' + partjson['results'][0]['items'][sel]['mpn'] + ' - Lifecycle status found, written to results')
	#Check to see whether that file exists
	return lifecycle_status

#Return the RoHS compliance for the selected part and download the compliance statement (if available)
#	partjson is the JSON file for the part returned by the API function
#	sel is the user's part selection (automatic if only 1 item returned)
def rohs_compliance(partjson, sel):
	
	#Number of compliance documents for the item
	numco = len(partjson['results'][0]['items'][sel]['compliance_documents'])
	#Empty array for document types
	types = []

	#Loop through all compliance documents
	for i in range(numco):
		#Append the type of document to the 'types' array
		types.append(partjson['results'][0]['items'][sel]['compliance_documents'][i]['subtypes'][0])
	#If the array contains a RoHS statement, find its index and url
	#	Also append 'Yes' to the cdo array
	#print(types)
	if 'rohs_statement' in types:
		log.append('PN ' + partjson['results'][0]['items'][sel]['mpn'] + ' - Found RoHS documentation')
		rl = types.index('rohs_statement')
		url = partjson['results'][0]['items'][sel]['compliance_documents'][rl]['url']
		download_file(partjson['results'][0]['items'][sel]['mpn'], 'RoHS', url)
		cdo.append('Yes')
	#If the RoHS statement isn't found, append 'No'
	else:
		log.append('PN ' + partjson['results'][0]['items'][sel]['mpn'] + ' - No URL found for compliance document')
		cdo.append('NO VALID URL FOUND')

	#A list of all of the specs
	spt = []

	#Loop through all specs and append them to 'spt'
	for key in partjson['results'][0]['items'][sel]['specs']:
		spt.append(key)
	
	#Check to see if there is RoHS status information in the specs
	if 'rohs_status' in spt:
		log.append('PN ' + partjson['results'][0]['items'][sel]['mpn'] + ' - Found RoHS key, checking...')
		#If there is, check to see if a valid value is given
		if partjson['results'][0]['items'][sel]['specs']['rohs_status']['value']:
			log.append('PN ' + partjson['results'][0]['items'][sel]['mpn'] + ' - RoHS status found, written to results')
			compliance = partjson['results'][0]['items'][sel]['specs']['rohs_status']['value'][0]
	#Otherwise, output 'NOT FOUND'
	else:
		log.append('PN ' + partjson['results'][0]['items'][sel]['mpn'] + ' - No RoHS information found')
		compliance = 'NOT FOUND'

	#Return the compliance status
	return compliance

#Download a file given a partname, the type of document, and the URL where it's located
#	e.g., 'SN74S74N', 'RoHS', 'http://www.example.com/SN74S74N_RoHS.pdf'
def download_file(pn, typ, url):
	#e.g., 'C:/Parts/SN74S74N/SN74S74N_Datasheet.pdf'
	filename = save_dir + '/' + typ + '/' + pn.replace("/", "-") + '_' + typ + '.pdf'
	if not os.path.exists(os.path.dirname(filename)):
		os.makedirs(os.path.dirname(filename))
	response = requests.get(url, stream=True)
	with open(filename, 'wb') as out_file:
		for chunk in response.iter_content(chunk_size=1024):
			out_file.write(chunk)
	log.append('PN ' + pn + ' - RoHS information downloaded')
	print(typ + ' document written to ' + filename + '\n')
	return

#Prompt the user for the proper part when multiple parts are returned from the API call
#	Returns an array index
def part_select(partjson):

	#Make a list of the parts
	pts = []

	print('Multiple part matches, please select correct part:\n')
	#Output the results of the search in an easy-to-read format
	for i in range(numi):
		#List index, starting at 1
		ind = str(i+1)
		#Find manufacturer for each item
		mfg = partjson['results'][0]['items'][i]['manufacturer']['name']
		#List the part number for each item (NOTE: Should be consistent across all items, but it can't hurt)
		mpn = partjson['results'][0]['items'][i]['mpn']
		#Find the description
		des = find_description(partjson, i)
		pts.append([ind, mfg, mpn, des])
	print(tabulate(pts, headers=['','Manufacturer','Part Number','Description'])+'\n')
	
	while True:
		cp = input('Enter number of correct part (or 0 if none are correct): ')
		print('\n')
		try: 
			#Convert the input to an integer and subtract 1 to return index
			a = int(cp)-1
			return a
		except(TypeError, ValueError, IndexError):
			print('Invalid selection, please enter number between 1 and ' + str(len(pts)) + ' or 0 if none are correct\n')
			continue

#Find the description for each part, called by part_select
def find_description(partjson, i):

	#Find the number of distributors for an item
	numdist = len(partjson['results'][0]['items'][i]['descriptions'])
	
	#Create an empty array for the descriptions
	descs = []
	
	#Loop through all descriptions
	for j in range(numdist):
		#If the distributor is Digi-Key, directly return the description
		if partjson['results'][0]['items'][i]['descriptions'][j]['attribution']['sources'][0]['name'] == 'Digi-Key':
			return partjson['results'][0]['items'][i]['descriptions'][j]['value']
		#For any other distributor, append the description to the descriptions array
		else:
			descs.append(partjson['results'][0]['items'][i]['descriptions'][j]['value'])
	if len(descs) == 0:
		return '***NO DESCRIPTION***'
	else:
		return min(descs, key=len)

#Write the data out to a spreadsheet
def write_spreadsheet(dlist):
	#Sets the output file location to the input directory
	#***WARNING*** It doesn't check if something's already there, and WILL overwrite it without prompting***
	wb_fn = save_dir + '/RoHS_Scrub_Results.xlsx'
	
	workbook = xlsxwriter.Workbook(wb_fn)
	worksheet = workbook.add_worksheet()
		
	#First write the headers
	headers = ['Item', 'Part Number', 'Lifecycle Status', 'RoHS Compliance', 'Certificate of Compliance Downloaded?']
	for i in range(len(headers)):
		worksheet.write(0, i, headers[i])

	for y in range(1, len(dlist)+1):
		for x in range(5):
			worksheet.write(y, x, dlist[y-1][x])	
	workbook.close()
	log.append('PN ' + pjs['results'][0]['items'][sel]['mpn'] + ' - Results spreadsheet saved.')
	return wb_fn

#Prompt the user for an API key if one wasn't provided at the command line
def get_api_key():
	log.append('No API key given, prompting user')
	return input('Please enter API key: ')

#Output the number of Bom items that were found
def coverage(dlist):
	covered = 0
	for i in range(len(dlist)):
		if dlist[i][3] != 'NOT FOUND' or 'CHECK P/N': covered += 1
		else: continue
	
	return covered

def write_log():
	today = str(datetime.now().replace(second=0, microsecond=0))
	filename = save_dir + '/RoHS_Scrub_LOGFILE.txt'
	log.append('Saving log...')
	with open(filename, 'w') as f:
		f.write(today + '\n')
		f.write('---START---\n')
		for line in log:
			f.write(line)
			f.write('\n')
		f.write('---END---')
		f.close()

if(len(sys.argv)) < 2:
	API_KEY = get_api_key()
else:
	API_KEY = sys.argv[1]

#Prompt the user for the text file of part numbers
parts = open_file()
#Make a list of integers from 1 to the number of parts found in the text file
indices = [i+1 for i in range(len(parts))]

#Get all data for each part
for item in range(len(parts)):
	#Fill a JSON with information from the API call
	pjs = make_API_call(parts[item])
	#Figure out how many items match the part number
	numi = len(pjs['results'][0]['items'])
	if numi == 0:
		lcs.append('NOT FOUND')
		rohsc.append('NOT FOUND')
		cdo.append('No')
		log.append('PN: ' + parts[item] + ' - No Results. Check part number')
	elif numi == 1:
		sel = 0
		#Find lifecycle status, append it to list
		lcs.append(lifecycle_status(pjs, sel))
		#Find RoHS status, append it to list
		rohsc.append(rohs_compliance(pjs, sel))
		log.append('PN: ' + parts[item] + ' - One match.')
	else:
		log.append('PN: ' + parts[item] + ' - Multiple part matches.')
		#Prompt the user for the proper part
		sel = part_select(pjs)
		if sel != -1:
			log.append('PN: ' + parts[item] + ' - User selected ' +  pjs['results'][0]['items'][sel]['manufacturer']['name'] + ' ' +  pjs['results'][0]['items'][sel]['mpn'] )
			#Find lifecycle status, append it to list
			lcs.append(lifecycle_status(pjs, sel))
			#Find RoHS status, append it to list
			rohsc.append(rohs_compliance(pjs, sel))
		else:
			log.append('PN: ' + parts[item] + ' - User selected no part match. Skipping.')
			lcs.append('CHECK P/N')
			rohsc.append('CHECK P/N')

#Output excel file of results
data_table = [list(x) for x in zip(indices, parts, lcs, rohsc, cdo)]
#Where the data spreadsheet was saved
data_loc = write_spreadsheet(data_table)

#Display path where files were saved
print('Results saved to ' + data_loc + '\n')

#Display BOM Coverage %
print('Total BOM Coverage:' + str(coverage(data_table)) + '/' + str(len(data_table)) + " items found\n")

write_log()
