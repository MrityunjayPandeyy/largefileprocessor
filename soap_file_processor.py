import os
import re
import xlsxwriter
import xlwt
import sys


flow_id = set([])
data_list = set([])
workbook=''

req_list = []
resp_list = []
src_file=""
dest_file = ""


if len(sys.argv) < 2:
    print "Enter a valid command. Either Source file or destination file is not found"
else:
	try:
		src_file = sys.argv[1]
		dest_file= sys.argv[2]	
		
		print "Source file is "+src_file
		print "Destination file is "+dest_file	
		
		if  dest_file == "":
			dest_file="output.xlsx"

	except Exception: 
		if  dest_file == "":
			dest_file="output.xlsx"		

# Create an new Excel file and add a worksheet.	
workbook = xlsxwriter.Workbook(dest_file)
worksheet = workbook.add_worksheet("BaseProcessedMessages")
bold = workbook.add_format({'bold': True, 'bg_color': 'green','border' : True})
wrap = workbook.add_format({'text_wrap': True}) 
worksheet.set_column('C:C', 80)
worksheet.set_column('D:D', 80)
worksheet.write('A1', 'Sr.No.', bold)
worksheet.write('B1', 'Line No.', bold)
worksheet.write('C1', 'Soap Request', bold)
worksheet.write('D1', 'Soap Response', bold)


def find_between( s, first, last ):
    try:
        start = s.index( first ) + len( first )
        end = s.index( last, start )
        return s[start:end]
    except ValueError:
        return ""
		
def seprate_data_sheet(sheet):
    print sheet
    if sheet:
        worksheet1 = workbook.add_worksheet(sheet)
        worksheet1.set_column('C:C', 80)
        worksheet1.set_column('D:D', 80)
        worksheet1.write('A1', 'Sr.No.', bold)
        worksheet1.write('B1', 'Line No.', bold)
        worksheet1.write('C1', 'Soap Request', bold)
        worksheet1.write('D1', 'Soap Response', bold)
		
def seprate_data_sheet():
    for sheet_name in flow_id:
        if sheet_name:
            worksheet1 = workbook.add_worksheet(sheet_name)
            worksheet1.set_column('C:C', 80)
            worksheet1.set_column('D:D', 80)
            worksheet1.write('A1', 'Sr.No.', bold)
            worksheet1.write('B1', 'Line No.', bold)
            worksheet1.write('C1', 'Soap Request', bold)
            worksheet1.write('D1', 'Soap Response', bold)
            row_req=0
            row_col=0
            line_regex_vc = re.compile("<body>"+sheet_name+"</body>")
            for index, req in enumerate(req_list, start=1):				
				if(line_regex_vc.search(req)):
				    line_regex_req = re.compile(find_between( req, "<soap:Body><", "Request xmlns=" )+"Response")
				    row_req=row_req+1
				    worksheet1.write(row_req, 0, row_req)
				    worksheet1.write(row_req, 1, index)					
				    worksheet1.write(row_req, 2, req,wrap)
				    for index, resp in enumerate(resp_list, start=1):
						if(line_regex_req.search(resp) and line_regex_vc.search(resp) ):
							#row_col=row_col+1
							worksheet1.write(row_req, 3, resp,wrap)		


def base_processing():
	i=0
	# Regex used to match relevant loglines (in this case, a specific soap address)
	line_regex_req = re.compile("<soap:Envelope xmlns:soap=")
	line_regex_resp=re.compile("Response>") 
	# Output file, where the matched loglines will be copied to
	output_filename = os.path.normpath("parsed_lines.log")
	# Overwrites the file, ensure we're starting out with a blank file
	with open(output_filename, "w") as out_file:
		out_file.write("")
	# Open output file in 'append' mode
	with open(output_filename, "a") as out_file:
		# Open input file in 'read' mode
		with open(src_file, "r") as in_file:
			# Loop over each log line
			#for line in in_file:
			 for index, line in enumerate(in_file, start=1):
				
				# If log line matches our regex, print to console, and output file
				if (line_regex_req.search(line)):
					i=i+1
					worksheet.write(i, 0, i)
					worksheet.write(i, 1, index)
					print line
					printReqLine="Request  : "+line
					flow_id.add(find_between( line, "<body>", "</body>" ))
					req_list.append(line)
					out_file.write(printReqLine)
					worksheet.write(i, 2, line,wrap)
				elif (line_regex_resp.search(line)):
					worksheet.write(i, 0, i)
					print line
					resp_list.append(line)
					printRespLine="Response : "+line
					flow_id.add(find_between( line, "<body>", "</body>" ))			
					out_file.write(printRespLine)
					worksheet.write(i, 3, line,wrap)


#Execution Started from here
base_processing()
#print flow_id
#print req_list
#print resp_list
seprate_data_sheet()	
workbook.close()
