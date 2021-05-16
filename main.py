from flask import Flask, render_template, request,send_from_directory
from members import member
import pandas as pd
import math
import random
import os

app = Flask(__name__)

@app.route('/')
@app.route('/index')
@app.route('/retest')
def index():
	return render_template('index.html',members=member)

@app.route('/retestCompute', methods = ['GET', 'POST'])
def retestCompute():
	if request.method == 'POST':
		os.remove("JioSaavn.xlsx")
		#get csv file from user.
		f = request.files['file']
		f.save(f.filename)
		absentMembers=request.form.getlist('members')
		members=member

		#remove absent members from list.
		for i in absentMembers:
			members.remove(i)	
			
		presentMembers=len(members)
		#init dataframe by taking values from csv.
		try:
			query_convertor="iconv -f utf-16 -t utf-8 "+f.filename+" > conv.csv"
			os.system(query_convertor)
			df=pd.read_csv('conv.csv',sep="\t")
			df = df[['Id', 'Status', 'Severity','Summary']].copy()
			#os.remove("JioSaavn.xlsx")
		except:
			print("Please check file name")
			return '<html><body><h1>Please check file name website is still in progress for validations for other file formats!</h1></body></html>'
		
		#df=pd.read_csv('conv.csv',sep="\t")
		#df1 = pd.read_csv(f.filename)
		#df = df[['Id', 'Status', 'Severity','Summary']].copy()

		###
		# create table for members.

		start_no=a=2
		end_no=(len(df.index)+1)
		emergency = request.form['emergency']
		
		ngrp=0
		if emergency == "no":
				prevBestCase=1000
				for group in range(3,6):
						spanGrp=math.ceil(presentMembers/group)
						bestCase=((spanGrp*group)-presentMembers)
						if bestCase < prevBestCase:
							prevBestCase=bestCase
							ngrp=group
		else:
				ngrp=2

		issue_grp=math.ceil((len(members))/ngrp)
		number_of_issues=end_no-start_no+1		

		span_grp=number_of_issues/issue_grp
		round_span_grp=math.ceil(span_grp)
		diff=issue_grp-(number_of_issues%issue_grp)	

		main_list=list()
		if diff != issue_grp:
				for i in range(1,issue_grp+1):
						if i <= (issue_grp-diff):
								main_list.append(str(a)+" to "+str(a+round_span_grp-1))
								a=a+round_span_grp
						else:
								main_list.append(str(a)+" to "+str(a+round_span_grp-2))
								a=a+round_span_grp-1
		else:
				for i in range(1,issue_grp+1):
						main_list.append(str(a)+" to "+str(a+round_span_grp-1))
						a=a+round_span_grp

		random.shuffle(members)

		header=list()
		for i in range(ngrp):
			header.append("Group"+str(i+1))

		memberTable={"Issue List":main_list}
		
		count=0
		for i in range(ngrp):
			rowData=[]
			for j in range(len(main_list)):
				if count < len(members):
					rowData.append(members[count])
				else:
					rowData.append(None)
				count+=1
			memberTable[header[i]]=rowData

		print(memberTable)
		df_member=pd.DataFrame(memberTable)
		print(df_member)

		###
		
		sheet_name = 'Retest Sheet'
		sheet_name2 = 'Assign Issue'
		writer= pd.ExcelWriter('JioSaavn.xlsx',engine='xlsxwriter')

		df.to_excel(writer, sheet_name=sheet_name,index=False)
		df_member.to_excel(writer, sheet_name=sheet_name2,index=False)
		
		workbook  = writer.book
		worksheet = writer.sheets[sheet_name]
		worksheet2 = writer.sheets[sheet_name2]

		header_format = workbook.add_format({
			'border': 1,
			'bg_color': '#FFFF00',
			'bold': True,
			'text_wrap': True,
			'align': 'center',
			'valign': 'vcenter',
			})

		issue_format = workbook.add_format({
			'border': 1,
			'text_wrap': True,
			'align' : 'center',
			'valign': 'vcenter'
			})

		summary_format = workbook.add_format({
			'border': 1,
			'text_wrap': True,
			'valign': 'vcenter',
			'align' : 'vjustify'
			})

		cond_persists = workbook.add_format({
			'bg_color': '#FFC7CE',
			'font_color': '#9C0006',
			'border': 1,
			'text_wrap': True
			})

		cond_resolved = workbook.add_format({
			'bg_color': '#C6EFCE',
      'font_color': '#006100',
			'border': 1,
			'text_wrap': True,
			'valign': 'vcenter',
			'align' : 'center'
			})	

		cond_cbt = workbook.add_format({
			'bg_color': '#FFEB9C',
    	'font_color': '#9C6500',
			'border': 1,
			'text_wrap': True,
			'valign': 'vcenter',
			'align' : 'center'
			})	

		worksheet.set_column('A:A', 15)
		worksheet.set_column('B:B', 15)
		worksheet.set_column('C:C', 15)
		worksheet.set_column('D:D', 85)
		worksheet.set_row(0, 25)
		worksheet.set_column(4, end_no-1, 16)
		cond_persists.set_align('center')
		cond_persists.set_align('vcenter')

		for col_num, value in enumerate(df.columns.values):
			worksheet.write(0, col_num, value, header_format)

		r = 1
		c = 0

		for index, row in df.iterrows():
			worksheet.write(r, c,row['Id'],issue_format)
			worksheet.write(r, c + 1,row['Status'],issue_format)
			worksheet.write(r, c + 2,row['Severity'],issue_format)
			worksheet.write(r, c + 3,row['Summary'],summary_format)
			r += 1

		##Conditional Formatting

		worksheet.data_validation(1, 4, end_no-1, ngrp+4-1, {'validate': 'list',
                                 'source': ['Persists', 'Resolved', 'Cannot be tested']})

		worksheet.conditional_format(1, 4, end_no-1, ngrp+4-1, {'type':     'cell',
                                    'criteria': 'equal to',
                                    'value':    '"Persists"',
                                    'format':   cond_persists})

		worksheet.conditional_format(1, 4, end_no-1, ngrp+4-1, {'type':     'cell',
                                    'criteria': 'equal to',
                                    'value':    '"Resolved"',
                                    'format':   cond_resolved})

		worksheet.conditional_format(1, 4, end_no-1, ngrp+4-1, {'type':     'cell',
                                    'criteria': 'equal to',
                                    'value':    '"Cannot be tested"',
                                    'format':   cond_cbt})

		c=0
		for i in range(4,ngrp+4):
			worksheet.write(0, i, header[c], header_format)
			c+=1
		
		##Assign Issue Formatting
		worksheet2.set_column('A:Z', 12)

		for col_num, value in enumerate(df_member.columns.values):
			worksheet2.write(0, col_num, value, header_format)

		r = 1
		c = 0

		for index, row in df_member.iterrows():
			worksheet2.write(r, c,row['Issue List'],header_format)
			for i in range(len(header)):
					worksheet2.write(r, c+i+1,row[header[i]],issue_format)
			r += 1

		writer.save()
		os.remove(f.filename)
		os.remove("conv.csv")
		return send_from_directory("/home/runner/JMA-Prod/","JioSaavn.xlsx", as_attachment=True)


@app.route('/nextMod')
def nextMod():
	return render_template('index.html')

@app.route('/test')
def test():
	return send_from_directory("/home/runner/JMA-Prod/","JioSaavn.xlsx", as_attachment=True)

app.run('0.0.0.0',8080)