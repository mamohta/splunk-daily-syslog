import xlsxwriter
import pandas
import pandas as pd
import numpy as np 
import requests
import urllib
from xml.dom import minidom
import re
import time
import json
from datetime import datetime, timedelta
import math
import os
import getpass
import sys



cpyName=sys.argv[1]
earliest=sys.argv[2]
latest=sys.argv[3]
todayDate=sys.argv[4]

item='All_Devices' #This is fro create report by Name to be put

reportingPeriod=earliest+" to "+latest




lastColInSummary='K' #So that merge options take place easily


# worksheet_summary = writer.sheets['undefined']

rowNum = 6

msgLink = {}
productFamily = []



def sendBGColor(sev):
	if(sev == '0'):
		return '#fabf8f'
	elif(sev == '1'):
		return '#FFD9D9'
	elif(sev == '2'):
		return '#FFE699'	
	elif(sev == '3'):
		return '#DDEBF7'	
	elif(sev == '4'):
		return '#CCC0DA'
	elif(sev == '5'):
		return '#D8E4BC'
	else:
		return '#F3EBF9'

def sendTextColor(sev):
	if(sev == '0'):
		return '#974706'
	elif(sev == '1'):
		return '#b32400'
	elif(sev == '2'):
		return '#999900'	
	elif(sev == '3'):
		return '#003d99'	
	elif(sev == '4'):
		return '#893BC3'
	elif(sev == '5'):
		return '#4F6228'
	else:
		return '#494529'


def addLinks(startRow,endRow,worksheet_summary,df):
	
	url_format = workbook.add_format({
		'font_color': 'blue',
		'underline':  1,
		'valign':'vcenter',
		'border': 1,
		'border_color' : 'black'
	})
	
	wrapText_format = workbook.add_format({
	# 'wrap':True,
	'valign':'vcenter',
	'border': 1,
	'border_color' : 'black',
	'align': 'left',
	'text_wrap':1})
	# wrapText_format.set_text_wrap()
	
	
	start = startRow
	dict = msgLink

	# print(df.loc[start-8]['Severity'],":::::::::::::::::")
	df = df.sort_values(['Severity','Device Family'],ascending=[True,True])
	for index, row in df.iterrows():
		color = sendBGColor(row['Severity'])
		name = row['Name']
		message = row['Message Name']
		content = row['Sample Message Format']
		sev = str(row['Severity'])
		regex = row['RegEx']
		
		
		worksheet_summary.write('A'+str(startRow),"internal:'Sev-"+str(sev)+"'!A"+str(dict[message+regex]))
		
		worksheet_summary.write('A'+str(startRow),name,workbook.add_format({
	'font_color': 'blue',
	'underline':  1,
	'bg_color':color,
	'valign':'vcenter',
	'border': 1,
	'border_color' : 'black',
	'align': 'left',
	'text_wrap':1}))
	
		worksheet_summary.write('B'+str(startRow),message+": "+content,workbook.add_format({
		'bg_color':color,
		'valign':'vcenter',
		'border': 1,
		'border_color' : 'black',
		'align': 'left',
		'text_wrap':1}))
		worksheet_summary.write('C'+str(startRow),row['Severity'],workbook.add_format({
	'bg_color':color,
	'valign':'vcenter',
	'border': 1,
	'border_color' : 'black',
	'align': 'left',
	'text_wrap':1}))
		worksheet_summary.write('D'+str(startRow),row['Device Family'],workbook.add_format({
	'bg_color':color,
	'valign':'vcenter',
	'border': 1,
	'border_color' : 'black',
	'align': 'left',
	'text_wrap':1})
	)
		worksheet_summary.write('E'+str(startRow),row['No of Devices'],workbook.add_format({
	'bg_color':color,
	'valign':'vcenter',
	'border': 1,
	'border_color' : 'black',
	'align': 'left',
	'text_wrap':1}))		
	
		worksheet_summary.write('F'+str(startRow),row['No of New Devices'],workbook.add_format({
	'bg_color':color,
	'valign':'vcenter',
	'border': 1,
	'border_color' : 'black',
	'align': 'left',
	'text_wrap':1}))
		worksheet_summary.write('G'+str(startRow),row['Message Count'],workbook.add_format({
	'bg_color':color,
	'valign':'vcenter',
	'border': 1,
	'border_color' : 'black',
	'align': 'left',
	'text_wrap':1}))		
	
		worksheet_summary.write('H'+str(startRow),row['7DayHistory'],workbook.add_format({
		'bg_color':color,
		'valign':'vcenter',
		'border': 1,
		'border_color' : 'black',
		'align': 'left',
		'text_wrap':1}))
		
		worksheet_summary.write('I'+str(startRow),row['30DayHistory'],workbook.add_format({
		'bg_color':color,
		'valign':'vcenter',
		'border': 1,
		'border_color' : 'black',
		'align': 'left',
		'text_wrap':1}))		
		worksheet_summary.write('J'+str(startRow),row['Recommended Action'],workbook.add_format({
		'bg_color':color,
		'valign':'vcenter',
		'border': 1,
		'border_color' : 'black',
		'align': 'left',
		'text_wrap':1}))		
		worksheet_summary.write('K'+str(startRow)," ",workbook.add_format({
		'bg_color':color,
		'valign':'vcenter',
		'border': 1,
		'border_color' : 'black',
		'align': 'left',
		'text_wrap':1}))
		
		startRow = startRow + 1



def createSummarySheet(workbook,df_Events,worksheet_summary):

	
	#Format of the title heading e.g. Azure Syslog Report
	heading_format = workbook.add_format({
		'bold': 1,
		'align': 'left',
		'valign': 'vcenter',
		'bottom':5,
		'bottom_color':'gray',
		'border': 1,
		'border_color' : 'black',
		'font_size':36,
		'font_name':'Calibri',
		'font_color':'#000080'})
	#New or old event format
	old_format = workbook.add_format({
		'bold' : 1,
		'italic' : 1,
		'align': 'left',	
		'valign': 'vcenter',
		'bottom_color':'gray',
		'border': 1,
		'border_color' : 'black',
		'font_size':20,
		'font_name':'Calibri',
		'font_color':'white',
		'bg_color':'#F79646'})


	#Reporting Period cells format
	reportingPeriod_format = workbook.add_format({
		'align': 'left',
		'valign': 'vcenter',
		'bottom':5,
		'bottom_color':'gray',
		'font_size':20,
		'border': 1,
		'border_color' : 'black',
		'font_name':'Calibri',
		'font_color':'#000080'})
		
	#New or old event format
	new_format = workbook.add_format({
		'bold' : 1,
		'italic' : 1,
		'align': 'left',	
		'valign': 'vcenter',
		'bottom_color':'gray',
		'font_size':20,
		'border': 1,
		'border_color' : 'black',
		'font_name':'Calibri',
		'font_color':'white',
		'bg_color':'#70AD47'})	
		
	colNames_format = workbook.add_format({
		'bold' : 1,
		'border' : 1,
		'border_color' : 'black',
		'align': 'left',	
		'valign': 'vcenter',
		'font_size':12,
		'font_name':'Calibri',
		'border': 1,
		'border_color' : 'black',
		'font_color':'black',
		'bg_color':'#dfdfdf',
		'text_wrap':1})
	
	
	

	#Give column widths
	worksheet_summary.set_column('A:A', 36.86)
	worksheet_summary.set_column('B:B', 52.14)
	worksheet_summary.set_column('C:C', 5.86)
	worksheet_summary.set_column('D:D', 14)
	worksheet_summary.set_column('E:E', 7.57)
	worksheet_summary.set_column('F:F', 7.71)
	worksheet_summary.set_column('G:G', 8.57)
	worksheet_summary.set_column('H:H', 11.14)
	worksheet_summary.set_column('I:I', 12.14)
	worksheet_summary.set_column('J:J', 45)
	

	
	
	#Headings
	worksheet_summary.merge_range('A1:'+lastColInSummary+'1', cpyName+' '+item+' Syslog Report', heading_format)
	worksheet_summary.merge_range('A2:'+lastColInSummary+'2', 'Reporting Period: ' +reportingPeriod, reportingPeriod_format)
	
	
	worksheet_summary.write('C3',"Sev 0",workbook.add_format({
		'align': 'left',	
		'font_name':'Calibri',
		'font_color':'#974706',
		'border': 1,
		'border_color' : 'black',
		'bg_color':'#fabf8f'}))
	
	worksheet_summary.write('D3',"Sev 1",workbook.add_format({
		'align': 'left',	
		'font_name':'Calibri',
		'font_color':'#b32400',
		'border': 1,
		'border_color' : 'black',
		'bg_color':'#FFD9D9'}))
	worksheet_summary.write('E3',"Sev 2",workbook.add_format({
		'align': 'left',	
		'font_name':'Calibri',
		'font_color':'#9C6500',
		'border': 1,
		'border_color' : 'black',
		'bg_color':'#FFE699'}))
	worksheet_summary.write('F3',"Sev 3",workbook.add_format({
		'bold' : 1,
		'align': 'left',	
		'font_name':'Calibri',
		'font_color':'#003d99',
		'bg_color':'#DDEBF7',
		'border': 1,
		'border_color' : 'black'}))
	worksheet_summary.write('G3',"Sev 4",workbook.add_format({
		'bold' : 1,
		'align': 'left',	
		'font_name':'Calibri',
		'font_color':'#893BC3',
		'border': 1,
		'border_color' : 'black',
		'bg_color':'#E6DCE0'}))
	worksheet_summary.write('H3',"Sev 5",workbook.add_format({
		'bold' : 1,
		'align': 'left',	
		'font_name':'Calibri',
		'font_color':'#4F6228',
		'border': 1,
		'border_color' : 'black',
		'bg_color':'#D8E4BC'}))
	worksheet_summary.write('I3',"Sev 6",workbook.add_format({
		'bold' : 1,
		'align': 'left',	
		'font_name':'Calibri',
		'font_color':'#494529',
		'border': 1,
		'border_color' : 'black',
		'bg_color':'#DDD9C4'}))
	
	
	worksheet_summary.merge_range('A5:'+lastColInSummary+'5', 'New Events', new_format)
	
	
	# df_newEvents.to_excel(writer,sheet_name = "Summary",header = True,index = False,startrow = 6)
	worksheet_summary.write('A7',"Message Name",colNames_format)
	worksheet_summary.write('B7',"Sample Message Content",colNames_format)
	worksheet_summary.write('C7',"Sev",colNames_format)
	worksheet_summary.write('D7',"Device Family",colNames_format)
	worksheet_summary.write('E7',"Num of Devices",colNames_format)
	worksheet_summary.write('F7',"Num of New Devices",colNames_format)
	worksheet_summary.write('G7',"Message Count",colNames_format)
	worksheet_summary.write('H7',"Num Occurrence Last 7 Days",colNames_format)
	worksheet_summary.write('I7',"Num Occurrence Last 30 Days",colNames_format)
	worksheet_summary.write('J7',"Cisco Recommended Action",colNames_format)
	worksheet_summary.write('K7',cpyName+" Action",colNames_format)
	

	# print(df)
	df_new = df_Events[df_Events['7DayHistory']==0]
	addLinks(8,df_new['Message Name'].count()+7,worksheet_summary,df_new)
	
	df_old = df_Events[df_Events['7DayHistory']>0]
	
	start_old = df_new['Message Name'].count() + 9
	worksheet_summary.merge_range('A'+str(start_old)+':'+lastColInSummary+str(start_old), 'Repeat Events(Last 7 Days)', old_format)
	
	worksheet_summary.write('A'+str(start_old+2),"Message Name",colNames_format)
	worksheet_summary.write('B'+str(start_old+2),"Sample Message Content",colNames_format)
	worksheet_summary.write('C'+str(start_old+2),"Sev",colNames_format)
	worksheet_summary.write('D'+str(start_old+2),"Device Family",colNames_format)
	worksheet_summary.write('E'+str(start_old+2),"Num of Devices",colNames_format)
	worksheet_summary.write('F'+str(start_old+2),"Num of new Devices",colNames_format)
	worksheet_summary.write('G'+str(start_old+2),"Message Count",colNames_format)
	worksheet_summary.write('H'+str(start_old+2),"Num Occurrence Last 7 Days",colNames_format)
	worksheet_summary.write('I'+str(start_old+2),"Num Occurrence Last 30 Days",colNames_format)
	worksheet_summary.write('J'+str(start_old+2),"Cisco Recommended Action",colNames_format)
	worksheet_summary.write('K'+str(start_old+2),cpyName+" Action",colNames_format)
	
	addLinks(start_old+3,df_old['Message Name'].count()+7,worksheet_summary,df_old)
	

		



def populateSheet(df1,worksheet,workbook,sev,df_Events):	
	global rowNum
	global rowNumSum
	global msgLink_new
	global msgLink_old
	
	colNames_format = workbook.add_format({
		'bold' : 1,
		'border' : 1,
		'border_color' : 'black',
		'align': 'left',	
		'valign': 'vcenter',
		'font_size':11.5,
		'font_name':'Calibri',
		'font_color':'black',
		'bg_color':'#dfdfdf',
		'text_wrap':1})	
	
	names_format = workbook.add_format({
		'bold' : 1,
		'border' : 1,
		'border_color' : 'black',
		'align': 'left',	
		'valign': 'vcenter',
		'font_size':11.5,
		'font_name':'Calibri',
		'font_color':'black'})

	text_format = workbook.add_format({
		'border' : 1,
		'border_color' : 'black',
		'align': 'left',	
		'valign': 'vcenter',
		'font_size':11.5,
		'font_name':'Calibri',
		'font_color':'black',
		'text_wrap':1})

	top_border = workbook.add_format({
		'bottom':2,
		'bg_color':sendTextColor(sev),
		'top_color':'black'})	
	cell_format = workbook.add_format({
		'bg_color':sendBGColor(sev)})
		
	
	

	
	regex = df1['RegEx'].unique()
	for eachRegex in regex:
		rowNum = rowNum + 1
		worksheet.merge_range('A'+str(rowNum)+':I'+str(rowNum), '', cell_format)

		
		rowNum = rowNum + 1
		
		df = df1[df1['RegEx']==eachRegex]
		name = df['Name'].iloc[0] #This is the message value taken from first row
		message = df['Message'].iloc[0]
		content = df['Message'].iloc[0] + ": " + df['Message Content'].iloc[0] #This is the message content value taken from first row
		content_summary = df['Message Content'].iloc[0]  #This is the message content value taken from first row
		description = df['Description'].iloc[0]
		action = df['Recommended Action'].iloc[0]
		regex = df['RegEx'].iloc[0]
		action = df['Recommended Action'].iloc[0]
		productFamilyList="" #String of all product families occurring on that message
		productFamily = df['Product Family'].unique()
		
		
		for product in productFamily:
			productFamilyList=productFamilyList + product +"\n\n"
		
		product=df['Product Family'].iloc[0]
		
		numNewDevices = df[df['7DayHistory']==0]["Device"].count()
		# numNewDevices = 0
	 
		# if (len(df.loc[df['Previously Reported'] == "Yes"]) >0):
			# previouslyReported = "Yes"
		# else:
			# previouslyReported = "No"
			
		


		

		df_Events.loc[rowNumSum] = [name, message, content_summary, sev,productFamilyList,df['Device'].count(),numNewDevices,df['Count'].sum(),regex,df['7DayHistory'].sum(),df['30DayHistory'].sum(),str(action)]
		rowNumSum = rowNumSum + 1
		msgLink[message+regex] = rowNum

			
		df = df.sort_values(['7DayHistory','Count'],ascending=[True,True])
		
		worksheet.write('A'+str(rowNum),"Name",colNames_format)
		worksheet.write('B'+str(rowNum),name,names_format)
		worksheet.merge_range('C'+str(rowNum)+':I'+str(rowNum), '', cell_format)
		rowNum = rowNum + 1
		
		worksheet.write('A'+str(rowNum),"Sample Message Content",colNames_format)
		worksheet.write('B'+str(rowNum),content,text_format)
		
		worksheet.merge_range('C'+str(rowNum)+':I'+str(rowNum), '', cell_format)
		rowNum = rowNum + 1
		
		worksheet.write('A'+str(rowNum),"Description",colNames_format)
		worksheet.write('B'+str(rowNum),str(description),text_format)

		worksheet.merge_range('C'+str(rowNum)+':I'+str(rowNum), '', cell_format)
		rowNum = rowNum + 1
		
		worksheet.write('A'+str(rowNum),"Recommended Action",colNames_format)
		worksheet.write('B'+str(rowNum),str(action),text_format)

		worksheet.merge_range('C'+str(rowNum)+':I'+str(rowNum), '', cell_format)
		rowNum = rowNum + 1
		
	
		worksheet.merge_range('A'+str(rowNum)+':I'+str(rowNum), '', cell_format)
		rowNum = rowNum + 1

		
		
		worksheet.write('A'+str(rowNum),"Device Name",colNames_format)
		worksheet.write('B'+str(rowNum),"Message Content",colNames_format)
		worksheet.write('C'+str(rowNum),"Count",colNames_format)
		worksheet.write('D'+str(rowNum),"TimeStamp",colNames_format)
		worksheet.write('E'+str(rowNum),"SW Version",colNames_format)
		worksheet.write('F'+str(rowNum),"Product ID",colNames_format)
		worksheet.write('G'+str(rowNum),"7DayHistory",colNames_format)
		worksheet.write('H'+str(rowNum),"30DayHistory",colNames_format)
		worksheet.write('I'+str(rowNum),"Custom Analysis",colNames_format)
		rowNum = rowNum + 1
		
		
		
		for index,row in df.iterrows():
			worksheet.write('A'+str(rowNum),str(row['Device']),text_format)
			worksheet.write('B'+str(rowNum),str(row['Message Content']),text_format)
			worksheet.write('C'+str(rowNum),str(row['Count']),text_format)
			worksheet.write('D'+str(rowNum),str(row['TimeStamp']),text_format)
			worksheet.write('E'+str(rowNum),str(row['SW Version']),text_format)
			worksheet.write('F'+str(rowNum),str(row['Product ID']),text_format)
			worksheet.write('G'+str(rowNum),str(row['7DayHistory']),text_format)
			worksheet.write('H'+str(rowNum),str(row['30DayHistory']),text_format)
			worksheet.write('I'+str(rowNum),"",text_format)
			rowNum = rowNum + 1
		

		worksheet.merge_range('A'+str(rowNum)+':I'+str(rowNum), '', cell_format)
		
	rowNum = rowNum + 2


def createSevSheet(df,sev,workbook,df_Events):
	
	global rowNum
	global groupSel
	
	
	# df_waste.to_excel(writer, sheet_name="Sev-"+str(sev))
	# worksheet = writer.sheets["Sev-"+str(sev)]
	worksheet=workbook.add_worksheet("Sev-"+str(sev))
	#Format of the title heading e.g. Azure Syslog Report
	heading_format = workbook.add_format({
		'bold': 1,
		'align': 'left',
		'valign': 'vcenter',
		'bottom':5,
		'bottom_color':'gray',
		'font_size':36,
		'font_name':'Calibri',
		'font_color':'#000080'})


	#Reporting Period cells format
	reportingPeriod_format = workbook.add_format({
		'align': 'left',
		'valign': 'vcenter',
		'bottom':5,
		'bottom_color':'gray',
		'font_size':20,
		'font_name':'Calibri',
		'font_color':'#000080'})
		
	#New or old event format
	new_format = workbook.add_format({
		'bold' : 1,
		'italic' : 1,
		'align': 'left',	
		'valign': 'vcenter',
		'bottom_color':'gray',
		'font_size':20,
		'font_name':'Calibri',
		'font_color':'white',
		'bg_color':'#70AD47'})		
	
	#New or old event format
	old_format = workbook.add_format({
		'bold' : 1,
		'italic' : 1,
		'align': 'left',	
		'valign': 'vcenter',
		'bottom_color':'gray',
		'font_size':20,
		'font_name':'Calibri',
		'font_color':'white',
		'bg_color':'#F79646'})	
		
	#Device Family format
	device_family = workbook.add_format({
		'bold' : 1,
		'align': 'left',	
		'valign': 'vcenter',
		'bottom_color':'gray',
		'font_size':17,
		'font_name':'Calibri',
		'font_color':'white',
		'bg_color':'#2F75B5'})
		
	sev_format = workbook.add_format({
		'bold' : 1,
		'bottom':2,
		'top':5,
		'align': 'left',	
		'valign': 'vcenter',
		'bottom_color':sendTextColor(sev),
		'top_color':sendTextColor(sev),
		'font_size':20,
		'font_name':'Calibri',
		'font_color':sendTextColor(sev),
		'bg_color':sendBGColor(sev)})
	

	

	#Give column widths
	worksheet.set_column('A:A', 28)
	worksheet.set_column('B:B', 90)
	worksheet.set_column('C:C', 8.43)
	worksheet.set_column('D:D', 18.29)
	worksheet.set_column('E:E', 13.71)
	worksheet.set_column('F:F', 13.71)
	worksheet.set_column('G:G', 13.71)
	worksheet.set_column('H:I', 13.71)
	worksheet.set_column('I:I', 40.43)
	
	#Headings
	worksheet.merge_range('A1:I1', cpyName+' '+item+' Syslog Report', heading_format)
	worksheet.merge_range('A2:I2', 'Reporting Period: '+ reportingPeriod, reportingPeriod_format)
	worksheet.merge_range('A4:I4', 'Severity Level : '+str(sev), sev_format)

	populateSheet(df,worksheet,workbook,sev,df_Events)


	




def create(dataFrame,workbook):
	global rowNum
	global rowNumSum
	global msgLink

	
	df_Events = pd.DataFrame(columns = ['Name','Message Name','Sample Message Format','Severity','Device Family','No of Devices','No of New Devices','Message Count','RegEx','7DayHistory','30DayHistory','Recommended Action'])
	# df_oldEvents = pd.DataFrame(columns = ['Name','Message Name','Sample Message Format','Severity','Device Family','No of Devices','Message Count','RegEx','Number of Times Occured Last Month'])
	# df_waste.to_excel(writer, sheet_name="Summary")
	worksheet_summary = workbook.add_worksheet("Summary")
	rowNumSum = 0
	msgLink = {}
	

	
	severity = dataFrame['Sev'].unique()
	severity.sort()
	for sev in severity:
		df_sev = dataFrame[(dataFrame['Sev']==sev) & (~dataFrame['RegEx'].str.contains("no match")) & (~dataFrame['Filter'].str.contains("Y"))] #This creates a data frame with all values of sev severity
		if not(sev == 'undefined'):
			sev = str(sev)
		if not df_sev.empty:
			createSevSheet(df_sev,sev,workbook,df_Events)
			# print(df_Events.to_string())
			rowNum = 6 #Since every time in new sheet we have to start populating from row 6
	
	

	#creating summary sheet
	createSummarySheet(workbook,df_Events,worksheet_summary)
	del df_Events
	
#########################################################################################################################
#############################THIS PART IS READING FROM A JSON FILE to get cpyName details##############################
#########################################################################################################################
#########################################################################################################################


folderName = os.path.join("Reports",cpyName,datetime.strftime(datetime.now(),"%m-%d"),earliest+' to '+latest)

data_df=pd.read_excel(os.path.join(folderName,cpyName+'_Raw_'+earliest+'_to_'+latest+'.xlsx'))

data_df = data_df.replace(np.nan, 'nan', regex=True)

# print(data_df)
data_df[['7DayHistory']]=data_df[['7DayHistory']].astype(int)
data_df[['30DayHistory']]=data_df[['30DayHistory']].astype(int)
data_df[['Count']]=data_df[['Count']].astype(int)
# # data_df[['Sev']]=data_df[['Sev']].astype(int)









# ########################################################################################################################
# ########################################################################################################################
# ########################################################################################################################
# ###########################THIS PART IS CREATING EXCEL OUT OF THE DATAFRAME#############################################
# ########################################################################################################################
# ########################################################################################################################
# ########################################################################################################################
# ########################################################################################################################
# print("Creating Excel")

createReportBy = data_df['createReportBy'].unique()

for item in createReportBy:
	# item=item.replace("_", "-")
	nameOfFile=cpyName+'_Syslog Analysis_'+earliest+' to '+latest+'_'+item+'.xlsx'
	writer = pd.ExcelWriter(os.path.join(folderName,nameOfFile), engine='xlsxwriter')
	workbook  = writer.book
	dataFrame=data_df[(~(data_df['Filter'].str.contains("Y"))) & ((data_df['createReportBy'].str.contains(item)))]
	create(dataFrame,workbook)

	data_df_product_new = data_df[(data_df['RegEx'].str.contains("no match"))& ((data_df['createReportBy'].str.contains(item)))]
	data_df_product_new = data_df_product_new.sort_values(['Sev'],ascending=[True])
	data_df_product_new.to_excel(writer,sheet_name = "OtherMessages",index = False,columns = ['Product Family','Device','SW Version','Product ID','Message','Message Content','Description','Recommended Action','Sev','Count','TimeStamp','RegEx'])
	worksheet = writer.sheets["OtherMessages"]
	# Give column widths
	worksheet.set_column('A:A', 5.57)
	worksheet.set_column('B:B', 8.71)
	worksheet.set_column('C:C', 6.29)
	worksheet.set_column('D:D', 11.86)
	worksheet.set_column('E:E', 29.14)
	worksheet.set_column('F:F', 37.71)
	worksheet.set_column('G:G', 33)
	worksheet.set_column('H:H', 41.86)
	worksheet.set_column('I:I', 5.86)
	worksheet.set_column('J:J',	7.86)
	worksheet.set_column('K:K',	7)
	worksheet.set_column('L:L',	15)
	
	data_df_product_filter = data_df[(data_df['Filter'].str.contains("Y"))& ((data_df['createReportBy'].str.contains(item)))]
	data_df_product_filter = data_df_product_filter.sort_values(['Sev'],ascending=[True])
	data_df_product_filter.to_excel(writer,sheet_name = "NonActionableMessages",index = False,columns = ['Product Family','Device','SW Version','Product ID','Message','Message Content','Description','Recommended Action','Sev','Count','TimeStamp','RegEx','30DayHistory','7DayHistory'])
	worksheet = writer.sheets["NonActionableMessages"]
	# Give column widths
	worksheet.set_column('A:A', 5.57)
	worksheet.set_column('B:B', 8.71)
	worksheet.set_column('C:C', 6.29)
	worksheet.set_column('D:D', 11.86)
	worksheet.set_column('E:E', 29.14)
	worksheet.set_column('F:F', 37.71)
	worksheet.set_column('G:G', 33)
	worksheet.set_column('H:H', 41.86)
	worksheet.set_column('I:I', 5.86)
	worksheet.set_column('J:J',	7.86)
	worksheet.set_column('K:K',	7)
	worksheet.set_column('L:L',	15)
	
	writer.save()
	writer.close()

