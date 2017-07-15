import smtplib
from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase 
import sys
import os
import pandas as pd

def sendBGColor(sev):
	if(sev == '0' or sev==0):
		return '#fabf8f'
	elif(sev == '1' or sev==1):
		return '#FFD9D9'
	elif(sev == '2' or sev==2):
		return '#FFE699'	
	elif(sev == '3' or sev==3):
		return '#DDEBF7'	
	elif(sev == '4' or sev==4):
		return '#CCC0DA'
	elif(sev == '5' or sev==5):
		return '#D8E4BC'
	else:
		return '#F3EBF9'

def sendTextColor(sev):
	if(sev == '0' or sev==0):
		return '#974706'
	elif(sev == '1' or sev==1):
		return '#b32400'
	elif(sev == '2' or sev==2):
		return '#999900'	
	elif(sev == '3' or sev==3):
		return '#003d99'	
	elif(sev == '4' or sev==4):
		return '#893BC3'
	elif(sev == '5' or sev==5):
		return '#4F6228'
	else:
		return '#494529'



to=sys.argv[1]
cc=sys.argv[2]
folderName=sys.argv[3]
earliest=sys.argv[4]
latest=sys.argv[5]
cpyName=sys.argv[6]

reportingPeriod=earliest+" to "+latest

me = "splunk-daily-syslog@cisco.com"


df_raw = pd.read_excel(os.path.join(folderName,cpyName+'_Raw_'+earliest+'_to_'+latest+'.xlsx'),keep_default_na=False)

for file in os.listdir(folderName):
	# html=" "
	if "Raw" in (file):
		# df_raw = pd.read_excel(os.path.join(folderName,file),keep_default_na=False)
		continue
	itemArray=file.split(latest+'_')
	item=itemArray[len(itemArray)-1][:-5]
	df = pd.read_excel(os.path.join(folderName,file),keep_default_na=False)
	# print(item)
	df_filtered = df_raw[(df_raw['createReportBy'].str.contains(item))& (~df_raw['RegEx'].str.contains("no match")) & (~df_raw['Filter'].str.contains("Y"))]
	html="<html><style>\
	.Heading\
		{color:black;\
		font-size:12.0pt;\
		font-weight:700;\
		text-align:left;\
		vertical-align:middle;\
		border:.5pt solid black;\
		background:#DFDFDF;\
		white-space:normal;}\
	.Content\
		{word-wrap:break-word;\
		text-align:left;\
		vertical-align:middle;\
		border:.5pt solid black;\
		white-space:normal;}\
	</style>\
	"

	# print(list(df.columns.values)[0])
	html = html + "\
	<table style='border-collapse:collapse;table-layout:fixed'>\
	  <tr height=62 style='height:46.5pt'>\
	  <td colspan=11 height=62 style='height:46.5pt; color:navy;\
		font-size:36.0pt;\
		font-weight:700;\
		text-align:left;\
		vertical-align:middle;\
		border:.5pt solid black;'><a name='Top'>"+cpyName+" "+item+" Syslog Report</a></td>\
	 </tr>\
	"
	for row in df.itertuples():
		# print(row[1])
		if "Reporting Period" in row[1]:
			html=html+"\
				<tr height=40 style='height:40.5pt'>\
				  <td colspan=11 height=40 style='height:40pt;color:navy;\
					font-size:20.0pt;\
					font-weight:400;\
					text-align:left;\
					vertical-align:middle;\
					border:.5pt solid black;'>Reporting Period : "+reportingPeriod+"</td>\
				 </tr> "
		elif "New Events" in row[1] or "Repeat Events" in row[1]:
			html=html + "\
				<tr height=16 style='height:16.0pt'>\
				  <td height=16 colspan=11 style='height:16pt;border:.5pt solid black;'></td>\
				</tr>"
			if("Repeat Events" in row[1]):
				bgcolor="#F79646"
			elif('New Events' in row[1]):
				bgcolor="#70AD47"
			html=html+"\
				<tr height=35 style='height:26.25pt'>\
				  <td colspan=11 height=35 bgcolor='"+bgcolor+"' \
					style='height:26.25pt;	color:white;\
					font-size:20.0pt;\
					font-weight:700;\
					font-style:italic;\
					text-align:left;\
					vertical-align:middle;\
					border:.5pt solid black;\'>"+row[1]+"</td>\
				 </tr>\
				 <tr height=16 style='height:16.0pt'>\
				  <td height=16 colspan=11 style='height:16.0pt;border:.5pt solid black;'></td>\
				</tr>"
		elif "Message Name" in row[1]:
			html=html +"\
				<tr>\
				  <td height=84 class='Heading'>Message Name</td>\
				  <td class='Heading' width=617 style='border-left:none;width:463pt'>Sample Message Content</td>\
				  <td class='Heading' width=72 style='border-left:none;width:54pt'>Sev</td>\
				  <td class='Heading' width=103 style='border-left:none;width:77pt'>Device Family</td>\
				  <td class='Heading' width=58 style='border-left:none;width:44pt'>Num of Devices</td>\
				  <td class='Heading' width=94 style='border-left:none;width:71pt'>Num of New Devices</td>\
				  <td class='Heading' width=65 style='border-left:none;width:49pt'>Message Count</td>\
				  <td class='Heading' width=83 style='border-left:none;width:62pt'>Num Occurrence Last 7 Days</td>\
				  <td class='Heading' width=90 style='border-left:none;width:68pt'>Num Occurrence Last 30 Days</td>\
				  <td class='Heading' width=90 style='border-left:none;width:68pt'>Cisco Recommended Action</td>\
				  <td class='Heading' width=90 style='border-left:none;width:68pt'>"+cpyName+" Action</td>\
				</tr>"
		elif len(row[1])>0:
			html=html +\
				"<tr height=40 style='height:30.0pt'>"
			for i in range(1,len(df.columns)+1):
				if(i==1):
					# print(row[i])
					html=html+\
						"<td height=40 class=Content width=345 bgcolor='"+sendBGColor(str(row[3]))+"'style='height:30.0pt;border-top:none;width:259pt'><a href='#"+str(row[i])+"'>" + str(row[i]) + "</a></td>"
				else:				
					html=html+\
						"<td height=40 class=Content width=345 bgcolor='"+sendBGColor(str(row[3]))+"'style='height:30.0pt;border-top:none;width:259pt'>" + str(row[i]) + "</td>"
					
			html=html + "</tr>"
	html=html+\
		"<tr height=16 style='height:16.0pt'>\
		  <td height=16 colspan=11 style='height:16.0pt;border:.5pt solid black;'></td>\
		</tr>"	
	sev=df_filtered['Sev'].unique()
	sev.sort()
	# print(df_raw)
	for severity in sev:
		html=html+"\
			<tr height=35 style='height:26.25pt'>\
			  <td colspan=11 height=35 bgcolor='"+sendBGColor(severity)+"' \
				style='height:26.25pt;	color:"+sendTextColor(severity)+";\
				font-size:20.0pt;\
				font-weight:700;\
				text-align:left;\
				vertical-align:middle;\
				border:3px solid black;\'>Severity:"+str(severity)+"</td>\
			 </tr>\
			 <tr height=16 style='height:16.0pt'>\
			  <td height=16 colspan=11 style='height:16.0pt;border:.5pt solid black;'></td>\
			</tr>"
		df_sev = df_filtered[(df_filtered['Sev']==severity) & (~df_filtered['RegEx'].str.contains("no match")) & (~df_filtered['Filter'].str.contains("Y"))] #This creates a data frame with all values of sev severity
		
		# print(df_sev)
		regex = df_sev['RegEx'].unique()
		for eachRegex in regex:
			# print(severity)
			df = df_sev[df_sev['RegEx']==eachRegex]
			# print(df)
			name = df['Name'].iloc[0] #This is the message value taken from first row
			message = df['Message'].iloc[0]
			content = df['Message'].iloc[0] + ": " + df['Message Content'].iloc[0] #This is the message content value taken from first row
			content_summary = df['Message Content'].iloc[0]  #This is the message content value taken from first row
			description = df['Description'].iloc[0]
			action = df['Recommended Action'].iloc[0]
			regex = df['RegEx'].iloc[0]
			action = df['Recommended Action'].iloc[0]
			# print("::::::::::",name)
			html=html+\
			"<tr>\
				  <td height=20 class='Heading' style='border-left:2px solid black;border-top:2px solid black'><a name='"+name+"'>Message Name</a></td>\
				  <td height=20 class=Content width=345><b>" + name + "</b></td>\
				  <td height=20 style='border-right:2px solid black;border-top:2px solid black'  colspan=9 bgcolor='"+sendBGColor(severity)+"'><a href=#Top>Click to return to top</a></td>\
			 <\tr>\
			 <tr>\
				  <td height=20 class='Heading' style='border-left:2px solid black'>Sample Mesage Content</td>\
				  <td height=20 class=Content width=345>" + content + "</td>\
				  <td colspan=9 style='border-right:2px solid black' bgcolor='"+sendBGColor(severity)+"'></td>\
			  <\tr>\
			  <tr>\
				  <td height=20 class='Heading' style='border-left:2px solid black'>Description</td>\
				  <td height=20 class=Content width=345>" + str(description) + "</td>\
				  <td colspan=9 style='border-right:2px solid black' bgcolor='"+sendBGColor(severity)+"'></td>\
			  <\tr>\
			  <tr>\
				  <td height=20 class='Heading' style='border-left:2px solid black'>Recommended Action</td>\
				  <td height=20 class=Content width=345>" + str(action) + "</td>\
				  <td colspan=9 style='border-right:2px solid black' bgcolor='"+sendBGColor(severity)+"'></td>\
			  <\tr>\
			  <tr>\
				<td colspan=11 style='border-left:2px solid black;border-right:2px solid black' bgcolor='"+sendBGColor(severity)+"'></td>\
			  </tr>\
			  <tr>\
				  <td class='Heading' style='border-left:2px solid black'>Device Name</td>\
				  <td class='Heading'>Message Content</td>\
				  <td class='Heading'>Count</td>\
				  <td class='Heading'>TimeStamp</td>\
				  <td class='Heading'>SW Version</td>\
				  <td class='Heading'>Product ID</td>\
				  <td class='Heading'>7 Day History</td>\
				  <td class='Heading'>30 Day History</td>\
				  <td class='Heading' colspan=3 style='border-right:2px solid black'>Custom Analysis</td>\
			  </tr>"
			for index,row in df.iterrows():
				html=html+"\
				<tr>\
				  <td class=Content style='border-left:2px solid black'>" + str(row['Device']) + "</td>\
				  <td class=Content>" + str(row['Message Content']) + "</td>\
				  <td class=Content>" + str(row['Count']) + "</td>\
				  <td class=Content>" + str(row['TimeStamp']) + "</td>\
				  <td class=Content>" + str(row['SW Version']) + "</td>\
				  <td class=Content>" + str(row['Product ID']) + "</td>\
				  <td class=Content>" + str(row['7DayHistory']) + "</td>\
				  <td class=Content>" + str(row['30DayHistory']) + "</td>\
				  <td class=Content colspan=3 style='border-right:2px solid black'></td>\
				<\tr>"
			html=html+"<tr>\
			<td height=16 colspan=11 style='border-top:2px solid black' bgcolor='"+sendBGColor(severity)+"'></td>\
			</tr>"
			
			

	# print(html.encode('utf-8'))
	
	# html=html.encode('utf-8')
	# text_file = open("testHtml.html", "ab")
	# text_file.write(html)
	# text_file.close()



	

	# Create message container - the correct MIME type is multipart/alternative.
	msg = MIMEMultipart('alternative')
	msg['Subject'] = cpyName+" Syslog analysis for "+item+" data between " + reportingPeriod
	msg['From'] = me
	msg['To'] = to
	msg['CC'] = cc
	
	rcpt=cc.split(",") + to.split(",")

	part = MIMEBase('application', "octet-stream")
	# part.set_payload( open(folderName+"/"+file,"rb").read() )
	part.set_payload( open(os.path.join(folderName,file),"rb").read() )
	encoders.encode_base64(part)
	part.add_header('Content-Disposition', 'attachment; filename='+file)
	msg.attach(part)

	# Create the body of the message (a plain-text and an HTML version).
	# text = "Hi!\nHow are you?\nHere is the link you wanted:\nhttp://www.python.org"

	# Record the MIME types of both parts - text/plain and text/html.
	# part1 = MIMEText(text, 'plain')
	part1 = MIMEText(html, 'html')

	# Attach parts into message container.
	# According to RFC 2046, the last part of a multipart message, in this case
	# the HTML message, is best and preferred.
	msg.attach(part1)
	# msg.attach(part2)

	# Send the message via local SMTP server.
	s = smtplib.SMTP('outbound.cisco.com')
	# sendmail function takes 3 arguments: sender's address, recipient's address
	# and message to send - here it is sent as one string.
	s.sendmail(me, rcpt, msg.as_string())
	s.quit()