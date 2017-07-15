import requests
import urllib
from xml.dom import minidom
import re
import time
import getpass
from datetime import datetime, timedelta
import pandas as pd
import os
import json
import xlsxwriter
import argparse


parser = argparse.ArgumentParser()

parser.add_argument('--earliest', action='store', dest='earliest',help='Give Earliest Date before today',type=int)
parser.add_argument('--latest', action='store', dest='latest',help='Give Latest Date before today',type=int)

results = parser.parse_args()
earliest=results.earliest
latest=results.latest



userName = "mamohta"
password = "Vatsal123456!"




def getSearchResults(searchQuery,baseurl):
	# Step 1: Get a session key
	response = requests.post(baseurl+'/services/auth/login',data=({'username':userName, 'password':password}),verify=False)
	sessionKey = minidom.parseString(response.text).getElementsByTagName('sessionKey')[0].childNodes[0].nodeValue

	# Step 2: Create a search job
	searchjob=requests.post(baseurl + '/services/search/jobs',headers={'Authorization': 'Splunk %s' % sessionKey},data=({'search': searchQuery}),verify=False)

	reg = re.compile(r'<sid>(?P<sid>.*?)<\/sid>', re.IGNORECASE)

	time.sleep(1)
	
	print(searchjob.text)
	sid = reg.findall(searchjob.text)[0]

	sid = minidom.parseString(searchjob.text).getElementsByTagName('sid')[0].childNodes[0].nodeValue

	
	# Step 3: Get the search status
	# myhttp.add_credentials(username, password)
	servicessearchstatusstr = '/services/search/jobs/%s/' % sid
	isnotdone = True
	while isnotdone:
		searchstatus = requests.post(baseurl + servicessearchstatusstr,auth=(userName,password),verify=False)
		isdonestatus = re.compile('isDone">(0|1)')
		isdonestatus = isdonestatus.search(searchstatus.text).groups()[0]
		if (isdonestatus == '1'):
			isnotdone = False
	print(isdonestatus)
	
	return sid
	
#Get today's date details
currentDT = datetime.now()
day=currentDT.strftime("%a")


cpyName='MS Azure'
cc=" "	
to="mamohta@cisco.com"
createReportBy="Product Family"
customOptions="NA"
filterByOptions="NA"
cpyKey="21101"
frequency="Daily"
useDefault="No"
server="02"


info_min_time=datetime.strftime(datetime.now() - timedelta(earliest), '%m/%d/%Y')+":00:00:00" 
info_max_time=datetime.strftime(datetime.now() - timedelta(latest), '%m/%d/%Y')+":23:59:59"
info_min_time_thirty=datetime.strftime(datetime.now() - timedelta(31+earliest), '%m/%d/%Y')+":00:00:00"
info_max_time_thirty=datetime.strftime(datetime.now() - timedelta(earliest+1), '%m/%d/%Y')+":23:59:00"



# Calculating 30 Day Query on the basis of  Default Catalog needs to be used or not

thirtyDaySearchQuery='search index=syslog-summary-* RegEx!="no match" Company='+cpyKey+' earliest="'+info_min_time_thirty+'" latest="'+info_max_time_thirty+'" \
| append \
    [ search index=syslog-summary-* RegEx="no match" Company='+cpyKey+' earliest="'+info_min_time_thirty+'" latest="'+info_max_time_thirty+'" \
    | eval MsgType_Desc=MsgType.MsgDesc \
    | lookup kvstore_company_catalog RegEx as MsgType_Desc cpyKey as Company OUTPUT Description Action,Sev,Name,RegEx,ThresholdPerMin,ThresholdPerDay,Filter \
    | fillnull value="no match" RegEx \
    | eval Sev=coalesce(Sev,Severity) \
    | fillnull value="N" Filter \
    | fillnull value=0 ThresholdPerDay ThresholdPerMin] \
| addinfo \
| eval age=(info_max_time-_time)/86400 \
| eval 30DayHistory=if(age<30 AND age>0,1,0) \
| eval 7DayHistory=if(age<7 AND age>0,1,0)\
|stats sum(7DayHistory) as 7DayHistory sum(30DayHistory) as 30DayHistory values(Company) as cpyKey by Device,MsgType,RegEx\
|outputlookup kvstore_syslog_temp'

baseurl="https://as-practice-"+server+".cisco.com:8089"


sid=getSearchResults(thirtyDaySearchQuery,baseurl)

print('Added to kvstore')



searchQuery='search (index=syslog-summary-* RegEx!="no match" Company='+cpyKey+') OR (index='+cpyKey+'-np sourcetype=devices) earliest="'+info_min_time+'" latest="'+info_max_time+'" \
| eval DEVICENAME=coalesce(Device,deviceName) \
|eval DEVICENAME=lower(DEVICENAME)\
| eventstats values(productFamily) as productFamily,values(swVersion) as swVersion,values(productId) as productId values(deviceId) as deviceId by DEVICENAME \
| search MsgType=* \
| rename DEVICENAME as Device \
| append \
[ search (index=syslog-summary-* RegEx="no match" Company='+cpyKey+') OR (index='+cpyKey+'-np sourcetype=devices) earliest="'+info_min_time+'" latest="'+info_max_time+'" \
| eval DEVICENAME=coalesce(Device,deviceName) \
|eval DEVICENAME=lower(DEVICENAME)\
| eventstats values(productFamily) as productFamily,values(swVersion) as swVersion,values(productId) as productId values(deviceId) as deviceId by DEVICENAME \
| search MsgType=* \
| rename DEVICENAME as Device \
| eval MsgType_Desc=MsgType.MsgDesc \
| lookup kvstore_company_catalog RegEx as MsgType_Desc cpyKey as Company OUTPUT Description Action,Sev,Name,RegEx,ThresholdPerMin,ThresholdPerDay,Filter \
| fillnull value="no match" RegEx \
| fillnull value="undefined" Sev \
| eval Name=coalesce(Name,MsgType) \
| fillnull value="N" Filter \
| fillnull value=0 ThresholdPerDay ThresholdPerMin] \
| stats sum(Count) as Count, first(MsgDesc) AS MsgDesc ,first(TimeStamp) AS TimeStamp,first(Description) as Description,first(Action) as Action,first(Sev) as Sev,first(Name) as Name first(ThresholdPerMin) as ThresholdPerMin first(ThresholdPerDay) as ThresholdPerDay first(Filter) as Filter first(collector) as Collector first(productFamily) as productFamily first(productId) as productId first(swVersion) as swVersion first(deviceId) as deviceId by Device MsgType RegEx Company\
| fillnull value="NA" swVersion,productId,productFamily,deviceId\
| lookup kvstore_syslog_temp Device,MsgType,RegEx OUTPUT 30DayHistory,7DayHistory \
| fillnull value=0 30DayHistory,7DayHistory\
| eval Device=upper(Device)\
| table Company,Name,productFamily,Device,swVersion,productId,MsgType,MsgDesc,Description,Action,Sev,Count,7DayHistory,30DayHistory,TimeStamp,RegEx,ThresholdPerMin,ThresholdPerDay,Filter\
| rename MsgType as Message MsgDesc as "Message Content" Action as "Recommended Action" productFamily as "Product Family" swVersion as "SW Version" productId as "Product ID" \
| eval "Product Family" = case(match(Device,"IBR"), "IBR",match(Device,"GMR"), "ASR1k",match(Device,"XCG"), "N3k9k-XCG", 1=1, "Azure_Phynet") \
| eval createReportBy = case(match(Device,"IBR"), "IBR",match(Device,"GMR"), "ASR1k",match(Device,"XCG"), "N3k9k-XCG", 1=1, "Azure_Phynet")'


print('Search Query Created')
print('Searching')
sid=getSearchResults(searchQuery,baseurl)

# Step 4: Get the search results
servicessearchstatusstr = '/services/search/jobs/%s/' % sid
i=1

services_search_results_str = '/services/search/jobs/'+str(sid)+'/results?output_mode=json&count=50000&offset=0'
searchresults = requests.get(baseurl + services_search_results_str,auth=(userName,password),verify=False)

jsonResult=[]
jsonResultTemp = json.loads(searchresults.text)['results']
jsonResult = jsonResult + jsonResultTemp

while(len(jsonResultTemp)==50000):
	jsonResultTemp=[]
	offset=50000*i
	services_search_results_str = '/services/search/jobs/'+str(sid)+'/results?output_mode=json&count=50000&offset='+str(offset)
	searchresults = requests.get(baseurl + services_search_results_str,auth=(userName,password),verify=False)
	jsonResultTemp = json.loads(searchresults.text)['results']
	jsonResult = jsonResult + jsonResultTemp
	# print(len(jsonResult))
	i=i+1

data_df=pd.DataFrame(jsonResult)

todayDate=currentDT.strftime("%m-%d")
earliest=info_min_time[:-14]
latest=info_max_time[:-14]
earliest=earliest.replace('/','-')
latest=latest.replace('/','-')

folderName = os.path.join(cpyName,datetime.strftime(datetime.now(),"%m-%d"),earliest+' to '+latest)
if not os.path.exists(folderName):
	os.makedirs(folderName)

writer = pd.ExcelWriter(os.path.join(folderName,cpyName+'_Raw_'+earliest+'_to_'+latest+'.xlsx'), engine='xlsxwriter')
workbook  = writer.book

data_df.to_excel(writer,sheet_name = "Sheet1",index = False,columns = ['Product Family','Device','SW Version','Product ID','Name','Message','Message Content','Description','Recommended Action','Sev','Count','7DayHistory','30DayHistory','TimeStamp','RegEx','Filter','createReportBy']) 
worksheet = writer.sheets["Sheet1"]

workbook.close()
	

# #Background the search to save it for 7 days
# requests.post(baseurl + '/services/search/jobs/%s/control' % sid,auth=(userName,password),data={'action':'save'},verify=False)

os.system('python -W ignore createReport.py "'+cpyName+'" '+earliest+' '+latest+' '+todayDate)

# os.system("python sendMails.py "+to+" "+cc+" "+folderName+" "+earliest+" "+latest+" '"+cpyName+"'")

		



	




