import requests
import urllib
from xml.dom import minidom
import json
import getpass
import sys
import re
import time


cpyKey=sys.argv[1]
useDefault=sys.argv[2]
server=sys.argv[3]

if('Yes' in useDefault):
	savedSearchName="syslog_summary_with_default"
	searchQuery='index='+cpyKey+'-syslog sourcetype = "np_syslog_v1" source="*.gz" \
| eval MsgType_Desc=MsgType.MsgDesc \
| lookup defaultCatalog RegEx as MsgType_Desc OUTPUT Description Action,Sev,Name,RegEx,ThresholdPerMin,ThresholdPerDay,Filter \
| rename Description as defaultDescription Action as defaultAction Sev as defaultSev, Name as defaultName, ThresholdPerMin as defaultThresholdPerMin,ThresholdPerDay as defaultThresholdPerDay,Filter as defaultFilter RegEx as defaultRegEx \
| lookup kvstore_company_catalog RegEx as MsgType_Desc cpyKey as Company OUTPUT Description Action,Sev,Name,RegEx,ThresholdPerMin,ThresholdPerDay,Filter \
| rename Description as customerDescription Action as customerAction Sev as customerSev, Name as customerName, ThresholdPerMin as customerThresholdPerMin,ThresholdPerDay as customerThresholdPerDay,Filter as customerFilter RegEx as customerRegEx \
| eval RegEx=coalesce(customerRegEx,defaultRegEx),Description=coalesce(customerDescription,defaultDescription),Action=coalesce(customerAction,defaultAction),Sev=coalesce(customerSev,defaultSev),Name=coalesce(customerName,defaultName),ThresholdPerMin=coalesce(customerThresholdPerMin,defaultThresholdPerMin),ThresholdPerDay=coalesce(customerThresholdPerDay,defaultThresholdPerDay),Filter=coalesce(customerFilter,defaultFilter) \
| fillnull value="no match" RegEx \
| eval Sev=coalesce(Sev,Severity) \
| fillnull value="N" Filter \
| fillnull value=0 ThresholdPerDay ThresholdPerMin \
| rex field=source ".*/(?<collector>.*)/.*" \
| bin _time span=1min \
|rex field=source ".*\/(?<lastTime>.*).gz" \
|eval lastTime=strptime(lastTime,"%Y%m%d%H%M%S")\
| stats max(lastTime) as lastTime count as CountPerMin, first(MsgDesc) AS MsgDesc ,first(TimeStamp) AS TimeStamp,first(Description) as Description,first(Action) as Action,first(Sev) as Sev,first(Name) as Name first(ThresholdPerMin) as ThresholdPerMin first(ThresholdPerDay) as ThresholdPerDay first(Filter) as Filter first(collector) as collector by Device MsgType RegEx _time Company\
| eval Filter=if( ((CountPerMin>=ThresholdPerMin) AND (Filter=="N")),"N","Y") \
| bin _time span=1day \
| stats max(lastTime) as lastTime sum(CountPerMin) as Count, first(MsgDesc) AS MsgDesc ,first(TimeStamp) AS TimeStamp,first(Description) as Description,first(Action) as Action,first(Sev) as Sev,first(Name) as Name first(ThresholdPerMin) as ThresholdPerMin first(ThresholdPerDay) as ThresholdPerDay first(Filter) as Filter first(collector) as Collector by Device MsgType RegEx _time Company\
| eval Filter=if(((Count>=ThresholdPerDay) AND (Filter=="N")),"N","Y") \
|collect index=syslog-summary-'+server+'\
|stats max(lastTime) as lastTime by Company\
|outputlookup kvstore_last_summarized_syslog'

else:
	savedSearchName="syslog_summary_without_default"
	searchQuery='index='+cpyKey+'-syslog sourcetype = "np_syslog_v1" source="*.gz" \
| eval MsgType_Desc=MsgType.MsgDesc \
| lookup kvstore_company_catalog RegEx as MsgType_Desc cpyKey as Company OUTPUT Description Action,Sev,Name,RegEx,ThresholdPerMin,ThresholdPerDay,Filter \
| fillnull value="no match" RegEx \
| eval Sev=coalesce(Sev,Severity) \
| fillnull value="N" Filter \
| fillnull value=0 ThresholdPerDay ThresholdPerMin \
| rex field=source ".*/(?<collector>.*)/.*" \
| bin _time span=1min \
|rex field=source ".*\/(?<lastTime>.*).gz" \
|eval lastTime=strptime(lastTime,"%Y%m%d%H%M%S")\
| stats max(lastTime) as lastTime count as CountPerMin, first(MsgDesc) AS MsgDesc ,first(TimeStamp) AS TimeStamp,first(Description) as Description,first(Action) as Action,first(Sev) as Sev,first(Name) as Name first(ThresholdPerMin) as ThresholdPerMin first(ThresholdPerDay) as ThresholdPerDay first(Filter) as Filter first(collector) as collector by Device MsgType RegEx _time Company\
| eval Filter=if( ((CountPerMin>=ThresholdPerMin) AND (Filter=="N")),"N","Y") \
| bin _time span=1day \
| stats max(lastTime) as lastTime sum(CountPerMin) as Count, first(MsgDesc) AS MsgDesc ,first(TimeStamp) AS TimeStamp,first(Description) as Description,first(Action) as Action,first(Sev) as Sev,first(Name) as Name first(ThresholdPerMin) as ThresholdPerMin first(ThresholdPerDay) as ThresholdPerDay first(Filter) as Filter first(collector) as Collector by Device MsgType RegEx _time Company\
| eval Filter=if(((Count>=ThresholdPerDay) AND (Filter=="N")),"N","Y") \
|collect index=syslog-summary-'+server+'\
|stats max(lastTime) as lastTime by Company\
|outputlookup kvstore_last_summarized_syslog'

baseurl="https://as-practice-"+server+".cisco.com:8089"
userName = "admin"
password = "Engi924+"

#get Saved Search Details first

#Step 1: Get a session key
response = requests.post(baseurl+'/services/auth/login',data=({'username':userName, 'password':password}),verify=False)
sessionKey = minidom.parseString(response.text).getElementsByTagName('sessionKey')[0].childNodes[0].nodeValue

print("Logged in. Searching")

#Step 2: Get Saved Search Details
searchjob=requests.get(baseurl + '/servicesNS/admin/as_data/saved/searches/'+savedSearchName,headers={'Authorization': 'Splunk %s' % sessionKey},verify=False)

#If Saved search is present
if searchjob.status_code != 404:
	print("Adding to current Saved Search")
	reg = re.compile(r'<s:key name="search">(?P<searchQuery>.*?)</s:key>', re.IGNORECASE)
	savedSearchQuery = reg.findall(searchjob.text)[0]
	searchQueryArray=savedSearchQuery.split('|')
	
	if "CDATA" in searchQueryArray[0]:
		searchQueryArray[0]=searchQueryArray[0].replace("<![CDATA[","")
		searchQueryArray[len(searchQueryArray)-1]=searchQueryArray[len(searchQueryArray)-1][:-3] #This is done to remove CDATA elements
	
	#Ensuring index is not added twice
	if not ("index="+cpyKey+"-syslog" in searchQueryArray[0]):
		searchQueryArray[0]='index='+cpyKey+'-syslog OR ' + searchQueryArray[0]
	
	
	newQuery='|'.join(searchQueryArray)
	#Step 1: Get a session key
	response = requests.post(baseurl+'/services/auth/login',data=({'username':userName, 'password':password}),verify=False)
	sessionKey = minidom.parseString(response.text).getElementsByTagName('sessionKey')[0].childNodes[0].nodeValue

	print("Logged in. Searching")

	#Step 2: Create a search job
	searchjob=requests.post(baseurl + '/servicesNS/admin/as_data/saved/searches/'+savedSearchName,headers={'Authorization': 'Splunk %s' % sessionKey},data=({'search':newQuery}),verify=False)
	
#If Saved Search is not present create a saved search
else:
	print("Creating Saved Search")

	#Step 1: Get a session key
	response = requests.post(baseurl+'/services/auth/login',data=({'username':userName, 'password':password}),verify=False)
	sessionKey = minidom.parseString(response.text).getElementsByTagName('sessionKey')[0].childNodes[0].nodeValue

	#Step 2: Create a search job
	searchjob=requests.post(baseurl + '/servicesNS/admin/as_data/saved/searches',headers={'Authorization': 'Splunk %s' % sessionKey},data=({'search': searchQuery,'actions':'email','action.email.to':'mamohta@cisco.com','description':'Populating Summary Syslogs','is_scheduled':'1','cron_schedule':'0 0 * * *','name':savedSearchName,'dispatch.earliest_time':'-2d','dispatch.latest_time':'now'}),verify=False)
	

#Populate last 30 days data into index
print("Populating Last 30 Days Data now")

#Step 1: Get a session key
response = requests.post(baseurl+'/services/auth/login',data=({'username':userName, 'password':password}),verify=False)
sessionKey = minidom.parseString(response.text).getElementsByTagName('sessionKey')[0].childNodes[0].nodeValue

#Step 2: Create a search job
searchjob=requests.post(baseurl + '/services/search/jobs',headers={'Authorization': 'Splunk %s' % sessionKey},data=({'search': 'search '+searchQuery,'earliest_time':'-45d','latest_time':'now'}),verify=False)

print(searchjob.text)

reg = re.compile(r'<sid>(?P<sid>.*?)<\/sid>', re.IGNORECASE)

time.sleep(1)
try:
	sid = reg.findall(searchjob.text)[0]
except IndexError:
	print("no results found")

# print(searchjob.text)
sid = minidom.parseString(searchjob.text).getElementsByTagName('sid')[0].childNodes[0].nodeValue

#Background the search
requests.post(baseurl + '/services/search/jobs/%s/control' % sid,auth=(userName,password),data={'action':'save'},verify=False)




	
	





