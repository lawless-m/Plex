import requests
from html.parser import HTMLParser
from collections import deque

import os
import sys

# set up sniffing
VerifySSL = True # don't proxy it
#VerifySSL = False # proxy it

if not VerifySSL:
	os.environ["HTTP_PROXY"] = "http://localhost:8080"
	os.environ["HTTPS_PROXY"] = "http://localhost:8080"

def Foggy(svInput):
	ivLength = len(svInput)
	if ivLength == 0 or ivLength > 158:
		svInput = svInput.replace('"',"&qt;")
		return svInput
	svOutput = ""
	for i in range(0,ivLength):
		ivRnd = 0
		if svInput[i] in (' ', '"', '>'):
			ivRnd = 1
		svOutput += chr(ivRnd + 97)
		svOutput += chr(ord(svInput[i]) + ivRnd)
	svOutput += chr(ivLength + 96)
	return svOutput

class WorkOrder:
	url = ""
	key = ""
	no = ""
	line = ""

	def __init__(self, url, key, no):
		self.url = url
		self.key = key
		self.no = no

class ParseWorkOrderList(HTMLParser):
	orders = deque()
	wo = None
	line = ""

	def handle_starttag(self, tag, attrs):
		def srch_a(key):
			for (k,v) in attrs :
				if k == key :
					return v
		if tag == "a":
			url = srch_a("href")
			if url[0:len("Work_Request_Form.asp?Do=Update&Work_Request_Key=")] == "Work_Request_Form.asp?Do=Update&Work_Request_Key=" :
				qs = url.split('&')
				self.wo = WorkOrder(url, qs[1].split('=')[1], qs[2].split('=')[1])
		if tag == "td":
			self.line = ""

	def handle_data(self, data) :
		self.line += data 

	def handle_endtag(self, tag) :
		if self.wo == None:
			return
		if tag == "td":
			if "\n" in self.line:
				self.line = ""
			else:
				self.wo.line = self.line
				self.orders.append(self.wo)
				self.wo = None

class ParseForm(HTMLParser):
	fields = {}
	chkcnt = {}
	select_name = ""
	in_field = False
	
	def handle_starttag(self, tag, attrs):
		def srch_a(key):
			for (k,v) in attrs :
				if k == key :
					if v == None:
						return True
					return v
					
		if tag == "input" :			
			n = srch_a("name")
			if srch_a("type") == "checkbox" :
				k = self.chkcnt.get(n, 1)
				self.chkcnt[n] = k+1
				n += "_%d" % k
				self.fields[n] = (srch_a("value"), srch_a("retval"))
			else:	
				self.fields[n] = srch_a("value")
			return
			
		if tag == "select" :
			self.fieldname = srch_a("name")
			return
			
		if tag == "option" :
			if srch_a("selected"):
				self.in_field = True
				self.fields[self.fieldname] = (srch_a("value"), "")
			return
			
		if tag == "textarea" :
			self.fieldname = srch_a("name")
			self.in_field = True
			self.fields[self.fieldname] = ""
			return
	
	def handle_data(self, data):
		if self.in_field:
			if type(self.fields[self.fieldname]) == type((None,None)) :
				self.fields[self.fieldname] = (self.fields[self.fieldname][0], self.fields[self.fieldname][1] + data)
			else:
				self.fields[self.fieldname] += data 
			return
			
	def handle_entityref(self, name): 
		if self.in_field :
			data = self.unescape('&amp;')
			self.fields[self.fieldname] += data
			self.fields[self.fieldname] += name
				
	def handle_endtag(self, tag):
		if tag == "select":
			self.fieldname = ""
			return
			
		if tag == "option":
			self.in_field = False
			return
		
		if tag == "textarea":
			self.in_field = False
			self.fieldname = ""
			return
			

class PyPlex:
	last_url = ""
	last_r = None
	viewstate = ""
	key = ""
	
	def __init__(self, un, pw, cc):
		self.session = requests.Session()
		self.get("/Modules/SystemAdministration/Login/Index.aspx")
		self.post("/Modules/SystemAdministration/Login/Index.aspx", {"__VIEWSTATE":self.viewstate,"txtUserID":un, "txtPassword":Foggy(pw), "txtCompanyCode":cc, "hdnUseSslAfterLogin":1})
		self.get("/Modules/SystemAdministration/Login/Login.aspx")
		self.post("/Modules/SystemAdministration/Login/Login.aspx", {"__VIEWSTATE":self.viewstate,"browserMinorVersion":"undefined", "screenHeight":"900", "screenWidth":"1600", "screenDepth":"24", "browserName":"Netscape", "browserVersion":"5.0 (Windows)", "platform":"Win64"})
		self.key = self.session.cookies["Session_Key"][1:-1].lower()
		
	def static(self, url, params={}, headers={}):
		if not "Referer" in headers:
			headers["Referer"] = self.last_url
		
		self.last_r = self.session.get("https://static.plexonline.com" + url, verify=VerifySSL, headers=headers, params=params)
		
	def script(self, url, params={}, headers={}):
		if not "Referer" in headers:
			headers["Referer"] = self.last_url
			
		self.last_r = self.session.get("https://www.plexonline.com" + url, verify=VerifySSL, headers=headers, params=params)
		self.last_url = url

	def get(self, url, params={}, headers={}):
	
		if not "Referer" in headers:
			headers["Referer"] = self.last_url
			
		self.last_r = self.session.get("https://www.plexonline.com" + url, verify=VerifySSL, headers=headers, params=params)
		self.last_url = url
		vs = self.find_viewstate()
		if vs != "":
			self.viewstate = vs

	def post(self, url, data, headers={}):
	
		if not "Referer" in headers:
			headers["Referer"] = self.last_url
			
		self.last_r = self.session.post("https://www.plexonline.com" + url, data, verify=VerifySSL, headers=headers)
		self.last_url = url
		vs = self.find_viewstate()
		if vs != "":
			self.viewstate = vs

	def find_viewstate(self):
		pl = ParseForm()
		pl.feed(self.last_r.text)
		return pl.fields.get("__VIEWSTATE", "")

	def work_request_list(self, params):
		self.get("/" + self.key + "/Equipment/Work_Request.asp")
		ps = {"hdnApplication_Filter_Default_Control_Application_Key":"", "hdnApplication_Filter_Default_Control_No_Delete":"", "hdnApplication_Filter_Default_Control_Allow_Empty_Default":"", "flttxtEquipment":"", "fltEquipment":"", "fltDescription":"", "flttxtWork_Request_Status":"", "fltWork_Request_Status":"", "flttxtEquipment_Type":"", "fltEquipment_Type":"", "flttxtAssigned_To":"", "fltAssigned_To":"", "flttxtRequested_By":"", "fltRequested_By":"", "fltPriority":"", "flttxtType":"", "fltType":"", "fltWR_Sort_Order":"", "fltWR_Sort_Ordering":"", "fltActive":"", "fltRequest_No":"", "fltBegin_Date_DTE":"", "fltEnd_Date_DTE":"", "fltDue_Date1_DTE":"", "fltDue_Date2_DTE":"", "flttxtBuildings":"", "fltBuildings":"", "fltEquipment_Group":"", "fltInclude_PM":"", "flttxtLocation":"", "fltLocation":"", "flttxtParts":"", "fltParts":"", "fltAsset_No":"", "flttxtAssignToUsersList":"", "fltAssignToUsersList":"", "hdnRecordsExist":""}
		for p in params:
			ps[p] = params[p]
		self.post("/" + self.key + "/Equipment/Work_Request.asp", ps)
		wos = ParseWorkOrderList()
		wos.feed(self.last_r.text)
		self.record_html("work_request_list")
		return wos.orders, self.last_r.text
	
	def work_request_csv(self, params):	
		self.get("/" + self.key + "/Equipment/Work_Request.asp")
		ps = {"hdnApplication_Filter_Default_Control_Application_Key":"", "hdnApplication_Filter_Default_Control_No_Delete":"", "hdnApplication_Filter_Default_Control_Allow_Empty_Default":"", "flttxtEquipment":"", "fltEquipment":"", "fltDescription":"", "flttxtWork_Request_Status":"", "fltWork_Request_Status":"", "flttxtEquipment_Type":"", "fltEquipment_Type":"", "flttxtAssigned_To":"", "fltAssigned_To":"", "flttxtRequested_By":"", "fltRequested_By":"", "fltPriority":"", "flttxtType":"", "fltType":"", "fltWR_Sort_Order":"", "fltWR_Sort_Ordering":"", "fltActive":"", "fltRequest_No":"", "fltBegin_Date_DTE":"", "fltEnd_Date_DTE":"", "fltDue_Date1_DTE":"", "fltDue_Date2_DTE":"", "flttxtBuildings":"", "fltBuildings":"", "fltEquipment_Group":"", "fltInclude_PM":"", "flttxtLocation":"", "fltLocation":"", "flttxtParts":"", "fltParts":"", "fltAsset_No":"", "flttxtAssignToUsersList":"", "fltAssignToUsersList":"", "hdnRecordsExist":""}
	
		for p in params:
			ps[p] = params[p]
		ps["hdnRecordsExist"] = "0"
		self.post("/" + self.key + "/Equipment/Work_Request.asp", ps)
		ps["hdnWork_Request_ID"] = "0"
		ps["hdnAssigned_To_PUN"] = "0"
		ps["txthdnAssigned_To_PUN"] = ""
		ps["hdnRequested_By_PUN"] = "0"
		ps["txthdnRequested_By_PUN"] = ""
		ps["hdnDue_Date_PUN_DTE"] = ""
		ps["hdnStatus_PUN"] = "0"
		ps["txthdnStatus_PUN"] = ""	
		ps["hdnRecordsExist"] = "1"
	
		self.post("/" + self.key + "/Download_Apps/CSV/Build_CSV3.asp", ps)
		return self.last_r.text

	def work_request(self, wourl):
		list_ref = "/" + self.key + "/Equipment/Work_Request.asp"
		self.get("/" + self.key + "/Equipment/" + wourl)
		self.last_url = list_ref
		f = ParseForm()
		f.feed(self.last_r.text)
		self.record_html("work_request")
		return f

	def userlist(self):
		params = {"Plexus_Customer_No":"60471", "Picker_Called_By":"https://www.plexonline.com/" + self.key + "/Equipment/Work_Request.asp","SQL":"dbo.Users_By_Role_Get_Picker 60471,'Maintenance User','',1,1", "DB":"Plexus_Control", "Fields":"Full_Name", "Title":"","ReturnField":"Full_Name", "hdnReturnField":"Plexus_User_No", "BlankOption":"true", "AutoSelect":"true", "ImageFolder":"", "ImageField":"", "Touchscreen":"false", "MultiSelect":"true", "StripHtml":"false", "CurrentValue":"", "DisplayCheckmark":"false", "DateFields":"", "ConnectPCN":"60471", "Session_Key":"{" + self.key + "}", "Dialog_ID":"0"}
		self.get("/" + self.key + "/Pickers/Modal_Super_Picker.asp", params=params)
		f = ParseForm()
		f.feed(self.last_r.text)
		return {f.fields[k][0] : f.fields[k][1] for k in f.fields if k != None and k[0:10] == "chkOption_"}

	def equiplist(self):
		params = {"Plexus_Customer_No":"60471","Picker_Called_By":"https://www.plexonline.com/" + self.key + "/Equipment/Equipment.asp","SQL":"EXEC Maintenance.dbo.Equipments_Get_All 60471,''","DB":"Maintenance","Fields":"Equipment_ID","Title":"","ReturnField":"ID_Only","hdnReturnField":"Equipment_Key","BlankOption":"true","AutoSelect":"true","ImageFolder":"","ImageField":"","Touchscreen":"false","MultiSelect":"true","StripHtml":"false","CurrentValue":"","DisplayCheckmark":"false","DateFields":"","ConnectPCN":"60471","Session_Key":"{" + self.key + "}","Dialog_ID":"0"}
		self.get("/" + self.key + "/Pickers/Modal_Super_Picker.asp", params=params)
		#return self.last_r.text
		f = ParseForm()
		f.feed(self.last_r.text)
		return {f.fields[k][0] : f.fields[k][1] for k in f.fields if k != None and k[0:10] == "chkOption_"}

	def record_html(self, fn):
		fid = open("C:\\Users\\C18610\\_cache\\%s.html" % fn, "w+")
		fid.write(self.last_r.text)
		fid.close()
		
	
	def pm_list(self):
		self.post("/" + self.key + "/Equipment/Maintenance.asp?ssAction=Back", {})
		return self.last_r.text
		
	def pm_maint_frm(self, url):
		self.get("/" + self.key + "/Equipment/" + url)
		return self.last_r.text
		
	def pm_report(self):
			
		self.get("/" + self.key + "/Rendering_Engine/default.aspx?Request=Show&RequestData=SourceType(Screen)SourceKey(10608)", headers={"Referer":"https://www.plexonline.com/" + self.key + "/Report_System/Report_List.asp"})
		
		headers = {"Pragma":"no-cache", "Content-Type":"application/x-www-form-urlencoded", "RenderType":"Partial", "PanelTargets":"FILTER_PANEL_2_28:True|GRID_PANEL_3_28:False|",  "Referer": "https://www.plexonline.com/" + self.key + "/Rendering_Engine/default.aspx?Request=Show&RequestData=SourceType(Screen)SourceKey(10608)"}

		self.post("/" + self.key + "/Rendering_Engine/default.aspx?Request=Show&RequestData=SourceType(Screen)SourceKey(10608)", "undefined__EVENTTARGET=&__EVENTARGUMENT=&__LASTFOCUS=&__VIEWSTATE=%2F" + self.viewstate[1:-1] + "%3D&hdnScreenTitle=PM%20Requirement%20Report&hdnFilterElementsKeyHandle=172744%2F%5CResponsibility%5B%5D172761%2F%5CEquipment_ID%5B%5D172762%2F%5CBuilding_Keys_new%5B%5D172765%2F%5CChecklist_Key_new%5B%5D534809%2F%5CActive&=&ScreenParameters=&RequestKey=0&Layout1$el_172761=&Layout1$el_172761_hf=&Layout1$el_172761_hf_last_valid=&Layout1$el_172765=&Layout1$el_172744=&Layout1$el_172744_hf=&Layout1$el_172744_hf_last_valid=&Layout1$el_534809=1&Layout1$el_172762=&Layout1$el_172762_hf=&Layout1$el_172762_hf_last_valid=&panel_row_count_3=0&", headers=headers)
		
		return self.last_r.text
