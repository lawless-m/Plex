module Plexus

if isfile(raw"c:\users\matt\ntuser.ini")
	lib = raw"C:\Users\matt\Documents\power\Lib"
	dbdir = raw"C:\Users\matt\Documents\power\SQLite.Data"
	outdir = raw"C:\Users\matt\Documents\power"
	
else
	lib = raw"Z:\Maintenance\Matt-Heath\Lib"
	dbdir = raw"Z:\Maintenance\Matt-Heath\power\SQLite.Data"
	outdir = raw"Z:\Maintenance\PPM"
end

push!(LOAD_PATH, lib)


push!(LOAD_PATH, raw"Z:\Maintenance\Matt-Heath\Lib")
using HTMLParser
using Cache

export Plex, WorkRequest, work_requests, work_request, pm_list, work_requests_pms, PMFilter, pm_report

using PyCall
unshift!(PyVector(pyimport("sys")["path"]), lib)

@pyimport PyPlex

@enum PMFilter WR_pms WR_not_pms WR_all

function txt2time(d, t)
	hm = DateTime(t, "HH:MM")
	DateTime(d, "dd/mm/yy") + Base.Dates.Year(2000) + Base.Dates.Hour(hm) + Base.Dates.Minute(hm)
end

struct Plex
	py::PyObject
	values::Dict{String,Dict{Any, Any}}
	function Plex(un, pw, branch)
		p = PyPlex.PyPlex(un, pw, branch)
		new(p, Dict{String, Dict{Any, Any}}())
	end
end

function users!(p::Plex)
	p.values["Users"] = p.py[:userlist]()
end

function equipment!(p::Plex)
	
	return p.py[:equiplist]()
	
	function proc(t::StartTag)
		attr(k) = get(t.attrs, k, "")
		vtable = Dict(
			"a"=>()->begin end,
			"td"=>()->begin end
			)
		get(vtable, t.name, ()->nothing)();
	end
	
	proc(t::Data) = begin end

	function proc(t::EndTag)
	end
	
	proc(b::Block) = nothing
	
	k = 1
	for blk in HTMLParser.HTML(cache("equiplist", p.py[:equiplist])).blks
		println("K $k ", blk)
		k += 1
	end

	#p.values["Equipment"] = p.py[:equiplist]()
end

struct WorkOrderListing
	url::String
	key::String
	no::String
	line::String
end

function work_order_list(t)
	url = ""
	key = ""
	no = ""
	line = ""
	
	wols = Vector{WorkOrderListing}()
	
	function proc(t::StartTag)
		attr(k) = get(t.attrs, k, "")
		vtable = Dict(
			"a"=>()->begin
					if startswith(attr("href"), "Work_Request_Form.asp?Do=Update&Work_Request_Key=")
						url = attr("href")
						qs = split(url, '&')
						key = split(qs[2], '=')[2]
						no = split(qs[3], '=')[2]
					end
				end, 
			"td"=>()->
				line = ""
			)
		get(vtable, t.name, ()->nothing)();
	end
	
	proc(t::Data) = line = "$line$(t.data)";

	function proc(t::EndTag)
		if length(url) > 0 && t.name == "td"
			if '\n' in line
				line = ""
			else
				push!(wols, WorkOrderListing(url, key, no, line))
				url = ""
			end
		end
	end
	
	proc(b::Block) = nothing
	
	foreach(proc, HTMLParser.HTML(t).blks)
	filter((w)->w.no!="", wols)
end

struct WorkRequest
	no::String
	line::String
	equipment::Tuple{String, String}
	assignedto::Tuple{String, String}
	assignedby::Tuple{String, String}
	descr::String
	note::String
	started::DateTime
	completed::DateTime
	py::PyObject

	function WorkRequest(no::String, line::String, wr::PyObject, users, equip)
		new(no, line, (wr[:fields]["lstEquipment_RQD"], equip[wr[:fields]["lstEquipment_RQD"]]), wr[:fields]["lstAssigned_To"], (wr[:fields]["fltRequested_By"], users[wr[:fields]["fltRequested_By"]]), wr[:fields]["txtDescription_TXA_1000"], wr[:fields]["txtNote_TXA_3000"], txt2time(wr[:fields]["txtRequest_Date_DTE"], wr[:fields]["txtRequest_Date_TME"]), txt2time(wr[:fields]["txtComplete_Date_DTE"], wr[:fields]["txtComplete_Date_TME"]), wr)
	end
end

function show(io::IO, w::WorkRequest)
	println(io, w.no)
	println(io, "Assigned to: ", w.assignedto[2], " (", w.assignedto[1], ")")
	println(io, "Assigned by: ", w.assignedby[2], " (", w.assignedby[1], ")")
	println(io, "Descr: ", w.descr)
	println(io, "Note: ", w.note)
	println(io, "Started: ", w.started)
	println(io, "Completed: ", w.completed)
	println(io, "TTR: ", Base.Dates.Minute(w.completed - w.started))
end

function work_requests(p::Plex, month::DateTime, by::Tuple{String, String})
	work_requests(p, Base.Dates.firstdayofmonth(month),  Base.Dates.lastdayofmonth(month), by::Tuple{String, String})
end

function work_requests(p::Plex, dstart::DateTime, dend::DateTime, by::Tuple{String, String})

	wos, html = p.py[:work_request_list](Dict("fltBegin_Date_DTE"=>Base.Dates.format(dstart, "d/m/yy"), "fltEnd_Date_DTE"=>Base.Dates.format(dend, "d/m/yy"), "hdnApplication_Filter_Default_Control_Application_Key"=>"1406", "hdnApplication_Filter_Default_Control_No_Delete"=>"0", "hdnApplication_Filter_Default_Control_Allow_Empty_Default"=>"1", "flttxtRequested_By"=>by[2], "fltRequested_By"=>by[1], "fltWR_Sort_Order"=>"Work_Request_No,Due_Date", "fltWR_Sort_Ordering"=>"ASC", "fltActive"=>"1", "hdnRecordsExist"=>"0"))
	
	users = get!(p.values, "Users", users!(p))
	equip = get!(p.values, "Equipment", equipment!(p))

	foreach(println, wos)
	
	h_wrl = work_order_list(html)
	
	foreach(println, h_wrl)

	h_wos = Dict(wo.no=>WorkRequest(wo.no, wo.line, p.py[:work_request](wo.url), users, equip) for wo in h_wrl)
	
	println(h_wos)

	p_wos = Dict(wo[:no]=>WorkRequest(wo[:no], wo[:line],  p.py[:work_request](wo[:url]), users, equip) for wo in wos)
	
	println(p_wos)
	
	p_wos
end


function work_requests_pms(p::Plex, dstart::Date, dend::Date, complete::Bool; filter::PMFilter=WR_all, line="")
	params = Dict(
	"hdnApplication_Filter_Default_Control_Application_Key"=>"1406", 
	"fltWR_Sort_Order"=>"Work_Request_No,Due_Date", 
	"fltWR_Sort_Ordering"=>"ASC", 
	"hdnApplication_Filter_Default_Control_No_Delete"=>"0", 
	"hdnApplication_Filter_Default_Control_Allow_Empty_Default"=>"1", 
	"fltBegin_Date_DTE"=>Base.Dates.format(dstart, "d/m/yy"), 
	"fltEnd_Date_DTE"=>Base.Dates.format(dend, "d/m/yy"))
	
	if line != ""
		params["fltEquipment_Group"] = line
	end
	if !complete
		params["fltActive"] = "1"
	end
	if filter == WR_pms
		params["fltInclude_PM"] = "1"
	elseif filter == WR_not_pms
		params["fltInclude_PM"] = "0"
	end
	
	cache("work_request_csv_$(Dates.value(dstart))_$(Dates.value(dend))_$(line)", ()->p.py[:work_request_csv](params))
end
		

struct Maint_frm
	priority
	hours
	sinstruct
	tasklist
end
	

function pm_maint_frm(p::Plex, url)

	inpsel = false
	inp = false
	insi = false
	intasks = false
	intask = false
	
	taskrow = 0
	tcell = 0
	
	priority = ""
	hours = ""
	sinstruct = ""
	task = ""
	tasklist = Vector{String}()
	
	function proc(b::StartTag)
		vt = Dict( "select"=>()->if get(b.attrs, "name", "") == "lstPriority" inpsel = true end
				 , "option"=>()->if inpsel && get(b.attrs, "selected", "") != "" inp=true end
				 , "input"=>()->if get(b.attrs, "name", "") == "txtScheduled_Hours_DEC" hours = parse(Float32, get(b.attrs, "value", "0")) end
				 , "textarea"=>()->if get(b.attrs, "name", "") == "txtSpecial_Instructions_TXA_500" insi = true end
				 , "th"=>()->if get(b.attrs, "class", "") == "Module_Title"
								intasks = true 
								taskrow = 0
							end
				, "tbody"=>()->begin taskrow = taskrow + 1 end
				, "tr"=>()->if taskrow > 0 tcell = 0 end
				, "td"=>()->if taskrow > 0 tcell += 1 end
				, "b"=>()->if taskrow > 1 intask=true end
				)
		get(vt, b.name, ()->nothing)()	
	end

	function proc(b::Data)
		if inp 
			priority = b.data
			inp = false
			inpsel = false
			return
		end
		
		if insi
			sinstruct = b.data == "\r\n\t\t\t" ? "" : b.data
			insi = false
			return
		end
		
		if intask
			task = b.data
			intask = false
			push!(tasklist, task)
			return
		end
	end
	
	function proc(b::EndTag)
	
	end
	
	proc(b::Block) = nothing
	maint_frm_html = cache("maint_frm_html_" * replace(Base.Base64.base64encode(url), ['=', '/'], ""), ()->p.py[:pm_maint_frm](url))
	for blk in HTMLParser.HTML(maint_frm_html).blks
		proc(blk)
	end
	

	Maint_frm(priority, hours, sinstruct, tasklist)
end

struct PM_listing
	checklist_no
	checklist_key
	pmtitle
	start
	freq
	equipkey
	equiptxt
	schedfrm
	maintfrm
end

function pm_list(p::Plex)
	inpms = false
	inbody = false
	inpm = false

	checklist_no = ""
	pmtitle = ""
	start = now()
	freq = ""
	equiptxt = ""
	schedfrm = ""
	maintfrm = ""

	k = 0
	
	function proc(b::StartTag)
		vt = Dict( "table"=>()->if get(b.attrs, "class", "") == "StandardGrid"
						inpms = true
					end
					,"tbody"=>()->begin inbody = inpms end
					,"td" =>()-> if inbody && length(b.attrs) == 0
						inpm = true
					end
			)
		get(vt, b.name, ()->nothing)()	
	end
	
	
	function proc(b::EndTag)
		vt = Dict("tr" =>()-> if inpms && inpm 
						inpm = false
						k = 0
					end
					,"table"=>()->begin inpms = false end
			)
		get(vt, b.name, ()->nothing)()	
	end
	
	proc(b::Block) = nothing
	
	pms = Vector{PM_listing}()
	
	pm_list_html = cache("pm_list_html", p.py[:pm_list])
	
	for blk in HTMLParser.HTML(pm_list_html).blks
		proc(blk)
		
		if inpm
			k = k + 1
			
			#println("K ", k, " ", blk)
	
			if k == 2
				equiptxt = blk.data
				continue
			end
			if k == 7
				schedfrm = blk.attrs["href"]
				continue
			end
			if k == 8
				 checklist_no = replace(blk.data, [' ', '\r', '\n'], "")
			end
			if k == 14
				pmtitle = blk.data
				continue
			end
			if k == 18
				start = DateTime(blk.data, "dd/mm/yy") + Dates.Year(2000)
				continue
			end
			if k == 23
				maintfrm = blk.attrs["href"]
				continue
			end
			if k == 24
				freq = split(blk.data, ' ', limit=2)[1]
				
				mbits = split(maintfrm, ['=', '&'])
				push!(pms, PM_listing(checklist_no, mbits[6], pmtitle, start, freq, mbits[4], equiptxt, schedfrm, pm_maint_frm(p, maintfrm)))
				continue
				
			end
		end
	end
	

	pms
end

function pm_report(p::Plex)
	inpm = false
	ChkNo = ""
	EquipKey = ""
	LastComplete = ""
	DueDate = ""
	dk = 0
	pms = Vector()
	
	function proc(t::StartTag)
		attr(k) = get(t.attrs, k, "")
		vtable = Dict(
			"span"=>()->begin
				if attr("class") == "TextBodySmall"
					inpm = true
					dk = 0
				end
				end, 
			"a"=>()->
				if inpm
					bits = split(attr("href"), "&amp;")
					if length(bits) > 3
						ChkNo = split(bits[2], '=')[2]
						EquipKey = split(bits[3], '=')[2]
					end					
				end
			)
		get(vtable, t.name, ()->nothing)();
	end
	
	function proc(t::Data)
		if inpm			
			dk += 1
			if dk == 5
				if t.data == "&nbsp;"
					LastComplete = 0
				else
					LastComplete = Dates.value(Date(t.data, "mm/dd/yyyy"))
				end
				return
			end
			if dk == 6
				DueDate = Dates.value(Date(t.data, "mm/dd/yyyy"))
				return
			end
		end
	end

	function proc(t::EndTag)
		if inpm && t.name == "tr"
			inpm = false
			
			push!(pms, (LastComplete, DueDate, ChkNo, EquipKey))
		end
	end
	
	proc(b::Block) = nothing
	
	rawhtml = cache("pm_report", ()->p.py[:pm_report]())

	foreach(proc, HTMLParser.HTML(rawhtml).blks)
	pms
end


function dump_fields(wr)
	function pfield(k::String, t::Tuple)
		println(k, " - ", t[1], ":", t[2])
	end

	function pfield(k, v)
		println(k, " - ", v)
	end

	for (k, v) in wr[:fields]
		if k != nothing
			pfield(k, v)
		end
	end
end



#####
end