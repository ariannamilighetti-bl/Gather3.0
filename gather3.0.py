# -*- coding: utf-8 -*-
"""
Created on Wed Dec 20 09:05:04 2023

@author: amilighe

Issues: not sure how to create bullet points
"""

from lxml import etree
from lxml.builder import ElementMaker
from lxml.etree import Comment
from datetime import datetime
from openpyxl import load_workbook

def get_header(ws):
    header = []
    for cell in ws[1]:  
        header.append(cell.value)
    return header
def StartRecord (rec_num): 
    return {"StartRecord":rec_num}
def tid (row,arg):
    global tid_num
    if row[arg].value != None:
        fixed_shelfmark = row[3].value.replace("/","_").replace(" ","_").replace(",_f_","_ f ")
        tid_full = fixed_shelfmark+"_"+str(tid_num)
        tid_num=tid_num+1
        return{"tid":tid_full}
    else:
        return""
def typed(arg):
    return{"type":arg}
def langcode (row):
    lang_code = row[45].value
    return{"langcode":lang_code}
def scriptcode (row):
    script_code = row[47].value
    return{"scriptcode":script_code}
def level(row):
    item_level = row[4].value
    return{"level":item_level}
def IAMSlabel(row):
    return{"label":"IAMS_label_NA"}
def identifier(row):
    return{"identifier":"ark_identifier"}
def label(arg,header_row):
    label_title=header_row[arg]
    return{"label":label_title}
def datechar(row):
    return{"datechar":"Creation"}
def calendar(row):
    calendar_type = row[18].value
    return{"calendar":calendar_type}
def era(row):
    era_type = row[17].value
    return{"era":era_type}
def normal(row):
    normal_type = str(row[15].value)+"/"+str(row[16].value)
    return{"normal":normal_type}
def mat_langcode(row):
    lang_code = row[40].value
    return{"langcode":lang_code}
def mat_scriptcode(row):
    script_code = row[42].value
    return{"scriptcode":script_code}
def source(row):
    source = "IAMS"
    return {"source":source}
def authfilenumber(row):
    return {"authfilenumber":"ark_id_IAMS"}
def role (row):
    return {"role":"IAMS_role"}
def source1 (row):
    return {"source":"IAMS_source"}
def altrender (row):
    return{"altrender":"IAMS_altrender"}
def content(row,arg):
    if row[arg].value:
        return row[arg].value    
    else:
       return ""
def pcontent(row, arg):
    content = []   
    if row[arg].value:      
        lines = row[arg].value.split("\n")           
        for line in lines:
            if line.startswith("-"):
            # Line is a bullet point 
                item = E.item(line.strip("-"), tid(row, arg))  
                lists.append(item)
                content.append(lists)
            else:
                p = E.p(line, tid(row,arg))   
                content.append(p)        
    else:
        p = E.p("")
        content.append(p)
    return content

def persnamecontent(row,arg):
    if row[arg].value:
        lines = row[arg].value.split("|")
        full_text = []
        for line in lines:
            text = E.persname(line,authfilenumber(row),role(row),source1(row),altrender(row),tid(row,arg))  
            full_text.append(text)
        return full_text
    else:
        return ""
def famnamecontent(row,arg):
    if row[arg].value:
        lines = row[arg].value.split("|")
        full_text = []
        for line in lines:
            text = E.famname(line,authfilenumber(row),role(row),source1(row),altrender(row),tid(row,arg))  
            full_text.append(text)
        return full_text
    else:
        return ""
def corpnamecontent(row,arg):
    if row[arg].value:
        lines = row[arg].value.split("|")
        full_text = []
        for line in lines:
            text = E.corpname(line,authfilenumber(row),role(row),source1(row),altrender(row),tid(row,arg))  
            full_text.append(text)
        return full_text
    else:
        return ""
def placecontent(row,arg):
    if row[arg].value:
        lines = row[arg].value.split("|")
        full_text = []
        for line in lines:
            text = E.places(line,authfilenumber(row),role(row),source1(row),altrender(row),tid(row,arg))  
            full_text.append(text)
        return full_text
    else:
        return ""
def subjectcontent(row,arg):
    if row[arg].value:
        lines = row[arg].value.split("|")
        full_text = []
        for line in lines:
            text = E.subject(line,authfilenumber(row),role(row),source1(row),altrender(row),tid(row,arg))  
            full_text.append(text)
        return full_text
    else:
        return ""
list_of_shelfmarks = input ('Please write the shelfmark refereneces to Gather. Enter shelfmarks separated by a ";": ')
shelfmarks = list_of_shelfmarks.split(';')
for shelfmark in shelfmarks:
    shelfmark_to_gather = shelfmark.replace("/","_").replace(" ","_").replace(",_f_","_ f ")
    print(shelfmark_to_gather)
    rec_num = 1
    tid_num=1
    
    wb = load_workbook('IAMS_Transl_Gather.xlsx', read_only=True)
    try:
        ws = wb[shelfmark_to_gather]
    except KeyError:
        print("Sheet not found")
        
    E = ElementMaker(namespace="urn:isbn:1-931666-22-9",nsmap={'ead': "urn:isbn:1-931666-22-9", 'xlink': "http://www.w3.org/1999/xlink", 'xsi': "http://www.w3.org/2001/XMLSchema-instance"})
    
    EAD = E.ead
    EADHEADER = E.eadheader
    EADID = E.eadid
    FILEDESC = E.filedesc
    TITLESTMT = E.titlestmt
    TITLEPROPER = E.titleproper
    PROFILEDESC = E.profiledesc
    CREATION = E.creation
    DATE = E.date
    LANGUSAGE = E.langusage
    LANGUAGE = E.language
    LANGUAGE1 = E.language1
    ARCHDESC = E.archdesc
    DID = E.did
    REPOSITORY = E.repository
    UNITID = E.unitid
    UNITTITLE = E.unittitle
    TITLE = E.title
    UNITDATE = E.unitdate
    LANGMATERIAL = E.langmaterial
    PHYSDESC = E.physdesc
    EXTENT = E.extent
    ACCESSRESTRICT = E.accessrestrict
    P = E.p
    LEGALSTATUS = E.legalstatus
    ACCRUALS = E.accruals
    BIOGHIST = E.bioghist
    APPRAISAL = E.appraisal
    ARRANGEMENT = E.arrangement
    PHYSTECH = E.phystech
    SCOPECONTENT = E.scopecontent
    LISTS = E.lists
    USERRESTRICT = E.userrestrict
    ODD = E.odd
    CONTROLACCESS = E.controlaccess
    GENREFORM = E.genreform
    PERSNAME = E.persname
    FAMNAME = E.famname
    CORPNAME = E.corpname
    SUBJECT = E.subject
    PLACES = E.places
    NOTE = E.note
    
    full_ead = EAD()
    header_row = get_header(ws)
    
    for row in ws.iter_rows(min_row=2, values_only=False):
       comment = Comment(f"New record starts here {row[5].value}")
       full_ead.append(comment)
       print (row[5].value)
       shelfmark = str(row[5].value)
       script = row[42]
       
       ead = EAD()
       eadheader = EADHEADER(StartRecord(str(rec_num)))
       eadid = EADID(str(shelfmark),tid(row,5))
       filedesc = FILEDESC()
       titlestmt = TITLESTMT()
       titleproper = TITLEPROPER()
       profiledesc = PROFILEDESC()
       creation = CREATION()
       date = DATE()
       for dates in range(0,2,1):
           if dates == 0:
               item = DATE(str(datetime.now()), typed("exported"),tid(row,5))
           else:
               item = DATE("15/12/2023", typed("modified"),tid(row,5))
           date.append(item)
       langusage = LANGUSAGE()
       language = LANGUAGE(content(row,40),langcode(row),scriptcode(row),tid(row,40))
       
       
       archdesc = ARCHDESC(level(row))
       did = DID()
       repository= REPOSITORY(row[0].value + ": " + row[1].value,tid(row,0))
       unitid = UNITID(shelfmark, IAMSlabel(row), identifier(row),tid(row,5))
       unittitle1 = UNITTITLE(label(10,header_row))
       title = TITLE(content(row,10),tid(row,10))
       unittitle2 = UNITTITLE(label(7,header_row))
       unittitle3 = UNITTITLE(label(6,header_row))
       unitdate = UNITDATE(content(row,14),datechar(row),calendar(row),era(row),normal(row),tid(row,14))
       langmaterial = LANGMATERIAL()
       language1 = LANGUAGE(content(row,41),mat_langcode(row),tid(row,41))
       language2 = LANGUAGE(content(row,43),mat_scriptcode(row),tid(row,43))
       physdesc = PHYSDESC()
       extent = EXTENT(content(row,19),tid(row,19))
       
       
       accessrestrict = ACCESSRESTRICT()
       for p in pcontent(row, 25):
           accessrestrict.append(p)
       accessrestrict1 = ACCESSRESTRICT()
       legalstatus = LEGALSTATUS(content(row,71),tid(row,71))
       accruals = ACCRUALS()
       for p in pcontent(row, 23):
           accruals.append(p)
       bioghist = BIOGHIST()
       for p in pcontent(row, 24):
           bioghist.append(p)
       appraisal = APPRAISAL()
       for p in pcontent(row, 22):
           appraisal.append(p)
       arrangement = ARRANGEMENT()
       for p in pcontent(row,31):
           arrangement.append(p)
       phystech = PHYSTECH()
       for p in pcontent(row,21):
           phystech.append(p)
       lists = E.lists()
       scopecontent = SCOPECONTENT()
       for p in pcontent(row,20):
           scopecontent.append(p)
       scopecontent_top = scopecontent[:scopecontent.index(lists)]  
       scopecontent_bottom = scopecontent[scopecontent.index(lists):]
       scopecontent.append(scopecontent_top)
       scopecontent.append(lists) 
       scopecontent.append(scopecontent_bottom)
       userestrict = USERRESTRICT()
       for p in pcontent(row, 27):
           userestrict.append(p)
       odd = ODD()
       for p in pcontent(row, 36):
           #unsure if this is the correct mapping for this field!
           odd.append(p)
       controlaccess = CONTROLACCESS()
       genreform = GENREFORM(content(row,79),source(row),tid(row,79))
       controlaccess.append(genreform)
       for persname in persnamecontent(row,48):
           controlaccess.append(persname)
       for famname in famnamecontent(row,49):
           controlaccess.append(famname)
       for corpname in corpnamecontent(row,50):
           controlaccess.append(corpname)
       for places in placecontent(row,51):
           controlaccess.append(places)
       for subject in subjectcontent(row,49):
           controlaccess.append(subject)
       note = NOTE(label(2,header_row))
       for p in pcontent(row,2):
           note.append(p)
                   
    #Header creation 
       ead.append(eadheader)
       eadheader.append(eadid)
       eadheader.append(filedesc)
       filedesc.append(titlestmt)
       titlestmt.append(titleproper)
       eadheader.append(profiledesc)
       profiledesc.append(creation)
       creation.append(date)
       creation.append(date)
       profiledesc.append(langusage)
       langusage.append(language)
    #Archdesc creation
       ead.append(archdesc)
       archdesc.append(did)
       did.append(repository)
       did.append(unitid)
       did.append(unittitle1)
       unittitle1.append(title)
       did.append(unittitle2)
       did.append(unittitle3)
       did.append(unitdate)
       did.append(langmaterial)
       langmaterial.append(language1)
       langmaterial.append(language2)
       did.append(physdesc)
       physdesc.append(extent)
       archdesc.append(accessrestrict)
       archdesc.append(accessrestrict1)
       accessrestrict1.append(legalstatus)
       archdesc.append(accruals)
       archdesc.append(bioghist)
       archdesc.append(appraisal)
       archdesc.append(arrangement)
       archdesc.append(phystech)
       archdesc.append(scopecontent)
       
       archdesc.append(userestrict)
       archdesc.append(odd)
       archdesc.append(controlaccess)
       archdesc.append(note)
       
       rec_num = rec_num+1
       
       full_ead.append(ead)
    
    
    
    with open(shelfmark_to_gather+'.xml', 'wb') as f:
        f.write(etree.tostring(full_ead, pretty_print=True))