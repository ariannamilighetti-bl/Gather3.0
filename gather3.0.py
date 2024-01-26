# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

from lxml import etree
from lxml.builder import ElementMaker
from datetime import datetime
from openpyxl import load_workbook

def get_header():
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
def content(row,arg):
    if row[arg].value:  
        if "\n" in row[arg].value:
            print ("yes")
            lines = row[arg].value.split("\n")
            full_text = []
            print(lines[1])
            for line in lines:
               text = E.p(line,tid(row,arg))
               full_text.append(text)
            return full_text
        else:
            return row[arg].value
        
    else:
        return ""
    
shelfmark_to_gather = input()
rec_num = 1
tid_num=1

wb = load_workbook('IAMS_import_template.xlsx', read_only=True)
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
UNITTITLE = E.unittile
TITLE = E.title
UNITDATE = E.unitdate
LANGMATERIAL = E.langmaterial
PHYSDESC = E.physdesc
EXTENT = E.extent
ACCESSRESTRICT = E.ACCESSRESTRICT
P = E.p
P1 = E.p1
LEGALSTATUS = E.legalstatus
ACCRUALS = E.accruals
BIOGHIST = E.bioghist
APPRAISAL = E.appraisal
ARRANGEMENT = E.arrangement
PHYSTECH = E.phystech
SCOPECONTENT = E.scopecontent
USERRESTRICT = E.userrestrict
ODD = E.odd
CONTROLACCESS = E.controlaccess
GENREFORM = E.genreform
NOTE = E.note


full_ead = EAD()
header_row = get_header()

for row in ws.iter_rows(min_row=2, values_only=False):
    print (row[5].value)
    shelfmark = str(row[5].value)
    script = row[42]
    current_shelfmark = EAD (
        EADHEADER(StartRecord(str(rec_num)),
            EADID(str(shelfmark),tid(row,5)),
            
            FILEDESC(
                TITLESTMT(
                    TITLEPROPER(),
                ),
            ),
            PROFILEDESC(
                CREATION(
                    DATE(str(datetime.now()), typed("exported"),tid(row,5)),
                    DATE("15/12/2023", typed("modified"),tid(row,5)),
                ),
                LANGUSAGE(
                    LANGUAGE(content(row,40),langcode(row),scriptcode(row),tid(row,40)),
                    ),
            ),
        ), 
        ARCHDESC(level(row),
            DID(
                REPOSITORY(row[0].value + ": " + row[1].value,tid(row,0)),
                UNITID(shelfmark, IAMSlabel(row), identifier(row),tid(row,5)),
                UNITTITLE(label(10,header_row),
                    TITLE(content(row,10),tid(row,10)),
                ),
                UNITTITLE(label(7,header_row)),
                UNITTITLE(label(6,header_row)),
                UNITDATE(content(row,14),datechar(row),calendar(row),era(row),normal(row),tid(row,14)),
                LANGMATERIAL(
                    LANGUAGE(content(row,41),mat_langcode(row),tid(row,41)),
                ),
                LANGMATERIAL(
                    LANGUAGE(content(row,43),mat_scriptcode(row),tid(row,43)),
                ),
                PHYSDESC(
                    EXTENT(content(row,19),tid(row,19)),
                ),
            ),
            ACCESSRESTRICT(
                P(content(row,25),tid(row,25)),    
            ),
            ACCESSRESTRICT(
                LEGALSTATUS(content(row,71),tid(row,71)),    
            ),
            ACCRUALS(
                P(),
            ),
            BIOGHIST(
                P(),
            ),
            APPRAISAL(
                P(),
            ),
            ARRANGEMENT(
                P(content(row,31),tid(row,31)),
            ),
            PHYSTECH(
                P(content(row,21)),
            ),
            SCOPECONTENT(
                P(content(row,20),tid(row,20)),
            ),
            USERRESTRICT(
                P(),
            ),
            ODD(
                P(),
            ),
            CONTROLACCESS(
                GENREFORM(content(row,79), tid(row,79)),
            #missing authority files
            ),
            NOTE(label(2,header_row),
                 P("India Office Records",tid(row,5)),
            ),
            NOTE(label(2,header_row),
                 P(content(row,2), tid(row,2)),
            ),
        ),         
    )
    rec_num = rec_num+1
    full_ead.append(current_shelfmark)



with open(shelfmark_to_gather+'.xml', 'wb') as f:
    f.write(etree.tostring(full_ead, pretty_print=True))

                                                           