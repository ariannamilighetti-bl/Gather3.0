# -*- coding: utf-8 -*-
"""
Created on Tue Jan 30 14:10:39 2024

@author: amilighe
"""
from tkinter import *
import tkinter as tk
from tkinter import filedialog

from datetime import datetime

from lxml import etree
from lxml.builder import ElementMaker
from lxml.etree import Comment
from openpyxl import load_workbook


# These are the column number for the fields used
# If the template changes, change the column numbers here
# IAMS template:
repository_clmn = 0
coll_area_clmn = 1
collection_clmn = 2
level_clmn = 3
reference_clmn = 4
ext_ref_clmn = 5
title_clmn = 6
date_rng_clmn = 7
era_clmn = 8
calendar_clmn = 9
extent_clmn = 10
scope_content_clmn = 11
phys_char_clmn = 12
access_cond_clmn = 13
arrangement_clmn = 14
mat_language_clmn = 22
mat_langcode_clmn = 23
mat_script_clmn = 24
mat_scriptcode_clmn = 25
descr_lang_clmn = 26
descr_langcode_clmn = 27
descr_script_clmn = 28
descr_scriptcode_clmn = 29
rel_persons_clmn = 30
rel_fams_clmn = 31
rel_corp_bds_clmn = 32
rel_places_clmn = 33
rel_subject_clmn = 34
coordinates_clmn = 37
scale_clmn = 38
scale_des_clmn = 39
orientation_clmn = 41
legal_sts_clmn = 42
mat_type_clmn = 49
# Authority files document
auth_name_clmn = 36
auth_ark_id_clmn = 9

# Gather definitions
# Definitions used to create the nodes:


def get_header(ws):
    header = []
    for cell in ws[1]:
        header.append(cell.value)
    return header


def start_record(rec_num):
    return {"StartRecord": rec_num}


def tid(row, arg, shelfmark_modified, row_num):
    # the tid_num should grow within the xml across child records
    global tid_num
    if row_num == 0:
        tid_num = 1
    if row[arg].value:
        tid_full = shelfmark_modified+"_"+str(tid_num)
        tid_num = tid_num+1
        return {"tid": tid_full}
    else:
        return {}


def content(row, arg):
    if row[arg].value:
        return str(row[arg].value)
    else:
        return {}


def labels(row, arg, label):
    content = str(row[arg].value)
    return {label: content}


def header_label(header_row, arg, label):
    label_title = header_row[arg]
    return {label: label_title}


def date_normal(row, arg):
    full_date = str(row[arg].value)
    if "-" in full_date:
        index = full_date.rfind("-")
        start_date = full_date[index-4:index]
        end_date = full_date[-4:]
        date = start_date+"/"+end_date
    else:
        date = full_date[-4:]+"/"+full_date[-4:]
    return {"normal": date}


def pcontent(row, arg, E, shelfmark_modified, row_num):
    global tid_num
    content = []
    lists = E.list()
    if row[arg].value:
        list_content = []
        paragraphs = row[arg].value.split("</p>")
        for chunk in paragraphs:
            lines = chunk.split("</item>")
            for line in lines:
                line = line.strip()
                if line:
                    tid_label = tid(row, arg, shelfmark_modified, row_num)
                    if line.find('<emph render="italic">') != -1:
                        top = line.split('<emph render="italic">')[0]
                        emph_all = line.split(
                            '<emph render="italic">')[1]
                        emph = emph_all.split('</emph>')[0]
                        bottom = line.split("</emph>")[1]
                        emph_tid = shelfmark_modified+"_"+str(tid_num)
                        tid_num += 1
                        line = top + '<ead:emph render="italic" tid="' + emph_tid + '">' + emph + '</ead:emph>' + bottom
                    line = line.replace("<list>", "").replace(
                        "</list>", "")
                    if line.startswith("<item>"):
                        line = line.replace("<item>", "")
                        list_content.append(line)
                        line_content = E.item(line, tid_label)
                        lists.append(line_content)
                        content.append(lists)
                    elif list_content == []:
                        line = line.replace("<p>", "")
                        top_p = E.p(line, tid_label)
                        content.append(top_p)
                    else:
                        line = line.replace("<p>", "")
                        bttm_p = E.p(line,
                                     tid_label)
                        content.append(bttm_p)
    else:
        p = E.p()
        content.append(p)
    return content


# Authority Files processing definitions
def gen_auth_lookup(auth_ws, auth_name_clmn, auth_ark_id_clmn):
    auth_lookup = {}
    for row_num, row in enumerate(auth_ws.iter_rows()):
        auth_name = str(row[auth_name_clmn].value).strip().lower()
        if row[auth_ark_id_clmn].value:
            ark_id = row[auth_ark_id_clmn].value
        else:
            ark_id = "not_allocated"
        auth_lookup[auth_name] = ark_id
    return auth_lookup


def auth_dets(arg, label):
    if arg == "not_allocated":
        return {}
    else:
        return {label: arg}


def authority_files(row, arg, auth_lookup, E, shelfmark_modified, row_num):
    if row[arg].value:
        lines = row[arg].value.split("|")
        full_text = []
        for line in lines:
            attributes = line.split("=")
            subject = attributes[0]
            role_type = "not_allocated"
            altrender_type = "not_allocated"
            if len(attributes) > 3 and attributes[3]:
                role_type = attributes[3]
            if len(attributes) > 4 and attributes[4]:
                altrender_type = attributes[4]
            element_dict = {rel_persons_clmn: E.persname,
                            rel_fams_clmn: E.famname,
                            rel_corp_bds_clmn: E.corpname,
                            rel_places_clmn: E.geogname,
                            rel_subject_clmn: E.subject}
            text = element_dict[arg](
                subject,
                {"authfilenumber":
                 auth_lookup.get(subject.strip().lower(), "not_found")},
                auth_dets(role_type, "role"),
                {"source": "IAMS"},
                auth_dets(altrender_type, "altrender"),
                tid(row, arg, shelfmark_modified, row_num))

            full_text.append(text)
        return full_text
    else:
        return ""


def template_verification(ws):
    validation_check = 0
    row_num = 0
    for row in ws.iter_rows(min_row=2):
        row_num += 1
        if row[repository_clmn].value:
            if row[coll_area_clmn].value:
                if row[collection_clmn].value:
                    if row[level_clmn].value:
                        if row[reference_clmn].value:
                            if row[title_clmn].value:
                                if row[date_rng_clmn].value:
                                    if row[era_clmn].value:
                                        if row[calendar_clmn].value:
                                            if row[access_cond_clmn].value:
                                                if row[mat_language_clmn].value:
                                                    if row[mat_langcode_clmn].value:
                                                        if row[mat_script_clmn].value:
                                                            if row[mat_scriptcode_clmn].value:
                                                                if row[descr_lang_clmn].value:
                                                                    if row[mat_type_clmn].value:
                                                                        if row[scope_content_clmn].value:
                                                                            if row[scope_content_clmn].value.find("<p>")!= -1:
                                                                                validation_check += 1
                                                                        else:
                                                                            validation_check += 1
    if validation_check == row_num:
        validated = True
    else:
        validated = False
    return validated


# Full Gather process
def QatarGather(IAMS_filename, Auth_filename, end_directory):
    # The actual code starts here.
    # This part defines where the authority files details are held.
    # auth_file_name = 'TB_Authorities.xlsx'
    auth_file_wb = load_workbook(Auth_filename, read_only=True)
    auth_ws = auth_file_wb["in"]
    auth_lookup = gen_auth_lookup(auth_ws, auth_name_clmn, auth_ark_id_clmn)

    # This selects the spreadsheet to gather
    # One tab each shelfmark to gather.
    # wb_input = input('Please write the name of the spreadsheet to Gather.
    # Omit ".xlsx": ')
    # wb_name = wb_input + '.xlsx'
    wb = load_workbook(IAMS_filename, read_only=True)
    shelfmarks = wb.sheetnames
    shm_num = 0

    for shelfmark_modified in shelfmarks:
        shm_num += 1
        rec_num = 1
        row_num = 0
        ws = wb[shelfmark_modified]
        header_row = get_header(ws)
        E = ElementMaker(namespace="urn:isbn:1-931666-22-9",
                         nsmap={'ead': "urn:isbn:1-931666-22-9",
                                'xlink': "http://www.w3.org/1999/xlink",
                                'xsi': "http://www.w3.org/2001/XMLSchema-instance"
                                })
        full_ead = E.ead()

        sh_label = tk.Label(master=run_frame, text=shelfmark_modified)
        sh_label.grid(row=1+shm_num, column=0, padx=5, pady=5, sticky="nsew")
        sh_verif_lbl = tk.Label(master=run_frame)
        sh_verif_lbl.grid(row=1+shm_num, column=1, padx=5, pady=5,
                          sticky="nsew")
        sh_complete_label = tk.Label(master=run_frame)
        sh_complete_label.grid(row=1+shm_num, column=2, padx=5, pady=5,
                               sticky="nsew")
        validation = template_verification(ws)
        print(validation)
        if validation is True:
            sh_verif_lbl.configure(text="Verified", bg="green", fg="white")
        # This part creates the tree for each child shelfmark.
            for row in ws.iter_rows(min_row=2, values_only=False):
                ead = E.ead()
                comment = Comment(
                    f"New record starts here {row[reference_clmn].value}")
                full_ead.append(comment)
                shelfmark = str(row[reference_clmn].value)

                # header
                eadheader = E.eadheader(start_record(str(rec_num)))
                ead.append(eadheader)

                eadid = E.eadid(str(shelfmark),
                                tid(row, reference_clmn,
                                    shelfmark_modified, row_num))
                eadheader.append(eadid)
                row_num += 1

                filedesc = E.filedesc()  # wrapper node, should not have info
                eadheader.append(filedesc)

                titlestmt = E.titlestmt()  # wrapper node, should not have info
                filedesc.append(titlestmt)

                titleproper = E.titleproper()  # not used in IAMS material
                titlestmt.append(titleproper)

                profiledesc = E.profiledesc()  # wrapper node,  no info
                eadheader.append(profiledesc)

                creation = E.creation()  # not used in Qatar(?)
                profiledesc.append(creation)

                date_exported = E.date(str(datetime.now()
                                           .strftime("%Y-%m-%dT%H:%M:%S")),
                                       {"type": "exported"},
                                       tid(row, reference_clmn,
                                           shelfmark_modified, row_num))
                creation.append(date_exported)

                date_modified = E.date(str(wb.properties.modified.strftime(
                    "%Y-%m-%dT%H:%M:%S")), {"type": "modified"},
                    tid(row, reference_clmn, shelfmark_modified, row_num))
                creation.append(date_modified)

                langusage = E.langusage()  # not used in IAMS material
                profiledesc.append(langusage)

                # this is language of the description
                language = E.language(content(row, descr_lang_clmn),
                                      labels(row, descr_langcode_clmn,
                                             "langcode"),
                                      labels(row, descr_scriptcode_clmn,
                                             "scriptcode"),
                                      tid(row, descr_lang_clmn,
                                          shelfmark_modified, row_num))
                langusage.append(language)

                # archdesc
                archdesc = E.archdesc(labels(row, level_clmn, "level"))
                ead.append(archdesc)

                did = E.did()  # wrapper node, should not have info
                archdesc.append(did)

                # British Library: Indian Office Records
                repository = E.repository(
                    row[repository_clmn].value + ": " + row[coll_area_clmn
                                                            ].value,
                    tid(row, repository_clmn, shelfmark_modified, row_num))
                did.append(repository)

                unitid = E.unitid(shelfmark, {"label": "IAMS_label_NA"},
                                  {"identifier": "ark_identifier"},
                                  tid(row, reference_clmn, shelfmark_modified,
                                      row_num))
                # These are the IAMS identifiers (ark and number)
                did.append(unitid)

                # this will say "title"
                unittitle = E.unittitle(header_label(header_row, title_clmn,
                                                     "label"))
                did.append(unittitle)

                # Item title
                title = E.title(content(row, title_clmn),
                                tid(row, title_clmn, shelfmark_modified,
                                    row_num))
                unittitle.append(title)

                if row[ext_ref_clmn].value:
                    unittitle = E.unittitle(content(
                        row, ext_ref_clmn), header_label(header_row,
                                                         ext_ref_clmn,
                                                         "label"),
                                            tid(row, ext_ref_clmn,
                                                shelfmark_modified, row_num))
                else:
                    unittitle = E.unittitle(
                        {"label": "Former external reference"})
                did.append(unittitle)  # Former external reference

                unittitle = E.unittitle({"label": "Former internal reference"})
                did.append(unittitle)  # Former internal reference (not used)

                unitdate = E.unitdate(content(row, date_rng_clmn), {
                    "datechar": "Creation"},
                    labels(row, calendar_clmn, "calendar"),
                    labels(row, era_clmn, "era"),
                    date_normal(row, date_rng_clmn),
                    tid(row, date_rng_clmn, shelfmark_modified, row_num))
                did.append(unitdate)  # Date of the material

                # langmaterial = E.langmaterial()  # This is language
                # did.append(langmaterial)

                # This allows for multiple languages and lang codes separated by |
                # (mat_language_clmn, mat_langcode_clmn) for language
                # (mat_script_clmn, mat_scriptcode_clmn) for script
                code_label_index = 0
                for r in ((mat_language_clmn, mat_langcode_clmn),
                          (mat_script_clmn, mat_scriptcode_clmn)):
                    langmaterial = E.langmaterial()
                    did.append(langmaterial)
                    languages = row[r[0]].value.split("|")
                    lang_codes = row[r[1]].value.split("|")
                    code_labels = ["langcode", "scriptcode"]
                    for lang, c in zip(languages, lang_codes):
                        language = E.language(
                            lang, {code_labels[code_label_index]: c},
                            tid(row, r[0], shelfmark_modified, row_num))
                        langmaterial.append(language)
                    code_label_index += 1

                physdesc = E.physdesc()  # wrapper node, should not have info
                did.append(physdesc)

                extent = E.extent(content(row, extent_clmn),
                                  tid(row, extent_clmn, shelfmark_modified,
                                      row_num))
                physdesc.append(extent)
        # Map details generated here
                if row[mat_type_clmn].value == 'Maps and Plans':
                    if row[scope_content_clmn].value:
                        materialspec = E.materialspec(content(row, scale_clmn),
                                                      {'type': "scale"},
                                                      tid(row, scale_clmn,
                                                          shelfmark_modified,
                                                          row_num))
                        did.append(materialspec)
                        materialspec = E.materialspec(
                            content(row, scale_des_clmn),
                            {'type': "scale designator"},
                            tid(row, scale_des_clmn, shelfmark_modified,
                                row_num))
                        did.append(materialspec)
                        materialspec = E.materialspec(
                            content(row, coordinates_clmn),
                            {'type': "coordinates"}, {'label': "decimal"},
                            tid(row, coordinates_clmn, shelfmark_modified,
                                row_num))
                        did.append(materialspec)
                        materialspec = E.materialspec(
                            content(row, orientation_clmn),
                            {'type': "orientation"}, tid(row, orientation_clmn,
                                                         shelfmark_modified,
                                                         row_num))
                        did.append(materialspec)

                accessrestrict = E.accessrestrict()
                for p in pcontent(row, access_cond_clmn, E, shelfmark_modified,
                                  row_num):
                    accessrestrict.append(p)
                archdesc.append(accessrestrict)

                accessrestrict = E.accessrestrict()
                # This second accessrestrict is a wrapper node
                archdesc.append(accessrestrict)

                legalstatus = E.legalstatus(content(row, legal_sts_clmn),
                                            tid(row, legal_sts_clmn,
                                                shelfmark_modified, row_num))
                accessrestrict.append(legalstatus)

                accruals = E.accruals()  # Empty node
                p = E.p()
                accruals.append(p)
                archdesc.append(accruals)

                bioghist = E.bioghist()  # Empty node
                p = E.p()
                bioghist.append(p)
                archdesc.append(bioghist)

                appraisal = E.appraisal()  # Empty node
                p = E.p()
                appraisal.append(p)
                archdesc.append(appraisal)

                arrangement = E.arrangement()
                for p in pcontent(row, arrangement_clmn, E, shelfmark_modified,
                                  row_num):
                    arrangement.append(p)
                archdesc.append(arrangement)

        # This allows to skip the node if item is part of a bigger volume
                if row_num == 1 or row[mat_type_clmn].value != "Archives and Manuscripts":
                    phystech = E.phystech()
                    for p in pcontent(
                            row, phys_char_clmn, E, shelfmark_modified,
                            row_num):
                        phystech.append(p)
                    archdesc.append(phystech)

                scopecontent = E.scopecontent()
                for p in pcontent(
                        row, scope_content_clmn, E, shelfmark_modified,
                        row_num):
                    scopecontent.append(p)
                archdesc.append(scopecontent)

                userestrict = E.userestrict()  # Empty node
                p = E.p()
                userestrict.append(p)
                archdesc.append(userestrict)

                odd = E.odd()  # Empty node
                p = E.p()
                odd.append(p)
                archdesc.append(odd)

                controlaccess = E.controlaccess()
                genreform = E.genreform(
                    content(row, mat_type_clmn), {"source": "IAMS"},
                    tid(row, mat_type_clmn, shelfmark_modified, row_num))
                controlaccess.append(genreform)

                # Authority files processing starts here:
                for arg in range(rel_persons_clmn, rel_subject_clmn+1, 1):
                    for af in authority_files(row, arg, auth_lookup, E,
                                              shelfmark_modified, row_num):
                        controlaccess.append(af)
                archdesc.append(controlaccess)
                # End of authority files

                note = E.note({"type": "project/collection"})
                for p in pcontent(row, coll_area_clmn, E, shelfmark_modified,
                                  row_num):
                    note.append(p)
                controlaccess.append(note)

                note = E.note({"type": "project/collection"})
                for p in pcontent(row, collection_clmn, E, shelfmark_modified,
                                  row_num):
                    note.append(p)
                controlaccess.append(note)

                rec_num += 1

    # This puts together two parts of each child's tree (header+description)
    # This will append as many children as there are in the Excel tab
                full_ead.append(eadheader)
                full_ead.append(archdesc)

        # This part writes out the XML file
            with open(end_directory+"/"+shelfmark_modified+"_"+str(
                    datetime.now().strftime("%Y%m%d_%H%M%S")
                    )+'.xml', 'wb') as f:
                f.write(etree.tostring(
                    full_ead, encoding="utf-8", xml_declaration=True,
                    pretty_print=True))
            sh_complete_label.config(text="Complete")
            print(shelfmark + ' complete \n')
        else:
            sh_verif_lbl.configure(text="Not recognised", bg="#cc0000",
                                   fg="white")
            sh_complete_label.config(text="Could not complete", bg="#cc0000",
                                     fg="white")
    wb.close()
    auth_file_wb.close()
    print('Gather complete!')
    status = "Complete!"
    complete_label.config(text=status)


# Editor Definitions
def UploadIAMS(event=None):
    global IAMS_filename
    IAMS_filename = filedialog.askopenfilename(
        filetypes=(("Excel files", "*.xlsx"), ("Any file", "*")))
    IAMS_filename_label.delete(0, END)
    IAMS_filename_label.insert(0, IAMS_filename)


def UploadAuth(event=None):
    global auth_filename
    auth_filename = filedialog.askopenfilename(
        filetypes=(("Excel files", "*.xlsx"), ("Any file", "*")))
    auth_filename_label.delete(0, END)
    auth_filename_label.insert(0, auth_filename)


def askDirectory(event=None):
    global end_directory
    end_directory = filedialog.askdirectory()
    end_directory_label.delete(0, END)
    end_directory_label.insert(0, end_directory)


app = tk.Tk()
app.title("Gather Renewed")

title_lbl = tk.Label(master=app, text="Gather Renewed", bg="#eeeeee",
                     fg="black", anchor="w", font=("calibri", 18, "bold"))
instr_lbl = tk.Label(master=app,
                     text="Please select the IAMS template to gather, the relevant Authority File sheet and a destination folder for the created files. Once ready, click 'Run'.",
                     anchor="w")

selection_frm = tk.LabelFrame(master=app, text="Select Files", bg="#eeeeee",
                              fg="black")
# IAMS Fields
IAMS_label = tk.Label(master=selection_frm, text="Import IAMS Template",
                      bg="#eeeeee", fg="black")
IAMS_button = tk.Button(selection_frm, text='Open', bg="#0b5394", fg="white",
                        command=UploadIAMS)
IAMS_filename_label = tk.Entry(master=selection_frm)
# AuthFiles Fields
auth_file_label = tk.Label(master=selection_frm,
                           text="Import Authorities Spreadsheet", bg="#eeeeee",
                           fg="black")
auth_file_button = tk.Button(selection_frm, text='Open', bg="#0b5394",
                             fg="white", command=UploadAuth)
auth_filename_label = tk.Entry(master=selection_frm)
# Destination Fields
dir_label = tk.Label(master=selection_frm,
                     text="Select directory to save files to", bg="#eeeeee",
                     fg="black")
dir_button = tk.Button(selection_frm, text='Open', bg="#0b5394", fg="white",
                       command=askDirectory)
end_directory_label = tk.Entry(master=selection_frm)
# Run button
run_button = tk.Button(master=app, text="Run", bg="green", fg="white",
                       command=lambda: QatarGather(IAMS_filename,
                                                   auth_filename,
                                                   end_directory))
# Running frame
run_frame = tk.LabelFrame(master=app, text="Running", bg="#eeeeee", fg="black")
run_shmark = tk.Label(master=run_frame, text="Shelfmark", bg="#eeeeee",
                      fg="black")
run_verification = tk.Label(master=run_frame, text="IAMS Template Validation",
                            bg="#eeeeee", fg="black")
run_status = tk.Label(master=run_frame, text="Status", bg="#eeeeee",
                      fg="black")

# complete_label = tk.Label(master=app, text=status)

title_lbl.grid(column=0, row=0, sticky="nsew", padx=5)
instr_lbl.grid(column=0, row=1, sticky="nsew", padx=5)
selection_frm.grid(column=0, row=2, sticky="nsew", padx=5)

IAMS_label.grid(column=0, row=0, columnspan=1, sticky="nsew", padx=5, pady=5)
IAMS_button.grid(column=1, row=0, columnspan=1, sticky="nsew", padx=5, pady=5)
IAMS_filename_label.grid(column=2, row=0, columnspan=3, sticky="nsew", padx=5,
                         pady=5)

auth_file_label.grid(column=0, row=1, columnspan=1, sticky="nsew", padx=5,
                     pady=5)
auth_file_button.grid(column=1, row=1, columnspan=1, sticky="nsew", padx=5,
                      pady=5)
auth_filename_label.grid(column=2, row=1, columnspan=3, sticky="nsew", padx=5,
                         pady=5)

dir_label.grid(column=0, row=2, columnspan=1, sticky="nsew", padx=5, pady=5)
dir_button.grid(column=1, row=2, columnspan=1, sticky="nsew", padx=5, pady=5)
end_directory_label.grid(column=2, row=2, columnspan=3, sticky="nsew", padx=5,
                         pady=5)

run_button.grid(column=0, row=3, sticky="nsew", padx=5, pady=5)

run_frame.grid(column=0, row=4, sticky="nsew", padx=5, pady=5)
run_shmark.grid(column=0, row=0, sticky="nsew", padx=5, pady=5)
run_verification.grid(column=1, row=0, sticky="nsew", padx=5, pady=5)
run_status.grid(column=2, row=0, sticky="nsew", padx=5, pady=5)

app.columnconfigure(0, weight=1)
selection_frm.columnconfigure(0, weight=1)
selection_frm.columnconfigure(1, weight=1)
selection_frm.columnconfigure(2, weight=10)
run_frame.columnconfigure(0, weight=1)
run_frame.columnconfigure(1, weight=2)
run_frame.columnconfigure(2, weight=3)

app.mainloop()