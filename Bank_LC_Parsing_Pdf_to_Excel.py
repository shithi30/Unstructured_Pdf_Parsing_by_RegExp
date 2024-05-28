#!/usr/bin/env python
# coding: utf-8

# In[1]:


# import
import glob
from pathlib import Path
import win32com.client
from win32com.client import Dispatch
import pandas as pd
import duckdb
import re
from pdfminer.high_level import extract_text
import pikepdf
import fitz
from pretty_html_table import build_table
import random
from datetime import datetime
import time


# In[2]:


# fetch LCs
def fetch_read_lc(rec_date_from, rec_date_to): 

    # output folder
    output_dir = Path.cwd() / 'Emailed LCs'
    output_dir.mkdir(parents=True, exist_ok=True)

    # outlook inbox
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    inbox = outlook.Folders.Item(1).Folders['Trade Alerts']

    # emails
    messages = inbox.Items
    for message in reversed(messages): 

        # time
        try: rec_date = str(message.SentOn)[0:10]
        except: continue
        if rec_date < rec_date_from or rec_date > rec_date_to: continue

        # attachments
        attachments = message.Attachments
        for attachment in attachments:
            
            # LCs
            filename = attachment.FileName
            if re.match(".+-T02.PDF", filename): attachment.SaveAsFile(output_dir / "HSBC LCs" / filename)      
            if re.match("^[0-9]+.pdf$", filename): attachment.SaveAsFile(output_dir / "SCB LCs" / filename)     
            if re.match(".+ACK.pdf", filename, re.IGNORECASE): attachment.SaveAsFile(output_dir / "MTB LCs" / filename)
            if re.match(".+SWIFT.+.pdf", filename, re.IGNORECASE): attachment.SaveAsFile(output_dir / "BBL LCs" / filename)
            if re.match(".+700.pdf", filename, re.IGNORECASE): attachment.SaveAsFile(output_dir / "PRB LCs" / filename)


# In[3]:


# HSBC
def parse_hsbc(f):

    # fetch datapoint
    def get_data_btn(text_str, sub1, sub2):
        text_str = text_str.replace(sub1, "*").replace(sub2, "*")
        datapoint = text_str.split("*")[1].strip()
        return datapoint

    # scrape
    pdf_name = []
    dc_no = []
    dc_curr = []
    dc_amt = []
    beneficiary = []
    issue_date = []
    insurance_no = []
    insurance_date = []
    material_name = []
    payment_term = []
    bb_ref = []
    inv_no = []
    hs_code_importer = []
    hs_code_exporter = []
    hs_code = []
    upas_tenor_1 = []
    upas_tenor_2 = []
    
    # doc name
    val = f.split("\\")[-1]
    print("Parsing: " + val)
    pdf_name.append(val)

    # all text
    text = extract_text(f)
    text = text.replace("\n", " ")
    text = re.sub("Page:\d+\s/\s\d+", " ", text)
    text = re.sub(" +", " ", text)

    # dc no.
    val = get_data_btn(text, "DC NO:", "DATE OF ISSUE:")
    dc_no.append(val)

    # dc value
    val = get_data_btn(text, "DC AMT: ", " AVAILABLE WITH/BY")
    # currency
    pattern = re.compile("^[A-Z]*")
    dc_curr.append(pattern.findall(val)[0])
    # amount
    pattern = re.compile("[0-9,]+")
    dc_amt.append(pattern.findall(val)[0].strip())

    # beneficiary
    val = get_data_btn(text, "BENEFICIARY: ", "DC AMT:")
    pattern = re.compile(".+(?:LTD|LIMITED)")
    try: beneficiary.append(pattern.findall(val)[0])
    except: beneficiary.append(val)

    # issue date
    val = get_data_btn(text, "DATE OF ISSUE:", "APPLICABLE RULES:")
    issue_date.append(val)

    # insurance date
    val = get_data_btn(text, "MENTIONING INSURANCE", "WITHIN")
    insurance_date.append(val.split()[-1])

    # insurance no.
    pattern = re.compile("UIC[\w/\-\(\s]+[\)0-9]\s")
    insurance_no.append(pattern.findall(val)[0].strip())

    # BB ref
    pattern = re.compile("BANK DC NO.\s*\d+")
    try: bb_ref.append(pattern.findall(text)[0])
    except: bb_ref.append(None)

    # invoice no. 
    val = get_data_btn(text, "GOODS: ", "DOCUMENTS REQUIRED: ")
    pattern = re.compile("(?:INDENT|INVOICE) NO.+\d\sDATE")
    try: inv_no.append(pattern.findall(val)[0][:-5])
    except: inv_no.append(None)

    # payment term
    pattern = re.compile(".+\+")
    try: payment_term.append(pattern.findall(val)[0][:-2])
    except: payment_term.append(None)

    # HS code - importer
    pattern = re.compile("\d+\.\d+\.\d+\s*\(IMP")
    try: hs_code_importer.append(pattern.findall(text)[0].split("(")[0].strip())
    except: hs_code_importer.append(None)

    # HS code - exporter
    pattern = re.compile("\d+\.\d+\.\d+\s*\(EXP")
    try: hs_code_exporter.append(pattern.findall(text)[0].split("(")[0].strip())
    except: hs_code_exporter.append(None)

    # HS code
    pattern = re.compile("\d+\.\d+\.\d+\s*\(")
    if len(pattern.findall(text)) > 0: hs_code.append(None)
    else: pattern = re.compile("\d+\.\d+\.\d+")
    if len(pdf_name) > len(hs_code): hs_code.append(pattern.findall(text)[0])

    # UPAS tenor
    val = get_data_btn(text, "DESPITE", "AT MATURITY")
    pattern = re.compile("\d+")
    try: upas_tenor_1.append(pattern.findall(val)[0])
    except: upas_tenor_1.append(None)
    try: upas_tenor_2.append(pattern.findall(val)[1])
    except: upas_tenor_2.append(None)

    # text
    text = get_data_btn(extract_text(f), "GOODS:", "DOCUMENTS REQUIRED:")
    text = re.sub("Page:\d+\s/\s\d+", " ", text)
    text = re.sub(" +", " ", text).split("\n")

    # material name
    material = []
    for t in text: 
        pattern = re.compile("^(?!\s*QUAN).+AT THE RATE")
        vals = pattern.findall(t)
        for v in vals: material.append(v[0:-11].strip())
    for t in text:
        if len(material) > 0: break
        pattern = re.compile("\+.+")
        vals = pattern.findall(t)
        for v in vals: material.append(v.strip())
    mat = ''
    for m in material:
        pattern = re.compile("[A-Z0-9].+")
        mat = mat + pattern.findall(m)[0] + ', '
    material_name.append(mat[:-2])

    # accumulate
    df_hsbc = pd.DataFrame()
    df_hsbc['pdf_name'] = pdf_name
    df_hsbc['dc_no'] = dc_no
    df_hsbc['dc_curr'] = dc_curr
    df_hsbc['dc_amt'] = dc_amt
    df_hsbc['beneficiary'] = beneficiary
    df_hsbc['issue_date'] = issue_date
    df_hsbc['issue_date_refined'] = [datetime.strptime(idate, "%y%m%d").strftime("%d-%b-%y") for idate in issue_date]
    df_hsbc['insurance_no'] = insurance_no
    df_hsbc['insurance_date'] = insurance_date
    df_hsbc['insurance_date_refined'] = [datetime.strptime(idate, "%d%b%Y").strftime("%d-%b-%y") for idate in insurance_date]
    df_hsbc['material_name'] = material_name
    df_hsbc['payment_term'] = payment_term
    df_hsbc['bb_ref'] = bb_ref
    df_hsbc['hs_code_importer'] = hs_code_importer
    df_hsbc['hs_code_exporter'] = hs_code_exporter
    df_hsbc['hs_code'] = hs_code
    df_hsbc['upas_tenor_1'] = upas_tenor_1
    df_hsbc['upas_tenor_2'] = upas_tenor_2
    df_hsbc['inv_no'] = inv_no
    df_hsbc['bank'] = 'HSBC'

    # return 
    return df_hsbc


# In[4]:


# BBL
def parse_bbl(f):

    # fetch datapoint
    def get_data_btn(text_str, sub1, sub2):
        text_str = text_str.replace(sub1, "*").replace(sub2, "*")
        datapoint = text_str.split("*")[1].strip()
        return datapoint

    # scrape
    pdf_name = []
    dc_no = []
    dc_curr = []
    dc_amt = []
    beneficiary = []
    issue_date = []
    insurance_no = []
    insurance_date = []
    material_name = []
    payment_term = []
    bb_ref = []
    inv_no = []
    hs_code_importer = []
    hs_code_exporter = []
    hs_code = []
    upas_tenor_1 = []
    upas_tenor_2 = []
    
    # doc name
    val = f.split("\\")[-1]
    print("Parsing: " + val)
    pdf_name.append(val)

    # all text
    text = extract_text(f)
    
    # dc no.
    val = get_data_btn(text, "Documentary Credit Number", "Date of Issue")
    dc_no.append(val)
    
    # dc value
    pattern = re.compile("(?:USD|EUR|BDT|INR)\s*[\d\.,]+")
    val = pattern.findall(text)[0]
    # currency
    pattern = re.compile("[A-Z]+")
    dc_curr.append(pattern.findall(val)[0])
    # amount
    pattern = re.compile("[\d\.,]+")
    dc_amt.append(pattern.findall(val)[0])
    
    # beneficiary
    val = get_data_btn(text, "Beneficiary", "Currency Code, Amount")
    pattern = re.compile(".+(?:LTD|LIMITED)")
    try: beneficiary.append(pattern.findall(val)[0])
    except: beneficiary.append(val.split("\n")[0])
    
    # issue date
    val = get_data_btn(text, "Date of Issue", "Applicable Rules")
    issue_date.append(val)
    
    # insurance
    val = get_data_btn(text, "NUMBER AND DATE ", " AND A COPY OF")
    # insurance no.
    pattern = re.compile("UIC[\w/\-\(\s]+[\)0-9]\s")
    insurance_no.append(pattern.findall(val)[0].strip())
    # insurance date
    val = val.split()[-1]
    insurance_date.append(val)
    
    # BB ref.
    val = get_data_btn(text, "BB. TIN       : ", "CC. BIN")
    bb_ref.append(val)
    
    # UPAS tenor
    val = get_data_btn(text, "1.DESPITE THE DC TENOR", "Sender to Receiver Information").split("\n")
    pattern = re.compile("\d+")
    try: upas_tenor_1.append(pattern.findall(val[0])[0])
    except: upas_tenor_1.append(None)
    try: upas_tenor_2.append(pattern.findall(val[1])[0])
    except: upas_tenor_2.append(None)
    
    # invoice no. 
    val = get_data_btn(text, "INVOICE NUMBER", "PRICE/DELIVERY TERMS")
    try: inv_no.append(val.split("DATED")[0].strip("\n "))
    except: inv_no.append(None)
        
    # payment term
    val = get_data_btn(text, "PRICE/DELIVERY TERMS:", "INCOTERMS 2020")
    try: payment_term.append(val)
    except: payment_term.append(None)
        
    # material name
    mat = ""
    val = get_data_btn(text, "Description of Goods and/or Services", "AND OTHER DETAILS").split("\n.\n")
    for v in val: mat = mat + v.split("\n")[0] + ", "
    material_name.append(mat[:-2])
    
    # text
    text = get_data_btn(text, "EE. H.S. CODE", ".\n3. COUNTRY OF ORIGIN,")
    text = text.replace(".", "").split("/")
        
    # HS code 
    code = []
    mode = []
    for t in text:
        # line
        pattern = re.compile("(?:PORT)*.*[0-9]{8}.*(?:PORT)*")
        try: val = pattern.findall(t)[0]
        except: continue
        # code
        pattern = re.compile("[0-9]{8}")
        vals = pattern.findall(val)
        for v in vals: code.append(v)
        # mode
        pattern = re.compile("[A-Z]{2}PORT")
        vals = pattern.findall(val)
        for v in vals: mode.append(v)
        # general
        mode_len = len(mode)
        for i in range(len(code)-len(mode)): 
            if mode_len == 0: mode.append("GENERAL")
            else: mode.append(mode[mode_len-1])
            
    # HS code - import/export/general
    imp = exp = gen = ''
    for i in range(0, len(code)):
        if mode[i] == 'IMPORT': imp = imp + code[i] + ', '
        elif mode[i] == 'EXPORT': exp = exp + code[i] + ', '
        else: gen = gen + code[i] + ', '    
    hs_code.append(gen[:-2])
    hs_code_importer.append(imp[:-2])
    hs_code_exporter.append(exp[:-2])
    
    # accumulate
    df_bbl = pd.DataFrame()
    df_bbl['pdf_name'] = pdf_name
    df_bbl['dc_no'] = dc_no
    df_bbl['dc_curr'] = dc_curr
    df_bbl['dc_amt'] = dc_amt
    df_bbl['beneficiary'] = beneficiary
    df_bbl['issue_date'] = issue_date
    df_bbl['issue_date_refined'] = [datetime.strptime(idate, "%y%m%d").strftime("%d-%b-%y") for idate in issue_date]
    df_bbl['insurance_no'] = insurance_no
    df_bbl['insurance_date'] = insurance_date
    df_bbl['insurance_date_refined'] = [datetime.strptime(idate, "%d-%b-%Y").strftime("%d-%b-%y") for idate in insurance_date]
    df_bbl['material_name'] = material_name
    df_bbl['payment_term'] = payment_term
    df_bbl['bb_ref'] = bb_ref
    df_bbl['hs_code_importer'] = hs_code_importer
    df_bbl['hs_code_exporter'] = hs_code_exporter
    df_bbl['hs_code'] = hs_code
    df_bbl['upas_tenor_1'] = upas_tenor_1
    df_bbl['upas_tenor_2'] = upas_tenor_2
    df_bbl['inv_no'] = inv_no
    df_bbl['bank'] = 'BBL'

    # return 
    return df_bbl


# In[5]:


# SCB
def breach_scb(f, path):

    # password
    pdf = pikepdf.open(f, password="csg@7865", allow_overwriting_input=True)
    pdf.save(path + f.split("\\")[-1])

    # attachments
    doc = fitz.open(path + f.split("\\")[-1])
    name_dict = {}
    for item in doc.embfile_names(): name_dict[item] = doc.embfile_info(item)["filename"]

    # save
    for item, file in name_dict.items():
        if "ADV.pdf" in file:
            with open(path + file, "wb") as outfile: outfile.write(doc.embfile_get(item))
    doc.close()

def parse_scb(f):

    # fetch datapoint
    def get_data_btn(text_str, sub1, sub2):
        text_str = text_str.replace(sub1, "!").replace(sub2, "!")
        datapoint = text_str.split("!")[1].strip()
        return datapoint

    # scrape
    pdf_name = []
    dc_no = []
    dc_curr = []
    dc_amt = []
    beneficiary = []
    issue_date = []
    insurance_no = []
    insurance_date = []
    material_name = []
    payment_term = []
    bb_ref = []
    inv_no = []
    hs_code_importer = []
    hs_code_exporter = []
    hs_code = []
    upas_tenor_1 = []
    upas_tenor_2 = []
    
    # doc name
    name = f.split("\\")[-1]
    print("Parsing: " + name)
    pdf_name.append(name)

    # all text
    text = extract_text(f)

    # dc no.
    val = get_data_btn(text, "CREDIT NUMBER", ":31C")
    dc_no.append(val)

    # dc value
    pattern = re.compile("(?:USD|EUR|BDT|INR)\s*[\d\.,]+")
    val = pattern.findall(text)[0]
    # currency
    pattern = re.compile("[A-Z]+")
    dc_curr.append(pattern.findall(val)[0])
    # amount
    pattern = re.compile("[\d\.,]+")
    dc_amt.append(pattern.findall(val)[0])

    # beneficiary
    val = get_data_btn(text, ":59:      BENEFICIARY", ":32B:").split("\n")[0]
    beneficiary.append(val)

    # issue date
    val = get_data_btn(text, "DATE OF ISSUE", ":40E")
    issue_date.append(val)

    # insurance
    val = "UIC" + get_data_btn(text, "UIC", "\n.\n+ ")
    # insurance no.
    pattern = re.compile("UIC[\w/\-\(\s]+[\)0-9]\s")
    try: insurance_no.append(pattern.findall(val)[0].strip())
    except: insurance_no.append(None)
    # insurance date
    pattern = re.compile("[\d]{2}\.[\d]{2}\.[\d]{4}")
    try: insurance_date.append(pattern.findall(val)[0])
    except: insurance_date.append(None)

    # material name
    val = get_data_btn(text, "OR SERVICES\n+", "QUANTITY").replace("\n", " ")
    material_name.append(val)

    # payment term
    val = get_data_btn(text, "INCOTERMS ", ":46A").replace("\n", " ")
    payment_term.append(val)

    # BB ref.
    val = get_data_btn(text, "DC REFERENCE NUMBER:", " AND LC")
    bb_ref.append(val)

    # invoice no. 
    inv = ''
    val = get_data_btn(text, "46A:", "APPLICANT'S BIN")
    pattern = re.compile("[\S]+\sDATED")
    vals = pattern.findall(val)
    for v in vals: inv = inv + v[:-6] + ', '
    inv_no.append(inv[:-2])

    # UPAS tenor
    try: val = get_data_btn(text, "TENOR BEING", "COMPLYING DOCUMENTS")
    except: val = ""
    pattern = re.compile("[0-9]+")
    try: upas_tenor_1.append(pattern.findall(val)[0])
    except: upas_tenor_1.append(None)
    try: upas_tenor_2.append(pattern.findall(val)[1])
    except: upas_tenor_2.append(None)

    # text
    text = get_data_btn(text, "46A:", "INSURANCE COVER NOTE")
    text = text.replace(".", "").split("\n")

    # HS code 
    code = []
    mode = []
    for t in text:
        # line
        pattern = re.compile("(?:PORT)*.*CODE.*[0-9]{8}.*(?:PORT)*")
        try: val = pattern.findall(t)[0]
        except: continue
        # code
        pattern = re.compile("[0-9]{8}")
        vals = pattern.findall(val)
        for v in vals: code.append(v)
        # mode
        pattern = re.compile("[A-Z]{2}PORT")
        vals = pattern.findall(val)
        for v in vals: mode.append(v)
        # general
        mode_len = len(mode)
        for i in range(len(code)-len(mode)): 
            if mode_len == 0: mode.append("GENERAL")
            else: mode.append(mode[mode_len-1])

    # HS code - import/export/general
    imp = ''
    exp = ''
    gen = ''
    for i in range(0, len(code)):
        if mode[i] == 'IMPORT': imp = imp + code[i] + ', '
        elif mode[i] == 'EXPORT': exp = exp + code[i] + ', '
        else: gen = gen + code[i] + ', '    
    hs_code.append(gen[:-2])
    hs_code_importer.append(imp[:-2])
    hs_code_exporter.append(exp[:-2])

    # accumulate
    df_scb = pd.DataFrame()
    df_scb['pdf_name'] = pdf_name
    df_scb['dc_no'] = dc_no
    df_scb['dc_curr'] = dc_curr
    df_scb['dc_amt'] = dc_amt
    df_scb['beneficiary'] = beneficiary
    df_scb['issue_date'] = issue_date
    df_scb['issue_date_refined'] = [datetime.strptime(idate, "%y%m%d").strftime("%d-%b-%y") for idate in issue_date]
    df_scb['insurance_no'] = insurance_no
    df_scb['insurance_date'] = insurance_date
    df_scb['insurance_date_refined'] = [datetime.strptime(idate, "%d.%m.%Y").strftime("%d-%b-%y") for idate in insurance_date]
    df_scb['material_name'] = material_name
    df_scb['payment_term'] = payment_term
    df_scb['bb_ref'] = bb_ref
    df_scb['hs_code_importer'] = hs_code_importer
    df_scb['hs_code_exporter'] = hs_code_exporter
    df_scb['hs_code'] = hs_code
    df_scb['upas_tenor_1'] = upas_tenor_1
    df_scb['upas_tenor_2'] = upas_tenor_2
    df_scb['inv_no'] = inv_no
    df_scb['bank'] = 'SCB'

    # return 
    return df_scb


# In[6]:


# PRB
def parse_prb(f):

    # fetch datapoint
    def get_data_btn(text_str, sub1, sub2):
        text_str = text_str.replace(sub1, "*").replace(sub2, "*")
        datapoint = text_str.split("*")[1].strip()
        return datapoint

    # scrape
    pdf_name = []
    dc_no = []
    dc_curr = []
    dc_amt = []
    beneficiary = []
    issue_date = []
    insurance_no = []
    insurance_date = []
    material_name = []
    payment_term = []
    bb_ref = []
    inv_no = []
    hs_code_importer = []
    hs_code_exporter = []
    hs_code = []
    upas_tenor_1 = []
    upas_tenor_2 = []
    
    # doc name
    val = f.split("\\")[-1]
    print("Parsing: " + val)
    pdf_name.append(val)

    # all text
    text = extract_text(f)
    collapsed_text = re.sub(" +", " ", text)
    compressed_text = text
    for s in["\n", " +"]: compressed_text = re.sub(s, " ", compressed_text)
    
    # dc no.
    val = get_data_btn(collapsed_text, "Credit Number : ", ":31C/Date of Issue")
    dc_no.append(val)
    
    # dc value
    pattern = re.compile("(?:USD|EUR|BDT|INR)\s*[\d\.,]+")
    val = pattern.findall(text)[0]
    # currency
    pattern = re.compile("[A-Z]+")
    dc_curr.append(pattern.findall(val)[0])
    # amount
    pattern = re.compile("[\d\.,]+")
    dc_amt.append(pattern.findall(val)[0])
    
    # beneficiary
    val = get_data_btn(collapsed_text, ":59/Beneficiary : ", ":32B/Currency Code, Amount")
    pattern = re.compile("[\w\W]+(?:LTD|LIMITED)")
    try: beneficiary.append(pattern.findall(val)[0])
    except: beneficiary.append(val.split("\n")[0])
    
    # issue date
    val = get_data_btn(collapsed_text, ":31C/Date of Issue : ", ":40E/Applicable Rules")
    issue_date.append(val)
    
    # insurance
    pattern = re.compile("UIC[\w/\-\(\s]+[\)0-9]\sDATED\s[\w\W]{10}")
    val = pattern.findall(compressed_text)[0].split(" DATED ")
    # insurance no.
    insurance_no.append(val[0])
    # insurance date
    insurance_date.append(val[1])
    
    # BB ref.
    bb_ref.append(None)
    
    # UPAS tenor
    val = get_data_btn(collapsed_text, ":42C/Drafts at … : ", ":42D/Drawee")
    pattern = re.compile("\d+")
    try: upas_tenor_1.append(pattern.findall(val)[0])
    except: upas_tenor_1.append(None)
    try: upas_tenor_2.append(pattern.findall(val)[1])
    except: upas_tenor_2.append(None)
    
    # invoice no. 
    val = get_data_btn(compressed_text, ":45A/Description of Goods and/or Service : ", ":46A/Documents Required")
    val = get_data_btn(val, "PROFORMA INVOICE NO.", "DATED")
    inv_no.append(val)
        
    # payment term
    val = get_data_btn(compressed_text, "TRADE TERM: ", ".. DESCRIPTION, QUALITY, QUANTITY")
    try: payment_term.append(val)
    except: payment_term.append(None)
        
    # material name
    mat = ""
    val = get_data_btn(compressed_text, ":45A/Description of Goods and/or Service : ", "TRADE TERM: ").split("..")[0:-1]
    for v in val: mat = mat + v.split(" QUANTITY")[0] + ", "
    material_name.append(mat[:-2])
    
    # text
    text = get_data_btn(collapsed_text, "B) ", "C) ")
    text = text.replace(".", "").split("\n")
    
    # HS code 
    code = []
    mode = []
    for t in text:
        # line
        pattern = re.compile("(?:PORT)*.*[0-9]{8}.*(?:PORT)*")
        try: val = pattern.findall(t)[0]
        except: continue
        # code
        pattern = re.compile("[0-9]{8}")
        vals = pattern.findall(val)
        for v in vals: code.append(v)
        # mode
        pattern = re.compile("[A-Z]{2}PORT")
        vals = pattern.findall(val)
        for v in vals: mode.append(v)
        # general
        mode_len = len(mode)
        for i in range(len(code)-len(mode)): 
            if mode_len == 0: mode.append("GENERAL")
            else: mode.append(mode[mode_len-1])
            
    # HS code - import/export/general
    imp = exp = gen = ''
    for i in range(0, len(code)):
        if mode[i] == 'IMPORT': imp = imp + code[i] + ', '
        elif mode[i] == 'EXPORT': exp = exp + code[i] + ', '
        else: gen = gen + code[i] + ', '    
    hs_code.append(gen[:-2])
    hs_code_importer.append(imp[:-2])
    hs_code_exporter.append(exp[:-2])
    
    # accumulate
    df_prb = pd.DataFrame()
    df_prb['pdf_name'] = pdf_name
    df_prb['dc_no'] = dc_no
    df_prb['dc_curr'] = dc_curr
    df_prb['dc_amt'] = dc_amt
    df_prb['beneficiary'] = beneficiary
    df_prb['issue_date'] = issue_date
    df_prb['issue_date_refined'] = ''
    df_prb['insurance_no'] = insurance_no
    df_prb['insurance_date'] = insurance_date
    df_prb['material_name'] = material_name
    df_prb['payment_term'] = payment_term
    df_prb['bb_ref'] = bb_ref
    df_prb['hs_code_importer'] = hs_code_importer
    df_prb['hs_code_exporter'] = hs_code_exporter
    df_prb['hs_code'] = hs_code
    df_prb['upas_tenor_1'] = upas_tenor_1
    df_prb['upas_tenor_2'] = upas_tenor_2
    df_prb['inv_no'] = inv_no
    df_prb['bank'] = 'PRB'

    # return 
    return df_prb


# In[7]:


# LCs
fetch_read_lc('2023-09-01', '2032-12-31')


# In[8]:


# accumulate HSBC
inc_docs = ""
df_hsbc = pd.DataFrame()
start_time = time.time()

# parse HSBC
files = glob.glob(r"C:/Users/Shithi.Maitra/Unilever Codes/Ad Hoc/PR Prioritization Procurement/Emailed LCs/HSBC LCs/*-T02.PDF")
for f in reversed(files):
    try: df_hsbc = df_hsbc.append(parse_hsbc(f))
    except: inc_docs = inc_docs + f.split("\\")[-1] + ", "
df_hsbc = df_hsbc.reset_index(drop=True)
display(df_hsbc)

# analyse HSBC
email_hsbc_df = pd.DataFrame()
email_hsbc_df['Bank Name'] = ['HSBC']
email_hsbc_df['LCs Received'] = [df_hsbc.shape[0] + inc_docs.count(",")]
email_hsbc_df['LCs Parsed'] = [df_hsbc.shape[0]]
email_hsbc_df['LCs Incomplete'] = [inc_docs.count(",")]
email_hsbc_df['Incomplete LC Docs'] = [",".join(inc_docs.split(",", 3)[:3]) + ", ..."]
email_hsbc_df['Sec to Parse'] = [round(time.time() - start_time)]


# In[9]:


# breach SCB
files = glob.glob(r"C:/Users/Shithi.Maitra/Unilever Codes/Ad Hoc/PR Prioritization Procurement/Emailed LCs/SCB LCs/*.pdf")
for f in reversed(files): breach_scb(f, "C:/Users/Shithi.Maitra/Unilever Codes/Ad Hoc/PR Prioritization Procurement/Emailed LCs/SCB LCs/SCB Breached LCs/")

# accumulate SCB
inc_docs = ""
df_scb = pd.DataFrame()
start_time = time.time()
    
# parse SCB
files = glob.glob(r"C:/Users/Shithi.Maitra/Unilever Codes/Ad Hoc/PR Prioritization Procurement/Emailed LCs/SCB LCs/SCB Breached LCs/*ADV.pdf")
for f in reversed(files):
    try: df_scb = df_scb.append(parse_scb(f))
    except: inc_docs = inc_docs + f.split("\\")[-1] + ", "
df_scb = df_scb.reset_index(drop=True)
display(df_scb)

# analyse SCB
email_scb_df = pd.DataFrame()
email_scb_df['Bank Name'] = ['SCB']
email_scb_df['LCs Received'] = [df_scb.shape[0] + inc_docs.count(",")]
email_scb_df['LCs Parsed'] = [df_scb.shape[0]]
email_scb_df['LCs Incomplete'] = [inc_docs.count(",")]
email_scb_df['Incomplete LC Docs'] = [",".join(inc_docs.split(",", 3)[:3]) + ", ..."]
email_scb_df['Sec to Parse'] = [round(time.time() - start_time)]


# In[10]:


# accumulate BBL
inc_docs = ""
df_bbl = pd.DataFrame()
start_time = time.time()

# parse BBL
files = glob.glob(r"C:/Users/Shithi.Maitra/Unilever Codes/Ad Hoc/PR Prioritization Procurement/Emailed LCs/BBL LCs/*SWIFT*.PDF")
for f in reversed(files):
    try: df_bbl = df_bbl.append(parse_bbl(f))
    except: inc_docs = inc_docs + f.split("\\")[-1] + ", "
df_bbl = df_bbl.reset_index(drop=True)
display(df_bbl)

# analyse BBL
email_bbl_df = pd.DataFrame()
email_bbl_df['Bank Name'] = ['BBL']
email_bbl_df['LCs Received'] = [df_bbl.shape[0] + inc_docs.count(",")]
email_bbl_df['LCs Parsed'] = [df_bbl.shape[0]]
email_bbl_df['LCs Incomplete'] = [inc_docs.count(",")]
email_bbl_df['Incomplete LC Docs'] = [",".join(inc_docs.split(",", 3)[:3]) + ", ..."]
email_bbl_df['Sec to Parse'] = [round(time.time() - start_time)]


# In[11]:


# accumulate PRB
inc_docs = ""
df_prb = pd.DataFrame()
start_time = time.time()

# parse PRB
files = glob.glob(r"C:/Users/Shithi.Maitra/Unilever Codes/Ad Hoc/PR Prioritization Procurement/Emailed LCs/PRB LCs/*700.pdf")
for f in reversed(files):
    try: df_prb = df_prb.append(parse_prb(f))
    except: inc_docs = inc_docs + f.split("\\")[-1] + ", "
df_prb = df_prb.reset_index(drop=True)
display(df_prb)

# analyse PRB
email_prb_df = pd.DataFrame()
email_prb_df['Bank Name'] = ['PRB']
email_prb_df['LCs Received'] = [df_prb.shape[0] + inc_docs.count(",")]
email_prb_df['LCs Parsed'] = [df_prb.shape[0]]
email_prb_df['LCs Incomplete'] = [inc_docs.count(",")]
email_prb_df['Incomplete LC Docs'] = [",".join(inc_docs.split(",", 3)[:3]) + ", ..."]
email_prb_df['Sec to Parse'] = [round(time.time() - start_time)]


# In[12]:


# OP

# file
qry = '''select * from df_hsbc union all select * from df_bbl union all select * from df_scb /*union all select * from df_prb*/'''
df = duckdb.query(qry).df()
df.to_excel("LCs_parsed_test.xlsx", index=False)

# analysis
qry = '''select * from email_hsbc_df union all select * from email_scb_df union all select * from email_bbl_df /*union all select * from email_prb_df*/'''
email_df = duckdb.query(qry).df()
display(email_df)


# In[13]:


# email
ol = win32com.client.Dispatch("outlook.application")
olmailitem = 0x0
newmail = ol.CreateItem(olmailitem)

# subject, recipients
newmail.Subject = 'Parsed LCs in Test'
# newmail.To = 'shithi.maitra@unilever.com'
newmail.To = 'ubl_tradealert@Unilever.com'
newmail.CC = 'mehedi.asif@unilever.com; asif.rezwan@unilever.com; md-ashiqur.akhand@unilever.com; shaik.hossen@unilever.com'

# body
newmail.HTMLbody = f'''
Dear concern,<br><br>
The service of automated parsing of datapoints from LCs is now developed and under testing. Please find (and verify) attached results from LCs shared <a href="mailto: ubl_tradealert@Unilever.com">@ubl_tradealert</a>.
''' + build_table(email_df, random.choice(['green_light', 'red_light', 'blue_light', 'grey_light', 'orange_light']), font_size='10px', text_align='left') + '''
Note that, the statistics presented above reflect LCs shared since 26-Sep-2023. More banks will be added to the test eventually. This is an auto email via <i>win32com</i>.<br><br>
Thanks,<br>
Shithi Maitra<br>
Asst. Manager, Cust. Service Excellence<br>
Unilever BD Ltd.<br>
'''

# attachment(s) 
folder = r"C:/Users/Shithi.Maitra/Unilever Codes/Ad Hoc/PR Prioritization Procurement/"
filename = folder + "LCs_parsed_test.xlsx"
newmail.Attachments.Add(filename)

# send
newmail.Send()
