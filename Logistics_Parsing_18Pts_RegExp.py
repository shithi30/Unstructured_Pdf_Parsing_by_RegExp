#!/usr/bin/env python
# coding: utf-8

# In[3]:


# HSBC

# import
import glob
import re
import pandas as pd
from pdfminer.high_level import extract_text

# fetch datapoint
def get_data_btn(text_str, sub1, sub2):
    text_str = text_str.replace(sub1, "*").replace(sub2, "*")
    datapoint = text_str.split("*")[1].strip()
    return datapoint

# files
files = glob.glob("*-T02.PDF")

# scrape
pdf_name = []
dc_no = []
dc_curr = []
dc_amt = []
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
applicant_name = []
applicant_address = []
consignee_bin = []
beneficiary_name = []
beneficiary_address = []
irc_no = []
port_load = []
port_discharge = []
for f in files:
    
    # doc name
    print("Parsing: " + f)
    pdf_name.append(f)

    # all text
    text_orig = extract_text(f)
    text = re.sub("Page:\d+\s/\s\d+", " ", text_orig)
    
    # applicant name
    vals = get_data_btn(text, "APPLICANT: ", "BENEFICIARY: ").split("\n")
    applicant_name.append(vals[0])
    
    # applicant address
    val = ""
    for v in vals[1:]: val = val + v
    val = re.sub(" +", " ", val).strip()
    applicant_address.append(val)
    
    # beneficiary name
    vals = get_data_btn(text, "BENEFICIARY: ", "DC AMT:").split("\n")
    beneficiary_name.append(vals[0])
    
    # beneficiary address
    val = ""
    for v in vals[1:]: val = val + v
    val = re.sub(" +", " ", val).strip()
    beneficiary_address.append(val)

    # modified text
    text = text.replace("\n", " ")
    text = re.sub(" +", " ", text)

    # dc no.
    val = get_data_btn(text, "DC NO:", "DATE OF ISSUE:")
    dc_no.append(val)
    
    # BIN
    pattern = re.compile("E-BIN.+\d{9}[\-\s]+\d{4}")
    try: consignee_bin.append(pattern.findall(text)[0][7:])
    except: consignee_bin.append(None)
    
    # IRC
    pattern = re.compile("IRC NO[\.\s\:]+\d+, E-TIN")
    try: val = pattern.findall(text)[0]
    except: val = None
    pattern = re.compile("\d+")
    try: irc_no.append(pattern.findall(val)[0])
    except: irc_no.append(None)
        
    # port of loading
    try: val = get_data_btn(text, "DEPART AIRPORT: ", "DISCHARGE PORT/")
    except: val = None
    port_load.append(val)
    
    # port of discharge
    try: val = get_data_btn(text, "DEST AIRPORT:", "LATEST DATE OF SHIPMENT:")
    except: val = None
    port_discharge.append(val)
    
    # dc value
    val = get_data_btn(text, "DC AMT: ", " AVAILABLE WITH/BY")
    # currency
    pattern = re.compile("^[A-Z]*")
    dc_curr.append(pattern.findall(val)[0])
    # amount
    pattern = re.compile("[0-9,]+")
    dc_amt.append(pattern.findall(val)[0].strip())
    
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
    text = get_data_btn(text_orig, "GOODS:", "DOCUMENTS REQUIRED:")
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
    mat = ""
    for m in material:
        pattern = re.compile("[A-Z0-9].+")
        mat = mat + pattern.findall(m)[0] + ", "
    material_name.append(mat[:-2])
    
# accumulate
df_hsbc = pd.DataFrame()
df_hsbc["pdf_name"] = pdf_name
df_hsbc["dc_no"] = dc_no
df_hsbc["dc_curr"] = dc_curr
df_hsbc["dc_amt"] = dc_amt
df_hsbc["issue_date"] = issue_date
df_hsbc["insurance_no"] = insurance_no
df_hsbc["insurance_date"] = insurance_date
df_hsbc["material_name"] = material_name
df_hsbc["payment_term"] = payment_term
df_hsbc["bb_ref"] = bb_ref
df_hsbc["hs_code_importer"] = hs_code_importer
df_hsbc["hs_code_exporter"] = hs_code_exporter
df_hsbc["hs_code"] = hs_code
df_hsbc["upas_tenor_1"] = upas_tenor_1
df_hsbc["upas_tenor_2"] = upas_tenor_2
df_hsbc["inv_no"] = inv_no
df_hsbc["applicant_name"] = applicant_name
df_hsbc["applicant_address"] = applicant_address
df_hsbc["consignee_bin"] = consignee_bin
df_hsbc["beneficiary_name"] = beneficiary_name
df_hsbc["beneficiary_address"] = beneficiary_address
df_hsbc["irc_no"] = irc_no
df_hsbc["port_load"] = port_load
df_hsbc["port_discharge"] = port_discharge
df_hsbc["bank"] = "HSBC"

# show
display(df_hsbc)

# save
df_hsbc.to_excel("C:/Users/Shithi.Maitra/Downloads/LC_data_logistics.xlsx", index=False)


# In[ ]:





# In[ ]:





# In[ ]:




