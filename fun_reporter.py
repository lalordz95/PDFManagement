import os
from fillpdf import fillpdfs
import PyPDF2
import re
from collections import Counter
import xlsxwriter
import glob
import traceback
from PyQt5 import QtCore
import subprocess

def init_conect(self): #Definitions of the window of the interface
    
    self.setWindowFlags(QtCore.Qt.WindowMinimizeButtonHint)
    self.setWindowFlag(QtCore.Qt.WindowCloseButtonHint, True)
    self.b_selectfolder.clicked.connect(self.bt_selectf)  
    self.b_runreport.clicked.connect(self.bt_runreporter)

def replacer(some_dict): #Not answered items in the dictionary come as blank strings, replaces "" for "Off"
    return { k: ('Off' if v is None else v) for k, v in some_dict.items() }

def report_ass(PDFfile_name, numerador, pdf_name_path, row, workbook, worksheet, path_forms): #main process
    
    while True: #reads the PDF file to identify the assessment
        #MINI1
        if PDFfile_name.find('AD.pdf') != -1: #PDFfile variable stores the name of the current file
            current_ass = "AD"
            break
        elif PDFfile_name.find('AMI.pdf') != -1:
            current_ass = "AMI"
            break
        elif PDFfile_name.find('ASI3.pdf') != -1:
            current_ass = "ASI3"
            break
        elif PDFfile_name.find('CAOPE.pdf') != -1:
            current_ass = "CAOPE"
            break
        elif PDFfile_name.find('CFSFI6.pdf') != -1:
            current_ass = "CFSFI6"
            break
        elif PDFfile_name.find('CR-DPSS.pdf') != -1:
            current_ass = "CR-DPSS"
            break
        elif PDFfile_name.find('CR-MADRS.pdf') != -1:
            current_ass = "CR-MADRS"
            break
        elif PDFfile_name.find('CSRH.pdf') != -1:
            current_ass = "CSRH"
            break
        elif PDFfile_name.find('CSSRS.pdf') != -1:
            current_ass = "CSSRS"
            break
        elif PDFfile_name.find('DMC.pdf') != -1:
            current_ass = "DMC"
            break
        elif PDFfile_name.find('DSLF.pdf') != -1:
            current_ass = "DSLF"
            break
        elif PDFfile_name.find('GADD.pdf') != -1:
            current_ass = "GADD"
            break
        elif PDFfile_name.find('GADS.pdf') != -1:
            current_ass = "GADS"
            break
        elif PDFfile_name.find('GD.pdf') != -1:
            current_ass = "GD"
            break
        elif PDFfile_name.find('IPGDS.pdf') != -1:
            current_ass = "IPGDS"
            break
        elif PDFfile_name.find('L2-ASRMS.pdf') != -1:
            current_ass = "L2-ASRMS"
            break
        elif PDFfile_name.find('L2-RTB.pdf') != -1:
            current_ass = "L2-RTB"
            break
        elif PDFfile_name.find('MANIC.pdf') != -1:
            current_ass = "MANIC"
            break
        elif PDFfile_name.find('MDDD.pdf') != -1:
            current_ass = "MDDD"
            break
        elif PDFfile_name.find('MECSS.pdf') != -1:
            current_ass = "MECSS"
            break
        elif PDFfile_name.find('MEDDC.pdf') != -1:
            current_ass = "MEDDC"
            break
        elif PDFfile_name.find('MIS.pdf') != -1:
            current_ass = "MIS"
            break
        elif PDFfile_name.find('OCD-BDD-SR.pdf') != -1:
            current_ass = "OCD-BDD-SR"
            break
        elif PDFfile_name.find('OCD-SESPDDC.pdf') != -1:
            current_ass = "OCD-SESPDDC"
            break
        elif PDFfile_name.find('OCD-SHDDC.pdf') != -1:
            current_ass = "OCD-SHDDC"
            break
        elif PDFfile_name.find('OCD-STHPDDC.pdf') != -1:
            current_ass = "OCD-STHPDDC"
            break
        elif PDFfile_name.find('PABS.pdf') != -1:
            current_ass = "PABS"
            break
        elif PDFfile_name.find('PANHS.pdf') != -1:
            current_ass = "PANHS"
            break
        elif PDFfile_name.find('PAS.pdf') != -1:
            current_ass = "PAS"
            break
        elif PDFfile_name.find('AS.pdf') != -1:
            current_ass = "AS"
            break
        elif PDFfile_name.find('PCGFL.pdf') != -1:
            current_ass = "PCGFL"
            break
        elif PDFfile_name.find('PDC.pdf') != -1:
            current_ass = "PDC"
            break
        elif PDFfile_name.find('PDDD.pdf') != -1:
            current_ass = "PDDD"
            break
        elif PDFfile_name.find('PDED.pdf') != -1:
            current_ass = "PDED"
            break
        elif PDFfile_name.find('PDMDWPF.pdf') != -1:
            current_ass = "PDMDWPF"
            break
        elif PDFfile_name.find('PDS.pdf') != -1:
            current_ass = "PDS"
            break
        elif PDFfile_name.find('PSST.pdf') != -1:
            current_ass = "PSST"
            break
        elif PDFfile_name.find('RLSS.pdf') != -1:
            current_ass = "RLSS"
            break
        elif PDFfile_name.find('RMS.pdf') != -1:
            current_ass = "RMS"
            break
        elif PDFfile_name.find('RSANHS.pdf') != -1:
            current_ass = "RSANHS"
            break
        elif PDFfile_name.find('SAFEMH.pdf') != -1:
            current_ass = "SAFEMH"
            break
        elif PDFfile_name.find('SEADD.pdf') != -1:
            current_ass = "SEADD"
            break
        elif PDFfile_name.find('SEADS.pdf') != -1:
            current_ass = "SEADS"
            break
        elif PDFfile_name.find('SF-L1-CCSM') != -1:
            current_ass = "SF-L1-CCSM"
            break
        elif PDFfile_name.find('SHPS.pdf') != -1:
            current_ass = "SHPS"
            break
        elif PDFfile_name.find('SOADD.pdf') != -1:
            current_ass = "SOADD"
            break
        elif PDFfile_name.find('SOADS.pdf') != -1:
            current_ass = "SOADS"
            break
        elif PDFfile_name.find('SOALS.pdf') != -1:
            current_ass = "SOALS"
            break
        elif PDFfile_name.find('SPHD.pdf') != -1:
            current_ass = "SPHD"
            break
        elif PDFfile_name.find('SPHS.pdf') != -1:
            current_ass = "SPHS"
            break
        elif PDFfile_name.find('SR-MADRS.pdf') != -1: 
            current_ass = "SR-MADRS"
            break
        elif PDFfile_name.find('CDRS.pdf') != -1:
            current_ass = "CDRS"
            break
        elif PDFfile_name.find('DRS.pdf') != -1:
            current_ass = "DRS"
            break
        elif PDFfile_name.find('SR-YBOCS-II.pdf') != -1:
            current_ass = "SR-YBOCS-II"
            break
        elif PDFfile_name.find('SSTS.pdf') != -1:
            current_ass = "SSTS"
            break
        elif PDFfile_name.find('UCLA.pdf') != -1:
            current_ass = "UCLA"
            break
        elif PDFfile_name.find('Y-BOCS-SC.pdf') != -1:
            current_ass = "Y-BOCS-SC"
            break
        #MINI2
        elif PDFfile_name.find('ACIPS.pdf') != -1:
            current_ass = "ACIPS"
            break
        elif PDFfile_name.find('ADQ.pdf') != -1:
            current_ass = "ADQ"
            break
        elif PDFfile_name.find('AN.pdf') != -1:
            current_ass = "AN"
            break
        elif PDFfile_name.find('ARM-R.pdf') != -1:
            current_ass = "ARM-R"
            break
        elif PDFfile_name.find('ASSF.pdf') != -1:
            current_ass = "ASSF"
            break
        elif PDFfile_name.find('BED.pdf') != -1:
            current_ass = "BED"
            break
        elif PDFfile_name.find('BFIS-I.pdf') != -1:
            current_ass = "BFIS-I"
            break
        elif PDFfile_name.find('BFIS-SR.pdf') != -1:
            current_ass = "BFIS-SR"
            break
        elif PDFfile_name.find('BN.pdf') != -1:
            current_ass = "BN"
            break
        elif PDFfile_name.find('BPM.pdf') != -1:
            current_ass = "BPM"
            break
        elif PDFfile_name.find('BPQ-SF.pdf') != -1:
            current_ass = "BPQ-SF"
            break
        elif PDFfile_name.find('CA.pdf') != -1:
            current_ass = "CA"
            break
        elif PDFfile_name.find('CATI.pdf') != -1:
            current_ass = "CATI"
            break
        elif PDFfile_name.find('CATQ.pdf') != -1:
            current_ass = "CATQ"
            break
        elif PDFfile_name.find('CBS.pdf') != -1:
            current_ass = "CBS"
            break
        elif PDFfile_name.find('CR-SNSI.pdf') != -1:
            current_ass = "CR-SNSI"
            break
        elif PDFfile_name.find('CR-SSSD.pdf') != -1:
            current_ass = "CR-SSSD"
            break
        elif PDFfile_name.find('CSI.pdf') != -1:
            current_ass = "CSI"
            break
        elif PDFfile_name.find('DAACI.pdf') != -1:
            current_ass = "DAACI"
            break
        elif PDFfile_name.find('DBMQ.pdf') != -1:
            current_ass = "DBMQ"
            break
        elif PDFfile_name.find('DCDD.pdf') != -1:
            current_ass = "DCDD"
            break
        elif PDFfile_name.find('DES-II.pdf') != -1:
            current_ass = "DES-II"
            break
        elif PDFfile_name.find('DSPS.pdf') != -1:
            current_ass = "DSPS"
            break
        elif PDFfile_name.find('SPS.pdf') != -1:
            current_ass = "SPS"
            break
        elif PDFfile_name.find('DST.pdf') != -1:
            current_ass = "DST"
            break
        elif PDFfile_name.find('DTD.pdf') != -1:
            current_ass = "DTD"
            break
        elif PDFfile_name.find('DTS.pdf') != -1:
            current_ass = "DTS"
            break
        elif PDFfile_name.find('EDAQA.pdf') != -1:
            current_ass = "EDAQA"
            break
        elif PDFfile_name.find('EDEQ.pdf') != -1:
            current_ass = "EDEQ"
            break
        elif PDFfile_name.find('ESS.pdf') != -1:
            current_ass = "ESS"
            break
        elif PDFfile_name.find('GQASC.pdf') != -1:
            current_ass = "GQASC"
            break
        elif PDFfile_name.find('HQ.pdf') != -1:
            current_ass = "HQ"
            break
        elif PDFfile_name.find('IEDD.pdf') != -1:
            current_ass = "IEDD"
            break
        elif PDFfile_name.find('IOPF.pdf') != -1:
            current_ass = "IOPF"
            break
        elif PDFfile_name.find('ITEM.pdf') != -1:
            current_ass = "ITEM"
            break
        elif PDFfile_name.find('ITQ.pdf') != -1:
            current_ass = "ITQ"
            break
        elif PDFfile_name.find('L2-SD.pdf') != -1:
            current_ass = "L2-SD"
            break
        elif PDFfile_name.find('L2-SS.pdf') != -1:
            current_ass = "L2-SS"
            break
        elif PDFfile_name.find('L2-SU.pdf') != -1:
            current_ass = "L2-SU"
            break
        elif PDFfile_name.find('LEC5S.pdf') != -1:
            current_ass = "LEC5S"
            break
        elif PDFfile_name.find('MAIA-2.pdf') != -1:
            current_ass = "MAIA-2"   
            break
        elif PDFfile_name.find('MPA.pdf') != -1:
            current_ass = "MPA"
            break
        elif PDFfile_name.find('MSI-BPD.pdf') != -1:
            current_ass = "MSI-BPD"
            break
        elif PDFfile_name.find('MSQ.pdf') != -1:
            current_ass = "MSQ"
            break
        elif PDFfile_name.find('NSCS-SF.pdf') != -1:
            current_ass = "NSCS-SF"
            break
        elif PDFfile_name.find('PAAQ.pdf') != -1: #1 Primero busca PAAQ, GPAQ, PAQ y AQ
            current_ass = "PAAQ"
            break
        elif PDFfile_name.find('GPAQ.pdf') != -1:
            current_ass = "GPAQ"
            break
        elif PDFfile_name.find('PAQ.pdf') != -1:
            current_ass = "PAQ"
            break
        elif  PDFfile_name.find('AQ.pdf') != -1:
            current_ass = "AQ"
            break
        elif PDFfile_name.find('PADUQ.pdf') != -1:
            current_ass = "PADUQ"
            break
        elif PDFfile_name.find('PDS-B.pdf') != -1:
            current_ass = "PDS-B"
            break
        elif PDFfile_name.find('PFRS.pdf') != -1:
            current_ass = "PFRS"
            break
        elif PDFfile_name.find('PHQ-15.pdf') != -1:
            current_ass = "PHQ-15"
            break
        elif PDFfile_name.find('PRBQ.pdf') != -1:
            current_ass = "PRBQ"
            break
        elif PDFfile_name.find('PSQI.pdf') != -1:
            current_ass = "PSQI"
            break
        elif PDFfile_name.find('QFCD.pdf') != -1:
            current_ass = "QFCD"
            break
        elif PDFfile_name.find('RBQ-2A.pdf') != -1:
            current_ass = "RBQ-2A"
            break
        elif PDFfile_name.find('RPFC.pdf') != -1:
            current_ass = "RPFC"
            break
        elif PDFfile_name.find('SODS-A.pdf') != -1:
            current_ass = "SODS-A"
            break
        elif PDFfile_name.find('SOPF.pdf') != -1:
            current_ass = "SOPF"
            break
        elif PDFfile_name.find('SOPSS.pdf') != -1:
            current_ass = "SOPSS"
            break
        elif PDFfile_name.find('SR-AADDSS.pdf') != -1:
            current_ass = "SR-AADDSS"
            break
        elif PDFfile_name.find('SR-ATQ.pdf') != -1:
            current_ass = "SR-ATQ"
            break
        elif PDFfile_name.find('SR-IED-SQ.pdf') != -1:
            current_ass = "SR-IED-SQ"
            break
        elif PDFfile_name.find('SR-IGDS-SF.pdf') != -1:
            current_ass = "SR-IGDS-SF"
            break
        elif PDFfile_name.find('SR-MTS.pdf') != -1:
            current_ass = "SR-MTS"
            break
        elif PDFfile_name.find('SR-PIDSM5.pdf') != -1:
            current_ass = "SR-PIDSM5"
            break
        elif PDFfile_name.find('SR-PTSD.pdf') != -1:
            current_ass = "SR-PTSD"
            break
        elif PDFfile_name.find('SR-RCDQ.pdf') != -1:
            current_ass = "SR-RCDQ"
            break
        elif PDFfile_name.find('SR-SCID-5-SPQ.pdf') != -1:
            current_ass = "SR-SCID-5-SPQ"
            break
        elif PDFfile_name.find('SR-VOPTS.pdf') != -1: #PRIMERO BUSCA SR-VOTPS y luego PTS
            current_ass = "SR-VOPTS"
            break
        elif PDFfile_name.find('PTS.pdf') != -1:
            current_ass = "PTS"
            break
        elif PDFfile_name.find('SR-WHODAS2.pdf') != -1:
            current_ass = "SR-WHODAS2"
            break
        elif PDFfile_name.find('STSI.pdf') != -1:
            current_ass = "STSI"
            break
        elif PDFfile_name.find('SUDN-A.pdf') != -1:
            current_ass = "SUDN-A"
            break
        elif PDFfile_name.find('TAS-20.pdf') != -1:
            current_ass = "TAS-20"
            break
        elif PDFfile_name.find('TEQ.pdf') != -1:
            current_ass = "TEQ"
            break
        elif PDFfile_name.find('TRCS.pdf') != -1:
            current_ass = "TRCS"
            break
        elif PDFfile_name.find('YGTSS.pdf') != -1:
            current_ass = "YGTSS"
            break
        else:     #if PDFfile_name is different from any acronym above, reads the internal text of the PDF to identify it
            uk_pdf_file = PyPDF2.PdfReader(pdf_name_path)

            # The next two vectors assess_v and acronyms_v are sets of two, when the program identifies the text in the PDF from assess_v, it associates the appropiate acronym
            assess_v = ["AGORAPHOBIA DIAGNOSTIC CRITERIA", "MOBILITY INVENTORY FOR AGORAPHOBIA", "SEVERITY MEASURE FOR AGORAPHOBIA", "ANXIETY SCALE - ASI-3", "COMMUNITY ASSESSMENT OF PSYCHOLOGICAL EXPERIENCES", "COMPONENTS OF THE FEMALE SEXUAL FUNCTION INDEX -6 QUESTIONNAIRE", "CLINICIAN - RATED DIMENSIONS OF PSYCHOSIS SYMPTOM SEVERITY", "MADRS  – CLINICIAN RATED", "COLUM BIA-SUICIDE SEVERITY RATING SCALE SCREENING", "RISK ASSESSMENT VERSION",
                          "DAILY MOOD CHART", "DISGUST REVISED SCALE", "DEPRESSION SCALE LONG FORM", "GE NERALIZED ANX IETY DISORDER DIAGNOSIS", "SEVERITY MEASURE FOR GENERALIZED ANXIETY DISORDER", "presentGENDER DYSPHORIA", "INTERNATIONAL PROLONGED GRIEF DISORDER SCALE", "ALTMAN SELF-RATING MANIA SCALE", "LEVEL 2 - REPETITIVE THOUGHTS AND BEHAVIORS – ADULT", "AND HY POMANIC DISORDERS", "MAJOR DEPRESSIVE DISORDER DIAGNOSIS",
                          "MORNING-E VENING CIRCADIAN STABILITY SCALE", "MOOD ELEVATED DISORDER DIAGNOSTIC CRITERIA", "THE MAGICAL IDEATION SCALE", "OCD BODY DYSMORPHIC DIRSORDER – SELF REPORT", "OCD SPECTRUM EXCORIATION", "OCD SPECTRUM HOARDING", "TRICHOTILLOMANIA", "THE PERCEPTUAL ABERRATION SCALE", "PHYSICAL ANHEDONIA SCALE",
                          "PANIC ATTACK SYMPTOMS AS A CO-O CCURRING SYMPTOM", "GRIEF OR  FELT LOSS", "PANIC DISORDER DIAGNOSTIC CRITERIA", "PREMENSTRUAL DYSPHORIC", "PERSISTENT DEPRESSIVE EPISODE: DYSTHYMIA", "DISORDER WITH PSYCHOTIC FEATURES", "SEVERITY MEASURE FOR PANIC DISORDER", "THE PREMENSTRUAL SYMPTOMS SCREENING TOOL", "PREMENSTRUAL TENSION SYNDROME", "RIVERSIDE LIFE SATISFACTION SCALE",
                          "RAPID MOOD SCREENER", "REVISED SOCIAL ANHEDONIA SCALE", "SCREENING ASSESSMENT FOR ELEVATED MOOD HISTORY", "SEPARATION ANXIETY DISORDER  DIAGNOSIS", "SEVERITY MEASURE FOR SEPARATION ANXIETY DISORDER", "DSM-5-TR SELF-RATED LEVEL 1 CROSS-CUTTING SYMPTOM MEASURE – ADULT", "THE SNAITH -HAMILTON PLEASURE SCALE", "SOCIAL ANXIETY DISORDER DIAGNOSIS", "SEVERITY MEASURE FOR SOCIAL ANXIETY", "LIEBOWITZ SOCIAL ANXIETY SCALE",
                          "PHOBIAS", "SEVERITY MEASURE FOR S PECIFIC PHOBIA — ADULT", "SOCIAL PROVISIONS SCALE", "MADRS  – SELF -RATED ASSESSMENT", "merchandise.Y-BO CS -II SELF REPORT VERSION", "SHEEHAN-SUICIDALITY TRACKING SCALE", "UCLA Loneliness S cale", "Y-BOCS SYMPTOMS CHECKLIST",
                          #MINI 2
                          "ANTICIPATORY AND CONSUMMATORY", "ADULT DYSPRAXIA QUESTIONNAIRE", "ANOREXIA NERVOSA", "AUTISM SPECTRUM QUOTIENT", "ADULT RESILENCE MEASURE -REVISED", "AUTONOMIC SYMPTOMS SHORT FORM", "BINGE EATING DISORDER", "BARKLEY FUNCTIONAL IMPAIRMENT SCALE - LF - INFORMANT", "BARKLEY FUNCTIONAL IMPAIRMENT SCALE - LF – SELF REPORT", "BULIMIA NERVOSA",
                          "BODY PAIN MAP", "BODY PERCPETIONS QUESTIONNAIRE SHORT FORM", "COGNITIVE ASSESSMENT", "COMPREHENSIVE AUTISTIC TRAIT INVENTORY", "THE CAMOUFLAGING AUTISTIC TRAITS", "THE CAMBRIDGE BEHAVIOUR SCALE", "THE CONNOR DAVIDSON RESILENCE SCALE", "CLINICIAN-RATED SEVERITY OF NONSUICIDAL SELF-INJURY", "CLINICIAN-RATED SEVERITY OF SOMATIC SYMPTOM DISORDER", "CITY STRESS INVENTORY",
                          "ADULT CONCENTRATION INVENTORY", "DISSOCIATIVE  BELIEFS  ABOUT  MEMORY  QUESTIONNAIRE", "inclusionP 1/2 DEVELOPMENTAL CO-", "Dissociative Experiences Scale-II", "DSPS", "DYSPRAXIA SCREENING TEST", "states .P 1/2 DEVELOPMENTAL TRAUMA", "DISTRESS TOLERANCE SCALE", "EXTREME DEMAND AVOIDANCE", "EATING DISORDER EXAMINATION",
                          "EPWORTH SLEEPINESS SCALE", "ACTIVTY QUESTIONNAIRE", "GIRLS' QUESTIONNAIRE FOR AUTISM SPECTRUM", "HEADACHE QUESTIONNAIRE", "INTERMITTENT EXPLOSIVE DISORDER DIAGNOSIS", "INVENTORY OF PSYCHOLOGICAL FUNCTIONING", "INTERNATIONAL TRAUMA EXPOSURE MEASURE", "THE INTERNATIONAL TRAUMA QU ESTIONNAIRE", "LEVEL 2 - SLEEP DISTURBANCE – ADULT", "LEVEL 2 - SOMATIC SYMPTOM - ADULT PATIENT",
                          "LEVEL 2 - SUBSTANCE USE – ADULT", "LEC-5 STANDARD", "MAIA-2", "MUNICH PARASOMNIAS ASSESSMENT", "MCLEAN SCREENING INSTRUMENT FOR BPD", "MIGRAINE SCREENING", "NEFF'S SELF-COMPASSION", "PHYSICAL ACTIVITY ADULT", "PAST ALCOHOL", "PERTH ALEXITHYMIA QUESTIONNAIRE",
                          "PATHOLOGICAL DISSOCIATION", "PROTECTIVE FACTORS RESILENCE SCALE", "PHYSICAL S YMPTOMS", "POSTTRAUMA RISKY", "PITTSBURGH SLEEP QUALITY INDEX", "QUESTIONNAIRE FOR CONCENTRATION DISORDERS", "RBQ-2A", "RESILIENCE PROTECTIVE  FACTORS", "SEVERITY OF DISSOCIATIVE SYMPTOMS – ADULT*", "SCALE OF PROTECTIVE FACTORS",
                          "SEVERITY OF POSTTRAUMATIC STRESS SYMPTOMS - ADULT", "ADULT ATTENTION DEFICIT DISORDER", "ADULT TIC QUESTIONNAIRE – SELF REPORT", "2 INTERMITTENT EXPLOSIVE", "INTERNET GAMING", "MOTOR TIC SYMTPOMS  – SELF REPORT", "THE  PERSONALITY INVENTORY FOR DSM-5", "PTSD DIAGNOSTIC SCALE", "COORDINATION DISORDER QUESTIONNAIRE", "1/11SCID-5-SPQ - SELF-REPORT",
                          "VOCAL OR PHONIC TIC SYMTPOMS", "WHODAS 2.0", "SCREENER TYPES OF SOUND", "SUBSTANCE USE DISORDER", "TAS-20", "THE EMPATHY QUOTIENT", "TRAUMA RELATED COGNITIONS SCALES", "YALE GLOBAL TIC SEVERITY SCALE"]
                          
            acronyms_v = ["AD", "AMI", "AS", "ASI3", "CAOPE", "CFSFI6", "CR-DPSS", "CR-MADRS", "CSSRS", "CSRH",
                          "DMC", "DRS", "DSLF", "GADD", "GADS", "GD", "IPGDS", "L2-ASRMS", "L2-RTB", "MANIC", "MDDD",
                          "MECSS", "MEDDC", "MIS", "OCD-BDD-SR", "OCD-SESPDDC", "OCD-SHDDC", "OCD-STHPDDC", "PABS", "PANHS",
                          "PAS", "PCGFL", "PDC", "PDDD", "PDED", "PDMDWPF", "PDS", "PSST", "PTS", "RLSS",
                          "RMS", "RSANHS", "SAFEMH", "SEADD", "SEADS", "SF-L1-CCSM", "SHPS", "SOADD", "SOADS", "SOALS",
                          "SPHD", "SPHS", "SPS", "SR-MADRS", "SR-YBOCS-II", "SSTS", "UCLA", "Y-BOCS-SC",
                          # MINI 2
                          "ACIPS", "ADQ", "AN", "AQ", "ARM-R", "ASSF", "BED", "BFIS-I", "BFIS-SR", "BN",
                          "BPM", "BPQ-SF", "CA", "CATI", "CATQ", "CBS", "CDRS", "CR-SNSI", "CR-SSSD", "CSI",
                          "DAACI", "DBMQ", "DCDD", "DES-II", "DSPS", "DST", "DTD", "DTS", "EDAQA", "EDEQ",
                          "ESS", "GPAQ", "GQASC", "HQ", "IEDD", "IOPF", "ITEM", "ITQ", "L2-SD", "L2-SS",
                          "L2-SU", "LEC5S", "MAIA-2", "MPA", "MSI-BPD", "MSQ", "NSCS-SF", "PAAQ", "PADUQ", "PAQ",
                          "PDS-B", "PFRS", "PHQ-15", "PRBQ", "PSQI", "QFCD", "RBQ-2A", "RPFC", "SODS-A", "SOPF",
                          "SOPSS", "SR-AADDSS", "SR-ATQ", "SR-IED-SQ", "SR-IGDS-SF", "SR-MTS", "SR-PIDSM5", "SR-PTSD", "SR-RCDQ", "SR-SCID-5-SPQ",
                          "SR-VOPTS", "SR-WHODAS2", "STSI", "SUDN-A", "TAS-20", "TEQ", "TRCS", "YGTSS"]
                            
            class Found(Exception): pass
            try:
                page = uk_pdf_file.pages[0] #Checks the first page of the PDF
                text = page.extract_text() #Extracts the text of the first page of the PDF
                for (miss_title, miss_acronym) in zip(assess_v, acronyms_v): #Starts the search for text coincidences
                    res_search = re.search(r'\b' + miss_title + r'\b', text) #Exact coincidence
                    if res_search != None: #breaks the for loop when a coincidence is found
                        raise Found
            except: Found
            
            str_numerador  = str(numerador) #Index of the current file
            rename_var = str_numerador + miss_acronym + ".pdf" #Defines the name of the file, with the found acronym
            os.rename(pdf_name_path,os.path.join(path_forms,rename_var)) #Renames the file with the previous definition
            pdf_name_path = os.path.join(path_forms,rename_var) #Defines the new path of the renamed file
            PDFfile_name = os.path.basename(pdf_name_path) #Redefines PDFfile_name variable
            
            #The while cycle starts again and the acronym of the PDF is now identified.
            
    newrow = fillpdfs.get_form_fields(pdf_name_path) #Extracts the raw data of the fields of the form, creating a dictionary
    newrow = replacer(newrow) #Calls replacer function to change "" to "Off".
    
    # print(newrow) #If tou print in the console the newrow variable, it will show you the dictionary of the current PDF 
    
    if numerador == 1: #The first time that we reach this point, creates the header of the report
        
        pdf_name = str(newrow['Name']) #Variables of data in the dictionary of current file
        pdf_lname = str(newrow['Last Name'])
        pdf_phn = str(newrow['PHN'])
        pdf_date = str(newrow['Date'])
    
        title_str = "Missing Answers Report: " + pdf_lname + " " + pdf_name + " - " + pdf_phn + " - " + pdf_date #Creates the title of the report: Last Name + First Name + PHN + Date of interview
        
        cell_format1 = workbook.add_format({'bold': True}) #Defining the format for the title
        cell_format1.set_font_size(20)
        worksheet.write("A1", title_str, cell_format1) #Writes the title in the excel file
        
        cell_format2 = workbook.add_format({'bold': True}) #Defining the format for the title of the columns
        cell_format2.set_font_size(12)
        cell_format2.set_bg_color("#4287f5")
        
        worksheet.write("A2", "Assessment", cell_format2) #Writing of the title of the columns with the defined format
        worksheet.write("B2", "Missing Answers?", cell_format2)
        worksheet.write("C2", "Items Missing Answers", cell_format2)
    
    flag_notcount = 0 #There are some PDFs that are intended to have missing answers, this flag is related to that case
    flag_keys = 0
    
    # MINI 1
    if current_ass == 'DMC':
        flag_notcount = 1 #For some assessments there will always be missing answers, a flag is defined in for these case
    elif current_ass == 'MDDD': #There are some PDFs that have fields that are intended to be empty. rem_keys define these items, this is different for each assessment
        rem_keys = ['1y/n','1increase','1y/ndecrease','25','C13','C14','C15','P13','P14','P15']
        flag_keys = 1
    elif current_ass == 'PDED':
        rem_keys = ['C4', 'C5', 'P4', 'P5', 'C6', 'C7', 'C8', 'C9', 'P6', 'P7', 'P8', 'P9']
        flag_keys = 1
    elif current_ass == 'GD':
        rem_keys = ['E', 'F', 'G', 'H', 'I', 'J']
        flag_keys = 1
    elif current_ass == 'PDDD':
        rem_keys = ['E', 'F', 'G']
        flag_keys = 1
    elif current_ass == 'SAFEMH':
        rem_keys = ['H']
        flag_keys = 1
    elif current_ass == 'MEDDC':
        rem_keys = ['CHKB11', 'DC1', 'DP1', 'DC2', 'DP2', 'EC', 'EP']
        flag_keys = 1
    elif current_ass == 'PDMDWPF':
        rem_keys = ['CBox8']
        flag_keys = 1
    elif current_ass == 'CSRH':
        flag_notcount = 1
    elif current_ass == 'GADD':
        rem_keys = ['EC', 'EP', 'FC', 'FP']
        flag_keys = 1
    elif current_ass == 'PDC':
        rem_keys = ['CC', 'CP', 'DC', 'DP']
        flag_keys = 1
    elif current_ass == 'PAS':
        rem_keys = ['17', '17/17', '18', '18/18', '19', '19/19']
        flag_keys = 1
    elif current_ass == 'SPHD':
        rem_keys = ['Group7']
        flag_keys = 1
    elif current_ass == 'AD':
        rem_keys = ['HC', 'HP', 'IC', 'IP']
        flag_keys = 1
    elif current_ass == 'SOALS':
        rem_keys = ['1-3', '2-3', '3-3', '4-3', '5-3', '6-3', '7-3', '8-3', '9-3', '10-3', '11-3', '12-3', '13-3', '14-3', '15-3', '16-3', '17-3', '18-3', '19-3', '20-3', '21-3', '22-3', '23-3', '24-3']
        flag_keys = 1
    elif current_ass == 'SOADD':
        rem_keys = ['I', 'I/1', 'J', 'J/1', 'K', 'K/1']
        flag_keys = 1
    elif current_ass == 'OCD-BDD-SR':
        rem_keys = ['12yn']
        flag_keys = 1
    elif current_ass == 'OCD-SHDDC':
        rem_keys = ['CE', 'PE', 'CF', 'PF']
        flag_keys = 1
    elif current_ass == 'OCD-SESPDDC':
        rem_keys = ['CD', 'PD', 'CE', 'PE']
        flag_keys = 1
    elif current_ass == 'OCD-STHPDDC':
        rem_keys = ['CD', 'PD', 'CE', 'PE']
        flag_keys = 1
    # MINI 2
    elif current_ass == 'DCDD':
        rem_keys = ['D']
        flag_keys = 1
    elif current_ass == 'IEDD':
        rem_keys = ['B', 'C', 'D', 'E', 'F']
        flag_keys = 1
    elif current_ass == 'SUDN-A':
        rem_keys = ['1 crystal meth']
        flag_keys = 1
    elif current_ass == 'BPM':
        flag_notcount = 1
    elif current_ass == 'MPA':
        rem_keys = ['C1', 'S11', 'O11']
        flag_keys = 1
    elif current_ass == 'PAAQ':
        rem_keys = ['2e', '2f']
        flag_keys = 1
        
    if flag_keys == 1: #Removes items defined since line 518
        for key in rem_keys:
            del newrow[key]
    
    flag_mi = 0
    if flag_notcount == 0: #Line 520 explains the reason of this flag
        
        mi_count = Counter(newrow.values())['Off'] #This function counts the number of "Off" items in the dictionary, "Off" meaning 'not answered'
        
        keys = [k for k, v in newrow.items() if v == 'Off'] #This function extracts the "Off" keys (keys = items = questions in the PDF)
        keys = ', '.join(keys) #Concatenates the "Off" items
        
        if mi_count > 0: #If the count of "Off" items is greater than 0, the PDF is missing answers
            mi_info = 'Yes' #A yes legend is assigned
            flag_mi = 1
        else:
            mi_info = 'No' #A no legend is assigned if the PDF is complete
            keys = "None"
        
    else:
        mi_info = '-' #Output for PDFs related to line 520
        mi_count = '-'
        keys = '-'
    
    write_row = [current_ass, mi_info, keys] #Creates a list where the first value is the acronym of the PDF
                                             #the second value a "Yes" or "No", (missing info)
                                             #the third value the list of items missing info
    
    for col_num, data in enumerate(write_row): #Writes the write_row data in the next empty row in excel sheet
        worksheet.write(row, col_num, data)
        if flag_mi == 1: #if the current assessment is missing information, it highilghts the row
            data_format1 = workbook.add_format({'bg_color': '#F22C3C'})
            worksheet.set_row(row, cell_format=data_format1)
 
def create_report(self, path_forms): #Main function     
    numerador = 0 
    row = 1 
    path_csv = os.path.join(path_forms, "Missing Answers Report.xlsx") #Defines the path of the report, same as selected folder
    workbook = xlsxwriter.Workbook(path_csv) #Creates/overwrites the excel report
    worksheet = workbook.add_worksheet() #Defines the sheet where the information will be written
    
    pdfCounter = 0 

    for folder, subfolders, files in os.walk(path_forms): 
        for file in files:
            if file.endswith(('.pdf', '.PDF')):
                pdfCounter = pdfCounter + 1 #Counts the total number of PDF files
                
    parcial = 100/pdfCounter #Variable to divide in equal parts the progress bar
    bar_val = 0 #Starting value of the progress bar
    
    self.valorstring.emit("Analyzing PDF files...") #Displays information for the user
    self.valorbar.emit(int(round(bar_val))) #Sets the value of the progress bar
    
    try:

        for folder, subfolders, files in os.walk(path_forms):
            for file in files:
                if file.endswith(('.pdf', '.PDF')):
                    bar_val = bar_val + parcial #Sums the value of the progress bar, file by file
                    PDFfile_name = os.path.basename(file) #gets the name of the current file
                    numerador = numerador + 1 #Enumerates the current PDF file
                    row = row + 1 #Enumerates the current row of the excel sheet
                    pdf_name_path = os.path.join(folder, PDFfile_name) #Defines the path of the current file
                    report_ass(PDFfile_name, numerador, pdf_name_path, row, workbook, worksheet, path_forms) #calls the main function
                    if bar_val > 99: #Sometimes the last value is greater than 99 but less than 100, the 100% is shown only at the end of the process
                        self.valorbar.emit(int(round(99))) #The current progressBar only admits int numbers
                    else:
                        self.valorbar.emit(int(round(bar_val)))
                        
        border_format = workbook.add_format({ 'border':1 }) #defines the format of the border of the cells
        worksheet.conditional_format( 'A2:C138' , { 'type' : 'no_blanks' , 'format' : border_format} )
        worksheet.autofit() #autofits the cells to the content
        width = 14
        width2 = 18
        worksheet.set_column(0, 0, width) #sets a predefined widht to the first two columns
        worksheet.set_column(1, 1, width2)
        workbook.close() #closes the excel file
        self.valorstring.emit("Process succesfully completed!")#Displays a message to the user, so it is known that the process is finished
        self.valorpath.emit("---Select a new folder---") 
        self.valorbar.emit(int(round(100)))
        subprocess.Popen([path_csv], shell=True) #Opens the generated report
    except Exception:
        print(traceback.format_exc()) #print any exception in the console