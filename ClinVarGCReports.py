from ftplib import FTP
import xml.etree.ElementTree as ET
import os
import sys
import datetime
import time
import gzip
import re
import pprint
import xlsxwriter

orgDict = {}
gcVarIDs = []
scvHash = {}
a2vHash = {}
HGVSHash = {}
EPHash = {}
sub = 'GenomeConnect_ClinGen'


def get_file(file, path):
    '''This function gets NCBI files from FTP'''

    domain = 'ftp.ncbi.nih.gov'
    user = 'anonymous'
    password = 'tsneddon@broadinstitute.org'

    ftp = FTP(domain)
    ftp.login(user, password)
    ftp.cwd(path)
    localfile = open(file, 'wb')
    ftp.retrbinary('RETR ' + file, localfile.write)
    raw_date = ftp.sendcmd('MDTM ' + file)
    date = datetime.datetime.strptime(raw_date[4:], "%Y%m%d%H%M%S").strftime("%m-%d-%Y")
    ftp.quit()
    localfile.close()

    return(date)


def make_directory(dir, date):
    '''This function makes a local directory for new files if directory does not already exist'''

    directory = dir + '/GCReports_' + date

    if not os.path.exists(directory):
        os.makedirs(directory)
    else:
        sys.exit('Program terminated, ' + directory + ' already exists.')

    return(directory)


def create_orgDict1(gzfile):
    '''This function adds lookup of OrgID to GTR submitter name to orgDict'''

    with gzip.open(gzfile) as input:
        for event, elem in ET.iterparse(input):

            for root in elem.iter(tag='GTRLabData'):
                for node0 in root.iter(tag='GTRLab'):
                    labCode = int(node0.attrib['id'])
                    for node1 in node0.iter(tag='Organization'):
                        for node2 in node1.iter(tag='Name'):
                            labName = node2.text
                            labName = re.sub('[^0-9a-zA-Z]+', '_', labName)
                            labName = labName[0:50]
                            orgDict[labCode] = [labName]

    input.close()
    os.remove(gzfile)
    return(orgDict)


def create_orgDict2(infile):
    '''This function adds lookup of OrgID to ClinVar submitter name to orgDict'''

    with open(infile) as input:
        line = input.readline()

        while line:
            line = input.readline()

            if not line.startswith('#'): #ignore lines that start with #
                col = re.split(r'\t', line) #split on tabs
                if not col[0] == '': #ignore empty lines
                    labName = col[0]
                    labName = re.sub('[^0-9a-zA-Z]+', '_', labName)
                    labName = labName[0:50]
                    labID = int(col[1])
                    if labID not in orgDict:
                        orgDict[labID] = [labName]
                    else:
                        orgDict[labID].append(labName)

    input.close()
    os.remove(infile)
    return(orgDict)


def create_scvHash(gzfile):
    '''This function makes a hash of each SCV in each VarID'''

    with gzip.open(gzfile, 'rt') as input:
        line = input.readline()

        while line:
            line = input.readline()

            if not line.startswith('#'): #ignore lines that start with #
                col = re.split(r'\t', line) #split on tabs
                if not col[0] == '': #ignore empty lines
                    varID = int(col[0])
                    clinSig = col[1]
                    rawDate = col[2]
                    dateLastEval = convert_date(rawDate) #convert date eg May 02, 2018 -> YYYYMMDD

                    conditionList = []
                    rawConditionList = col[5].split(';')
                    for item in rawConditionList:
                        if ':'in item:
                            item = item.split(':', 1)[1]
                        conditionList.append(item)

                    condition = '; '.join(sorted(set(conditionList)))

                    revStat = col[6]
                    colMeth = col[7]

                    submitter = col[9]
                    submitter = re.sub('[^0-9a-zA-Z]+', '_', submitter)
                    submitter = submitter[0:50]

                    SCV = col[10]

                    if (revStat == 'reviewed by expert panel' or revStat == 'practice guideline') and 'PharmGKB' not in submitter: #-- to exclude PharmGKB records
                        EPHash[varID] = {'ClinSig':clinSig, 'Submitter':submitter, 'DateLastEval':dateLastEval}

                    if submitter == sub and varID not in gcVarIDs:
                        gcVarIDs.append(varID)

                    if varID not in scvHash.keys():
                        scvHash[varID] = {}

                    scvHash[varID][SCV] = {'ClinSig':clinSig, 'DateLastEval':dateLastEval, 'Submitter':submitter, 'ReviewStatus':revStat, 'ColMeth':colMeth, 'Condition':condition}

    input.close()
    os.remove(gzfile)
    return(scvHash, EPHash)


def add_labdata(gzfile):
    '''This function adds data to the scvHash from the ClinVar beta XML'''

    with gzip.open(gzfile) as input:
        for event, elem in ET.iterparse(input):

            varID = ''
            scv = ''
            orgID = ''
            DLE = ''
            labCode = ''
            labName = []
            clinSig = ''

            if elem.tag == 'VariationArchive':
                varID = int(elem.attrib['VariationID'])
                if varID in gcVarIDs:
                    #This starts the assertion block
                    for ClinAss in elem.iter(tag='ClinicalAssertion'):
                        for ClinAcc in ClinAss.iter(tag='ClinVarAccession'):
                            scv = ClinAcc.attrib['Accession'] + '.' + ClinAcc.attrib['Version']
                            if scv in scvHash[varID]:
                                for ObsMeth in ClinAss.iter(tag='ObsMethodAttribute'):
                                    for Attr in ObsMeth.iter(tag='Attribute'):
                                        if Attr.attrib['Type'] == 'TestingLaboratory':
                                            if 'dateValue' in Attr.attrib:
                                                DLE = Attr.attrib['dateValue']
                                            else:
                                                DLE = 'None'
                                            scvHash[varID][scv].update({'DateLastEval':DLE})

                                            if 'integerValue' in Attr.attrib:
                                                labCode = int(Attr.attrib['integerValue'])
                                            else:
                                                labCode = 'None'

                                            scvHash[varID][scv].update({'LabCode':labCode})
                                            
                                            if Attr.text != None:
                                                lab = Attr.text
                                                lab = re.sub('[^0-9a-zA-Z]+', '_', lab)
                                                lab = lab[0:50]
                                                labName.append(lab)
                                            else:
                                                try:
                                                    if orgDict[labCode]:
                                                        labName.extend(orgDict[labCode])
                                                except:
                                                    labName.append('None')
                                            scvHash[varID][scv].update({'LabName':labName})

                                            if labCode == 'None' and labName[0] != 'None':
                                                for id in orgDict:
                                                    for name in labName:
                                                        if name in orgDict[id]:
                                                            labCode = int(id)
                                                            scvHash[varID][scv].update({'LabCode':labCode})

                                    for Comment in ObsMeth.iter(tag='Comment'):
                                        if Comment.text != None:
                                            clinSig = Comment.text

                                    if clinSig != '':
                                        scvHash[varID][scv].update({'ClinSig':clinSig})
                                    else:
                                        scvHash[varID][scv].update({'ClinSig':'None'})

                else:
                    elem.clear()

    input.close()
    os.remove(gzfile)
    return(scvHash)


def create_a2vHash(gzfile):
    '''This function makes a dictionary of VarID to AlleleID'''

    with gzip.open(gzfile, 'rt') as input:
        line = input.readline()

        while line:
            line = input.readline()
            if not line.startswith('#'): #ignore lines that start with #
                col = re.split(r'\t', line) #split on tabs
                if not col[0] == '': #ignore empty lines
                    varID = int(col[0])
                    varType = col[1]
                    alleleID = int(col[2])

                    #Ignore rows that are not Variant (simple type)
                    #This excludes Haplotype, CompoundHeterozygote, Complex, Phase unknown, Distinct chromosomes
                    if varType == 'Variant':
                        a2vHash[alleleID] = varID

    input.close()
    os.remove(gzfile)
    return(a2vHash)


def create_HGVSHash(gzfile):
    '''This function makes a hash of metadata for each VarID'''

    with gzip.open(gzfile, 'rt') as input:
        line = input.readline()

        while line:
            line = input.readline()

            if not line.startswith('#'): #ignore lines that start with #
                col = re.split(r'\t', line) #split on tabs
                if not col[0] == '': #ignore empty lines
                    alleleID = int(col[0])
                    varType = col[1]
                    HGVSname = col[2]
                    geneSym = col[4]
                    clinSig = col[6]
                    phenotype = col[13]

                    if alleleID in a2vHash:
                        HGVSHash[a2vHash[alleleID]] = {'VarType':varType, 'HGVSname':HGVSname, 'GeneSym':geneSym, 'ClinSig':clinSig, 'Phenotype':phenotype}

    input.close()
    os.remove(gzfile)
    return(HGVSHash)


def create_files(ExcelDir, excelFile, date):
    '''This function creates an Excel file'''

    dir = ExcelDir

    sub_output_file = dir + '/' + excelFile

    workbook = xlsxwriter.Workbook(sub_output_file)
    worksheet0 = workbook.add_worksheet('README')

    worksheet0.write(0, 0, "Date of ClinVar FTP file: " + date)
    worksheet0.write(2, 0, "Clinical submitter: " + sub)
    worksheet0.write(4, 0, "This Excel file is the output of a script that takes the most recent submission_summary.txt file from the ClinVar FTP site and outputs all the variants for " + sub)
    worksheet0.write(5, 0, 'Each tab is the result of a different set of parameters as outlined below:')
    worksheet0.write(6, 0, '#Variants:')
    worksheet0.write(7, 1, '1. All_subs: All ClinVar variants where there is a GenomeConnect submission.')
    worksheet0.write(8, 1, '2. All_novel: All ClinVar variants where the only submission is from GenomeConnect.')
    worksheet0.write(9, 1, '3. Lab_Conflict: ClinVar variants where the GenomeConnect testing lab clinical significance [P] vs [LP] vs [VUS] vs [LB] vs [B] differs from the clinical lab with same name.')
    worksheet0.write(10, 1, '4. Lab_Consensus: ClinVar variants where the GenomeConnect testing lab clinical significance [P] vs [LP] vs [VUS] vs [LB] vs [B] is the same as that from the clinical lab with same name.')
    worksheet0.write(11, 1, '5. EP_Conflict: ClinVar variants where the GenomeConnect testing lab clinical significance [P/LP] vs [VUS] vs [LB/B] differs from an Expert Panel or Practice Guideline.')
    worksheet0.write(12, 1, '6. Outlier: ClinVar variants where the GenomeConnect testing lab clinical significance [P/LP] vs [VUS] vs [LB/B] differs from at least one 1-star or above (or clinical testing) submitter.')

    worksheet0.write(14, 0, 'Note: Tab classification counts are for unique submissions only i.e. if the same variant is submitted twice as Pathogenic by the same submitter, it will only be counted once')
    worksheet0.write(15, 0, 'Note: A variant can occur in multiple tabs i.e. if the same variant is submitted twice, once as Pathogenic and once as Benign by the same submitter, the variant could be both an outlier and the consensus')

    tabList = [create_tab1, create_tab2, create_tab3, create_tab4, create_tab5, create_tab6]
    for tab in tabList:
        tab(workbook, worksheet0)

    workbook.close()


def convert_date(date):
    '''This function converts a ClinVar date eg May 02, 2018 -> YYYYMMDD'''

    mon2num = dict(Jan='01', Feb='02', Mar='03', Apr='04', May='05', Jun='06',\
                   Jul='07', Aug='08', Sep='09', Oct='10', Nov='11', Dec='12')

    if '-' not in date:
        newDate = re.split(', | ',date)
        newMonth = mon2num[newDate[0]]
        convertDate = (newDate[2] + newMonth + newDate[1]) #YYYYMMDD, an integer for date comparisons
    else:
        convertDate = date

    return(convertDate)


def print_date(date):
    '''This function converts a date eg YYYYMMDD -> YYYY-MM-DD'''

    printDate = date[0:4] + "-" + date[4:6] + "-" + date[6:8] #MM/DD/YYYY, for printing to file
    return(printDate)


def create_tab1(workbook, worksheet0):
    '''This function creates the Tab#1 (All_subs) in the Excel file'''

    worksheet1 = workbook.add_worksheet('1.AllSubs')

    tab = 1
    row = 0
    p2fileVarIDs = {}
    headerSubs = []

    for varID in gcVarIDs:

        subSignificance, submitters, p, lp, plp, vus, lb, b, lbb, vlbb, total, other = get_pathCounts(varID, tab)

        if varID not in p2fileVarIDs.keys():
            p2fileVarIDs[varID] = {}

        p2fileVarIDs[varID] = {'Total':total, 'PLP':plp, 'VUS':vus, 'LBB':lbb, 'Misc':other}

        for SCV in scvHash[varID]:
            if scvHash[varID][SCV]['Submitter'] != sub:
                headerSubs.append(scvHash[varID][SCV]['Submitter'])

    headerSubs = sorted(set(headerSubs))

    print_header(gcVarIDs, headerSubs, worksheet1, tab)

    for varID in gcVarIDs:
        varSubs = get_varSubs(varID)
        row = print_variants(worksheet1, row, varID, headerSubs, varSubs, p2fileVarIDs, tab)

    print_stats(worksheet0, 7, 0, row)


def create_tab2(workbook, worksheet0):
    '''This function creates the Tab#2 (All_novel_GC) in the Excel file'''

    worksheet2 = workbook.add_worksheet('2.AllNovelGC')

    tab = 2
    row = 0
    p2fileVarIDs = []
    headerSubs = []
    varSubs = []

    for varID in gcVarIDs:
        if len(scvHash[varID].values()) == 1:
            p2fileVarIDs.append(varID)

    print_header(p2fileVarIDs, headerSubs, worksheet2, tab)

    for varID in p2fileVarIDs:
        row = print_variants(worksheet2, row, varID, headerSubs, varSubs, p2fileVarIDs, tab)

    print_stats(worksheet0, 8, 0, row)


def create_tab3(workbook, worksheet0):
    '''This function creates the Tab#3 (Lab_Conflicts) in the Excel file'''

    worksheet3 = workbook.add_worksheet('3.LabConflicts')

    tab = 3
    row = 0
    p2fileVarIDs = []
    headerSubs = []

    for varID in gcVarIDs:

        lab = []
        sig = ''
        for SCV in scvHash[varID]:
            if sub == scvHash[varID][SCV]['Submitter']:
                lab = scvHash[varID][SCV]['LabName']
                sig = scvHash[varID][SCV]['ClinSig']

        for SCV in scvHash[varID]:
            if sub != scvHash[varID][SCV]['Submitter'] and scvHash[varID][SCV]['Submitter'] in lab and scvHash[varID][SCV]['ClinSig'] != sig:
                headerSubs.append(scvHash[varID][SCV]['Submitter'])
                if varID not in p2fileVarIDs:
                    p2fileVarIDs.append(varID)

    headerSubs = sorted(set(headerSubs))

    print_header(p2fileVarIDs, headerSubs, worksheet3, tab)

    for varID in p2fileVarIDs:
        varSubs = get_varSubs(varID)
        row = print_variants(worksheet3, row, varID, headerSubs, varSubs, p2fileVarIDs, tab)

    print_stats(worksheet0, 9, 0, row)


def create_tab4(workbook, worksheet0):
    '''This function creates the Tab#4 (Lab_Consensus) in the Excel file'''

    worksheet4 = workbook.add_worksheet('4.LabConsensus')

    tab = 4
    row = 0
    p2fileVarIDs = []
    headerSubs = []

    for varID in gcVarIDs:

        lab = []
        sig = ''
        for SCV in scvHash[varID]:
            if sub == scvHash[varID][SCV]['Submitter']:
                lab = scvHash[varID][SCV]['LabName']
                sig = scvHash[varID][SCV]['ClinSig']

        for SCV in scvHash[varID]:
            if sub != scvHash[varID][SCV]['Submitter'] and scvHash[varID][SCV]['Submitter'] in lab and scvHash[varID][SCV]['ClinSig'] == sig:
                headerSubs.append(scvHash[varID][SCV]['Submitter'])
                if varID not in p2fileVarIDs:
                    p2fileVarIDs.append(varID)

    headerSubs = sorted(set(headerSubs))

    print_header(p2fileVarIDs, headerSubs, worksheet4, tab)

    for varID in p2fileVarIDs:
        varSubs = get_varSubs(varID)
        row = print_variants(worksheet4, row, varID, headerSubs, varSubs, p2fileVarIDs, tab)

    print_stats(worksheet0, 10, 0, row)


def create_tab5(workbook, worksheet0):
    '''This function creates the Tab#5 (VCEP _Conflict) in the Excel file'''

    worksheet5 = workbook.add_worksheet('5.VCEPConflict')

    tab = 5
    row = 0
    p2fileVarIDs = {}
    headerSubs = []

    for varID in gcVarIDs:
        p2fileVarIDs, headerSubs = EP_outlier(varID, headerSubs, p2fileVarIDs, tab)

    print_header(p2fileVarIDs, headerSubs, worksheet5, tab)

    for varID in p2fileVarIDs:
        varSubs = get_varSubs(varID)
        row = print_variants(worksheet5, row, varID, headerSubs, varSubs, p2fileVarIDs, tab)

    print_stats(worksheet0, 11, 0, row)


def create_tab6(workbook, worksheet0):
    '''This function creates the Tab#6 (Outlier) in the Excel file'''

    worksheet6 = workbook.add_worksheet('6.Outlier_P.VUS.B')

    tab = 6
    row = 0
    p2fileVarIDs = {}
    headerSubs = []

    for varID in gcVarIDs:
        p2fileVarIDs, headerSubs = outlier(varID, headerSubs, p2fileVarIDs, tab)

    print_header(p2fileVarIDs, headerSubs, worksheet6, tab)

    for varID in p2fileVarIDs:
        varSubs = get_varSubs(varID)
        row = print_variants(worksheet6, row, varID, headerSubs, varSubs, p2fileVarIDs, tab)

    print_stats(worksheet0, 12, 0, row)


def EP_outlier(varID, headerSubs, p2fileVarIDs, tab):
    '''This function returns the submitters where the clinical significance is discrepant from an Expert Panel or Practice Guideline'''

    clinSig = ''
    EPconflict = ''

    if varID in EPHash:
        for SCV in scvHash[varID]:
            if scvHash[varID][SCV]['Submitter'] == sub and (scvHash[varID][SCV]['ClinSig'] == 'Pathogenic' or scvHash[varID][SCV]['ClinSig'] == 'Likely pathogenic' or \
               scvHash[varID][SCV]['ClinSig'] == 'Uncertain significance' or scvHash[varID][SCV]['ClinSig'] == 'Likely benign' or scvHash[varID][SCV]['ClinSig'] == 'Benign'):
                clinSig = scvHash[varID][SCV]['ClinSig']

                if (((EPHash[varID]['ClinSig'] == 'Pathogenic' or EPHash[varID]['ClinSig'] == 'Likely pathogenic') and scvHash[varID][SCV]['ClinSig'] == 'Uncertain significance') or \
                   (EPHash[varID]['ClinSig'] == 'Uncertain significance' and (scvHash[varID][SCV]['ClinSig'] == 'Pathogenic' or scvHash[varID][SCV]['ClinSig'] == 'Likely pathogenic'))):
                    EPconflict = 'P/LP vs VUS'
                if ((EPHash[varID]['ClinSig'] == 'Uncertain significance' and (scvHash[varID][SCV]['ClinSig'] == 'Likely benign' or scvHash[varID][SCV]['ClinSig'] == 'Benign')) or \
                    (scvHash[varID][SCV]['ClinSig'] == 'Uncertain significance' and (EPHash[varID]['ClinSig'] == 'Likely benign' or EPHash[varID]['ClinSig'] == 'Benign'))):
                    EPconflict = 'VUS vs LB/B'
                if (((EPHash[varID]['ClinSig'] == 'Likely benign' or EPHash[varID]['ClinSig'] == 'Benign') and (scvHash[varID][SCV]['ClinSig'] == 'Pathogenic' or scvHash[varID][SCV]['ClinSig'] == 'Likely pathogenic')) or \
                   ((EPHash[varID]['ClinSig'] == 'Pathogenic' or EPHash[varID]['ClinSig'] == 'Likely pathogenic') and (scvHash[varID][SCV]['ClinSig'] == 'Likely benign' or scvHash[varID][SCV]['ClinSig'] == 'Likely benign'))):
                    EPconflict = 'P/LP vs LB/B'

        if clinSig != '' and clinSig != EPHash[varID]['ClinSig']:
            if EPconflict != '':
                if varID not in p2fileVarIDs.keys():
                    p2fileVarIDs[varID] = {}

                subSignificance, submitters, p, lp, plp, vus, lb, b, lbb, vlbb, total, other = get_pathCounts(varID, tab)

                p2fileVarIDs[varID] = {'Total':total, 'PLP':plp, 'VUS':vus, 'LBB':lbb, 'Misc':other, 'EP':EPHash[varID]['Submitter'], 'EP_clinSig':EPHash[varID]['ClinSig']}
                p2fileVarIDs[varID].update({'EPConflict':EPconflict})

                if submitters:
                    headerSubs.extend(submitters)

    headerSubs = sorted(set(headerSubs))
    if sub in headerSubs:
        headerSubs.remove(sub)

    return(p2fileVarIDs, headerSubs)


def outlier(varID, headerSubs, p2fileVarIDs, tab):
    '''This function returns the outlier submitters in a medically significant VarID'''

    subSignificance, submitters, p, lp, plp, vus, lb, b, lbb, vlbb, total, other = get_pathCounts(varID, tab)

    conflict = ''

    if ('P' in subSignificance and vlbb != 0) or ('VUS' in subSignificance and (plp != 0 or lbb != 0)) or ('B' in subSignificance and plp+vus != 0):
        if varID not in p2fileVarIDs.keys():
            p2fileVarIDs[varID] = {}

        p2fileVarIDs[varID] = {'Total':total, 'PLP':plp, 'VUS':vus, 'LBB':lbb, 'Misc':other}

        if submitters:
            headerSubs.extend(submitters)

    headerSubs = sorted(set(headerSubs))
    if sub in headerSubs:
        headerSubs.remove(sub)

    if 'P' in subSignificance:
        if vus == 0 and lbb != 0:
            conflict = 'B'

        if vus != 0 and lbb == 0:
            conflict = 'VUS'

        if vus != 0 and lbb != 0:
            conflict = 'VUS/B'

    if 'VUS' in subSignificance:
        if plp != 0 and lbb == 0:
            conflict = 'P'

        if plp == 0 and lbb != 0:
            conflict = 'B'

        if plp != 0 and lbb != 0:
            conflict = 'P/B'

    if 'B' in subSignificance:
        if plp != 0 and vus == 0:
            conflict = 'P'

        if plp == 0 and vus != 0:
            conflict = 'VUS'

        if plp != 0 and vus != 0:
            conflict = 'P/VUS'

    if conflict != '':
        p2fileVarIDs[varID].update({'Conflict':conflict})

    return(p2fileVarIDs, headerSubs)


def get_pathCounts(varID, tab):
    '''This function returns the counts of ACMG pathogenicities for each VarID'''

    submitters = []
    p = 0
    lp = 0
    vus = 0
    lb = 0
    b = 0
    other = 0
    subSignificance = []
    unique_subs = []

    for SCV in scvHash[varID]:

        if SCV in scvHash[varID] and scvHash[varID][SCV]['ClinSig'] == 'Pathogenic':
            current_sub = scvHash[varID][SCV]['Submitter'] + scvHash[varID][SCV]['ClinSig']
            if scvHash[varID][SCV]['Submitter'] == sub and current_sub not in unique_subs:
                unique_subs.append(current_sub)
                subSignificance.append('P')
                p += 1

            if scvHash[varID][SCV]['Submitter'] != sub and current_sub not in unique_subs and (tab == 1 or \
               (tab != 1 and (scvHash[varID][SCV]['ReviewStatus'] == 'criteria provided, single submitter' or 'clinical testing' in scvHash[varID][SCV]['ColMeth']))):
                unique_subs.append(current_sub)
                submitters.append(scvHash[varID][SCV]['Submitter'])
                p += 1

        if SCV in scvHash[varID] and scvHash[varID][SCV]['ClinSig'] == 'Likely pathogenic':
            current_sub = scvHash[varID][SCV]['Submitter'] + scvHash[varID][SCV]['ClinSig']
            if scvHash[varID][SCV]['Submitter'] == sub and current_sub not in unique_subs:
                unique_subs.append(current_sub)
                subSignificance.append('P')
                lp += 1

            if scvHash[varID][SCV]['Submitter'] != sub and current_sub not in unique_subs and (tab == 1 or \
               (tab != 1 and (scvHash[varID][SCV]['ReviewStatus'] == 'criteria provided, single submitter' or 'clinical testing' in scvHash[varID][SCV]['ColMeth']))):
                unique_subs.append(current_sub)
                submitters.append(scvHash[varID][SCV]['Submitter'])
                lp += 1

        if SCV in scvHash[varID] and scvHash[varID][SCV]['ClinSig'] == 'Uncertain significance':
            current_sub = scvHash[varID][SCV]['Submitter'] + scvHash[varID][SCV]['ClinSig']
            if scvHash[varID][SCV]['Submitter'] == sub and current_sub not in unique_subs:
                unique_subs.append(current_sub)
                subSignificance.append('VUS')
                vus += 1

            if scvHash[varID][SCV]['Submitter'] != sub and current_sub not in unique_subs and (tab == 1 or \
               (tab != 1 and (scvHash[varID][SCV]['ReviewStatus'] == 'criteria provided, single submitter' or 'clinical testing' in scvHash[varID][SCV]['ColMeth']))):
                unique_subs.append(current_sub)
                submitters.append(scvHash[varID][SCV]['Submitter'])
                vus += 1

        if SCV in scvHash[varID] and scvHash[varID][SCV]['ClinSig'] == 'Likely benign':
            current_sub = scvHash[varID][SCV]['Submitter'] + scvHash[varID][SCV]['ClinSig']
            if scvHash[varID][SCV]['Submitter'] == sub and current_sub not in unique_subs:
                unique_subs.append(current_sub)
                subSignificance.append('B')
                lb += 1

            if scvHash[varID][SCV]['Submitter'] != sub and current_sub not in unique_subs and (tab == 1 or (tab != 1 and (scvHash[varID][SCV]['ReviewStatus'] == 'criteria provided, single submitter' and 'clinical testing' in scvHash[varID][SCV]['ColMeth']))):
                unique_subs.append(current_sub)
                submitters.append(scvHash[varID][SCV]['Submitter'])
                lb += 1

        if SCV in scvHash[varID] and scvHash[varID][SCV]['ClinSig'] == 'Benign':
            current_sub = scvHash[varID][SCV]['Submitter'] + scvHash[varID][SCV]['ClinSig']
            if scvHash[varID][SCV]['Submitter'] == sub and current_sub not in unique_subs:
                unique_subs.append(current_sub)
                subSignificance.append('B')
                b += 1

            if scvHash[varID][SCV]['Submitter'] != sub and current_sub not in unique_subs and (tab == 1 or (tab != 1 and (scvHash[varID][SCV]['ReviewStatus'] == 'criteria provided, single submitter' and 'clinical testing' in scvHash[varID][SCV]['ColMeth']))):
                unique_subs.append(current_sub)
                submitters.append(scvHash[varID][SCV]['Submitter'])
                b += 1

        else:
            current_sub = scvHash[varID][SCV]['Submitter'] + scvHash[varID][SCV]['ClinSig']
            if scvHash[varID][SCV]['Submitter'] == sub and current_sub not in unique_subs:
                unique_subs.append(current_sub)
                other += 1

            if scvHash[varID][SCV]['Submitter'] != sub and current_sub not in unique_subs and (tab == 1 or (tab != 1 and (scvHash[varID][SCV]['ReviewStatus'] == 'criteria provided, single submitter' and 'clinical testing' in scvHash[varID][SCV]['ColMeth']))):
                unique_subs.append(current_sub)
                submitters.append(scvHash[varID][SCV]['Submitter'])
                other += 1

    plp = p+lp
    lbb = lb+b
    vlbb = vus+lb+b
    total = plp+vus+lbb+other

    subSignificance = sorted(set(subSignificance))

    return(subSignificance, submitters, p, lp, plp, vus, lb, b, lbb, vlbb, total, other)


def get_varSubs(varID):
    '''This function returns the list of 1-star variant submitters'''

    varSubs = []
    if varID in scvHash:
        for SCV in scvHash[varID]:
            if scvHash[varID][SCV]['Submitter'] != sub:
                if scvHash[varID][SCV]['DateLastEval'] != '-':
                    #Convert date from YYYYMMDD -> YYYY-MM-DD
                    subPrintDate = print_date(scvHash[varID][SCV]['DateLastEval'])
                    varSubs.append(scvHash[varID][SCV]['Submitter'] + ' [' + scvHash[varID][SCV]['ClinSig'] + ' (' + subPrintDate + ')]')
                else:
                    varSubs.append(scvHash[varID][SCV]['Submitter'] + ' [' + scvHash[varID][SCV]['ClinSig'] + ' (No DLE)]')

    varSubs = sorted(set(varSubs))

    return(varSubs)


def print_header(gcVarIDs, headerSubs, worksheet, tab):
    '''This function prints all the header titles to the Excel tabs'''

    k = 0
    if gcVarIDs != []:
        worksheet.write(0, k, 'VarID')
        k+=1
        worksheet.write(0, k, 'Gene_symbol')
        k+=1
        worksheet.write(0, k, 'All_conditions')
        k+=1
        worksheet.write(0, k, 'HGVS_name')
        k+=1
        worksheet.write(0, k, 'Aggregate_clinical_significance')
        k+=1
        if tab == 5:
            worksheet.write(0, k, 'EP')
            k+=1
            worksheet.write(0, k, 'EP_significance')
            k+=1
            worksheet.write(0, k, 'EP_conflict')
            k+=1
        else:
            worksheet.write(0, k, 'EP')
            k+=1

        if tab == 6:
            worksheet.write(0, k, 'Conflict_significance')
            k+=1

        worksheet.write(0, k, sub + '_testing_lab(s)')
        k+=1
        worksheet.write(0, k, sub + '_testing_lab(s)_significance(s)')
        k+=1
        worksheet.write(0, k, sub + '_SCVid(s)')
        k+=1
        worksheet.write(0, k, sub + '_testing_lab(s)_Date(s)LastEvaluated')
        k+=1
        worksheet.write(0, k, sub + '_condition(s)')
        k+=1

        if tab != 5:
            for head in headerSubs:
                if head != sub:
                    worksheet.write(0, k, head)
                    k+=1

        if tab == 1 or tab == 6:
            worksheet.write(0, k, 'Total_submissions')
            k+=1
            worksheet.write(0, k, 'Total_PLP')
            k+=1
            worksheet.write(0, k, 'Total_VUS')
            k+=1
            worksheet.write(0, k, 'Total_LBB')
            k+=1
            worksheet.write(0, k, 'Total_Misc')
            k+=1
        if tab != 2 and tab != 5:
            worksheet.write(0, k, 'Submitting_labs')
    else:
        worksheet.write(0, 0, 'No variants found')


def print_variants(worksheet, row, varID, headerSubs, varSubs, p2fileVarIDs, tab):
    '''This function prints all the variants to the Excel tabs'''

    row += 1
    k = 0
    allList = []

    worksheet.write(row, k, varID)
    k+=1

    if HGVSHash[varID]['GeneSym']:
        worksheet.write(row, k, HGVSHash[varID]['GeneSym'])
    k+=1

    if HGVSHash[varID]['Phenotype']:
        worksheet.write(row, k, HGVSHash[varID]['Phenotype'])
    k+=1

    if HGVSHash[varID]['HGVSname']:
        worksheet.write(row, k, HGVSHash[varID]['HGVSname'])
    k+=1

    if HGVSHash[varID]['ClinSig']:
        worksheet.write(row, k, HGVSHash[varID]['ClinSig'])
    k+=1

    if tab != 5:
       if varID in EPHash.keys():
           worksheet.write(row, k, EPHash[varID]['Submitter'] + ' (' + EPHash[varID]['ClinSig'] + ')')
       else:
           worksheet.write(row, k, 'N/A')
       k+=1
    if tab == 5:
        worksheet.write(row, k, EPHash[varID]['Submitter'])
        k+=1
        if EPHash[varID]['DateLastEval'] != '-':
            #Convert date from YYYYMMDD -> YYYY-MM-DD
            subPrintDate = print_date(EPHash[varID]['DateLastEval'])
            worksheet.write(row, k, EPHash[varID]['ClinSig'] + ' (' + subPrintDate + ')')
        else:
            worksheet.write(row, k, EPHash[varID]['ClinSig'] + ' (No DLE)')
        k+=1
        worksheet.write(row, k, p2fileVarIDs[varID]['EPConflict'])
        k+=1
    if tab == 6:
        worksheet.write(row, k, p2fileVarIDs[varID]['Conflict'])
        k+=1

    labs = []
    clinSig = []
    scvs= []
    dle = []
    conditions = []

    for scv in scvHash[varID]:
        if scvHash[varID][scv]['Submitter'] == sub:
            labs.extend(scvHash[varID][scv]['LabName'])
            clinSig.append(scvHash[varID][scv]['ClinSig'])
            scvs.append(scv)
            #dle.append(print_date(scvHash[varID][scv]['DateLastEval']))
            dle.append(scvHash[varID][scv]['DateLastEval'])
            conditions.append(scvHash[varID][scv]['Condition'])

    labList = ' | '. join(sorted(set(labs)))
    clinSigList = ' | '. join(sorted(set(clinSig)))
    scvList = ' | '. join(sorted(set(scvs)))
    dleList = ' | '. join(sorted(set(dle)))
    conditionsList = ' | '. join(sorted(set(conditions)))

    worksheet.write(row, k, labList)
    k+=1
    worksheet.write(row, k, clinSigList)
    k+=1
    worksheet.write(row, k, scvList)
    k+=1
    worksheet.write(row, k, dleList)
    k+=1
    worksheet.write(row, k, conditionsList)
    k+=1

    if tab == 1 or tab == 6:
        for headerSub in headerSubs:
            p2file = 'no'
            for varSub in varSubs:
                if headerSub in varSub:
                    p2file = varSub[varSub.find("[")+1:varSub.find("]")]
            if p2file != 'no':
                worksheet.write(row, k, p2file)
                allList.append(headerSub + ' [' + p2file+ ']')
                k += 1
            else:
                k += 1

        if varID in p2fileVarIDs:
           worksheet.write(row, k, p2fileVarIDs[varID]['Total'])
           k+=1
           worksheet.write(row, k, p2fileVarIDs[varID]['PLP'])
           k+=1
           worksheet.write(row, k, p2fileVarIDs[varID]['VUS'])
           k+=1
           worksheet.write(row, k, p2fileVarIDs[varID]['LBB'])
           k+=1
           worksheet.write(row, k, p2fileVarIDs[varID]['Misc'])
           k+=1

    if tab == 3 or tab == 4:
        for headerSub in headerSubs:
            p2file = 'no'
            for varSub in varSubs:
                if headerSub in varSub:
                    p2file = varSub[varSub.find("[")+1:varSub.find("]")]
            if p2file != 'no':
                for scv in scvHash[varID]:
                    if 'LabName' in scvHash[varID][scv] and headerSub in scvHash[varID][scv]['LabName']:# and scvHash[varID][scv]['ClinSig'] not in clinSig:
                        worksheet.write(row, k, p2file)
                        allList.append(headerSub + ' [' + p2file+ ']')
                k += 1
            else:
                k += 1

    allList = ' | '. join(sorted(set(allList)))
    worksheet.write(row, k, allList)

    return(row)


def print_stats(worksheet0, line, column, row):
    '''This function prints the total variant count to the README Excel tab'''

    worksheet0.write(line, column, row)


def main():

    inputFile1 = 'gtr_ftp.xml.gz' #path: /pub/GTR/data/
    inputFile2 = 'organization_summary.txt' #path: /pub/clinvar/tab_delimited/
    inputFile3 = 'submission_summary.txt.gz' #path: /pub/clinvar/tab_delimited/
    inputFile4 = 'variation_archive_20190225.xml.gz' #path: /pub/clinvar/xml/clinvar_variation/beta/
    inputFile5 = 'variation_allele.txt.gz' #path: /pub/clinvar/tab_delimited/
    inputFile6 = 'variant_summary.txt.gz' #path: /pub/clinvar/tab_delimited/

    dir = 'ClinVarGCReports'

    get_file(inputFile1, '/pub/GTR/data/')
    get_file(inputFile2, '/pub/clinvar/tab_delimited/')
    date = get_file(inputFile3, '/pub/clinvar/tab_delimited/')
    get_file(inputFile4, '/pub/clinvar/xml/clinvar_variation/beta/')
    get_file(inputFile5, '/pub/clinvar/tab_delimited/')
    get_file(inputFile6, '/pub/clinvar/tab_delimited/')

    ExcelDir = make_directory(dir, date)

    create_orgDict1(inputFile1)
    create_orgDict2(inputFile2)
    create_scvHash(inputFile3)
    add_labdata(inputFile4)

    tabFile = ExcelDir + '/GCSummary_' + date + '.txt'
    excelFile = 'GenomeConnectReport_' + date + '.xlsx'

    create_a2vHash(inputFile5)
    create_HGVSHash(inputFile6)

    create_files(ExcelDir, excelFile, date)

main()
