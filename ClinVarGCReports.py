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

gtrHash = {}
orgDict = {}
gcHash = {}

gcVarIDs = []
scvHash = {}
a2vHash = {}
HGVSHash = {}
EPHash = {}
subList = ['GenomeConnect_ClinGen']


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


def create_gtrHash(gzfile):
    '''This function makes a lookup of OrgID to GTR submitter name'''

    with gzip.open(gzfile) as input:
        for event, elem in ET.iterparse(input):

            labCode = ''
            labName = ''

            for root in elem.iter(tag='GTRLabData'):
                for node0 in root.iter(tag='GTRLab'):
                    labCode = int(node0.attrib['id'])
                    for node1 in node0.iter(tag='Organization'):
                        for node2 in node1.iter(tag='Name'):
                            labName = node2.text
                            gtrHash[labCode] = labName

    os.remove(gzfile)
    return(gtrHash)


def create_orgDict(infile):
    '''This function makes a lookup of OrgID to ClinVar submitter name'''

    with open(infile) as input:
        line = input.readline()

        while line:
            line = input.readline()

            if not line.startswith('#'): #ignore lines that start with #
                col = re.split(r'\t', line) #split on tabs
                if not col[0] == '': #ignore empty lines
                    labName = col[0]
                    labID = int(col[1])

                    orgDict[labID] = labName

    input.close()
    os.remove(infile)
    return(orgDict)


def create_gcHash(gzfile):
    '''This function makes a hash of each SCV (submitter == GenomeConnect) in each VarID from ClinVar submission_summary.txt.gz'''

    with gzip.open(gzfile, 'rt') as input:
        line = input.readline()

        while line:
            line = input.readline()

            if not line.startswith('#'): #ignore lines that start with #
                col = re.split(r'\t', line) #split on tabs
                if not col[0] == '': #ignore empty lines
                    varID = int(col[0])

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
                    submitter = re.sub(r'\s+', '_', submitter) #replace all spaces with an underscore
                    submitter = re.sub(r'/', '-', submitter) # replace all slashes with a hyphen
                    submitter = re.sub(r'\W+', '', submitter) #remove all non-alphanumerics

                    submitter = submitter[0:50]

                    SCV = col[10]

                    if 'GenomeConnect' in submitter and varID not in gcHash.keys():
                        gcHash[varID] = {}

                    if 'GenomeConnect' in submitter:
                        gcHash[varID][SCV] = {'Submitter':submitter, 'Condition':condition, 'ReviewStatus':revStat, 'ColMeth':colMeth}

    input.close()
    return(gcHash)


def add_labdata(gzfile):
    '''This function adds data to the gcHash from the ClinVar beta XML'''

    with gzip.open(gzfile) as input:
        for event, elem in ET.iterparse(input):

            varID = ''
            scv = ''
            orgID = ''
            DLE = ''
            labCode = ''
            labName = ''
            clinSig = ''

            if elem.tag == 'VariationArchive':
                varID = int(elem.attrib['VariationID'])
                if varID in gcHash:
                    #This starts the assertion block
                    for ClinAss in elem.iter(tag='ClinicalAssertion'):
                        for ClinAcc in ClinAss.iter(tag='ClinVarAccession'):
                            scv = ClinAcc.attrib['Accession'] + '.' + ClinAcc.attrib['Version']
                            if scv in gcHash[varID]:
                                orgID = ClinAcc.attrib['OrgID']
                                gcHash[varID][scv].update({'OrgID':orgID})

                                for ObsMeth in ClinAss.iter(tag='ObsMethodAttribute'):
                                    for Attr in ObsMeth.iter(tag='Attribute'):
                                        if Attr.attrib['Type'] == 'TestingLaboratory':
                                            if 'dateValue' in Attr.attrib:
                                                DLE = Attr.attrib['dateValue']
                                            else:
                                                DLE = 'None'
                                            gcHash[varID][scv].update({'DLE':DLE})

                                            if Attr.text != None:
                                                labName = Attr.text
                                            else:
                                                labName = 'None'
                                            gcHash[varID][scv].update({'LabName':labName})

                                            if 'integerValue' in Attr.attrib:
                                                labCode = int(Attr.attrib['integerValue'])
                                            else:
                                                for id in orgDict:
                                                    if orgDict[id] == labName:
                                                        labCode = orgDict[id]
                                                    else:
                                                        labCode = 'None'

                                            #Hardcode HudsonAlpha as already represented in ClinVar by 505530
                                            if labCode == 505801:
                                                gcHash[varID][scv].update({'LabCode':505530})
                                            else:
                                                gcHash[varID][scv].update({'LabCode':labCode})
                                            #################################################################

                                    for Comment in ObsMeth.iter(tag='Comment'):
                                        if Comment.text != None:
                                            clinSig = Comment.text

                                    if clinSig != '':
                                        gcHash[varID][scv].update({'ClinSig':clinSig})
                                    else:
                                        gcHash[varID][scv].update({'ClinSig':'None'})

                else:
                    elem.clear()

    os.remove(gzfile)
    return(gcHash)


def write2file(output):
    '''This function writes the gcHash to a tab-delimted file'''

    with open(output, 'w') as GC:

        GC.write('#VarID' + '\t' + 'SCV' + '\t' + 'Testing lab Clinical Significance' + '\t' + 'Testing lab Date Last Evaluated' + '\t' + 'Condition(s)' + '\t' + 'Testing lab Name' + '\t' + 'Testing lab OrgID' + '\t' + 'Review Status' + '\t' + 'Collection Method' + '\t' + 'Submitter' + '\n')

        for varID in gcHash:
            for scv in gcHash[varID]:

                GC.write(str(varID) + '\t' + scv + '\t' + gcHash[varID][scv]['ClinSig'] + '\t' + gcHash[varID][scv]['DLE'] + '\t' + gcHash[varID][scv]['Condition'] + '\t')

                if gcHash[varID][scv]['LabName'] != 'None':
                    GC.write(gcHash[varID][scv]['LabName'] + '\t')
                else:
                    try:
                        GC.write(str(orgDict[gcHash[varID][scv]['LabCode']]) + '\t')
                    except KeyError:
                        try:
                            GC.write(str(gtrHash[gcHash[varID][scv]['LabCode']]) + '\t')
                        except KeyError:
                            GC.write('Unknown LabCode: ' + str(gcHash[varID][scv]['LabCode']) + '\t')
                            print(scv, '- Unknown LabCode: ', gcHash[varID][scv]['LabCode'])

                GC.write(str(gcHash[varID][scv]['LabCode']) + '\t' + gcHash[varID][scv]['ReviewStatus'] + '\t' + gcHash[varID][scv]['ColMeth'] + '\t' + gcHash[varID][scv]['Submitter'] + '\n')

    GC.close()
    return(GC)


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


def create_scvHash(gzfile):
    '''This function makes a hash of each SCV in each VarID'''

    global subList

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

                    submitter = re.sub(r'\s+', '_', submitter) #replace all spaces with an underscore
                    submitter = re.sub(r'/', '-', submitter) # replace all slashes with a hyphen
                    submitter = re.sub(r'\W+', '', submitter) #remove all non-alphanumerics

                    submitter = submitter[0:50]

                    SCV = col[10]

                    if (revStat == 'reviewed by expert panel' or revStat == 'practice guideline') and 'PharmGKB' not in submitter: #-- to exclude PharmGKB records
                        EPHash[varID] = {'ClinSig':clinSig, 'Submitter':submitter, 'DateLastEval':dateLastEval}

                    if 'GenomeConnect' in submitter and varID not in gcVarIDs:
                        gcVarIDs.append(varID)

                    if varID not in scvHash.keys():
                        scvHash[varID] = {}

                    scvHash[varID][SCV] = {'ClinSig':clinSig, 'DateLastEval':dateLastEval, 'Submitter':submitter, 'ReviewStatus':revStat, 'ColMeth':colMeth, 'Condition':condition}#, 'SubCode':subCode}

    os.remove(gzfile)
    return(scvHash, EPHash, subList)


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
                    type = col[1]
                    alleleID = int(col[2])

                    #Ignore rows that are not Variant (simple type)
                    #This excludes Haplotype, CompoundHeterozygote, Complex, Phase unknown, Distinct chromosomes
                    if type == 'Variant':
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
                    type = col[1]
                    HGVSname = col[2]
                    geneSym = col[4]
                    phenotype = col[13]
                    guidelines = col[26]

                    if alleleID in a2vHash:
                        HGVSHash[a2vHash[alleleID]] = {'VarType':type, 'HGVSname':HGVSname, 'GeneSym':geneSym,'Phenotype':phenotype,'Guidelines':guidelines}

    input.close()
    os.remove(gzfile)
    return(HGVSHash)


def add_GCdata(infile):
    '''This function adss GenomeConnect data to the scv hash'''

    with open(infile, 'rt') as input:
        line = input.readline()

        while line:
            line = input.readline()

            if not line.startswith('#'): #ignore lines that start with #
                col = re.split(r'\t', line) #split on tabs
                if not col[0] == '': #ignore empty lines
                    varID = int(col[0])
                    SCV= col[1]
                    clinSig = col[2]
                    dle = col[3] #2017-06-20 -> YYYYMMDD
                    dle = re.sub('-','',dle)
                    condition = col[4]
                    labName = col[5]
                    labCode = col[6]
                    revStat = col[7]
                    colMeth = col[8]
                    submitter = col[9].rstrip()

                    submitter = re.sub(r'\s+', '_', submitter) #replace all spaces with an underscore
                    submitter = re.sub(r'/', '-', submitter) # replace all slashes with a hyphen
                    submitter = re.sub(r'\W+', '', submitter) #remove all non-alphanumerics

                    submitter = submitter[0:50]

                    labName = re.sub(r'; ', '', labName) #remove all semi colons
                    labName = re.sub(r'\s+', '_', labName) #replace all spaces with an underscore
                    labName = re.sub(r'/', '-', labName) # replace all slashes with a hyphen
                    labName = re.sub(r'\W+', '', labName) #remove all non-alphanumerics

                    labName = labName[0:50]

                    scvHash[varID][SCV] = {'ClinSig':clinSig, 'DateLastEval':dle, 'LabName':labName, 'LabCode':labCode, 'ReviewStatus':revStat, 'ColMeth':colMeth, 'Condition':condition, 'Submitter':submitter}

    return(scvHash)


def create_files(ExcelDir, excelFile, date):
    '''This function creates an Excel file for each sub in the subList'''

    dir = ExcelDir
    count = 0

    for sub in subList:

        count += 1

        sub_output_file = dir + '/' + sub + '_' + excelFile

        workbook = xlsxwriter.Workbook(sub_output_file)
        worksheet0 = workbook.add_worksheet('README')

        worksheet0.write(0, 0, "Date of ClinVar FTP file: " + date)
        worksheet0.write(2, 0, "Clinical submitter: " + sub)
        worksheet0.write(4, 0, "This Excel file is the output of a script that takes the most recent submission_summary.txt file from the ClinVar FTP site and outputs all the variants for " + sub)
        worksheet0.write(5, 0, 'Each tab is the result of a different set of parameters as outlined below:')
        worksheet0.write(6, 0, '#Variants:')
        worksheet0.write(7, 1, '1. All_subs: All ClinVar variants where there is a GenomeConnect submission.')
        worksheet0.write(8, 1, '2. All_novel: All ClinVar variants where the only submission is from GenomeConnect.')
        worksheet0.write(9, 1, '3. Lab_Conflict: ClinVar variants where the GenomeConnect testing lab clinical significance [P/LP] vs [VUS] vs [LB/B] differs from the clinical lab with same name.')
        worksheet0.write(10, 1, '4. EP_Conflict: ClinVar variants where the GenomeConnect testing lab clinical significance [P/LP] vs [VUS] vs [LB/B] differs from an Expert Panel or Practice Guideline.')
        worksheet0.write(11, 1, '5. Outlier: ClinVar variants where the GenomeConnect testing lab clinical significance [P/LP] vs [VUS] vs [LB/B] differs from at least one 1-star or above (or clinical testing) submitter.')

        worksheet0.write(13, 0, 'Note: Tab classification counts are for unique submissions only i.e. if the same variant is submitted twice as Pathogenic by the same submitter, it will only be counted once')
        worksheet0.write(14, 0, 'Note: A variant can occur in multiple tabs i.e. if the same variant is submitted twice, once as Pathogenic and once as Benign by the same submitter, the variant could be both an outlier and the consensus')

        tabList = [create_tab1, create_tab2, create_tab3, create_tab4, create_tab5]
        for tab in tabList:
            tab(sub, workbook, worksheet0, count)

    workbook.close()


def create_tab1(sub, workbook, worksheet0, count):
    '''This function creates the Tab#1 (All_subs) in the Excel file'''

    worksheet1 = workbook.add_worksheet('1.AllSubs')

    tab = 1
    row = 0
    p2fileVarIDs = {}
    headerSubs = []

    for varID in gcVarIDs:

        subSignificance, submitters, p, lp, plp, vus, lb, b, lbb, vlbb, total, other = get_pathCounts(sub, varID, tab)

        if varID not in p2fileVarIDs.keys():
            p2fileVarIDs[varID] = {}

        p2fileVarIDs[varID] = {'Total':total, 'PLP':plp, 'VUS':vus, 'LBB':lbb, 'Misc':other}

        for SCV in scvHash[varID]:
            if scvHash[varID][SCV]['Submitter'] != sub:
                headerSubs.append(scvHash[varID][SCV]['Submitter'])

    headerSubs = sorted(set(headerSubs))

    print_header(sub, gcVarIDs, headerSubs, worksheet1, tab)

    for varID in gcVarIDs:
        varSubs = get_varSubs(sub, varID)
        row = print_variants(sub, worksheet1, row, varID, headerSubs, varSubs, p2fileVarIDs, tab)

    print_stats(worksheet0, 7, 0, row)


def create_tab2(sub, workbook, worksheet0, count):
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

    print_header(sub, p2fileVarIDs, headerSubs, worksheet2, tab)

    for varID in p2fileVarIDs:
        row = print_variants(sub, worksheet2, row, varID, headerSubs, varSubs, p2fileVarIDs, tab)

    print_stats(worksheet0, 8, 0, row)


def create_tab3(sub, workbook, worksheet0, count):
    '''This function creates the Tab#3 (Lab_Conflicts) in the Excel file'''

    worksheet3 = workbook.add_worksheet('3.LabConflicts')

    tab = 3
    row = 0
    p2fileVarIDs = []
    headerSubs = []

    for varID in gcVarIDs:

        lab = ''
        sig = ''
        for SCV in scvHash[varID]:
            if sub == scvHash[varID][SCV]['Submitter']:
                lab = scvHash[varID][SCV]['LabName']
                sig = scvHash[varID][SCV]['ClinSig']

                for SCV in scvHash[varID]:
                    if sub != scvHash[varID][SCV]['Submitter'] and scvHash[varID][SCV]['Submitter'] == lab and scvHash[varID][SCV]['ClinSig'] != sig:# and (scvHash[varID]['ClinSig'] == 'Pathogenic' or scvHash[varID]['ClinSig'] == 'Likely pathogenic' or \
                    #scvHash[varID]['ClinSig'] == 'Uncertain significance' or scvHash[varID]['ClinSig'] == 'Likely benign' or scvHash[varID]['ClinSig'] == 'Benign'):
                        headerSubs.append(scvHash[varID][SCV]['Submitter'])
                        if varID not in p2fileVarIDs:
                            p2fileVarIDs.append(varID)

    headerSubs = sorted(set(headerSubs))

    print_header(sub, p2fileVarIDs, headerSubs, worksheet3, tab)

    for varID in p2fileVarIDs:
        varSubs = get_varSubs(sub, varID)
        row = print_variants(sub, worksheet3, row, varID, headerSubs, varSubs, p2fileVarIDs, tab)

    print_stats(worksheet0, 9, 0, row)


def create_tab4(sub, workbook, worksheet0, count):
    '''This function creates the Tab#4 (VCEP _Conflict) in the Excel file'''

    worksheet4 = workbook.add_worksheet('4.VCEPConflict')

    tab = 4
    row = 0
    p2fileVarIDs = {}
    headerSubs = []

    for varID in gcVarIDs:
        p2fileVarIDs, headerSubs = Outlier_EP(varID, sub, headerSubs, p2fileVarIDs , tab)

    print_header(sub, p2fileVarIDs, headerSubs, worksheet4, tab)

    for varID in p2fileVarIDs:
        varSubs = get_varSubs(sub, varID)
        row = print_variants(sub, worksheet4, row, varID, headerSubs, varSubs, p2fileVarIDs, tab)

    print_stats(worksheet0, 10, 0, row)


def create_tab5(sub, workbook, worksheet0, count):
    '''This function creates the Tab#5 (Outlier) in the Excel file'''

    worksheet5 = workbook.add_worksheet('5.Outlier_P.VUS.B')

    tab = 5
    row = 0
    p2fileVarIDs = {}
    headerSubs = []

    for varID in gcVarIDs:
        p2fileVarIDs, headerSubs = outlier(varID, sub, headerSubs, p2fileVarIDs, tab)

    print_header(sub, p2fileVarIDs, headerSubs, worksheet5, tab)

    for varID in p2fileVarIDs:
        varSubs = get_varSubs(sub, varID)
        row = print_variants(sub, worksheet5, row, varID, headerSubs, varSubs, p2fileVarIDs, tab)

    print_stats(worksheet0, 11, 0, row)


def outlier(varID, sub, headerSubs, p2fileVarIDs, tab):
    '''This function returns the outlier submitters in a medically significant VarID'''

    subSignificance, submitters, p, lp, plp, vus, lb, b, lbb, vlbb, total, other = get_pathCounts(sub, varID, tab)

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


def Outlier_EP(varID, sub, headerSubs, p2fileVarIDs, tab):
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

                subSignificance, submitters, p, lp, plp, vus, lb, b, lbb, vlbb, total, other = get_pathCounts(sub, varID, tab)

                p2fileVarIDs[varID] = {'Total':total, 'PLP':plp, 'VUS':vus, 'LBB':lbb, 'Misc':other, 'EP':EPHash[varID]['Submitter'], 'EP_clinSig':EPHash[varID]['ClinSig']}
                p2fileVarIDs[varID].update({'EPConflict':EPconflict})

                if submitters:
                    headerSubs.extend(submitters)

    headerSubs = sorted(set(headerSubs))
    if sub in headerSubs:
        headerSubs.remove(sub)

    return(p2fileVarIDs, headerSubs)


def get_pathCounts(sub, varID, tab):
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

        if SCV in scvHash[varID] and scvHash[varID][SCV]['ClinSig'] == 'Pathogenic':
           #Don't double count (Illumina's) duplicate submissions!!!
            current_sub = scvHash[varID][SCV]['Submitter'] + scvHash[varID][SCV]['ClinSig']
            if scvHash[varID][SCV]['Submitter'] != sub and current_sub not in unique_subs and tab == 1:
                unique_subs.append(current_sub)
                submitters.append(scvHash[varID][SCV]['Submitter'])
                p += 1
            if scvHash[varID][SCV]['Submitter'] != sub and current_sub not in unique_subs and tab != 1 and (scvHash[varID][SCV]['ReviewStatus'] == 'criteria provided, single submitter' or 'clinical testing' in scvHash[varID][SCV]['ColMeth']):
                unique_subs.append(current_sub)
                submitters.append(scvHash[varID][SCV]['Submitter'])
                p += 1

        if SCV in scvHash[varID] and scvHash[varID][SCV]['ClinSig'] == 'Likely pathogenic':
            current_sub = scvHash[varID][SCV]['Submitter'] + scvHash[varID][SCV]['ClinSig']
            if scvHash[varID][SCV]['Submitter'] == sub and current_sub not in unique_subs:
                unique_subs.append(current_sub)
                subSignificance.append('P')
                lp += 1

        if SCV in scvHash[varID] and scvHash[varID][SCV]['ClinSig'] == 'Likely pathogenic':
           #Don't double count (Illumina's) duplicate submissions!!!
            current_sub = scvHash[varID][SCV]['Submitter'] + scvHash[varID][SCV]['ClinSig']
            if scvHash[varID][SCV]['Submitter'] != sub and current_sub not in unique_subs and tab == 1:
                unique_subs.append(current_sub)
                submitters.append(scvHash[varID][SCV]['Submitter'])
                lp += 1
            if scvHash[varID][SCV]['Submitter'] != sub and current_sub not in unique_subs and tab != 1 and (scvHash[varID][SCV]['ReviewStatus'] == 'criteria provided, single submitter' or 'clinical testing' in scvHash[varID][SCV]['ColMeth']):
                unique_subs.append(current_sub)
                submitters.append(scvHash[varID][SCV]['Submitter'])
                lp += 1

        if SCV in scvHash[varID] and scvHash[varID][SCV]['ClinSig'] == 'Uncertain significance':
            current_sub = scvHash[varID][SCV]['Submitter'] + scvHash[varID][SCV]['ClinSig']
            if scvHash[varID][SCV]['Submitter'] == sub and current_sub not in unique_subs:
                unique_subs.append(current_sub)
                subSignificance.append('VUS')
                vus += 1

        if SCV in scvHash[varID] and scvHash[varID][SCV]['ClinSig'] == 'Uncertain significance':
           #Don't double count (Illumina's) duplicate submissions!!!
            current_sub = scvHash[varID][SCV]['Submitter'] + scvHash[varID][SCV]['ClinSig']
            if scvHash[varID][SCV]['Submitter'] != sub and current_sub not in unique_subs and tab == 1:
                unique_subs.append(current_sub)
                submitters.append(scvHash[varID][SCV]['Submitter'])
                vus += 1
            if scvHash[varID][SCV]['Submitter'] != sub and current_sub not in unique_subs and tab != 1 and (scvHash[varID][SCV]['ReviewStatus'] == 'criteria provided, single submitter' or 'clinical testing' in scvHash[varID][SCV]['ColMeth']):
                unique_subs.append(current_sub)
                submitters.append(scvHash[varID][SCV]['Submitter'])
                vus += 1

        if SCV in scvHash[varID] and scvHash[varID][SCV]['ClinSig'] == 'Likely benign':
            current_sub = scvHash[varID][SCV]['Submitter'] + scvHash[varID][SCV]['ClinSig']
            if scvHash[varID][SCV]['Submitter'] == sub and current_sub not in unique_subs:
                unique_subs.append(current_sub)
                subSignificance.append('B')
                lb += 1

        if SCV in scvHash[varID] and scvHash[varID][SCV]['ClinSig'] == 'Likely benign':
           #Don't double count (Illumina's) duplicate submissions!!!
            current_sub = scvHash[varID][SCV]['Submitter'] + scvHash[varID][SCV]['ClinSig']
            if scvHash[varID][SCV]['Submitter'] != sub and current_sub not in unique_subs and tab == 1:
                unique_subs.append(current_sub)
                submitters.append(scvHash[varID][SCV]['Submitter'])
                lb += 1
            if scvHash[varID][SCV]['Submitter'] != sub and current_sub not in unique_subs and tab != 1 and (scvHash[varID][SCV]['ReviewStatus'] == 'criteria provided, single submitter' and 'clinical testing' in scvHash[varID][SCV]['ColMeth']):
                unique_subs.append(current_sub)
                submitters.append(scvHash[varID][SCV]['Submitter'])
                lb += 1

        if SCV in scvHash[varID] and scvHash[varID][SCV]['ClinSig'] == 'Benign':
            current_sub = scvHash[varID][SCV]['Submitter'] + scvHash[varID][SCV]['ClinSig']
            if scvHash[varID][SCV]['Submitter'] == sub and current_sub not in unique_subs:
                unique_subs.append(current_sub)
                subSignificance.append('B')
                b += 1

        if SCV in scvHash[varID] and scvHash[varID][SCV]['ClinSig'] == 'Benign':
           #Don't double count (Illumina's) duplicate submissions!!!
            current_sub = scvHash[varID][SCV]['Submitter'] + scvHash[varID][SCV]['ClinSig']
            if scvHash[varID][SCV]['Submitter'] != sub and current_sub not in unique_subs and tab == 1:
                unique_subs.append(current_sub)
                submitters.append(scvHash[varID][SCV]['Submitter'])
                b += 1
            if scvHash[varID][SCV]['Submitter'] != sub and current_sub not in unique_subs and tab != 1 and (scvHash[varID][SCV]['ReviewStatus'] == 'criteria provided, single submitter' and 'clinical testing' in scvHash[varID][SCV]['ColMeth']):
                unique_subs.append(current_sub)
                submitters.append(scvHash[varID][SCV]['Submitter'])
                b += 1

        else:
            current_sub = scvHash[varID][SCV]['Submitter'] + scvHash[varID][SCV]['ClinSig']
            if scvHash[varID][SCV]['Submitter'] == sub and current_sub not in unique_subs:
                unique_subs.append(current_sub)
                other += 1
            #Don't double count (Illumina's) duplicate submissions!!!
            if scvHash[varID][SCV]['Submitter'] != sub and current_sub not in unique_subs and tab == 1:
                unique_subs.append(current_sub)
                submitters.append(scvHash[varID][SCV]['Submitter'])
                other += 1

            if scvHash[varID][SCV]['Submitter'] != sub and current_sub not in unique_subs and tab != 1 and (scvHash[varID][SCV]['ReviewStatus'] == 'criteria provided, single submitter' and 'clinical testing' in scvHash[varID][SCV]['ColMeth']):
                unique_subs.append(current_sub)
                submitters.append(scvHash[varID][SCV]['Submitter'])
                other += 1

    plp = p+lp
    lbb = lb+b
    vlbb = vus+lb+b
    total = plp+vus+lbb+other

    subSignificance = sorted(set(subSignificance))

    return(subSignificance, submitters, p, lp, plp, vus, lb, b, lbb, vlbb, total, other)


def get_varSubs(sub, varID):
    '''This function returns the list of 1-star variant submitters'''

    varSubs = []
    if varID in scvHash:
        for SCV in scvHash[varID]:
            if scvHash[varID][SCV]['Submitter'] != sub: #and scvHash[varID][SCV]['ReviewStatus'] == 'criteria provided, single submitter' and 'clinical testing' in scvHash[varID][SCV]['ColMeth']:
                if scvHash[varID][SCV]['DateLastEval'] != '-':
                    #Convert date from YYYYMMDD -> YYYY-MM-DD
                    subPrintDate = print_date(scvHash[varID][SCV]['DateLastEval'])
                    varSubs.append(scvHash[varID][SCV]['Submitter'] + ' [' + scvHash[varID][SCV]['ClinSig'] + ' (' + subPrintDate + ')]')
                else:
                    varSubs.append(scvHash[varID][SCV]['Submitter'] + ' [' + scvHash[varID][SCV]['ClinSig'] + ' (No DLE)]')

    varSubs = sorted(set(varSubs))

    return(varSubs)


def print_header(sub, gcVarIDs, headerSubs, worksheet, tab):
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
        if tab == 4:
            worksheet.write(0, k, 'EP')
            k+=1
            worksheet.write(0, k, 'EP_significance')
            k+=1
            worksheet.write(0, k, 'EP_conflict')
            k+=1
        else:
            worksheet.write(0, k, 'EP')
            k+=1

        if tab == 5:
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

        for head in headerSubs:
            if head != sub:
                worksheet.write(0, k, head)
                k+=1

        if tab != 2 and tab != 3:
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
        worksheet.write(0, k, 'Submitting_labs')
    else:
        worksheet.write(0, 0, 'No variants found')


def print_variants(sub, worksheet, row, varID, headerSubs, varSubs, p2fileVarIDs, tab):
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

    if tab != 4:
       if varID in EPHash.keys():
           worksheet.write(row, k, EPHash[varID]['Submitter'] + ' (' + EPHash[varID]['ClinSig'] + ')')
       else:
           worksheet.write(row, k, 'N/A')
       k+=1
    if tab == 4:
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
    if tab == 5:
        worksheet.write(row, k, p2fileVarIDs[varID]['Conflict'])
        k+=1

    labs = []
    clinSig = []
    scvs= []
    dle = []
    conditions = []

    for scv in scvHash[varID]:
        if scvHash[varID][scv]['Submitter'] == sub:
            labs.append(scvHash[varID][scv]['LabName'])
            clinSig.append(scvHash[varID][scv]['ClinSig'])
            scvs.append(scv)
            dle.append(print_date(scvHash[varID][scv]['DateLastEval']))
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

    if tab != 2 and tab != 3:
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

    if tab == 2 or tab == 3:
        for headerSub in headerSubs:
            p2file = 'no'
            for varSub in varSubs:
                if headerSub in varSub:
                    p2file = varSub[varSub.find("[")+1:varSub.find("]")]
            if p2file != 'no':
                for scv in scvHash[varID]:
                    if 'LabName' in scvHash[varID][scv] and scvHash[varID][scv]['LabName'] == headerSub:# and scvHash[varID][scv]['ClinSig'] not in clinSig:
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

    create_gtrHash(inputFile1)
    create_orgDict(inputFile2)
    create_gcHash(inputFile3)
    add_labdata(inputFile4)

    tabFile = ExcelDir + '/GCSummary_' + date + '.txt'
    excelFile = 'GCReport_' + date + '.xlsx'

    write2file(tabFile)

    create_scvHash(inputFile3)
    create_a2vHash(inputFile5)
    create_HGVSHash(inputFile6)
    add_GCdata(tabFile)

    create_files(ExcelDir, excelFile, date)

main()
