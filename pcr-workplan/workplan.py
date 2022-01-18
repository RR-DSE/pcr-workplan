# workplan.py (utf-8)
# 
# Edited by: RR-DSE
# Timestamp: 22-01-18 20:39:23

# ------------
# Dependencies
# ------------

import fpdf
import io
import os
import sys
import shutil
import datetime
import xlrd
import glob
import re
import barcode
import importlib
import importlib.util
import tools
import listools
from tkinter import *
from tkinter.ttk import *
from tkcalendar import DateEntry
import tkinter.filedialog
import mysql.connector

# ---------
# Constants
# ---------

lsPCRDefKeys = [
  "protocol",
  "plate_type",
  "well",
  "row",
  "column",
  "well_title",
  "well_ref",
  "type",
  "id",
  "amplicon",
  "amplicon_title",
  "channel",
  "channel_title"
  ]

lsMixIngredientsDefKeys = [
  "protocol",
  "plate_type",
  "type",
  "type_title",
  "well",
  "well_title",
  "ingredient",
  "ingredient_title",
  "volume"]

lsIngredientsDefKeys = [
  "protocol",
  "plate_type",
  "ingredient",
  "title",
  "volume"]

lsSampleKeys = [
  "sample",
  "lis_id",
  "sample_date",
  "name",
  "gender",
  "birthday",
  "record",
  "department"]

sCode = "utf-8"

bLogVerbose = False

sLogFolder = "logs"
sConfigFolder = "config"
sUploadFolder = "data\\upload"
sExperimentsFolder = "data\\experiments"
sCommandsFolder = "commands"
sTemporaryFolder = "temp"

sInputDateFormat = "%Y-%m-%d"
sWorklistSelectorsFile = "{}\\lis_sel_worklist.csv".format(sConfigFolder)
uWorklistSourceStartLine = 16
uWorklistSourceTrimLines = 0
sWorklistInputDateTimeFormat = "%Y-%m-%d %H:%M:%S"
sWorklistSampleRE = r"\d{6}"

sLogFileTimestampFormat = "%Y_%m_%d_%H_%M_%S"
sDateFormat = "%d/%m/%Y"
sDateTimeFormat = "%Y-%m-%d %H:%M:%S"
sFileDateTimeFormat = "%Y_%m_%d"

sHost = "DBHOST"
sUser = "DBUSER"
sPassword = "DBPASSWORD"
sDatabase = "DBTABLESPACE"

# --------------
# Global objects
# --------------

dMySqlDB = None
dMySqlCursor = None

dMemLog = io.StringIO(newline=None)
oLogTimeStamp = None

dConfigReport = None
dConfigOptions = None

sSampleRE = ".*"

oReport = None

dProtocols = None
dPlateTypes = None
dProtocol = None
dPlateType = None
dPCRSlots = dict()
dIngredients = dict()
ldMixIngredientsDef = list()
dIngredientsDef = list()
dMixIngredients = dict()
ldPCRDef = list()

oDateTime = None
sFile = None
bOpened = False

dWorklist = dict()
dSamples = dict()
lsSamples = list()
lsProtocols = list()

dMainDialog = None

dExtraScripts = dict()

# -----------
# Log methods
# -----------

def Log(sText, bSetStatus = False):
  global dMemLog, bLogVerbose, oLogTimeStamp
  if not oLogTimeStamp:
    oLogTimeStamp = datetime.datetime.now()
  dMemLog.write(sText + os.linesep)
  if bLogVerbose:
    print(sText)
  if bSetStatus:
    SetStatusText(sText)

def SaveLog():
  global dMemLog, sLogFolder, sLogFileTimestampFormat, oLogTimeStamp
  if not oLogTimeStamp:
    oLogTimeStamp = datetime.datetime.now()
  dMemLog.seek(0)
  sTimeStamp = oLogTimeStamp.strftime(sLogFileTimestampFormat)
  dFile = open("{}\\{}".format(sLogFolder, sTimeStamp + "_workplan.log"), "w", encoding = "utf-8")
  dFile.write(dMemLog.read())
  dFile.close()

# -------------------
# Auxiliary functions
# -------------------

def GetName(sSource):
  sRes = listools.CleanStr(sSource).upper()
  return sRes

def GetSample(sSource):
  global sSampleRE
  if not sSource:
    return None
  oMatch = re.match(sSampleRE, sSource, flags=re.IGNORECASE)
  if oMatch:
    return oMatch.group(0).strip().upper()
  else:
    return None

def GetDate(sSource = None):
  global sDateFormat, sInputDateFormat
  if not sSource or sSource == "":
    return datetime.datetime.now().strftime(sDateFormat)
  else:
    return datetime.datetime.strptime(sSource.replace("/", "-").replace("\\", "-"), sInputDateFormat).strftime(sDateFormat)

def ConvertDateTime(sSource, sSourceFormat, sDestFormat):
  return datetime.datetime.strptime(sSource, sSourceFormat).strftime(sDestFormat)

def ImportModuleFromPath(sName, sPath):
  oSpec = importlib.util.spec_from_file_location(sName, "{}\\{}.py".format(sPath, sName))
  oModule = importlib.util.module_from_spec(oSpec)
  oSpec.loader.exec_module(oModule)
  return oModule

# -----------------------
# Configuration functions
# -----------------------

def LoadPlateTypes():
  global dPlateTypes, sConfigFolder
  dPlateTypes = dict()
  try:
    lsPlateTypesFiles = glob.glob("{}\\plate_*.csv".format(sConfigFolder))
    for sPlateTypesFile in lsPlateTypesFiles:
      dPlateTypesTable = tools.LoadTable(sPlateTypesFile)
      for dRow in dPlateTypesTable[1]:
        if dRow['type'] not in dPlateTypes:
          dPlateTypes[dRow['type']] = dict()
          dPlateTypes[dRow['type']]['id'] = dRow['type']
        if dRow['parameter'].lower() == 'title':
          dPlateTypes[dRow['type']]['title'] = dRow['title']
          dPlateTypes[dRow['type']]['abbreviation'] = dRow['abbreviation'] if dRow['abbreviation'] != None else ""
        if dRow['parameter'].lower() == 'detector':
          dPlateTypes[dRow['type']]['detector_title'] = dRow['title']
          dPlateTypes[dRow['type']]['detector_abbreviation'] = dRow['abbreviation'] if dRow['abbreviation'] != None else ""
        if dRow['parameter'].lower() == 'corner_pos':
          dPlateTypes[dRow['type']]['corner_pos'] = dRow['ref']
        if dRow['parameter'].lower() == 'row':
          if "rows" not in dPlateTypes[dRow['type']]:
            dPlateTypes[dRow['type']]['rows'] = dict()
          dPlateTypes[dRow['type']]['rows'][dRow['id']] = {
            'title': dRow['title'],
            'abbreviation': dRow['abbreviation'] if dRow['abbreviation'] != None else "",
            'row': int(dRow['row']),
            'ref': dRow['ref']}
        if dRow['parameter'].lower() == 'col':
          if "columns" not in dPlateTypes[dRow['type']]:
            dPlateTypes[dRow['type']]['columns'] = dict()
          dPlateTypes[dRow['type']]['columns'][dRow['id']] = {
            'title': dRow['title'],
            'abbreviation': dRow['abbreviation'] if dRow['abbreviation'] != None else "",
            'column': int(dRow['column']),
            'ref': dRow['ref']}
        if dRow['parameter'].lower() == 'well':
          if "wells" not in dPlateTypes[dRow['type']]:
            dPlateTypes[dRow['type']]['wells'] = dict()
          dPlateTypes[dRow['type']]['wells'][dRow['id']] = {
            'title': dRow['title'],
            'abbreviation': dRow['abbreviation'] if dRow['abbreviation'] != None else "",
            'row': int(dRow['row']),
            'column': int(dRow['column']),
            'ref': dRow['ref']}
  except Exception as dError:
    Log("Erro: Não foi possível carregar pelo menos um dos ficheiros de configuração para tipos de placas de PCR.")
    Log("Mensagem de erro:\n" + str(dError))
    raise
  return

def LoadProtocols():
  global dProtocols, sConfigFolder
  dProtocols = dict()
  try:
    lsProtocolsFiles = glob.glob("{}\\protocol_*.csv".format(sConfigFolder))
    for sProtocolsFile in lsProtocolsFiles:
      dProtocolsTable = tools.LoadTable(sProtocolsFile)
      for dRow in dProtocolsTable[1]:
        if dRow['protocol'] not in dProtocols:
          dProtocols[dRow['protocol']] = dict()
          dProtocols[dRow['protocol']]['id'] = dRow['protocol']
        if dRow['parameter'].lower() == 'title':
          dProtocols[dRow['protocol']]['title'] = dRow['title']
          dProtocols[dRow['protocol']]['abbreviation'] = dRow['abbreviation'] if dRow['abbreviation'] != None else ""
        if dRow['parameter'].lower() == 'plate_type':
          dProtocols[dRow['protocol']]['plate_type'] = dRow['ref']
        if dRow['parameter'].lower() == 'type':
          if "types" not in dProtocols[dRow['protocol']]:
            dProtocols[dRow['protocol']]['types'] = dict()
          dProtocols[dRow['protocol']]['types'][dRow['type']] = {
            'title': dRow['title'],
            'abbreviation': dRow['abbreviation'] if dRow['abbreviation'] != None else ""}
        if dRow['parameter'].lower() == 'ingredient':
          if "ingredients" not in dProtocols[dRow['protocol']]:
            dProtocols[dRow['protocol']]['ingredients'] = dict()
          if int(dRow['id']) not in dProtocols[dRow['protocol']]['ingredients']:
            dProtocols[dRow['protocol']]['ingredients'][int(dRow['id'])] = dict()
          dProtocols[dRow['protocol']]['ingredients'][int(dRow['id'])]['title'] = dRow['title']
          dProtocols[dRow['protocol']]['ingredients'][int(dRow['id'])]['abbreviation'] = dRow['abbreviation'] if dRow['abbreviation'] != None else ""
          dProtocols[dRow['protocol']]['ingredients'][int(dRow['id'])]['safety_abs'] = 0.0 if dRow['safety_abs'] == None else float(dRow['safety_abs'])
          dProtocols[dRow['protocol']]['ingredients'][int(dRow['id'])]['safety_perc'] = 0.0 if dRow['safety_perc'] == None else float(dRow['safety_perc'])
          dProtocols[dRow['protocol']]['ingredients'][int(dRow['id'])]['type'] = dRow['type'] if dRow['type'] != None else ""
        if dRow['parameter'].lower() == 'mix':
          if "mixes" not in dProtocols[dRow['protocol']]:
            dProtocols[dRow['protocol']]['mixes'] = dict()
          if dRow['type'] not in dProtocols[dRow['protocol']]['mixes']:
            dProtocols[dRow['protocol']]['mixes'][dRow['type']] = dict()
          if int(dRow['well']) not in dProtocols[dRow['protocol']]['mixes'][dRow['type']]:
            dProtocols[dRow['protocol']]['mixes'][dRow['type']][int(dRow['well'])] = dict()
          if "ingredients" not in dProtocols[dRow['protocol']]['mixes'][dRow['type']][int(dRow['well'])]:
            dProtocols[dRow['protocol']]['mixes'][dRow['type']][int(dRow['well'])]['ingredients'] = dict()
          if int(dRow['id']) not in dProtocols[dRow['protocol']]['mixes'][dRow['type']][int(dRow['well'])]['ingredients']:
            dProtocols[dRow['protocol']]['mixes'][dRow['type']][int(dRow['well'])]['ingredients'][int(dRow['id'])] = dict()
          dProtocols[dRow['protocol']]['mixes'][dRow['type']][int(dRow['well'])]['ingredients'][int(dRow['id'])]['ingredient'] = int(dRow['ref'])
          dProtocols[dRow['protocol']]['mixes'][dRow['type']][int(dRow['well'])]['ingredients'][int(dRow['id'])]['volume'] = 0.0 if dRow['volume'] == None else float(dRow['volume'])
          dProtocols[dRow['protocol']]['mixes'][dRow['type']][int(dRow['well'])]['ingredients'][int(dRow['id'])]['safety_abs'] = 0.0 if dRow['safety_abs'] == None else float(dRow['safety_abs'])
          dProtocols[dRow['protocol']]['mixes'][dRow['type']][int(dRow['well'])]['ingredients'][int(dRow['id'])]['safety_perc'] = 0.0 if dRow['safety_perc'] == None else float(dRow['safety_perc'])
        if dRow['parameter'].lower() == 'amplicon':
          if "mixes" not in dProtocols[dRow['protocol']]:
            dProtocols[dRow['protocol']]['mixes'] = dict()
          if dRow['type'] not in dProtocols[dRow['protocol']]['mixes']:
            dProtocols[dRow['protocol']]['mixes'][dRow['type']] = dict()
          if int(dRow['well']) not in dProtocols[dRow['protocol']]['mixes'][dRow['type']]:
            dProtocols[dRow['protocol']]['mixes'][dRow['type']][int(dRow['well'])] = dict()
          if "amplicons" not in dProtocols[dRow['protocol']]['mixes'][dRow['type']][int(dRow['well'])]:
            dProtocols[dRow['protocol']]['mixes'][dRow['type']][int(dRow['well'])]['amplicons'] = dict()
          if str(dRow['ref']) not in dProtocols[dRow['protocol']]['mixes'][dRow['type']][int(dRow['well'])]['amplicons']:
            dProtocols[dRow['protocol']]['mixes'][dRow['type']][int(dRow['well'])]['amplicons'][str(dRow['ref'])] = dict()
          dProtocols[dRow['protocol']]['mixes'][dRow['type']][int(dRow['well'])]['amplicons'][str(dRow['ref'])]['channel'] = dRow['channel']
          dProtocols[dRow['protocol']]['mixes'][dRow['type']][int(dRow['well'])]['amplicons'][str(dRow['ref'])]['channel_title'] = dRow['channel_title'] if dRow['channel_title'] != None else ""
        if dRow['parameter'].lower() == 'mix_title':
          if "mixes" not in dProtocols[dRow['protocol']]:
            dProtocols[dRow['protocol']]['mixes'] = dict()
          if dRow['type'] not in dProtocols[dRow['protocol']]['mixes']:
            dProtocols[dRow['protocol']]['mixes'][dRow['type']] = dict()
          if int(dRow['well']) not in dProtocols[dRow['protocol']]['mixes'][dRow['type']]:
            dProtocols[dRow['protocol']]['mixes'][dRow['type']][int(dRow['well'])] = dict()
          dProtocols[dRow['protocol']]['mixes'][dRow['type']][int(dRow['well'])]['title'] = dRow['title']
          dProtocols[dRow['protocol']]['mixes'][dRow['type']][int(dRow['well'])]['abbreviation'] = dRow['abbreviation'] if dRow['abbreviation'] != None else ""
        if dRow['parameter'].lower() == 'amplicon_title':
          if "amplicons" not in dProtocols[dRow['protocol']]:
            dProtocols[dRow['protocol']]['amplicons'] = dict()
          if str(dRow['id']) not in dProtocols[dRow['protocol']]['amplicons']:
            dProtocols[dRow['protocol']]['amplicons'][str(dRow['id'])] = dict()
          dProtocols[dRow['protocol']]['amplicons'][str(dRow['id'])]['title'] = dRow['title']
          dProtocols[dRow['protocol']]['amplicons'][str(dRow['id'])]['abbreviation'] = dRow['abbreviation'] if dRow['abbreviation'] != None else ""
          dProtocols[dRow['protocol']]['amplicons'][str(dRow['id'])]['ref'] = dRow['ref']
        if dRow['parameter'].lower() == 'plate_order':
          if "plate_orders" not in dProtocols[dRow['protocol']]:
            dProtocols[dRow['protocol']]['plate_orders'] = dict()
          if dRow['type'] not in dProtocols[dRow['protocol']]['plate_orders']:
            dProtocols[dRow['protocol']]['plate_orders'][dRow['type']] = dict()
          if int(dRow['id']) not in dProtocols[dRow['protocol']]['plate_orders'][dRow['type']]:
            dProtocols[dRow['protocol']]['plate_orders'][dRow['type']][int(dRow['id'])] = dict()
          dProtocols[dRow['protocol']]['plate_orders'][dRow['type']][int(dRow['id'])][int(dRow['well'])] = dRow['ref']
        if dRow['parameter'].lower() == 'extra_reaction_sample_factor':
          if "extra_reactions" not in dProtocols[dRow['protocol']]:
            dProtocols[dRow['protocol']]['extra_reactions'] = dict()
          dProtocols[dRow['protocol']]['extra_reactions'][int(dRow['id'])] = dict()
          dProtocols[dRow['protocol']]['extra_reactions'][int(dRow['id'])]['type'] = dRow['type']
          dProtocols[dRow['protocol']]['extra_reactions'][int(dRow['id'])]['title'] = dRow['title']
          dProtocols[dRow['protocol']]['extra_reactions'][int(dRow['id'])]['sample_factor'] = int(dRow['volume']) if dRow['volume'] != None else 0
          dProtocols[dRow['protocol']]['extra_reactions'][int(dRow['id'])]['count'] = int(dRow['safety_abs']) if dRow['safety_abs'] else 1
        if dRow['parameter'].lower() == 'instruction':
          if "instructions" not in dProtocols[dRow['protocol']]:
            dProtocols[dRow['protocol']]['instructions'] = dict()
          dProtocols[dRow['protocol']]['instructions'][int(dRow['id'])] = dict()
          dProtocols[dRow['protocol']]['instructions'][int(dRow['id'])]['title'] = dRow['title']
        if dRow['parameter'].lower() == 'max_unknown':
          dProtocols[dRow['protocol']]['max_unknown'] = int(dRow['volume']) if dRow['volume'] != None else 0
        if dRow['parameter'].lower() == 'reaction_volume':
          dProtocols[dRow['protocol']]['reaction_volume'] = int(dRow['volume']) if dRow['volume'] != None else 0
        if dRow['parameter'].lower() == 'lot':
          dProtocols[dRow['protocol']]['lot'] = dRow['title'] if dRow['title'] != None else ""
        if dRow['parameter'].lower() == 'lot_expiration_date':
          dProtocols[dRow['protocol']]['lot_expiration_date'] = dRow['title'] if dRow['title'] != None else ""
        if dRow['parameter'].lower() == 'filter_options':
          if "filter_options" not in dProtocols[dRow['protocol']]:
            dProtocols[dRow['protocol']]['filter_options'] = dict()
          dProtocols[dRow['protocol']]['filter_options'][int(dRow['id'])] = dict()
          dProtocols[dRow['protocol']]['filter_options'][int(dRow['id'])]['data'] = dRow['title']
          dProtocols[dRow['protocol']]['filter_options'][int(dRow['id'])]['channel'] = dRow['channel']
          dProtocols[dRow['protocol']]['filter_options'][int(dRow['id'])]['channel_title'] = dRow['channel_title']
        if dRow['parameter'].lower() == 'workplan_extra':
          dProtocols[dRow['protocol']]['workplan_extra'] = dRow['title'] if dRow['title'] != None else ""
  except Exception as dError:
    Log("Erro: Não foi possível carregar pelo menos um dos ficheiros de configuração para protocolos.")
    Log("Mensagem de erro:\n" + str(dError))
    raise
  return

def LoadConfigReport():
  global dConfigReport, sConfigFolder
  try:
    dConfigReport = tools.LoadDictSimple("{}\\report.csv".format(sConfigFolder), "parameter", "value")
  except Exception as dError:
    Log("Erro: Não foi possível carregar o ficheiro de configuração dos relatórios PDF.")
    Log("Mensagem de erro:\n" + str(dError))
    raise
  return

def LoadConfigOptions():
  global \
    sConfigFolder,\
    dConfigOptions,\
    sUploadFolder,\
    sExperimentsFolder,\
    sSampleRE,\
    sHost,\
    sUser,\
    sPassword,\
    sDatabase
  try:
    dConfigOptions = tools.LoadDictSimple("{}\\options.csv".format(sConfigFolder), "parameter", "value")
    if dConfigOptions['experiments_folder'] == None:
      dConfigOptions['experiments_folder'] = ""
    if dConfigOptions['worklist_folder'] == None:
      dConfigOptions['worklist_folder'] = ""
    if dConfigOptions['sample_expression'] == None or dConfigOptions['sample_expression'] == "":
      dConfigOptions['sample_expression'] = ".*"
    sUploadFolder = dConfigOptions['worklist_folder']
    sExperimentsFolder = dConfigOptions['experiments_folder']
    sSampleRE = dConfigOptions['sample_expression']
    sHost = dConfigOptions['db_host']
    sDatabase = dConfigOptions['db_database']
    sUser = dConfigOptions['db_user']
    sPassword = dConfigOptions['db_password']
  except Exception as dError:
    Log("Erro: Não foi possível carregar o ficheiro de configuração para opções.")
    Log("Mensagem de erro:\n" + str(dError))
    raise
  return

def LoadWorklist():
  global \
    dWorklist,\
    sUploadFolder,\
    uWorklistSourceStartLine,\
    uWorklistSourceTrimLines,\
    sWorklistSelectorsFile,\
    dConfigOptions
  try:
    dWorklist = dict()
    sSourceFile = "{}\\lista_trabalho.tsv".format(sUploadFolder)
    dSource = listools.GetSelectTable(sSourceFile, uWorklistSourceStartLine, uWorklistSourceTrimLines, sWorklistSelectorsFile)[1]
    for dRow in dSource:
      sResult = str(dRow['result']).strip() if dRow['result'] else ""
      if dRow['test_code'] == "SCOV2" and (sResult == "" or sResult == None):
        sLISID = str(dRow['lisid']).strip().upper()
        sSample = GetSample(str(dRow['sample']))
        sBirthday = GetDate(str(dRow['birthday']).strip().upper())
        sGender = str(dRow['gender']).strip().upper()
        sRecord = str(dRow['record']).strip().upper() if dRow['record'] else ""
        sDepartment = str(dRow['department']).strip().upper() if dRow['department'] else ""
        sName = GetName(str(dRow['name']).strip().upper())
        sSampleDate = GetDate(str(dRow['sample_date']).strip().upper())
        dNewDict = dict()
        dNewDict['sample'] = sSample
        dNewDict['lis_id'] = sLISID
        dNewDict['sample_date'] = sSampleDate
        dNewDict['name'] = sName
        dNewDict['gender'] = sGender
        dNewDict['birthday'] = sBirthday
        dNewDict['record'] = sRecord
        dNewDict['department'] = sDepartment
        dWorklist[sSample] = dNewDict
  except Exception as dError:
    Log("Erro: Não foi possível carregar o ficheiro da lista de trabalho.", True)
    Log("Mensagem de erro:\n" + str(dError))
  return

# ------------------
# Database functions
# ------------------

def MySqlCursorStart():
  global dMySqlDB, dMySqlCursor
  global sHost, sUser, sPassword, sDatabase
  dMySqlDB = mysql.connector.connect(host = sHost, user = sUser, passwd = sPassword, database = sDatabase)
  dMySqlCursor = dMySqlDB.cursor(dictionary = True)
  return

def MySqlCursorClose():
  global dMySqlDB, dMySqlCursor
  dMySqlCursor.close()
  dMySqlDB.close()
  return

def RunMySqlCursor(sSQLQuery, bFetch = False):
  global dMySqlCursor, sDateFormat, sDateTimeFormat
  dRes = list()
  dMySqlCursor.execute(sSQLQuery)
  if dMySqlCursor.rowcount == 0 or not bFetch:
    return dRes
  else:
    dRes = dMySqlCursor.fetchall()
    for dRow in dRes:
      for sKey in dRow:
        if isinstance(dRow[sKey], int):
          dRow[sKey] = str(dRow[sKey])
        elif isinstance(dRow[sKey], float):
          dRow[sKey] = str(int(dRow[sKey]))
        elif isinstance(dRow[sKey], datetime.datetime):
          dRow[sKey] = dRow[sKey].strftime(sDateTimeFormat)
        elif isinstance(dRow[sKey], datetime.date):
          dRow[sKey] = dRow[sKey].strftime(sDateFormat)
        elif dRow[sKey] == None:
          dRow[sKey] = "NULL"
        else:
          dRow[sKey] = str(dRow[sKey])
  return dRes

# -----------------
# General functions
# -----------------

def FindPCRSlot(sType):
  global dProtocol, dPCRSlots
  uSlot = None
  luKeys = list(dProtocol['plate_orders'][sType].keys())
  luKeys.sort()
  for uNewSlot in luKeys:
    if uNewSlot not in dPCRSlots[sType]:
      uSlot = uNewSlot
      break
  return uSlot

def SearchPCRSlot(sType, sID):
  global dPCRSlots
  dRes = None
  if not dPCRSlots or sType not in dPCRSlots:
    return None
  for uSlot, dSlot in dPCRSlots[sType].items():
    if dSlot['id'].lower() == sID.lower():
      dRes = dSlot
      break
  return dRes

def AddPCR(sType, sID):
  global \
    dProtocol,\
    dPlateType,\
    dPCRSlots,\
    dMixIngredients,\
    dIngredients
  if sType == "unknown":
    if "extra_reactions" in dProtocol:
      luIDs = list(dProtocol['extra_reactions'].keys())
      luIDs.sort()
      for uID in luIDs:
        sExtraType = dProtocol['extra_reactions'][uID]['type']
        sExtraTitle = dProtocol['extra_reactions'][uID]['title']
        uSampleFactor = dProtocol['extra_reactions'][uID]['sample_factor']
        uAddCount = dProtocol['extra_reactions'][uID]['count']
        uSampleCount = len(dPCRSlots['unknown']) + 1 if "unknown" in dPCRSlots else 1
        uExtraCount = len(dPCRSlots[sExtraType].keys()) if sExtraType in dPCRSlots else 0
        if uSampleFactor == 0:
          if uExtraCount < uAddCount:
            for uAddIndex in range(0, uAddCount):
              bRes = AddPCR(sExtraType, sExtraTitle)
              if not bRes:
                 return False
        elif float(uSampleCount) / float(uSampleFactor) > float(uExtraCount):
          bRes = AddPCR(sExtraType, sExtraTitle + "_" + str(uExtraCount + 1))
          if not bRes:
            return False
  if sType not in dPCRSlots:
    dPCRSlots[sType] = dict()
  uSlot = FindPCRSlot(sType)
  if not uSlot:
    return False
  if uSlot not in dPCRSlots[sType]:
    dPCRSlots[sType][uSlot] = dict()
  dPCRSlots[sType][uSlot]['id'] = sID
  dPCRSlots[sType][uSlot]['wells'] = dict()
  for uWell, sWellRef in dProtocol['plate_orders'][sType][uSlot].items():
    dPCRSlots[sType][uSlot]['wells'][uWell] = dict()
    dPCRSlots[sType][uSlot]['wells'][uWell]['well'] = sWellRef
    dPCRSlots[sType][uSlot]['wells'][uWell]['title'] = dPlateType['wells'][sWellRef]['title']
    dPCRSlots[sType][uSlot]['wells'][uWell]['abbreviation'] = dPlateType['wells'][sWellRef]['abbreviation']
    dPCRSlots[sType][uSlot]['wells'][uWell]['row'] = dPlateType['wells'][sWellRef]['row']
    dPCRSlots[sType][uSlot]['wells'][uWell]['column'] = dPlateType['wells'][sWellRef]['column']
    dPCRSlots[sType][uSlot]['wells'][uWell]['ref'] = dPlateType['wells'][sWellRef]['ref']
    dPCRSlots[sType][uSlot]['wells'][uWell]['mix_title'] = dProtocol['mixes'][sType][uWell]['title']
    dPCRSlots[sType][uSlot]['wells'][uWell]['mix_abbreviation'] = dProtocol['mixes'][sType][uWell]['abbreviation']
    dPCRSlots[sType][uSlot]['wells'][uWell]['type_title'] = dProtocol['types'][sType]['title']
    dPCRSlots[sType][uSlot]['wells'][uWell]['type_abbreviation'] = dProtocol['types'][sType]['abbreviation']
  if sType not in dMixIngredients:
    dMixIngredients[sType] = dict()
  for uWell in dProtocol['mixes'][sType]:
    if uWell not in dMixIngredients[sType]:
      dMixIngredients[sType][uWell] = dict()
    for uIngredient, dIngredient in dProtocol['mixes'][sType][uWell]['ingredients'].items():
      dMixIngredients[sType][uWell][uIngredient]=dict()
      dMixIngredients[sType][uWell][uIngredient]['ingredient'] = dIngredient['ingredient']
      dMixIngredients[sType][uWell][uIngredient]['title'] = dProtocol['ingredients'][dIngredient['ingredient']]['title']
      dMixIngredients[sType][uWell][uIngredient]['abbreviation'] = dProtocol['ingredients'][dIngredient['ingredient']]['abbreviation']
      dMixIngredients[sType][uWell][uIngredient]['volume'] = dIngredient['volume']
      dMixIngredients[sType][uWell][uIngredient]['safety_abs'] = dIngredient['safety_abs']
      dMixIngredients[sType][uWell][uIngredient]['safety_perc'] = dIngredient['safety_perc']
      if dIngredient['ingredient'] not in dIngredients:
        dIngredients[dIngredient['ingredient']] = dict()
        dIngredients[dIngredient['ingredient']]['volume'] = 0.0
        dIngredients[dIngredient['ingredient']]['type'] = dProtocol['ingredients'][dIngredient['ingredient']]['type']
        dIngredients[dIngredient['ingredient']]['title'] = dProtocol['ingredients'][dIngredient['ingredient']]['title']
        dIngredients[dIngredient['ingredient']]['abbreviation'] = dProtocol['ingredients'][dIngredient['ingredient']]['abbreviation']
        dIngredients[dIngredient['ingredient']]['safety_abs'] = dProtocol['ingredients'][dIngredient['ingredient']]['safety_abs']
        dIngredients[dIngredient['ingredient']]['safety_perc'] = dProtocol['ingredients'][dIngredient['ingredient']]['safety_perc']
      dIngredients[dIngredient['ingredient']]['volume'] = dIngredients[dIngredient['ingredient']]['volume'] + dIngredient['volume'] + dIngredient['safety_abs'] + (dIngredient['volume'] * dIngredient['safety_perc'] / 100.0)
  return True

def GetPlateWell(uRow, uColumn):
  global dPlateType
  sRes = None
  for sWell, dWell in dPlateType['wells'].items():
    if uRow == dWell['row'] and uColumn == dWell['column']:
      sRes = sWell
      break
  return sRes
   
def GetPlateRow(uRow):
  global dPlateType
  sRes = None
  for sRow, dRow in dPlateType['rows'].items():
    if uRow == dRow['row']:
      sRes = sRow
      break
  return sRes

def GetPlateColumn(uColumn):
  global dPlateType
  sRes = None
  for sColumn, dColumn in dPlateType['columns'].items():
    if uColumn == dColumn['column']:
      sRes = sColumn
      break
  return sRes

def SetPCRDef():
  global dPCRSlots, ldPCRDef, dProtocol, dPlateType
  ldPCRDef = list()
  for sType, dType in dPCRSlots.items():
    for uSlot, dSlot in dType.items():
      for uWell, dWell in dSlot['wells'].items():
        for sAmplicon, dAmplicon in dProtocol['mixes'][sType][uWell]['amplicons'].items():
          sAmpliconTitle = dProtocol['amplicons'][sAmplicon]['title']
          sAmpliconChannel = dAmplicon['channel']
          sAmpliconChannelTitle = dAmplicon['channel_title']
          dNewDict = dict()
          dNewDict['protocol'] = dProtocol['id']
          dNewDict['plate_type'] = dPlateType['id']
          dNewDict['well'] = dWell['well']
          dNewDict['row'] = dWell['row']
          dNewDict['column'] = dWell['column']
          dNewDict['well_title'] = dWell['title']
          dNewDict['well_ref'] = dWell['ref']
          dNewDict['type'] = sType
          dNewDict['id'] = dSlot['id']
          dNewDict['amplicon'] = sAmplicon
          dNewDict['amplicon_title'] = sAmpliconTitle
          dNewDict['channel'] = sAmpliconChannel
          dNewDict['channel_title'] = sAmpliconChannelTitle
          ldPCRDef.append(dNewDict)
  return

def RunExtraScript():
  global dExtraScripts, dProtocol, dPlateType, sFile, oDateTime, sConfigFolder, sExperimentsFolder, sUploadFolder, sCommandsFolder, sTemporaryFolder, sLogFolder, dSamples, lsSamples, dPCRSlots, ldPCRDef, dConfigReport, dConfigOptions
  dExtraData = dict()
  dExtraData['experiments_folder'] = sExperimentsFolder
  dExtraData['config_folder'] = sConfigFolder
  dExtraData['upload_folder'] = sUploadFolder
  dExtraData['commands_folder'] = sCommandsFolder
  dExtraData['temporary_folder'] = sTemporaryFolder
  dExtraData['log_folder'] = sLogFolder
  dExtraData['protocol'] = dProtocol
  dExtraData['plate_type'] = dPlateType
  dExtraData['file'] = sFile
  dExtraData['date_time'] = oDateTime
  dExtraData['samples'] = dSamples
  dExtraData['sample_list'] = lsSamples
  dExtraData['pcr_slots'] = dPCRSlots
  dExtraData['pcr_def'] = ldPCRDef
  dExtraData['config_report'] = dConfigReport
  dExtraData['config_options'] = dConfigOptions
  if dProtocol and "workplan_extra" in dProtocol:
    if dProtocol['workplan_extra'] not in dExtraScripts:
      dExtraScripts[dProtocol['workplan_extra']] = ImportModuleFromPath(dProtocol['workplan_extra'], sConfigFolder)
    dExtraScripts[dProtocol['workplan_extra']].Workplan(dExtraData)

def GetPlateCode():
  global dExtraScripts, dProtocol, dPlateType, sFile, oDateTime, sConfigFolder, sExperimentsFolder, sUploadFolder, sCommandsFolder, sTemporaryFolder, sLogFolder, dSamples, lsSamples, dPCRSlots, ldPCRDef, dConfigReport, dConfigOptions
  dExtraData = dict()
  dExtraData['experiments_folder'] = sExperimentsFolder
  dExtraData['config_folder'] = sConfigFolder
  dExtraData['upload_folder'] = sUploadFolder
  dExtraData['commands_folder'] = sCommandsFolder
  dExtraData['temporary_folder'] = sTemporaryFolder
  dExtraData['log_folder'] = sLogFolder
  dExtraData['protocol'] = dProtocol
  dExtraData['plate_type'] = dPlateType
  dExtraData['file'] = sFile
  dExtraData['date_time'] = oDateTime
  dExtraData['samples'] = dSamples
  dExtraData['sample_list'] = lsSamples
  dExtraData['pcr_slots'] = dPCRSlots
  dExtraData['pcr_def'] = ldPCRDef
  dExtraData['config_report'] = dConfigReport
  dExtraData['config_options'] = dConfigOptions
  if dProtocol and "workplan_extra" in dProtocol:
    if dProtocol['workplan_extra'] not in dExtraScripts:
      dExtraScripts[dProtocol['workplan_extra']] = ImportModuleFromPath(dProtocol['workplan_extra'], sConfigFolder)
    return dExtraScripts[dProtocol['workplan_extra']].GetPlateCode(dExtraData)
  else:
    return None

#--------------
# PDF functions
#-------------

def StartReport():
  global oReport, dConfigReport
  oReport = fpdf.FPDF(
    orientation = dConfigReport['page_orientation'],
    unit = dConfigReport['page_unit'],
    format = dConfigReport['page_format'])
  oReport.set_margins(
    dConfigReport['page_margin_left'],
    dConfigReport['page_margin_top'],
    dConfigReport['page_margin_right'])
  return

def AddPage():
  global oReport, dConfigReport
  oReport.add_page()
  return

def SetFont(sStyle):
  global oReport, dConfigReport
  oReport.set_font(
    dConfigReport["font_family_" + sStyle.lower()] if dConfigReport["font_family_" + sStyle.lower()] != None else "",
    dConfigReport["font_style_" + sStyle.lower()] if dConfigReport["font_style_" + sStyle.lower()] != None else "",
    dConfigReport["font_size_" + sStyle.lower()] if dConfigReport["font_size_" + sStyle.lower()] != None else 1)
  oReport.set_text_color(
    dConfigReport["font_color_r_" + sStyle.lower()] if dConfigReport["font_color_r_" + sStyle.lower()] != None else 0,
    dConfigReport["font_color_g_" + sStyle.lower()] if dConfigReport["font_color_g_" + sStyle.lower()] != None else 0,
    dConfigReport["font_color_b_" + sStyle.lower()] if dConfigReport["font_color_b_" + sStyle.lower()] != None else 0)
  return

def SetPos(fX, fY):
  global oReport, dConfigReport
  oReport.set_xy(dConfigReport['page_margin_left'] + fX, dConfigReport['page_margin_top'] + fY)
  return

def GetPos():
  global oReport, dConfigReport
  fCurrX = oReport.get_x() - dConfigReport['page_margin_left']
  fCurrY = oReport.get_y() - dConfigReport['page_margin_top']
  return (fCurrX, fCurrY)

def MovePos(fX, fY):
  global oReport, dConfigReport
  fCurrX = oReport.get_x()
  fCurrY = oReport.get_y()
  oReport.set_xy(fCurrX + fX, fCurrY + fY)
  return GetPos()

def GetStringWidth(sText, sStyle = None):
  global oReport, dConfigReport
  if sStyle:
    SetFont(sStyle)
  return oReport.get_string_width(sText)

def Write(sText, sStyle):
  global oReport, dConfigReport
  SetFont(sStyle)
  oReport.write(dConfigReport["font_height_" + sStyle.lower()], sText)
  return

def WriteLine(sText, sStyle, uAdvance = 0, fPosX = 0.0):
  global oReport, dConfigReport
  MovePos(dConfigReport["page_section_advance"] * float(uAdvance) + fPosX, 0.0)
  Write(sText + "\n", sStyle)
  MovePos(0.0, dConfigReport["font_sep_" + sStyle.lower()])
  return

def DrawCell(sText, sStyle, fWidth, fHeight, sAlign, sBorder, bFill = False, uAdvance = 0, fPosX = 0.0, tuFillColor = None):
  global oReport, dConfigReport
  if not fHeight:
    fActualHeight = dConfigReport["font_height_" + sStyle.lower()]
  else:
    fActualHeight = fHeight
  MovePos(dConfigReport["page_section_advance"] * float(uAdvance) + fPosX, 0.0)
  SetFont(sStyle)
  if bFill:
    if tuFillColor:
      oReport.set_fill_color(tuFillColor[0], tuFillColor[1], tuFillColor[2])
  oReport.cell(fWidth, h = fActualHeight, txt = sText, border = sBorder.upper(), ln = 1, align = sAlign, fill = bFill)
  return

def DrawRule(fWidth, sStyle = None):
  global oReport, dConfigReport
  fActualX1 = oReport.get_x() + dConfigReport['page_char_advance']
  fActualY1 = oReport.get_y()
  fActualX2 = fActualX1 + fWidth
  fActualY2 = fActualY1
  oReport.line(fActualX1, fActualY1, fActualX2, fActualY2)
  if sStyle:
    MovePos(0.0, dConfigReport["font_sep_" + sStyle.lower()])
  return

def DrawLine(fX1, fY1, fX2, fY2):
  global oReport, dConfigReport
  if not fX1:
    fActualX1 = oReport.get_x() + dConfigReport['page_char_advance']
  else:
    fActualX1 = dConfigReport['page_margin_left'] + fX1 + dConfigReport['page_char_advance']
  if not fY1:
    fActualY1 = oReport.get_y()
  else:
    fActualY1 = dConfigReport['page_margin_top'] + fY1
  if not fX2:
    fActualX2 = oReport.get_x() + dConfigReport['page_char_advance']
  else:
    fActualX2 = dConfigReport['page_margin_left'] + fX2 + dConfigReport['page_char_advance']
  if not fY2:
    fActualY2 = oReport.get_y()
  else:
    fActualY2 = dConfigReport['page_margin_top'] + fY2
  oReport.line(fActualX1, fActualY1, fActualX2, fActualY2)
  return

def DrawCircle(fX1, fY1, fRadius):
  global oReport, dConfigReport
  if not fX1:
    fActualX1 = oReport.get_x() + dConfigReport['page_char_advance']
  else:
    fActualX1 = dConfigReport['page_margin_left'] + fX1 + dConfigReport['page_char_advance']
  if not fY1:
    fActualY1 = oReport.get_y()
  else:
    fActualY1 = dConfigReport['page_margin_top'] + fY1
  oReport.ellipse(fActualX1 - fRadius, fActualY1 - fRadius, fRadius * 2.0, fRadius * 2.0)
  return

def Image(sFile, fX, fY, fWidth = None, fHeight = None):
  global oReport, dConfigReport
  if not fX:
    fActualX = oReport.get_x() + dConfigReport['page_char_advance']
  else:
    fActualX = dConfigReport['page_margin_left'] + fX + dConfigReport['page_char_advance']
  if not fY:
    fActualY = oReport.get_y()
  else:
    fActualY = dConfigReport['page_margin_top'] + fY
  oReport.image(sFile, fActualX, fActualY, w=fWidth, h=fHeight)
  return

def InfoLine(sCaption, sText, uStyle, uAdvance = 0, fPosX = 0.0):
  global oReport, dConfigReport
  MovePos(dConfigReport["page_section_advance"] * float(uAdvance) + fPosX, 0.0)
  Write(sCaption + ": ", "info_caption_" + str(uStyle))
  WriteLine(sText, "info_text_" + str(uStyle))
  return

def GenerateWorkplanReport():
  global \
    oReport,\
    dConfigReport,\
    dSamples,\
    lsSamples,\
    dProtocol,\
    dPlateType,\
    dPCRSlots,\
    dMixIngredients,\
    dIngredients,\
    sFile,\
    oDateTime,\
    sFileDateTimeFormat,\
    ldPCRDef,\
    sTemporaryFolder,\
    sDateFormat
  StartReport()
  AddPage()
  SetPos(0,0)
  # Report header
  WriteLine("#INSTITUTION# | #DEPARTMENT#", "header_1")
  WriteLine("Plano de trabalho para protocolo qPCR / RT-qPCR", "title_1")
  InfoLine("Protocolo", dProtocol['title'], 1, 0)
  InfoLine("Data e hora", oDateTime.strftime(sDateTimeFormat), 1, 0)
  InfoLine("Experiência", sFile, 1, 0)
  DrawRule(dConfigReport['page_body_width'], None)
  MovePos(0.0, dConfigReport['page_header_sep'])
  # Samples section
  # WriteLine("Amostras", "section_1")
  sSampleLine = ""
  fMaxWidth = dConfigReport['page_body_width'] - dConfigReport['page_section_advance']
  # for sSample in lsSamples:
  #  if GetStringWidth(sSampleLine + " | " + sSample, "normal_bold") >= fMaxWidth:
  #    WriteLine(sSampleLine, "normal_bold", 1)
  #    sSampleLine = sSample
  #  else:
  #    if sSampleLine == "":
  #      sSampleLine = sSample
  #    else:
  #      sSampleLine = sSampleLine + " | " + sSample
  # if sSampleLine != "":
  #  WriteLine(sSampleLine, "normal_bold", 1)
  # MovePos(0.0, dConfigReport['page_section_sep'])
  # Report definitions section
  WriteLine("Definições", "section_1")
  InfoLine("Experiência", sFile, 2, 1)
  InfoLine("Detetor", dPlateType['detector_title'], 2, 1)
  if "reaction_volume" in dProtocol:
    InfoLine("Volume de reação", str(dProtocol['reaction_volume']) + "uL", 2, 1)
  MovePos(0.0, dConfigReport['page_section_sep'])
  # Report instructions section
  if "instructions" in dProtocol:
    WriteLine("Especificidades e outras notas", "section_1")
    luInstructionIDs = list(dProtocol['instructions'].keys())
    luInstructionIDs.sort()
    for uID in luInstructionIDs:
      sTitle = dProtocol['instructions'][uID]['title']
      sTitle = str(uID) + ". " + sTitle.replace("%F%",sFile)
      WriteLine(sTitle, "small_1", 1)
    MovePos(0.0, dConfigReport['page_section_sep'])
  # Report mixes section
  WriteLine("Misturas de reação", "section_1")
  luIngredients = list(dIngredients.keys())
  luIngredients.sort()
  for uIngredient in luIngredients:
    if dIngredients[uIngredient]['type'].lower() == "hidden":
      continue
    dIngredients[uIngredient]['volume'] = dIngredients[uIngredient]['volume'] + dIngredients[uIngredient]['safety_abs'] + (dIngredients[uIngredient]['volume'] * dIngredients[uIngredient]['safety_perc'] / 100.0)
    InfoLine(dIngredients[uIngredient]['title']+" (\""+dIngredients[uIngredient]['abbreviation']+"\")", str(dIngredients[uIngredient]['volume']) + "uL", 3, 1)
  fPosX = 0.0
  uPos = 0
  fMaxPosY = GetPos()[1]
  SetPos(0.0, fMaxPosY)
  fPosY = fMaxPosY
  for sType, dType in dProtocol['types'].items():
    if sType not in dMixIngredients:
      continue
    sTypeTitle = dType['title']
    luWells = list(dProtocol['mixes'][sType].keys())
    luWells.sort()
    for uWell in luWells:
      if uPos == 0:
        fPosX = 0.0
        fMaxPosY = fMaxPosY + dConfigReport['page_subsection_sep']
        fPosY = fMaxPosY
        SetPos(0.0, fMaxPosY)
      else:
        fPosX = (dConfigReport['page_body_width'] - dConfigReport['page_section_advance']) / dConfigReport['ingredients_columns'] * uPos
        SetPos(0.0, fPosY)
      sWellTitle = dProtocol['mixes'][sType][uWell]['abbreviation']
      luIngredients = list(dMixIngredients[sType][uWell].keys())
      luIngredients.sort()
      WriteLine(sTypeTitle + " | " + sWellTitle, "section_3", 1, fPosX)
      for uIngredient in luIngredients:
        sIngredientTitle = str(uIngredient) + ". " + dMixIngredients[sType][uWell][uIngredient]['abbreviation']
        sVolume = str(dMixIngredients[sType][uWell][uIngredient]['volume']) + "uL"
        InfoLine(sIngredientTitle, sVolume, 4, 1, fPosX)
      fCurrPosY = GetPos()[1]
      if fCurrPosY > fMaxPosY:
          fMaxPosY = fCurrPosY
      uPos = uPos + 1 if uPos < dConfigReport['ingredients_columns'] - 1 else 0
  MovePos(0.0, dConfigReport['page_section_sep'])
  # Report plate layout section
  WriteLine("Configuração da placa", "section_1")
  InfoLine("Tipo de placa", dPlateType['title'], 2, 1)
  sPlateID = GetPlateCode()
  if sPlateID:
    InfoLine("ID da placa", sPlateID, 2, 1)
    tCurrPos = GetPos()
    SetPos(dConfigReport['plate_margin_left'] + tCurrPos[0] + dConfigReport['plate_width'] / 2, tCurrPos[1] - dConfigReport['barcode_image_height'])
    oBarcode = barcode.get("code128", sPlateID, writer=barcode.writer.ImageWriter())
    oBarcode.save("{}\\barcode_plate_id".format(sTemporaryFolder), {"module_width": dConfigReport['barcode_module_width'], "module_height": dConfigReport['barcode_module_height'], "quiet_zone": dConfigReport['barcode_quiet_zone'], "text_distance": dConfigReport['barcode_text_distance'], "font_size": int(dConfigReport['barcode_font_size'])})
    Image("{}\\barcode_plate_id.png".format(sTemporaryFolder), None, None, 0, float(dConfigReport['barcode_image_height']))
    SetPos(tCurrPos[0], tCurrPos[1])
  MovePos(0.0, dConfigReport['page_subsection_sep'])
  fWellWidth = (dConfigReport['plate_width'] - dConfigReport['plate_margin_left']) / len(dPlateType['columns'].keys())
  fWellRadius = fWellWidth / 2.0 - dConfigReport['plate_well_sep'] / 2.0
  fStartPosX = dConfigReport['page_section_advance']
  fStartPosY = GetPos()[1]
  fPosX = fStartPosX + dConfigReport['plate_margin_left'] + fWellWidth / 2.0
  fPosY = fStartPosY
  for uIndex in range(1, len(dPlateType['columns'].keys())+1):
    sColumn = dPlateType['columns'][GetPlateColumn(uIndex)]['abbreviation']
    fPosX = fStartPosX + dConfigReport['plate_margin_left'] + float(uIndex - 1) * fWellWidth + dConfigReport['page_char_advance']
    SetPos(fPosX, fPosY)
    DrawCell(sColumn, "label_1", fWellWidth, None, "C", "")
  fPosX = fStartPosX + dConfigReport['plate_column_label_disp_h']
  fCellWidth = dConfigReport['plate_margin_left']
  fPosY = fStartPosY + dConfigReport['plate_margin_top'] - dConfigReport['plate_column_label_disp_v']
  for uIndex in range(1, len(dPlateType['rows'].keys())+1):
    sRow = dPlateType['rows'][GetPlateRow(uIndex)]['abbreviation']
    SetPos(fPosX, fPosY)
    DrawCell(sRow, "label_1", fCellWidth, fWellWidth, "L", "")
    fPosY = fPosY + fWellWidth
  fPosX = fStartPosX + dConfigReport['plate_margin_left'] + fWellWidth / 2.0
  fPosY = fStartPosY + dConfigReport['plate_margin_top'] + fWellWidth / 2.0
  for uIndex in range(1, len(dPlateType['rows'].keys())+1):
    for uIndex in range(1, len(dPlateType['columns'].keys())+1):
      DrawCircle(fPosX, fPosY, fWellRadius)
      fPosX = fPosX + fWellWidth
    fPosX = fStartPosX + dConfigReport['plate_margin_left'] + fWellWidth / 2.0
    fPosY = fPosY + fWellWidth
  fPlateHeight = dConfigReport['plate_margin_top'] + (fWellWidth * len(dPlateType['rows'].keys()))
  if dPlateType['corner_pos'] == "UL":
    tfCorner1 = (fStartPosX + dConfigReport['plate_corner_width'], fStartPosY)
    tfCorner2 = (fStartPosX, fStartPosY + dConfigReport['plate_corner_width'])
    tfVert2 = (fStartPosX + dConfigReport['plate_width'], fStartPosY)
    tfVert3 = (fStartPosX, fStartPosY + fPlateHeight)
    tfVert4 = (fStartPosX + dConfigReport['plate_width'], fStartPosY + fPlateHeight)
  elif dPlateType['corner_pos'] == "UR":
    tfCorner1 = (fStartPosX + dConfigReport['plate_width'] - dConfigReport['plate_corner_width'], fStartPosY)
    tfCorner2 = (fStartPosX + dConfigReport['plate_width'], fStartPosY + dConfigReport['plate_corner_width'])
    tfVert2 = (fStartPosX, fStartPosY)
    tfVert3 = (fStartPosX + dConfigReport['plate_width'], fStartPosY + fPlateHeight)
    tfVert4 = (fStartPosX, fStartPosY + fPlateHeight)
  elif dPlateType['corner_pos'] == "DL":
    tfCorner1 = (fStartPosX, fStartPosY + fPlateHeight - dConfigReport['plate_corner_width'])
    tfCorner2 = (fStartPosX + dConfigReport['plate_corner_width'], fStartPosY + fPlateHeight)
    tfVert2 = (fStartPosX, fStartPosY)
    tfVert3 = (fStartPosX + dConfigReport['plate_width'], fStartPosY + fPlateHeight)
    tfVert4 = (fStartPosX + dConfigReport['plate_width'], fStartPosY)
  elif dPlateType['corner_pos'] == "DR":
    tfCorner1 = (fStartPosX + dConfigReport['plate_width'], fStartPosY + fPlateHeight - dConfigReport['plate_corner_width'])
    tfCorner2 = (fStartPosX + dConfigReport['plate_width'] - dConfigReport['plate_corner_width'], fStartPosY + fPlateHeight)
    tfVert2 = (fStartPosX + dConfigReport['plate_width'], fStartPosY)
    tfVert3 = (fStartPosX, fStartPosY + fPlateHeight)
    tfVert4 = (fStartPosX, fStartPosY)
  DrawLine(tfCorner1[0], tfCorner1[1], tfVert2[0], tfVert2[1])
  DrawLine(tfCorner2[0], tfCorner2[1], tfVert3[0], tfVert3[1])
  DrawLine(tfCorner1[0], tfCorner1[1], tfCorner2[0], tfCorner2[1])
  DrawLine(tfVert3[0], tfVert3[1], tfVert4[0], tfVert4[1])
  DrawLine(tfVert2[0], tfVert2[1], tfVert4[0], tfVert4[1])
  for sType, dType in dPCRSlots.items():
    for uSlot, dSlot in dType.items():
      sID = dSlot['id']
      for uWell, dWell in dSlot['wells'].items():
        sWellCaption = dWell['mix_abbreviation']
        sTypeCaption = dWell['type_abbreviation']
        fPosX = fStartPosX + dConfigReport['plate_margin_left'] + float(dWell['column'] - 1) * fWellWidth + dConfigReport['page_char_advance']
        fPosY = fStartPosY + dConfigReport['plate_margin_top'] + (fWellWidth / 2.0) + (float(dWell['row'] - 1) * fWellWidth) - dConfigReport['plate_well_label_disp_v_1']
        SetPos(fPosX, fPosY)
        DrawCell(sID, "label_2", fWellWidth, None, "C", "")
        fPosY = fStartPosY + dConfigReport['plate_margin_top'] + (fWellWidth / 2.0) + (float(dWell['row'] - 1) * fWellWidth) + dConfigReport['plate_well_label_disp_v_2']
        SetPos(fPosX, fPosY)
        DrawCell(sWellCaption, "label_3", fWellWidth, None, "C", "")
        fPosY = fStartPosY + dConfigReport['plate_margin_top'] + (fWellWidth / 2.0) + (float(dWell['row'] - 1) * fWellWidth) - dConfigReport['plate_well_label_disp_v_3']
        SetPos(fPosX, fPosY)
        DrawCell(sTypeCaption, "label_4", fWellWidth, None, "C", "")
  fXStartPos = dConfigReport['page_section_advance'] + dConfigReport['page_char_advance']
  fYStartPos = dConfigReport['page_body_height'] - dConfigReport['worklist_row_height'] - dConfigReport['worklist_detection_flags_v_disp']
  fXDisp = 0
  SetPos(fXStartPos + fXDisp, fYStartPos)
  DrawCell("A detetar", "small_1", dConfigReport['worklist_width_detection_flag_caption'], dConfigReport['worklist_row_height'], "C", "LTB")
  fXDisp = fXDisp + dConfigReport['worklist_width_detection_flag_caption']
  SetPos(fXStartPos + fXDisp, fYStartPos)
  DrawCell("", "normal", dConfigReport['worklist_width_detection_flag'], dConfigReport['worklist_row_height'], "C", "LTBR")
  fXDisp = fXDisp + dConfigReport['worklist_width_detection_flag'] + dConfigReport['worklist_detection_flags_sep']
  SetPos(fXStartPos + fXDisp, fYStartPos)
  DrawCell("Deteção concluída", "small_1", dConfigReport['worklist_width_detection_flag_caption'], dConfigReport['worklist_row_height'], "C", "LTB")
  fXDisp = fXDisp + dConfigReport['worklist_width_detection_flag_caption']
  SetPos(fXStartPos + fXDisp, fYStartPos)
  DrawCell("", "normal", dConfigReport['worklist_width_flag'], dConfigReport['worklist_row_height'], "C", "LTBR")
  fXDisp = fXDisp + dConfigReport['worklist_width_detection_flag'] + dConfigReport['worklist_detection_flags_sep']
  SetPos(fXStartPos + fXDisp, fYStartPos)
  DrawCell("Resultados no LIS", "small_1", dConfigReport['worklist_width_detection_flag_caption'], dConfigReport['worklist_row_height'], "C", "LTB")
  fXDisp = fXDisp + dConfigReport['worklist_width_detection_flag_caption']
  SetPos(fXStartPos + fXDisp, fYStartPos)
  DrawCell("", "normal", dConfigReport['worklist_width_flag'], dConfigReport['worklist_row_height'], "C", "LTBR")
  fXDisp = fXDisp + dConfigReport['worklist_width_detection_flag'] + dConfigReport['worklist_detection_flags_sep']
  SetPos(fXStartPos + fXDisp, fYStartPos)
  DrawCell("Alertas ou inconclusos", "small_1", dConfigReport['worklist_width_detection_flag_caption'], dConfigReport['worklist_row_height'], "C", "LTB")
  fXDisp = fXDisp + dConfigReport['worklist_width_detection_flag_caption']
  SetPos(fXStartPos + fXDisp, fYStartPos)
  DrawCell("", "normal", dConfigReport['worklist_width_flag'], dConfigReport['worklist_row_height'], "C", "LTBR")
  AddPage()
  dHistory = dict()
  bDBOK = True
  try:
    MySqlCursorStart()
    sRecordFilter = ""
    for sSample in lsSamples:
      dSample = dSamples[sSample]
      if dSample['name'] and dSample['birthday'] and \
        dSample['name'].upper() != "NULL" and dSample['birthday'].upper() != "NULL" and \
        dSample['name'] != "" and dSample['birthday'] != "":
        sDBRecord = dSample['record'].upper() if dSample['record'] and dSample['record'].upper() != "NULL" else ""
        sDBName = dSample['name'].upper()
        sDBBirthday = ConvertDateTime(dSample['birthday'], sDateFormat, "%Y-%m-%d")
        if sRecordFilter != "":
          sRecordFilter = sRecordFilter + " OR "
        sRecordFilter = sRecordFilter + "(UPPER(record) = '{}' OR (UPPER(name) = '{}' AND birthday = '{}'))".format(sDBRecord, sDBName, sDBBirthday)
    if sRecordFilter != "":
      sQuery = """
        SELECT
          UPPER(record) AS Record,
          UPPER(name) AS Name,
          DATE_FORMAT(birthday, '{}') AS Birthday,
          DATE_FORMAT(result_datetime, '%d/%m') AS ResultDate,
          result_code AS Result
        FROM tests
        WHERE
          status <> 2
          AND result_code != 'waiting' AND result_code != 'notest'
          AND DATEDIFF(NOW(), result_datetime) < 120
          AND ({})
        ORDER BY result_datetime DESC
        """.format(sDateFormat, sRecordFilter);
      dDBQueryRes = RunMySqlCursor(sQuery, True)
      for dRow in dDBQueryRes:
        if dRow['ResultDate'] and dRow['ResultDate'].upper() != "NULL":
          sRecord = dRow['Record'].upper() if dRow['Record'] and dRow['Record'].upper() != "NULL" else ""
          tKey = (sRecord, dRow['Name'].upper(), dRow['Birthday'].upper())
          for tSearch in dHistory:
            if tSearch[1] == dRow['Name'].upper() and tSearch[2] == dRow['Birthday'].upper():
              tKey = tSearch
              break
          if tKey not in dHistory:
            dHistory[tKey] = dict()
            dHistory[tKey]['results'] = list()
          sResultCode = dRow['Result'].lower()
          sResult = "E"
          dResTrans = {
            'negative': "N",
            'positive': "P",
            'error': "E",
            'inconclusive': "I"
            }
          if sResultCode in dResTrans:
            sResult = dResTrans[sResultCode]
          dHistory[tKey]['results'].append({
            'date': dRow['ResultDate'],
            'result': sResult
            })
      for tKey, dEntry in dHistory.items():
        sHistory = ""
        for dResult in dEntry['results'][0:5]:
          sHistory = sHistory + "{}{} ({})".format(
            ", " if sHistory != "" else "",
            dResult['result'],
            dResult['date']
            )
        dHistory[tKey]['history'] = sHistory
    MySqlCursorClose()
  except Exception as dError:
    bDBOK = False
  SetPos(0,0)
  WriteLine("#INSTITUTION# | #DEPARTMENT#", "header_1")
  WriteLine("Plano de trabalho para protocolo qPCR / RT-qPCR", "title_1")
  InfoLine("Protocolo", dProtocol['title'], 1, 0)
  InfoLine("Data e hora", oDateTime.strftime(sDateTimeFormat), 1, 0)
  InfoLine("Ficheiro", sFile, 1, 0)
  DrawRule(dConfigReport['page_body_width'], None)
  MovePos(0.0, dConfigReport['page_header_sep'])
  uExtraRows = 0
  if "extra_reactions" in dProtocol:
    luIDs = list(dProtocol['extra_reactions'].keys())
    luIDs.sort()
    uColumn = 1
    fXStartPos = dConfigReport['page_char_advance']
    fYStartPos = GetPos()[1] + dConfigReport['worklist_extra_sep']
    fXDisp = 0
    uExtraRows = uExtraRows + 1
    for uID in luIDs:
      sExtraTitle = dProtocol['extra_reactions'][uID]['title']
      if uColumn > dConfigReport['worklist_extras_columns']:
        uColumn = 1
        fXDisp = 0
        fYStartPos = fYStartPos + 2 * dConfigReport['worklist_row_height']
        uExtraRows = uExtraRows + 1
      SetPos(fXStartPos + fXDisp, fYStartPos)
      DrawCell(sExtraTitle, "small_1", dConfigReport['worklist_width_extra_flag_caption'], dConfigReport['worklist_row_height'], "C", "LTB")
      fXDisp = fXDisp + dConfigReport['worklist_width_extra_flag_caption']
      SetPos(fXStartPos + fXDisp, fYStartPos)
      DrawCell("", "normal", dConfigReport['worklist_width_extra_flag'], dConfigReport['worklist_row_height'], "C", "LTBR")
      fXDisp = fXDisp + dConfigReport['worklist_width_extra_flag'] + dConfigReport['worklist_width_extra_sep']
      uColumn = uColumn + 1
    fYStartPos = fYStartPos + dConfigReport['worklist_extra_sep'] + dConfigReport['worklist_row_height']
    fXStartPos = 0
    SetPos(fXStartPos, fYStartPos)
  WriteLine("Amostras", "section_1")
  fMaxWidth = dConfigReport['page_body_width'] - dConfigReport['page_section_advance']
  uSampleIndex = 1
  uSampleCount = 0
  for sSample in lsSamples:
    dSample = dSamples[sSample]
    if uSampleCount + uExtraRows >= dConfigReport['worklist_samples_per_page']:
      AddPage()
      SetPos(0,0)
      WriteLine("#INSTITUTION# | #DEPARTMENT#", "header_1")
      WriteLine("Plano de trabalho para protocolo qPCR / RT-qPCR", "title_1")
      InfoLine("Protocolo", dProtocol['title'], 1, 0)
      InfoLine("Data e hora", oDateTime.strftime(sDateTimeFormat), 1, 0)
      InfoLine("Ficheiro", sFile, 1, 0)
      DrawRule(dConfigReport['page_body_width'], None)
      MovePos(0.0, dConfigReport['page_header_sep'])
      WriteLine("Amostras", "section_1")
      fMaxWidth = dConfigReport['page_body_width'] - dConfigReport['page_section_advance']
      uExtraRows = 0
      uSampleCount = 0
    fXStartPos = GetPos()[0]
    fYStartPos = GetPos()[1]
    oBarcode = barcode.get("code128", dSample['sample'], writer=barcode.writer.ImageWriter())
    oBarcode.save("{}\\barcode_{}".format(sTemporaryFolder, str(uSampleIndex)), {"module_width": dConfigReport['barcode_module_width'], "module_height": dConfigReport['barcode_module_height'], "quiet_zone": dConfigReport['barcode_quiet_zone'], "text_distance": dConfigReport['barcode_text_distance'], "font_size": int(dConfigReport['barcode_font_size'])})
    Image("{}\\barcode_{}.png".format(sTemporaryFolder, str(uSampleIndex)), None, None, 0, float(dConfigReport['barcode_image_height']))
    fXDisp = dConfigReport['worklist_x_disp_barcode']
    SetPos(fXStartPos + fXDisp, fYStartPos)
    DrawCell("R", "normal", dConfigReport['worklist_width_flag_caption'], dConfigReport['worklist_row_height'], "C", "LTBR")
    fXDisp = fXDisp + dConfigReport['worklist_width_flag_caption']
    SetPos(fXStartPos + fXDisp, fYStartPos)
    DrawCell("", "normal", dConfigReport['worklist_width_flag'], dConfigReport['worklist_row_height'], "C", "LTBR")
    fXDisp = fXDisp + dConfigReport['worklist_width_flag'] + dConfigReport['worklist_x_disp_flags']
    SetPos(fXStartPos + fXDisp, fYStartPos)
    DrawCell("A", "normal", dConfigReport['worklist_width_flag_caption'], dConfigReport['worklist_row_height'], "C", "LTBR")
    fXDisp = fXDisp + dConfigReport['worklist_width_flag_caption']
    SetPos(fXStartPos + fXDisp, fYStartPos)
    DrawCell("", "normal", dConfigReport['worklist_width_flag'], dConfigReport['worklist_row_height'], "L", "LTBR")
    fXDisp = fXDisp + dConfigReport['worklist_width_flag'] + dConfigReport['worklist_x_disp_flags']
    fXDispSection2 = fXDisp
    SetPos(fXStartPos + fXDisp, fYStartPos)
    DrawCell(dSample["sample"], "small_1", dConfigReport['worklist_width_sample'], dConfigReport['worklist_row_height'], "C", "LTB")
    fXDisp = fXDisp + dConfigReport['worklist_width_sample']
    SetPos(fXStartPos + fXDisp, fYStartPos)
    DrawCell(dSample['sample_date'], "small_2", dConfigReport['worklist_width_date'], dConfigReport['worklist_row_height'], "C", "RTB")
    fXDisp = fXDisp + dConfigReport['worklist_width_date']
    SetPos(fXStartPos + fXDisp, fYStartPos)
    uNameWidth = GetStringWidth(dSample['name'], "small_1")
    DrawCell(dSample['name'] if uNameWidth < dConfigReport['worklist_width_name'] else dSample['name'].split()[0] + " " + dSample['name'].split()[1] + " " + dSample['name'].split()[-1], "small_1", dConfigReport['worklist_width_name'], dConfigReport['worklist_row_height'], "C", "LTB")
    fXDisp = fXDisp + dConfigReport['worklist_width_name']
    SetPos(fXStartPos + fXDisp, fYStartPos)
    DrawCell(dSample['birthday'], "small_2", dConfigReport['worklist_width_date'], dConfigReport['worklist_row_height'], "C", "RTB")
    fXDisp = fXDisp + dConfigReport['worklist_width_date']
    SetPos(fXStartPos + fXDisp, fYStartPos)
    DrawCell(dSample['record'], "small_2", dConfigReport['worklist_width_record'], dConfigReport['worklist_row_height'], "C", "LTB")
    fXDisp = fXDisp + dConfigReport['worklist_width_record']
    SetPos(fXStartPos + fXDisp, fYStartPos)
    uDepartmentWidth = GetStringWidth(dSample['department'], "small_2")
    DrawCell(dSample['department'] if uDepartmentWidth < dConfigReport['worklist_width_department_tolerance'] else dSample['department'][0:20], "small_2", dConfigReport['worklist_width_department'], dConfigReport['worklist_row_height'], "C", "RTB")
    fXDisp = fXDisp + dConfigReport['worklist_width_department'] + dConfigReport['worklist_x_disp_flags']
    SetPos(fXStartPos + fXDisp, fYStartPos)
    fXDisp = dConfigReport['worklist_x_disp_barcode']
    fXDisp = fXDispSection2
    SetPos(fXStartPos + fXDisp, fYStartPos + dConfigReport['worklist_row_height'])
    dSampleSlot = SearchPCRSlot("unknown", sSample)
    lsCoordinates = list()
    for uWell, dWell in dSampleSlot['wells'].items():
      lsCoordinates.append(dWell['abbreviation'])
    sCoordinates = ", ".join(lsCoordinates)
    DrawCell(sCoordinates, "small_2", dConfigReport['worklist_width_coordinates'], dConfigReport['worklist_row_height'], "C", "LTBR")
    fXDisp = fXDisp + dConfigReport['worklist_width_coordinates']
    SetPos(fXStartPos + fXDisp, fYStartPos + dConfigReport['worklist_row_height'])
    fHistoryWidth = 0
    tKey = None
    if dSample['record'] and dSample['record'].upper() != "NULL" and dSample['record'] != "":
      for tSearch in dHistory:
        if tSearch[0].upper() == dSample['record'].upper():
          tKey = tSearch
          break
    if not tKey and dSample['name'] and dSample['name'].upper() != "NULL" and dSample['name'] != "" and dSample['birthday'] and dSample['birthday'].upper() != "NULL" and dSample['birthday'] != "":
      for tSearch in dHistory:
        if tSearch[1].upper() == dSample['name'].upper() and tSearch[2].upper() == dSample['birthday'].upper():
          tKey = tSearch
          break
    if tKey:
      for dEntry in dHistory[tKey]['results'][0:6]:
        sEntry = "{} ({})".format(dEntry['result'], dEntry['date'])
        if dEntry['result'].lower() == "p":
          bSelFill = True
          tuSelFillColor = eval(dConfigReport['cell_warning_fill_color'])
          sSelStyle = "small_2"
        elif dEntry['result'].lower() == "i" or dEntry['result'].lower() == "e":
          bSelFill = True
          tuSelFillColor = eval(dConfigReport['cell_attention_fill_color'])
          sSelStyle = "small_2"
        else:
          bSelFill = True
          tuSelFillColor = eval(dConfigReport['cell_normal_fill_color'])
          sSelStyle = "small_2"
        DrawCell("", "small_1", dConfigReport['width_history_sep'], dConfigReport['worklist_row_height'], "L", "BTL" if fHistoryWidth == 0 else "BT")
        fXDisp = fXDisp + dConfigReport['width_history_sep']
        SetPos(fXStartPos + fXDisp, fYStartPos + dConfigReport['worklist_row_height'])
        DrawCell(sEntry, sSelStyle, dConfigReport['width_result'], dConfigReport['worklist_row_height'], "C", "BT", bFill = bSelFill, tuFillColor = tuSelFillColor)
        fXDisp = fXDisp + dConfigReport['width_result']
        SetPos(fXStartPos + fXDisp, fYStartPos + dConfigReport['worklist_row_height'])
        fHistoryWidth = fHistoryWidth + dConfigReport['width_result'] + dConfigReport['width_history_sep']
      if len(dHistory[tKey]['results']) == 7:
        dEntry = dHistory[tKey]['results'][6]
        sEntry = "{} ({})".format(dEntry['result'], dEntry['date'])
        if dEntry['result'].lower() == "p":
          bSelFill = True
          tuSelFillColor = eval(dConfigReport['cell_warning_fill_color'])
          sSelStyle = "small_2"
        elif dEntry['result'].lower() == "i" or dEntry['result'].lower() == "e":
          bSelFill = True
          tuSelFillColor = eval(dConfigReport['cell_attention_fill_color'])
          sSelStyle = "small_2"
        else:
          bSelFill = True
          tuSelFillColor = eval(dConfigReport['cell_normal_fill_color'])
          sSelStyle = "small_2"
        DrawCell("", "small_1", dConfigReport['width_history_sep'], dConfigReport['worklist_row_height'], "L", "BTL" if fHistoryWidth == 0 else "BT")
        fXDisp = fXDisp + dConfigReport['width_history_sep']
        SetPos(fXStartPos + fXDisp, fYStartPos + dConfigReport['worklist_row_height'])
        DrawCell(sEntry, sSelStyle, dConfigReport['width_result'], dConfigReport['worklist_row_height'], "C", "BT", bFill = bSelFill, tuFillColor = tuSelFillColor)
        fXDisp = fXDisp + dConfigReport['width_result']
        SetPos(fXStartPos + fXDisp, fYStartPos + dConfigReport['worklist_row_height'])
        fHistoryWidth = fHistoryWidth + dConfigReport['width_result'] + dConfigReport['width_history_sep']
      if len(dHistory[tKey]['results']) > 7:
        bSelFill = False
        tuSelFillColor = eval(dConfigReport['cell_normal_fill_color'])
        sSelStyle = "small_2"
        DrawCell("", "small_1", dConfigReport['width_history_sep'], dConfigReport['worklist_row_height'], "L", "BTL" if fHistoryWidth == 0 else "BT")
        fXDisp = fXDisp + dConfigReport['width_history_sep']
        SetPos(fXStartPos + fXDisp, fYStartPos + dConfigReport['worklist_row_height'])
        DrawCell("...", sSelStyle, dConfigReport['width_result'], dConfigReport['worklist_row_height'], "C", "BT", bFill = bSelFill, tuFillColor = tuSelFillColor)
        fXDisp = fXDisp + dConfigReport['width_result']
        SetPos(fXStartPos + fXDisp, fYStartPos + dConfigReport['worklist_row_height'])
        fHistoryWidth = fHistoryWidth + dConfigReport['width_result'] + dConfigReport['width_history_sep']
        dEntry = dHistory[tKey]['results'][-1]
        sEntry = "{} ({})".format(dEntry['result'], dEntry['date'])
        if dEntry['result'].lower() == "p":
          bSelFill = True
          tuSelFillColor = eval(dConfigReport['cell_warning_fill_color'])
          sSelStyle = "small_2"
        elif dEntry['result'].lower() == "i" or dEntry['result'].lower() == "e":
          bSelFill = True
          tuSelFillColor = eval(dConfigReport['cell_attention_fill_color'])
          sSelStyle = "small_2"
        else:
          bSelFill = True
          tuSelFillColor = eval(dConfigReport['cell_normal_fill_color'])
          sSelStyle = "small_2"
        DrawCell("", "small_1", dConfigReport['width_history_sep'], dConfigReport['worklist_row_height'], "L", "BTL" if fHistoryWidth == 0 else "BT")
        fXDisp = fXDisp + dConfigReport['width_history_sep']
        SetPos(fXStartPos + fXDisp, fYStartPos + dConfigReport['worklist_row_height'])
        DrawCell(sEntry, sSelStyle, dConfigReport['width_result'], dConfigReport['worklist_row_height'], "C", "BT", bFill = bSelFill, tuFillColor = tuSelFillColor)
        fXDisp = fXDisp + dConfigReport['width_result']
        SetPos(fXStartPos + fXDisp, fYStartPos + dConfigReport['worklist_row_height'])
        fHistoryWidth = fHistoryWidth + dConfigReport['width_result'] + dConfigReport['width_history_sep']
    fCellWidth = dConfigReport['worklist_width_sample'] + dConfigReport['worklist_width_date'] + dConfigReport['worklist_width_name'] + dConfigReport['worklist_width_date'] + dConfigReport['worklist_width_record'] + dConfigReport['worklist_width_department'] - dConfigReport['worklist_width_coordinates'] - fHistoryWidth
    sDBErrText = ""
    if not bDBOK:
      sDBErrText = "Atenção: erro de BD"
    DrawCell(sDBErrText, "normal", fCellWidth, dConfigReport['worklist_row_height'], "L", "TBR")
    SetPos(fXStartPos, fYStartPos + dConfigReport['worklist_row_height'] * 2 + dConfigReport['worklist_row_disp'])
    uSampleIndex = uSampleIndex + 1
    uSampleCount = uSampleCount + 1
  MovePos(0, dConfigReport['worklist_row_height'])
  InfoLine("Resultados (R)", "N - Não detetado; P - Detetado; I - Inconclusivo; RC - Repetir colheita (indeterminado); F - Inválido ou falha.", 4)
  InfoLine("Ações (A)", "EA - Repetir extração automática; EM - Repetir extração manual; GX - Realizar no GeneXpert; RD - Repetir deteção.", 4)
  return

def SaveWorkplanFiles():
  global \
    sExperimentsFolder,\
    sFile,\
    lsPCRDefKeys,\
    ldPCRDef,\
    oReport,\
    lsSampleKeys,\
    dSamples,\
    lsSamples,\
    sCommandsFolder
  if not os.path.isdir("{}\\{}".format(sExperimentsFolder, sFile)):
    os.mkdir("{}\\{}".format(sExperimentsFolder, sFile))
  tools.SaveTable((lsPCRDefKeys, ldPCRDef), "{}\\{}\\{}_pcr.csv".format(sExperimentsFolder, sFile, sFile))
  ldSamples = list()
  for sSample in lsSamples:
    ldSamples.append(dSamples[sSample])
  tools.SaveTable((lsSampleKeys, ldSamples), "{}\\{}\\{}_samples.csv".format(sExperimentsFolder, sFile, sFile))
  oReport.output("{}\\{}\\{}_workplan.pdf".format(sExperimentsFolder, sFile, sFile))
  os.startfile("{}\\{}\\{}_workplan.pdf".format(sExperimentsFolder, sFile, sFile))
  return

def GetFile():
  global sExperimentsFolder, sFileDateTimeFormat
  sFolderGroup = datetime.datetime.now().strftime(sFileDateTimeFormat)
  lsFolders = glob.glob("{}\\{}_*".format(sExperimentsFolder, sFolderGroup))
  uCurrIndex = len(lsFolders) + 1
  bIndexFound = False
  for uIndex in range(0, 1000):
    uCurrIndex = uCurrIndex + uIndex
    if not os.path.isdir("{}\\{}_{}".format(sExperimentsFolder, sFolderGroup, str(uCurrIndex))):
      bIndexFound = True
      break
  if not bIndexFound:
    return None
  return "{}_{}".format(sFolderGroup, str(uCurrIndex))

# ----------------
# Dialog functions
# ----------------

def SetStatusText(sText):
  global dMainDialog
  dMainDialog['LabelStatus'].config(text = sText)

def OnListBoxWorklistSelect(oEvent):
  SetStatusText("")
  tSelection = oEvent.widget.curselection()
  if len(tSelection) > 0:
    uSelection = int(tSelection[0])
    sSelection = oEvent.widget.get(uSelection)
    dMainDialog['EntryWorklist'].delete(0, END)
    dMainDialog['EntryWorklist'].insert(END, sSelection)

def OnListBoxSamplesSelect(oEvent):
  SetStatusText("")
  tSelection = oEvent.widget.curselection()
  if len(tSelection) > 0:
    uSelection = int(tSelection[0])
    sSample = oEvent.widget.get(uSelection)
    SampleDataUpdate(sSample)

def SampleDataUpdate(sSample = None):
  global dMainDialog, dSamples, dWorklist, sDateFormat
  dMainDialog['EntrySample'].config(state = NORMAL)
  dMainDialog['EntryLISID'].config(state = NORMAL)
  dMainDialog['DateEntrySampleDate'].config(state = NORMAL)
  dMainDialog['EntryName'].config(state = NORMAL)
  dMainDialog['EntryRecord'].config(state = NORMAL)
  dMainDialog['EntryDepartment'].config(state = NORMAL)
  dMainDialog['DateEntryBirthday'].config(state = NORMAL)
  dMainDialog['ComboBoxGender'].config(state = NORMAL)
  dMainDialog['ButtonSampleUpdate'].config(state = NORMAL)
  dMainDialog['EntrySample'].delete(0, END)
  dMainDialog['EntryLISID'].delete(0, END)
  oCurrDate = datetime.datetime.now().date()
  dMainDialog['DateEntrySampleDate'].set_date(oCurrDate)
  dMainDialog['EntryName'].delete(0, END)
  dMainDialog['EntryRecord'].delete(0, END)
  dMainDialog['EntryDepartment'].delete(0, END)
  dMainDialog['DateEntryBirthday'].set_date(oCurrDate)
  dMainDialog['ComboBoxGender'].current(0)
  if sSample and sSample != "":
    dMainDialog['EntrySample'].insert(END, dSamples[sSample]['sample'])
    dMainDialog['EntryLISID'].insert(END, dSamples[sSample]['lis_id'])
    oSampleDate = datetime.datetime.strptime(dSamples[sSample]['sample_date'], sDateFormat)
    dMainDialog['DateEntrySampleDate'].set_date(oSampleDate.date())
    dMainDialog['EntryName'].insert(END, dSamples[sSample]['name'])
    dMainDialog['EntryRecord'].insert(END, dSamples[sSample]['record'])
    dMainDialog['EntryDepartment'].insert(END, dSamples[sSample]['department'])
    oBirthday = datetime.datetime.strptime(dSamples[sSample]['birthday'], sDateFormat)
    dMainDialog['DateEntryBirthday'].set_date(oBirthday.date())
    dMainDialog['ComboBoxGender'].current(0 if dSamples[sSample]['gender'] == "M" else 1)
  if not sSample or sSample == "" or sSample in dWorklist:
    dMainDialog['EntrySample'].config(state = "readonly")
    dMainDialog['EntryLISID'].config(state = "readonly")
    dMainDialog['DateEntrySampleDate'].config(state = DISABLED)
    dMainDialog['EntryName'].config(state = "readonly")
    dMainDialog['EntryRecord'].config(state = "readonly")
    dMainDialog['EntryDepartment'].config(state = "readonly")
    dMainDialog['DateEntryBirthday'].config(state = DISABLED)
    dMainDialog['ComboBoxGender'].config(state = DISABLED)
    dMainDialog['ButtonSampleUpdate'].config(state = DISABLED)

def OnComboBoxProtocolSelected(oEvent):
  global dMainDialog, lsProtocols, dProtocols
  SetStatusText("")
  sSelection = oEvent.widget.get()
  sSelProtocol = None
  for sCurrProtocol, dCurrProtocol in dProtocols.items():
    if dCurrProtocol['abbreviation'] == sSelection:
      sSelProtocol = sCurrProtocol
      break
  SelectProtocol(sSelProtocol)
  
def SelectProtocol(sProtocol = None):
  global \
    dProtocol,\
    dProtocols,\
    dPlateType,\
    dPlateTypes,\
    dPCRSlots,\
    dMixIngredients,\
    dIngredients,\
    ldPCRDef,\
    lsSamples,\
    dSamples,\
    dMainDialog
  dProtocol = None
  dPlateType = None
  dPCRSlots = dict()
  dMixIngredients = dict()
  dIngredients = dict()
  ldPCRDef = list()
  if sProtocol:
    dProtocol = dProtocols[sProtocol]
    dPlateType = dPlateTypes[dProtocol['plate_type']]
  if dProtocol and len(lsSamples) > dProtocol['max_unknown']:
    dSamples = dict()
    lsSamples = list()
    ListBoxWorklistUpdate()
    ListBoxSamplesUpdate()
    SampleDataUpdate()
    SetStatusText("Atenção: O número de amostras da lista ultrapassou o limite do protocolo.")
  if not dProtocol:
    dMainDialog['ComboBoxProtocol'].set("")
  UpdateSampleCount()

def OnCommandWorklistUpdate():
  global dWorklist
  SetStatusText("")
  dWorklist = dict()
  LoadWorklist()
  ListBoxWorklistUpdate()

def OnListBoxWorklistDoubleButton1(oEvent):
  SetStatusText("")
  tSelection = oEvent.widget.curselection()
  if len(tSelection) > 0:
    OnCommandWorklistAdd()

def ListBoxWorklistUpdate():
  global dMainDialog, dWorklist, dSamples
  SetStatusText("")
  dMainDialog['EntryWorklist'].delete(0, END)
  dMainDialog['ListBoxWorklist'].delete(0, END)
  lsKeys = list(dWorklist.keys())
  lsKeys.sort()
  for sKey in lsKeys:
    if sKey not in dSamples:
      dMainDialog['ListBoxWorklist'].insert(END, sKey)

def OnCommandWorklistAdd():
  global dMainDialog, dWorklist, dSamples, lsSamples
  SetStatusText("")
  tSelection = dMainDialog['ListBoxWorklist'].curselection()
  if len(tSelection) == 0:
    SetStatusText("Atenção: Sem amostra selecionada.")
    return
  uSelection = int(tSelection[0])
  sSelection = dMainDialog['ListBoxWorklist'].get(uSelection)
  SelectWorklistSample(sSelection)

def ListBoxSamplesUpdate():
  global dMainDialog, dSamples, lsSamples
  dMainDialog['ListBoxSamples'].delete(0, END)
  for sSample in lsSamples:
    dMainDialog['ListBoxSamples'].insert(END, sSample)
  SampleDataUpdate(None)

def OnEntryWorklistKey(oEvent):
  global dMainDialog
  SetStatusText("")
  dMainDialog["ListBoxWorklist"].selection_clear(0, END)
  
def OnEntryWorklistEnter(oEvent):
  global dMainDialog, dWorklist, dSamples
  SetStatusText("")
  sSource = GetSample(oEvent.widget.get().strip())
  dMainDialog["ListBoxWorklist"].selection_clear(0, END)
  if not sSource:
    SetStatusText("Atenção: amostra inválida.")
    oEvent.widget.icursor(0)
    oEvent.widget.selection_range(0, END)
    return
  if sSource not in dWorklist:
    SetStatusText("Atenção: amostra não encontrada.")
    oEvent.widget.icursor(0)
    oEvent.widget.selection_range(0, END)
    return
  if sSource in dSamples:
    SetStatusText("Atenção: amostra já selecionada.")
    oEvent.widget.icursor(0)
    oEvent.widget.selection_range(0, END)
    return
  SelectWorklistSample(sSource)

def SelectWorklistSample(sSample):
  global dWorklist, dSamples, lsSamples, dProtocol
  SetStatusText("")
  if dProtocol and len(lsSamples) + 1 > dProtocol['max_unknown']:
    SetStatusText("Atenção: O número de amostras da lista já atingiu o limite do protocolo.")
    return
  dSamples[sSample] = dWorklist[sSample]
  lsSamples.append(sSample)
  ListBoxWorklistUpdate()
  ListBoxSamplesUpdate()
  UpdateSampleCount()
  SelectSample(sSample)

def OnCommandSampleRemove():
  global dMainDialog, dSamples, lsSamples
  SetStatusText("")
  tSelection = dMainDialog['ListBoxSamples'].curselection()
  if len(tSelection) == 0:
    return
  uSelection = int(tSelection[0])
  sSample = dMainDialog['ListBoxSamples'].get(uSelection)
  uIndex = lsSamples.index(sSample)
  lsSamples.pop(uIndex)
  dSamples.pop(sSample)
  ListBoxWorklistUpdate()
  ListBoxSamplesUpdate()
  UpdateSampleCount()
  if len(lsSamples) > 0:
    if uIndex >= len(lsSamples):
      SelectSample(lsSamples[uIndex - 1])
    else:
      SelectSample(lsSamples[uIndex])

def OnCommandSampleUp():
  global dMainDialog, lsSamples
  SetStatusText("")
  tSelection = dMainDialog['ListBoxSamples'].curselection()
  if len(tSelection) == 0:
    return
  uSelection = int(tSelection[0])
  sSample = dMainDialog['ListBoxSamples'].get(uSelection)
  uIndex = lsSamples.index(sSample)
  if uIndex == 0:
    return
  lsSamples.pop(uIndex)
  lsSamples.insert(uIndex - 1, sSample)
  ListBoxSamplesUpdate()
  SelectSample(sSample)

def OnCommandSampleDown():
  global dMainDialog, lsSamples
  SetStatusText("")
  tSelection = dMainDialog['ListBoxSamples'].curselection()
  if len(tSelection) == 0:
    return
  uSelection = int(tSelection[0])
  sSample = dMainDialog['ListBoxSamples'].get(uSelection)
  uIndex = lsSamples.index(sSample)
  if uIndex + 1 >= len(lsSamples):
    return
  lsSamples.pop(uIndex)
  lsSamples.insert(uIndex + 1, sSample)
  ListBoxSamplesUpdate()
  SelectSample(sSample)

def OnCommandRepeatSample():
  global dMainDialog, dProtocol, lsSamples, ldSamples
  if dProtocol and len(lsSamples) + 1 > dProtocol['max_unknown']:
    SetStatusText("Atenção: O número de amostras da lista já atingiu o limite do protocolo.")
    return
  SetStatusText("")
  tSelection = dMainDialog['ListBoxSamples'].curselection()
  if len(tSelection) == 0:
    return
  uSelection = int(tSelection[0])
  sSample = dMainDialog['ListBoxSamples'].get(uSelection)
  RepeatSample(sSample)

def OnCommandNewSample():
  global dMainDialog, dProtocol, lsSamples
  if dProtocol and len(lsSamples) + 1 > dProtocol['max_unknown']:
    SetStatusText("Atenção: O número de amostras da lista já atingiu o limite do protocolo.")
    return
  SetStatusText("")
  NewSample("NOVA")

def NewSample(sSample):
  global dMainDialog, dSamples, lsSamples, dWorklist, sDateFormat
  if sSample in dSamples:
    SetStatusText("Atenção: modifique primeiro a amostra {}.".format(sSample))
    SelectSample(sSample)
    dMainDialog['EntrySample'].focus_set()
    dMainDialog['EntrySample'].icursor(0)
    dMainDialog['EntrySample'].selection_range(0, END)
    return None
  if sSample in dWorklist:
    SetStatusText("Atenção: já existe a amostra {} na lista de trabalho.".format(sSample))
    return None
  dNewDict = dict()
  dNewDict['sample'] = sSample
  dNewDict['lis_id'] = ""
  dNewDict['sample_date'] = datetime.datetime.now().strftime(sDateFormat)
  dNewDict['name'] = ""
  dNewDict['birthday'] = datetime.datetime.now().strftime(sDateFormat)
  dNewDict['gender'] = "M"
  dNewDict['record'] = ""
  dNewDict['department'] = ""
  dSamples[sSample] = dNewDict
  lsSamples.append(sSample)
  ListBoxSamplesUpdate()
  UpdateSampleCount()
  SelectSample(sSample)
  dMainDialog['EntrySample'].focus_set()
  dMainDialog['EntrySample'].icursor(0)
  dMainDialog['EntrySample'].selection_range(0, END)
  return dNewDict

def RepeatSample(sSample):
  global dMainDialog, dSamples, lsSamples, dWorklist, sDateFormat
  sNewSample = "R{}".format(sSample)
  if sNewSample in dSamples:
    SetStatusText("Atenção: já existe a repetição para a amostra {}.".format(sSample))
    SelectSample(sNewSample)
    dMainDialog['EntrySample'].focus_set()
    dMainDialog['EntrySample'].icursor(0)
    dMainDialog['EntrySample'].selection_range(0, END)
    return None
  if sNewSample in dWorklist:
    SetStatusText("Atenção: já existe a amostra {} na lista de trabalho.".format(sNewSample))
    return None
  dNewDict = dict()
  dNewDict['sample'] = sNewSample
  dNewDict['lis_id'] = dSamples[sSample]['lis_id']
  dNewDict['sample_date'] = dSamples[sSample]['sample_date']
  dNewDict['name'] = dSamples[sSample]['name']
  dNewDict['birthday'] = dSamples[sSample]['birthday']
  dNewDict['gender'] = dSamples[sSample]['gender']
  dNewDict['record'] = dSamples[sSample]['record']
  dNewDict['department'] = dSamples[sSample]['department']
  dSamples[sNewSample] = dNewDict
  lsSamples.append(sNewSample)
  ListBoxSamplesUpdate()
  UpdateSampleCount()
  #SelectSample(sNewSample)
  #dMainDialog['EntrySample'].focus_set()
  #dMainDialog['EntrySample'].icursor(0)
  #dMainDialog['EntrySample'].selection_range(0, END)
  SelectSample(lsSamples[lsSamples.index(sSample) + 1])
  return dNewDict

def SelectSample(sSample):
  global dMainDialog, lsSamples
  uIndex = lsSamples.index(sSample)
  dMainDialog['ListBoxSamples'].selection_clear(0, END)
  dMainDialog['ListBoxSamples'].selection_set(uIndex)
  dMainDialog['ListBoxSamples'].see(uIndex)
  SampleDataUpdate(sSample)

def OnCommandUpdateSample():
  global dMainDialog, dSamples, lsSamples, dWorklist, sDateFormat
  SetStatusText("")
  tSelection = dMainDialog['ListBoxSamples'].curselection()
  if len(tSelection) == 0:
    SampleDataUpdate(None)
    return
  uSelection = int(tSelection[0])
  sOldSample = dMainDialog['ListBoxSamples'].get(uSelection)
  sSample = dMainDialog['EntrySample'].get().strip().upper()
  if sSample and (sSample in dWorklist or sOldSample in dWorklist):
    SetStatusText("Atenção: Não pode modificar uma amostra da lista de trabalho.".format(sSample))
    tSelection = dMainDialog['ListBoxSamples'].curselection()
    if len(tSelection) > 0:
      uSelection = int(tSelection[0])
      sCurrSample = dMainDialog['ListBoxSamples'].get(uSelection)
      SampleDataUpdate(sCurrSample)
      dMainDialog['EntrySample'].focus_set()
      dMainDialog['EntrySample'].icursor(0)
      dMainDialog['EntrySample'].selection_range(0, END)
    else:
      SampleDataUpdate(None)
    return
  sLISID = "" if not dMainDialog['EntryLISID'].get() else dMainDialog['EntryLISID'].get().strip()
  sRecord = "" if not dMainDialog['EntryRecord'].get() else dMainDialog['EntryRecord'].get().strip()
  sName = "" if not dMainDialog['EntryName'].get() else GetName(dMainDialog['EntryName'].get().strip())
  sDepartment = "" if not dMainDialog['EntryDepartment'].get() else dMainDialog['EntryDepartment'].get().strip()
  sGender = "M" if not dMainDialog['ComboBoxGender'].get() else dMainDialog['ComboBoxGender'].get().strip()
  if sGender == "Masculino":
    sGender = "M"
  else:
    sGender = "F"
  sBirthday = dMainDialog['DateEntryBirthday'].get()
  sSampleDate = dMainDialog['DateEntrySampleDate'].get()
  bValidSample = True
  if not sSample or sSample == "":
    sSample = sOldSample
    SetStatusText("Atenção: Indique um ID de amostra válido.".format(sSample))
    bValidSample = False
  elif sSample != sOldSample and sSample in dSamples:
    sSample = sOldSample
    SetStatusText("Atenção: Indique um ID de amostra diferente.".format(sSample))
    bValidSample = False
  dSamples[sOldSample]['lis_id'] = sLISID
  dSamples[sOldSample]['sample'] = sSample
  dSamples[sOldSample]['record'] = sRecord
  dSamples[sOldSample]['name'] = sName
  dSamples[sOldSample]['department'] = sDepartment
  dSamples[sOldSample]['gender'] = sGender
  dSamples[sOldSample]['birthday'] = sBirthday
  dSamples[sOldSample]['sample_date'] = sSampleDate
  if sSample != sOldSample:
    dSamples[sSample] = dSamples.pop(sOldSample)
    lsSamples[lsSamples.index(sOldSample)] = sSample
  ListBoxSamplesUpdate()
  SelectSample(sSample)
  if not bValidSample:
    dMainDialog['EntrySample'].delete(0, END)
    dMainDialog['EntrySample'].insert(END, sSample)
    dMainDialog['EntrySample'].focus_set()
    dMainDialog['EntrySample'].icursor(0)
    dMainDialog['EntrySample'].selection_range(0, END)

def OnCommandOpen():
  global \
    dMainDialog, \
    sExperimentsFolder, \
    lsSamples, \
    dSamples, \
    sFile, \
    bOpened, \
    sDateFormat
  SetStatusText("")
  if bOpened:
    Reset()
    return
  sFileName = tkinter.filedialog.askdirectory(
    title = "Selecione a pasta da experiência",
    initialdir = sExperimentsFolder,
    mustexist = True)
  if not sFileName:
    return
  sExperiment = os.path.basename(os.path.normpath(sFileName))
  sSamplesFileName = "{}_samples.csv".format(sExperiment)
  sSamplesFilePath = "{}\\{}\\{}".format(sExperimentsFolder, sExperiment, sSamplesFileName)
  if not os.path.isfile(sSamplesFilePath):
    SetStatusText("Atenção: Pasta de experiência inválida.")
    return
  dSamplesTable = tools.LoadTable(sSamplesFilePath)[1]
  Reset()
  for dSample in dSamplesTable:
    sSample = dSample['sample']
    lsSamples.append(sSample)
    dNewDict = dict()
    dNewDict['sample'] = sSample
    dNewDict['lis_id'] = "" if not dSample['lis_id'] else dSample['lis_id']
    dNewDict['sample_date'] = datetime.datetime.now().strftime(sDateFormat) if not dSample['sample_date'] else dSample['sample_date']
    dNewDict['name'] = "" if not dSample['name'] else dSample['name']
    dNewDict['birthday'] = datetime.datetime.now().strftime(sDateFormat) if not dSample['birthday'] else dSample['birthday']
    dNewDict['gender'] = "M" if not dSample['gender'] else dSample['gender']
    dNewDict['record'] = "" if not dSample['record'] else dSample['record']
    dNewDict['department'] = "" if not dSample['department'] else dSample['department']
    dSamples[sSample] = dNewDict
  del dSamplesTable
  ListBoxWorklistUpdate()
  ListBoxSamplesUpdate()
  UpdateSampleCount()
  SampleDataUpdate(None)
  sFile = sExperiment
  bOpened = True
  dMainDialog['ButtonCreate'].config(text = "Modificar")
  dMainDialog['ButtonOpen'].config(text = "Reiniciar")
  SetStatusText("Informação: foi aberto o plano da experiência {}.".format(sFile))

def OnCommandCreate():
  global \
    dMainDialog,\
    sFile,\
    lsSamples,\
    dSamples,\
    oDateTime,\
    dProtocol,\
    bOpened
  SetStatusText("")
  if not dProtocol:
    SetStatusText("Atenção: Selecione primeiro o protocolo.")
    return
  if len(lsSamples) == 0:
    SetStatusText("Atenção: Adicione pelo menos uma amostra.")
    return
  if bOpened:
    SetStatusText("A criar o plano de trabalho...")
  else:
    SetStatusText("A modificar o plano de trabalho...")
  try:
    oDateTime = datetime.datetime.now()
    if not bOpened:
      sFile = GetFile()
    for sSample in lsSamples:
      AddPCR("unknown", sSample)
    SetPCRDef()
    GenerateWorkplanReport()
    SaveWorkplanFiles()
    RunExtraScript()
  except Exception as dError:
    Log("Erro: Ocorreu uma exceção.", False)
    Log("Mensagem de erro:\n" + str(dError))
    Reset()
    SetStatusText("Erro: Não foi possível criar o plano. Verifique os ficheiros de configuração e a pasta da experiência.")
    return
  SetStatusText("")
  Reset()

def Reset():
  global \
    lsSamples, \
    dSamples, \
    bOpened
  SelectProtocol(None)
  dSamples = dict()
  lsSamples = list()
  ListBoxWorklistUpdate()
  ListBoxSamplesUpdate()
  UpdateSampleCount()
  SampleDataUpdate(None)
  bOpened = False
  dMainDialog['ButtonCreate'].config(text = "Criar plano")
  dMainDialog['ButtonOpen'].config(text = "Abrir plano")

def OnCommandClose():
  global dMainDialog
  dMainDialog['MainWindow'].quit()

def UpdateSampleCount():
  global dMainDialog, lsSamples, dProtocol
  if dProtocol:
    sText = "Amostras ({}/{})".format(str(len(lsSamples)), str(dProtocol['max_unknown']))
  else:
    sText = "Amostras ({})".format(str(len(lsSamples)))
  dMainDialog['LabelSamples'].config(text = sText)

def CreateMainDialog():
  global dMainDialog
  dMainDialog = dict()
  dMainDialog['MainWindow'] = Tk()
  oMainWindow = dMainDialog['MainWindow']
  oMainWindow.title("Plano de trabalho para PCR")
  oMainWindow.resizable(width = False, height = False)
  oMainWindow.columnconfigure(0, minsize=10)
  oMainWindow.columnconfigure(1, minsize=140)
  oMainWindow.columnconfigure(3, minsize=15)
  oMainWindow.columnconfigure(4, minsize=140)
  oMainWindow.columnconfigure(6, minsize=15)
  oMainWindow.columnconfigure(8, minsize=10)
  oMainWindow.columnconfigure(10, minsize=10)
  oMainWindow.rowconfigure(0, minsize=5)
  oMainWindow.rowconfigure(1, pad=2)
  oMainWindow.rowconfigure(2, pad=2)
  oMainWindow.rowconfigure(3, minsize=5)
  oMainWindow.rowconfigure(4, pad=2)
  oMainWindow.rowconfigure(5, pad=2)
  oMainWindow.rowconfigure(6, minsize=5)
  oMainWindow.rowconfigure(7, pad=2)
  oMainWindow.rowconfigure(8, pad=2)
  oMainWindow.rowconfigure(9, minsize=5)
  oMainWindow.rowconfigure(10, pad=2)
  oMainWindow.rowconfigure(11, pad=2)
  oMainWindow.rowconfigure(12, minsize=5)
  oMainWindow.rowconfigure(13, pad=2)
  oMainWindow.rowconfigure(14, pad=2)
  oMainWindow.rowconfigure(15, minsize=5)
  oMainWindow.rowconfigure(16, pad=2)
  oMainWindow.rowconfigure(17, pad=2)
  oMainWindow.rowconfigure(18, minsize=30)
  oMainWindow.rowconfigure(19, minsize=5)
  oMainWindow.rowconfigure(20, pad=2)
  oMainWindow.rowconfigure(21, minsize=20)
  oMainWindow.rowconfigure(22, minsize=15)
  oMainWindow.rowconfigure(23, pad=0)
  oMainWindow.rowconfigure(24, minsize=10)
  dMainDialog['LabelWorklist'] = Label(oMainWindow, text = "Lista de trabalho:")
  dMainDialog['LabelWorklist'].grid(row = 1, column = 1, sticky = W)
  dMainDialog['ListBoxWorklist'] = Listbox(oMainWindow, selectmode = SINGLE, exportselection = False)
  dMainDialog['ListBoxWorklist'].bind("<<ListboxSelect>>", OnListBoxWorklistSelect)
  dMainDialog['ListBoxWorklist'].bind("<Double-Button-1>", OnListBoxWorklistDoubleButton1)
  dMainDialog['ListBoxWorklist'].grid(row = 2, column = 1, rowspan = 16, sticky = N + E + W + S)
  dMainDialog['ScrollYWorklist'] = Scrollbar(oMainWindow, orient = VERTICAL)
  dMainDialog['ScrollYWorklist'].grid(row = 2, column = 2, rowspan = 16, sticky = W + N + S)
  dMainDialog['ListBoxWorklist'].config(yscrollcommand = dMainDialog['ScrollYWorklist'].set)
  dMainDialog['ScrollYWorklist'].config(command = dMainDialog['ListBoxWorklist'].yview)
  dMainDialog['EntryWorklist'] = Entry(oMainWindow, width = 12)
  dMainDialog['EntryWorklist'].bind("<Return>", OnEntryWorklistEnter)
  dMainDialog['EntryWorklist'].bind("<Key>", OnEntryWorklistKey)
  dMainDialog['EntryWorklist'].grid(row = 18, column = 1, columnspan = 2, sticky = W + E)
  oFrame1 = Frame(oMainWindow, borderwidth = 0, relief = FLAT)
  oFrame1.columnconfigure(0, weight = 1)
  oFrame1.columnconfigure(1, weight = 1)
  dMainDialog['Frame1'] = oFrame1
  dMainDialog['Frame1'].grid(row = 20, column = 1, columnspan = 2, sticky = W + E)
  dMainDialog['ButtonWorklistUpdate'] = Button(oFrame1, text = "Atualizar", width = 10, command = OnCommandWorklistUpdate)
  dMainDialog['ButtonWorklistUpdate'].grid(row = 0, column = 0, sticky = W)
  dMainDialog['ButtonWorklistAdd'] = Button(oFrame1, text = "Adicionar", width = 10, command = OnCommandWorklistAdd)
  dMainDialog['ButtonWorklistAdd'].grid(row = 0, column = 1, sticky = E)
  dMainDialog['LabelSamples'] = Label(oMainWindow, text = "Amostras:")
  dMainDialog['LabelSamples'].grid(row = 1, column = 4, sticky = W)
  dMainDialog['ListBoxSamples'] = Listbox(oMainWindow, selectmode = SINGLE, exportselection = False)
  dMainDialog['ListBoxSamples'].bind("<<ListboxSelect>>", OnListBoxSamplesSelect)
  dMainDialog['ListBoxSamples'].grid(row = 2, column = 4, rowspan = 17, sticky = N + E + W + S)
  dMainDialog['ScrollYSamples'] = Scrollbar(oMainWindow, orient = VERTICAL)
  dMainDialog['ScrollYSamples'].grid(row = 2, column = 5, rowspan = 17, sticky = W + N + S)
  dMainDialog['ListBoxSamples'].config(yscrollcommand = dMainDialog['ScrollYSamples'].set)
  dMainDialog['ScrollYSamples'].config(command = dMainDialog['ListBoxSamples'].yview)
  oFrame2 = Frame(oMainWindow, borderwidth = 0, relief = FLAT)
  oFrame2.columnconfigure(1, weight = 1)
  oFrame2.columnconfigure(3, minsize=5)
  dMainDialog['Frame2'] = oFrame2
  dMainDialog['Frame2'].grid(row = 20, column = 4, columnspan = 2, sticky = E + W)
  dMainDialog['ButtonSampleRemove'] = Button(oFrame2, text = "Remover", width = 10, command = OnCommandSampleRemove)
  dMainDialog['ButtonSampleRemove'].grid(row = 0, column = 0, sticky = W)
  dMainDialog['ButtonSampleUp'] = Button(oFrame2, text = "S", width = 3, command = OnCommandSampleUp)
  dMainDialog['ButtonSampleUp'].grid(row = 0, column = 2, sticky = E)
  dMainDialog['ButtonSampleDown'] = Button(oFrame2, text = "D", width = 3, command = OnCommandSampleDown)
  dMainDialog['ButtonSampleDown'].grid(row = 0, column = 4, sticky = E)
  dMainDialog['LabelProtocol'] = Label(oMainWindow, text = "Protocolo:")
  dMainDialog['LabelProtocol'].grid(row = 1, column = 7, sticky = W)
  dMainDialog['ComboBoxProtocol'] = Combobox(oMainWindow, state = "readonly")
  dMainDialog['ComboBoxProtocol'].bind("<<ComboboxSelected>>", OnComboBoxProtocolSelected)
  dMainDialog['ComboBoxProtocol'].grid(row = 2, column = 7, columnspan = 3, sticky = W + E)
  dMainDialog['LabelSample'] = Label(oMainWindow, text = "Amostra:")
  dMainDialog['LabelSample'].grid(row = 4, column = 7, sticky = W)
  dMainDialog['EntrySample'] = Entry(oMainWindow, width = 30)
  dMainDialog['EntrySample'].grid(row = 5, column = 7, sticky = W + E)
  dMainDialog['LabelLISID'] = Label(oMainWindow, text = "ID do SIL:")
  dMainDialog['LabelLISID'].grid(row = 4, column = 9, sticky = W)
  dMainDialog['EntryLISID'] = Entry(oMainWindow, width = 30)
  dMainDialog['EntryLISID'].grid(row = 5, column = 9, sticky = W + E)
  dMainDialog['LabelSampleDate'] = Label(oMainWindow, text = "Data de colheita:")
  dMainDialog['LabelSampleDate'].grid(row = 7, column = 7, sticky = W)
  dMainDialog['DateEntrySampleDate'] = DateEntry(oMainWindow, date_pattern = "dd/mm/yyyy", selectmode = "day", maxdate = datetime.datetime.now(), mindate = datetime.datetime.strptime("2020-01-01", "%Y-%m-%d"), locale = "pt_PT", width = 12)
  dMainDialog['DateEntrySampleDate'].grid(row = 8, column = 7, sticky = W + E)
  dMainDialog['LabelName'] = Label(oMainWindow, text = "Nome:")
  dMainDialog['LabelName'].grid(row = 10, column = 7, sticky = W)
  dMainDialog['EntryName'] = Entry(oMainWindow)
  dMainDialog['EntryName'].grid(row = 11, column = 7, columnspan = 3, sticky = W + E)
  dMainDialog['LabelBirthday'] = Label(oMainWindow, text = "Data de nascimento:")
  dMainDialog['LabelBirthday'].grid(row = 13, column = 7, sticky = W)
  dMainDialog['DateEntryBirthday'] = DateEntry(oMainWindow, date_pattern = "dd/mm/yyyy", selectmode = "day", maxdate = datetime.datetime.now(), mindate = datetime.datetime.strptime("1900-01-01", "%Y-%m-%d"), locale = "pt_PT")
  dMainDialog['DateEntryBirthday'].grid(row = 14, column = 7, sticky = W + E)
  dMainDialog['LabelGender'] = Label(oMainWindow, text = "Sexo:")
  dMainDialog['LabelGender'].grid(row = 13, column = 9, sticky = W)
  dMainDialog['ComboBoxGender'] = Combobox(oMainWindow, width = 5, values = ("Masculino", "Feminino"), state = "readonly")
  dMainDialog['ComboBoxGender'].current(0)
  dMainDialog['ComboBoxGender'].grid(row = 14, column = 9, sticky = W + E)
  dMainDialog['LabelRecord'] = Label(oMainWindow, text = "Processo:")
  dMainDialog['LabelRecord'].grid(row = 16, column = 7, sticky = W)
  dMainDialog['EntryRecord'] = Entry(oMainWindow, width = 12)
  dMainDialog['EntryRecord'].grid(row = 17, column = 7, sticky = W + E)
  dMainDialog['LabelDepartment'] = Label(oMainWindow, text = "Serviço:")
  dMainDialog['LabelDepartment'].grid(row = 16, column = 9, sticky = W)
  dMainDialog['EntryDepartment'] = Entry(oMainWindow, width = 12)
  dMainDialog['EntryDepartment'].grid(row = 17, column = 9, sticky = W + E)
  dMainDialog['SepFrame1'] = Frame(oMainWindow, height = 2, borderwidth = 2, relief = GROOVE)
  dMainDialog['SepFrame1'].grid(row = 18, rowspan = 2, column = 7, columnspan = 3, sticky = E + W)
  oFrame3 = Frame(oMainWindow, borderwidth = 0, relief = FLAT)
  oFrame3.columnconfigure(1, minsize=10)
  oFrame3.columnconfigure(2, minsize=10)
  oFrame3.columnconfigure(3, minsize=10)
  dMainDialog['Frame3'] = oFrame3
  dMainDialog['Frame3'].grid(row = 20, column = 7, columnspan = 3, sticky = E)
  dMainDialog['ButtonRepeatSample'] = Button(oFrame3, text = "Repetir amostra", width = 15, command = OnCommandRepeatSample)
  dMainDialog['ButtonRepeatSample'].grid(row = 0, column = 0, sticky = E)
  dMainDialog['ButtonNewSample'] = Button(oFrame3, text = "Nova amostra", width = 15, command = OnCommandNewSample)
  dMainDialog['ButtonNewSample'].grid(row = 0, column = 2, sticky = E)
  dMainDialog['ButtonSampleUpdate'] = Button(oFrame3, text = "Modificar", width = 10, command = OnCommandUpdateSample)
  dMainDialog['ButtonSampleUpdate'].grid(row = 0, column = 4, sticky = E)
  dMainDialog['SepFrame2'] = Frame(oMainWindow, height = 2, borderwidth = 2, relief = GROOVE)
  dMainDialog['SepFrame2'].grid(row = 22, column = 0, columnspan = 10, sticky = E + W)
  oFrame4 = Frame(oMainWindow, borderwidth = 0, relief = FLAT)
  oFrame4.columnconfigure(1, weight = 1)
  oFrame4.columnconfigure(3, minsize = 10)
  oFrame4.columnconfigure(5, minsize = 10)
  dMainDialog['Frame4'] = oFrame4
  dMainDialog['Frame4'].grid(row = 23, column = 1, columnspan = 9, sticky = E + W)
  dMainDialog['LabelStatus'] = Label(oFrame4, text = "", width = 80, anchor = W, justify = LEFT, relief = GROOVE, borderwidth = 1)
  dMainDialog['LabelStatus'].grid(row = 0, column = 0, sticky = E + W + N + S)
  dMainDialog['ButtonCreate'] = Button(oFrame4, text = "Criar plano", width = 10, command = OnCommandCreate)
  dMainDialog['ButtonCreate'].grid(row = 0, column = 2, sticky = E)
  dMainDialog['ButtonOpen'] = Button(oFrame4, text = "Abrir plano", width = 10, command = OnCommandOpen)
  dMainDialog['ButtonOpen'].grid(row = 0, column = 4, sticky = E)
  dMainDialog['ButtonClose'] = Button(oFrame4, text = "Fechar", width = 10, command = OnCommandClose)
  dMainDialog['ButtonClose'].grid(row = 0, column = 6, sticky = E)

# ------------
# Main routine
# ------------

try:
  LoadConfigOptions()
  LoadConfigReport()
  LoadProtocols()
  LoadPlateTypes()
  CreateMainDialog()
  LoadWorklist()
  lsProtocols = sorted(dProtocols.keys(), key = lambda sKey: dProtocols[sKey]['abbreviation'])
  lsProtocolTitles = list()
  for sProtocol in lsProtocols:
    lsProtocolTitles.append(dProtocols[sProtocol]['abbreviation'])
    dMainDialog['ComboBoxProtocol'].config(values = lsProtocolTitles)
  ListBoxWorklistUpdate()
  mainloop()
except Exception as dError:
  Log("Erro: Ocorreu uma exceção.", True)
  Log("Mensagem de erro:\n" + str(dError))
SaveLog()
sys.exit()
