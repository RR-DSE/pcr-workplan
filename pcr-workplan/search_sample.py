# search_sample.py (utf-8)
# 
# Edited by: RR-DSE
# Timestamp: 22-01-18 20:40:16

# ------------
# Dependencies
# ------------

import unicodedata
import io
import os
import glob
import datetime
import tools
import sys
import re
import listools

# --------------
# Global objects
# --------------

sConfigFolder = "config"
sLogFolder = "logs"
sExperimentsFolder = "data\\experiments"
sSampleRE = ".*"
dMainDialog = dict()
dExperiments = dict()
dConfigOptions = None
bLogVerbose = True
sLogFileTimestampFormat = "%Y_%m_%d_%H_%M_%S"
dMemLog = io.StringIO(newline=None)
oLogTimeStamp = None
dSamples = dict()

# -----------
# Log methods
# -----------

def Log(sText):
  global dMemLog, bLogVerbose, oLogTimeStamp
  if not oLogTimeStamp:
    oLogTimeStamp = datetime.datetime.now()
  dMemLog.write(sText + os.linesep)
  if bLogVerbose:
    print(sText)

def SaveLog():
  global dMemLog, sLogFolder, sLogFileTimestampFormat, oLogTimeStamp
  if not oLogTimeStamp:
    oLogTimeStamp = datetime.datetime.now()
  dMemLog.seek(0)
  sTimeStamp = oLogTimeStamp.strftime(sLogFileTimestampFormat)
  dFile = open("{}\\{}".format(sLogFolder, sTimeStamp + "_search_sample.log"), "w", encoding = "utf-8")
  dFile.write(dMemLog.read())
  dFile.close()

# -----------------------
# Configuration functions
# -----------------------

def LoadConfigOptions():
  global \
    sConfigFolder,\
    dConfigOptions,\
    sExperimentsFolder,\
    sSampleRE
  try:
    dConfigOptions = tools.LoadDictSimple("{}\\options.csv".format(sConfigFolder), "parameter", "value")
    if dConfigOptions['experiments_folder'] == None:
      dConfigOptions['experiments_folder'] = ""
    if dConfigOptions['worklist_folder'] == None:
      dConfigOptions['worklist_folder'] = ""
    if dConfigOptions['sample_expression'] == None or dConfigOptions['sample_expression'] == "":
      dConfigOptions['sample_expression'] = ".*"
    sExperimentsFolder = dConfigOptions['experiments_folder']
    sSampleRE = dConfigOptions['sample_expression']
  except Exception as dError:
    Log("Erro: Não foi possível carregar o ficheiro de configuração para opções.")
    Log("Mensagem de erro:\n" + str(dError))
    raise
  return

# -------------------
# Auxiliary functions
# -------------------

def RemoveAccents(sInput):
  sNFKD = unicodedata.normalize('NFKD', sInput)
  sRes = u"".join([c for c in sNFKD if not unicodedata.combining(c)])
  sRes = sRes.lower()
  sRes = listools.CleanStr(sRes)
  return sRes

def GetTable(sFile):
  dTable = tools.LoadTable(sFile)[1]
  return dTable

def UpdateSamples():
  global sExperimentsFolder, dSamples
  dSamples = dict()
  lsFolders = glob.glob("{}\\*".format(sExperimentsFolder))
  for sFolder in lsFolders:
    try:
      sExperiment = os.path.basename(os.path.normpath(sFolder))
      sSamplesFileName = "{}_samples.csv".format(sExperiment)
      sPCRFileName = "{}_pcr.csv".format(sExperiment)
      sSamplesFile = "{}\\{}\\{}".format(sExperimentsFolder, sExperiment, sSamplesFileName)
      sPCRFile = "{}\\{}\\{}".format(sExperimentsFolder, sExperiment, sPCRFileName)
      if os.path.exists(sSamplesFile):
        dSamplesTable = GetTable(sSamplesFile)
        for dRow in dSamplesTable:
          sSample = dRow['sample'].upper()
          if sSample not in dSamples:
            dSamples[sSample] = dict()
            dSamples[sSample]['experiments'] = dict()
          dSamples[sSample]['name'] = dRow['name']
          dSamples[sSample]['birthday'] = dRow['birthday']
          dSamples[sSample]['gender'] = dRow['gender']
          dSamples[sSample]['record'] = dRow['record']
          dSamples[sSample]['department'] = dRow['department']
          dSamples[sSample]['sample_date'] = dRow['sample_date']
      if os.path.exists(sPCRFile) and os.path.exists(sSamplesFile):
        dPCRTable = GetTable(sPCRFile)
        for dRow in dPCRTable:
          if dRow['type'] == "unknown":
            sSample = dRow['id'].upper()
            if sExperiment not in dSamples[sSample]['experiments']:
              dSamples[sSample]['experiments'][sExperiment] = list()
            if dRow['well_title'] not in dSamples[sSample]['experiments'][sExperiment]:
              dSamples[sSample]['experiments'][sExperiment].append(dRow['well_title'])
    except Exception as dError:
      Log("\nErro: Ocorreu um erro no processamento da pasta {}.".format(sExperiment))
      Log("Mensagem de erro:\n" + str(dError))
      raise
  return

def PrintSampleData(sInput):
  global dSamples
  print("")
  print("Dados da amostra {} -".format(sInput))
  print("- Nome: {}".format(dSamples[sInput]['name']))
  print("- Data de nascimento: {}".format(dSamples[sInput]['birthday']))
  print("- Processo: {}".format(dSamples[sInput]['record']))
  print("- Data de colheita: {}".format(dSamples[sInput]['sample_date']))
  print("- Servico: {}".format(dSamples[sInput]['department']))
  print("- Experiencias:")
  for sRow, dRow in dSamples[sInput]['experiments'].items():
    print("  - {} (pos. \"{}\")".format(sRow, "\",  \"".join(dRow)))
  print("")

try:
  LoadConfigOptions()
  UpdateSamples()
except Exception as dError:
  Log("\nErro: Occorreu um erro durante a inicialização.")
  Log("Mensagem de erro:\n" + str(dError))
  SaveLog()
  sys.exit()

os.system("cls")

bInput = True
while bInput:
  sInput = input("Amostra ou nome (\'C para cancelar\'):\n  ").strip().upper()
  if sInput.lower() == "c":
    bInput = False
  elif sInput == "":
    continue
  else:
    if sInput in dSamples:
      PrintSampleData(sInput)
    else:
      bFound = False
      sPattern = ".*"+".*".join(RemoveAccents(sInput).split())+".*"
      for sRow, dRow in dSamples.items():
        if dRow['name'] and re.match(sPattern, dRow['name'], flags = re.IGNORECASE):
          bFound = True
          PrintSampleData(sRow.upper())
      if not bFound:
        print("")
        print("Atencao: Amostra nao encontrada.")
        print("")
