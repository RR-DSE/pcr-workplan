# listools.py (utf-8)
# 
# Edited by: RR-DSE
# Timestamp: 22-01-18 20:28:00

import io
import os
import sys
import csv
import datetime
import re
import sqlite3

# Global constants

C_SELECTOR_NONE = 0
C_SELECTOR_STR = 1
C_SELECTOR_STRNAME = 2
C_SELECTOR_STRINT = 3
C_SELECTOR_STRFLOAT = 4
C_SELECTOR_STRDATE = 5
C_SELECTOR_STRTIME = 6
C_SELECTOR_STRDATETIME = 7
C_SELECTOR_INT = 8
C_SELECTOR_FLOAT = 9

C_VERBOSE = False
oLog = None

sTimestampFormat = "%Y-%m-%d %H:%M:%S"

# Standard table definitions

sEncoding = "utf8"
sQuote = '"'
sDelimiter = ","

# Source table definitions and regional settings

sSrcEncoding = "cp1252"
sSrcQuote = '"'
sSrcDelimiter = "\t"
lsNewLine = ["\r", "\n"]
lsWhitespace = [" ", "\t"]
lsDigits = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "+", "-", ".", "e", "E"]
sDecSep = ","
sDateFormatInput = "%d-%m-%Y"
sTimeFormatInput = "%H:%M:%S"
sDateTimeFormatInput = "%d-%m-%Y %H:%M:%S"
sDateFormatOutput = "%Y-%m-%d"
sTimeFormatOutput = "%H:%M:%S"
sDateTimeFormatOutput = "%Y-%m-%d %H:%M:%S"
lsArticles = ["de", "do", "da", "dos", "das", "o", "a", "os", "as"]

# String constants

sInvQuoteRepl = " "
sInvNewLineRepl = " "

# Classes

class Error(Exception):
  pass

# Functions

def Log(sSrc):
  global oLog
  if oLog:
    oLog.write(sSrc + "\n")
  if C_VERBOSE:
    print(sSrc)
  return True

def IsInt(sSrc):
  try:
    nRes = int(sSrc)
    return True
  except:
    return False

def GetInt(sSrc):
  try:
    nRes = int(sSrc)
    return nRes
  except:
    return 0

def IsFloat(sSrc):
  try:
    fRes = float(sSrc)
    return True
  except:
    return False

def GetFloat(sSrc):
  try:
    fRes = float(sSrc)
    return fRes
  except:
    return None

def CleanStr(sSrc):
  sRes = re.sub(r"^\s+", "", str(sSrc))
  sRes = re.sub(r"\s+$", "", sRes)
  sRes = re.sub(r"\s+", " ", sRes)
  return sRes

def CleanStrName(sSrc):
  global lsArticles
  sRes = CleanStr(sSrc)
  sRes = sRes.title()
  lsWord = re.split(" ", sRes)
  lsRes = [lsWord[0].capitalize()]
  for s in lsWord[1:]:
    if s.casefold() in lsArticles:
      lsRes.append(s.lower())
    else:
      lsRes.append(s)
  sRes = " ".join(lsRes)
  return sRes

def CleanStrInt(sSrc):
  global lsDigits, sDecSep
  sSrc1 = CleanStr(sSrc)
  lsWord = []
  for c in sSrc1:
    if c in lsDigits:
      lsWord.append(c)
    elif c == sDecSep:
      lsWord.append(".")
  fRes = GetFloat("".join(lsWord))
  sRes = str(round(fRes))
  return sRes

def CleanStrFloat(sSrc):
  global lsDigits, sDecSep
  sSrc1 = CleanStr(sSrc)
  lsWord = []
  for c in sSrc1:
    if c in lsDigits:
      lsWord.append(c)
    elif c == sDecSep:
      lsWord.append(".")
  sRes = str(GetFloat("".join(lsWord)))
  return sRes

def CleanStrDate(sSrc):
  global sDateFormatInput, sDateFormatOutput
  sSrc1 = CleanStr(sSrc)
  try:
    dRes = datetime.datetime.strptime(sSrc1, sDateFormatInput)
    sRes = dRes.strftime(sDateFormatOutput)
  except:
    sRes = ''
  return sRes

def CleanStrTime(sSrc):
  global sTimeFormatInput, sTimeFormatOutput
  sSrc1 = CleanStr(sSrc)
  try:
    dRes = datetime.datetime.strptime(sSrc1, sTimeFormatInput)
    sRes = dRes.strftime(sTimeFormatOutput)
  except:
    sRes = ''
  return sRes

def CleanStrDateTime(sSrc):
  global sDateTimeFormatInput, sDateTimeFormatOutput
  sSrc1 = CleanStr(sSrc)
  try:
    dRes = datetime.datetime.strptime(sSrc1, sDateTimeFormatInput)
    sRes = dRes.strftime(sDateTimeFormatOutput)
  except:
    sRes = ''
  return sRes

def CleanInt(sSrc):
  global lsDigits, sDecSep
  sSrc1 = CleanStr(sSrc)
  lsWord = []
  for c in sSrc1:
    if c in lsDigits:
      lsWord.append(c)
    elif c == sDecSep:
      lsWord.append(".")
  fRes = GetFloat("".join(lsWord))
  nRes = round(fRes)
  return nRes

def CleanFloat(sSrc):
  global lsDigits, sDecSep
  sSrc1 = CleanStr(sSrc)
  lsWord = []
  for c in sSrc1:
    if c in lsDigits:
      lsWord.append(c)
    elif c == sDecSep:
      lsWord.append(".")
  fRes = GetFloat("".join(lsWord))
  return fRes

def GetTable(sSrcFile, uSrcFileStartLine, uSrcFileTrimLines):
  global lsWhitespace, lsNewLine, sInvQuoteRepl, sInvNewLineRepl, sSrcEncoding, sSrcDelimiter, sSrcQuote
  Log("GetTable/Info: Importing from {0}...".format(sSrcFile))
  uSrcFileStartLine1 = uSrcFileStartLine
  if uSrcFileStartLine < 1:
    uSrcFileStartLine1 = 1
  uSrcFileTrimLines1 = uSrcFileTrimLines
  if uSrcFileTrimLines1 < 0:
    uSrcFileTrimLines1 = 0
  try:
    oSrcFile = open(sSrcFile, "r", newline = "", encoding = sSrcEncoding)
  except IOError:
    Log("GetTable/Error: Unable to open {0}.".format(sSrcFile))
    raise
  for i in range(uSrcFileStartLine1 - 1):
    oSrcFile.readline()
  sSrcStream = oSrcFile.read()
  oSrcFile.close()
  Log("GetTable/Info: Pre-processing the file...")
  oMemFile = io.StringIO(newline = "")
  uQuoteMode = 0
  for c in sSrcStream:
    if uQuoteMode == 0:
      if c == sSrcDelimiter:
        uQuoteMode = 4
      elif c in lsWhitespace:
        uQuoteMode = 0
      elif c == sSrcQuote:
        uQuoteMode = 2
        oMemFile.write(c)
      elif c in lsNewLine:
        uQuoteMode = 0
        oMemFile.write(c)
      else:
        uQuoteMode = 1
        oMemFile.write(c)
    elif uQuoteMode == 1:
      if c == sSrcDelimiter:
        uQuoteMode = 4
      elif c in lsNewLine:
        uQuoteMode = 0
        oMemFile.write(c)
      elif c == sSrcQuote:
        uQuoteMode = 1
        oMemFile.write(sInvQuoteRepl)
      else:
        uQuoteMode = 1
        oMemFile.write(c)
    elif uQuoteMode == 2:
      if c in lsNewLine:
        uQuoteMode = 2
        oMemFile.write(sInvNewLineRepl)
      elif c == sSrcQuote:
        uQuoteMode = 3
      else:
        uQuoteMode = 2
        oMemFile.write(c)
    elif uQuoteMode == 3:
      if c == sSrcDelimiter:
        uQuoteMode = 4
        oMemFile.write(sSrcQuote)
      elif c in lsWhitespace:
        uQuoteMode = 3
      elif c in lsNewLine:
        uQuoteMode = 0
        oMemFile.write(sSrcQuote)
        oMemFile.write(c)
      elif c == sSrcQuote:
        uQuoteMode = 3
        oMemFile.write(sInvQuoteRepl)
      else:
        uQuoteMode = 2
        oMemFile.write(sInvQuoteRepl)
        oMemFile.write(c)
    elif uQuoteMode == 4:
      if c == sSrcDelimiter:
        uQuoteMode = 4
        oMemFile.write(c)
      elif c in lsWhitespace:
        uQuoteMode = 0
        oMemFile.write(sSrcDelimiter)
      elif c == sSrcQuote:
        uQuoteMode = 2
        oMemFile.write(sSrcDelimiter)
        oMemFile.write(c)
      elif c in lsNewLine:
        uQuoteMode = 0
        oMemFile.write(c)
      else:
        uQuoteMode = 1
        oMemFile.write(sSrcDelimiter)
        oMemFile.write(c)
  del sSrcStream
  oMemFile.seek(0)
  oReaderSrcKeys = csv.reader(oMemFile, delimiter = sSrcDelimiter, quotechar = sSrcQuote, skipinitialspace = True)
  try:
    lsKeys = next(oReaderSrcKeys)
  except StopIteration:
    Log("GetTable/Error: Source file {0} is empty.".format(sSrcFile))
    del oReaderSrcKeys
    oMemFile.close()
    raise
  del oReaderSrcKeys
  oMemFile.seek(0)
  oReaderSrc = csv.DictReader(oMemFile, delimiter = sSrcDelimiter, quotechar = sSrcQuote, doublequote = True, escapechar = None, skipinitialspace = True)
  if uSrcFileTrimLines1 == 0:
    ldTable = list(oReaderSrc)
  else:
    ldTable = list(oReaderSrc)[:-uSrcFileTrimLines1]
  for ldTableElem in ldTable:
    for lsTableKey in lsKeys:
      if isinstance(ldTableElem[lsTableKey], str) and ldTableElem[lsTableKey].lower() == "null":
        ldTableElem[lsTableKey] = None
  del oReaderSrc
  oMemFile.close()
  Log("GetTable/Info: Process completed.")
  return (lsKeys, ldTable)

def GetTableNoPreprocess(sSrcFile, uSrcFileStartLine, uSrcFileTrimLines):
  global lsWhitespace, lsNewLine, sInvQuoteRepl, sInvNewLineRepl, sSrcEncoding, sSrcDelimiter, sSrcQuote
  Log("GetTableNoPreprocess/Info: Importing from {0}...".format(sSrcFile))
  uSrcFileStartLine1 = uSrcFileStartLine
  if uSrcFileStartLine < 1:
    uSrcFileStartLine1 = 1
  uSrcFileTrimLines1 = uSrcFileTrimLines
  if uSrcFileTrimLines1 < 0:
    uSrcFileTrimLines1 = 0
  try:
    oSrcFile = open(sSrcFile, "r", newline = "", encoding = sSrcEncoding)
  except IOError:
    Log("GetTableNoPreprocess/Error: Unable to open {0}.".format(sSrcFile))
    raise
  for i in range(uSrcFileStartLine1 - 1):
    oSrcFile.readline()
  sSrcStream = oSrcFile.read()
  oSrcFile.close()
  oMemFile = io.StringIO(newline = "")
  oMemFile.write(sSrcStream)
  del sSrcStream
  oMemFile.seek(0)
  oReaderSrcKeys = csv.reader(oMemFile, delimiter = sSrcDelimiter, quotechar = sSrcQuote, skipinitialspace = True)
  try:
    lsKeys = next(oReaderSrcKeys)
  except StopIteration:
    Log("GetTableNoPreprocess/Error: Source file {0} is empty.".format(sSrcFile))
    del oReaderSrcKeys
    oMemFile.close()
    raise
  del oReaderSrcKeys
  oMemFile.seek(0)
  oReaderSrc = csv.DictReader(oMemFile, delimiter = sSrcDelimiter, quotechar = sSrcQuote, doublequote = True, escapechar = None, skipinitialspace = True)
  if uSrcFileTrimLines1 == 0:
    ldTable = list(oReaderSrc)
  else:
    ldTable = list(oReaderSrc)[:-uSrcFileTrimLines1]
  for ldTableElem in ldTable:
    for lsTableKey in lsKeys:
      if isinstance(ldTableElem[lsTableKey], str) and ldTableElem[lsTableKey].lower() == "null":
        ldTableElem[lsTableKey] = None
  del oReaderSrc
  oMemFile.close()
  Log("GetTableNoPreprocess/Info: Process completed.")
  return (lsKeys, ldTable)

def LoadTable(sSrcFile):
  global sDelimiter, sQuote, sEncoding
  Log("LoadTable/Info: Loading from {0}...".format(sSrcFile))
  try:
    oSrcFile = open(sSrcFile, "r", newline = "", encoding = sEncoding)
  except IOError:
    Log("LoadTable/Error: Unable to open {0}.".format(sSrcFile))
    raise
  sSrcStream = oSrcFile.read()
  oSrcFile.close()
  oMemFile = io.StringIO(newline = "")
  oMemFile.write(sSrcStream)
  del sSrcStream
  oMemFile.seek(0)
  oReaderSrcKeys = csv.reader(oMemFile, delimiter = sDelimiter, quotechar = sQuote, skipinitialspace = True)
  try:
    lsKeys = next(oReaderSrcKeys)
  except:
    Log("LoadTable/Error: Unable to read {0}.".format(sSrcFile))
    del oReaderSrcKeys
    oMemFile.close()
    raise
  del oReaderSrcKeys
  oMemFile.seek(0)
  oMemFile.readline()
  oReaderSrc = csv.DictReader(oMemFile, fieldnames = lsKeys, delimiter = sDelimiter, quotechar = sQuote, quoting = csv.QUOTE_NONNUMERIC, doublequote = True, escapechar = "\\", skipinitialspace = True)
  try:
    ldTable = list(oReaderSrc)
    for ldTableElem in ldTable:
      for lsTableKey in lsKeys:
        if isinstance(ldTableElem[lsTableKey], str) and (ldTableElem[lsTableKey].lower() == "null" or ldTableElem[lsTableKey] == ""):
          ldTableElem[lsTableKey] = None
  except:
    Log("LoadTable/Error: Invalid format for {0}.".format(sSrcFile))
    del oRreaderSrc
    oMemFile.close()
    raise
  del oReaderSrc
  oMemFile.close()
  Log("LoadTable/Info: Process completed.")
  return (lsKeys, ldTable)

def LoadTableMem(oMemString):
  global sDelimiter, sQuote, sEncoding
  Log("LoadTableMem/Info: Loading from StringIO memory object...")
  try:
    oMemString.seek(0)
  except IOError:
    Log("LoadTableMem/Error: Unable to open {0}.".format(sSrcFile))
    raise
  sSrcStream = oMemString.read()
  oMemFile = io.StringIO(newline = "")
  oMemFile.write(sSrcStream)
  del sSrcStream
  oMemFile.seek(0)
  oReaderSrcKeys = csv.reader(oMemFile, delimiter = sDelimiter, quotechar = sQuote, skipinitialspace = True)
  try:
    lsKeys = next(oReaderSrcKeys)
  except:
    Log("LoadTableMem/Error: Unable to read from StringIO memory object.")
    del oReaderSrcKeys
    oMemFile.close()
    raise
  del oReaderSrcKeys
  oMemFile.seek(0)
  oMemFile.readline()
  oReaderSrc = csv.DictReader(oMemFile, fieldnames = lsKeys, delimiter = sDelimiter, quotechar = sQuote, quoting = csv.QUOTE_NONNUMERIC, doublequote = True, escapechar = "\\", skipinitialspace = True)
  try:
    ldTable = list(oReaderSrc)
    for ldTableElem in ldTable:
      for lsTableKey in lsKeys:
        if isinstance(ldTableElem[lsTableKey], str) and ldTableElem[lsTableKey].lower() == "null":
          ldTableElem[lsTableKey] = None
  except:
    Log("LoadTableMem/Error: Invalid format for StringIO memory object.")
    del oRreaderSrc
    oMemFile.close()
    raise
  del oReaderSrc
  oMemFile.close()
  Log("LoadTableMem/Info: Process completed.")
  return (lsKeys, ldTable)

def LoadTableEx(sSrcFile, sDelimiter, sQuote, sEncoding):
  Log("LoadTableEx/Info: Loading from {0}...".format(sSrcFile))
  try:
    oSrcFile = open(sSrcFile, "r", newline = "", encoding = sEncoding)
  except IOError:
    Log("LoadTableEx/Error: Unable to open {0}.".format(sSrcFile))
    raise
  sSrcStream = oSrcFile.read()
  oSrcFile.close()
  oMemFile = io.StringIO(newline = "")
  oMemFile.write(sSrcStream)
  del sSrcStream
  oMemFile.seek(0)
  oReaderSrcKeys = csv.reader(oMemFile, delimiter = sDelimiter, quotechar = sQuote, skipinitialspace = True)
  try:
    lsKeys = next(oReaderSrcKeys)
  except:
    Log("LoadTableEx/Error: Unable to read {0}.".format(sSrcFile))
    del oReaderSrcKeys
    oMemFile.close()
    raise
  del oReaderSrcKeys
  oMemFile.seek(0)
  oMemFile.readline()
  oReaderSrc = csv.DictReader(oMemFile, fieldnames = lsKeys, delimiter = sDelimiter, quotechar = sQuote, quoting = csv.QUOTE_MINIMAL, doublequote = True, escapechar = "\\", skipinitialspace = True)
  try:
    ldTable = list(oReaderSrc)
    for ldTableElem in ldTable:
      for lsTableKey in lsKeys:
        if isinstance(ldTableElem[lsTableKey], str) and ldTableElem[lsTableKey].lower() == "null":
          ldTableElem[lsTableKey] = None
  except:
    Log("LoadTableEx/Error: Invalid format for {0}.".format(sSrcFile))
    del oRreaderSrc
    oMemFile.close()
    raise
  del oReaderSrc
  oMemFile.close()
  Log("LoadTableEx/Info: Process completed.")
  return (lsKeys, ldTable)

def LoadSelectors(sSrcFile):
  Log("LoadSelectors/Info: Loading from {0}...".format(sSrcFile))
  try:
    dSrcTable = LoadTable(sSrcFile)
  except:
    Log("LoadSelectors/Error: Unable to read {0}.".format(sSrcFile))
    raise
  if not "source" in dSrcTable[0]:
    Log("LoadSelectors/Error: Invalid format while reading {0}.".format(sSrcFile))
    raise Error
  if not "field" in dSrcTable[0]:
    Log("LoadSelectors/Error: Invalid format while reading {0}.".format(sSrcFile))
    raise Error
  if not "type" in dSrcTable[0]:
    Log("LoadSelectors/Error: Invalid format while reading {0}.".format(sSrcFile))
    raise Error
  dRes = []
  for dRow in dSrcTable[1]:
    dRes.append((dRow["source"], dRow["field"], dRow["type"]))
  del dSrcTable
  Log("LoadSelectors/Info: Process completed.")
  return dRes
  
def SelectTable(ldTable, ldSelectors):
  ldRes = []
  lsNewKeys = [] 
  Log("SelectTable/Info: Selecting a table...")
  try:
    for dSelector in ldSelectors:
      lsNewKeys.append(dSelector[1])
    for dRow in ldTable[1]:
      dNewRow = {}
      for dSelector in ldSelectors:
        if dSelector[2] == 0:
          dNewRow[dSelector[1]] = dRow[dSelector[0]]
        elif dSelector[2] == 1:
          dNewRow[dSelector[1]] = CleanStr(dRow[dSelector[0]])
        elif dSelector[2] == 2:
          dNewRow[dSelector[1]] = CleanStrName(dRow[dSelector[0]])
        elif dSelector[2] == 3:
          dNewRow[dSelector[1]] = CleanStrInt(dRow[dSelector[0]])
        elif dSelector[2] == 4:
          dNewRow[dSelector[1]] = CleanStrFloat(dRow[dSelector[0]])
        elif dSelector[2] == 5:
          dNewRow[dSelector[1]] = CleanStrDate(dRow[dSelector[0]])
        elif dSelector[2] == 6:
          dNewRow[dSelector[1]] = CleanStrTime(dRow[dSelector[0]])
        elif dSelector[2] == 7:
          dNewRow[dSelector[1]] = CleanStrDateTime(dRow[dSelector[0]])
        elif dSelector[2] == 8:
          dNewRow[dSelector[1]] = CleanInt(dRow[dSelector[0]])
        elif dSelector[2] == 9:
          dNewRow[dSelector[1]] = CleanFloat(dRow[dSelector[0]])
        else:
          dNewRow[dSelector[1]] = dRow[dSelector[0]]
      ldRes.append(dNewRow)
  except:
    Log("SelectTable/Error: Invalid format.")
    del ldRes
    del lsNewKeys
    raise
  Log("SelectTable/Info: Process completed.")
  return (lsNewKeys, ldRes)

def GetSelectTable(sSrcFile, uSrcFileStartLine, uSrcFileTrimLines, sSelectorsFile):
  Log("GetSelectTable/Info: Loading table from {0}, using selectors from {1}...".format(sSrcFile, sSelectorsFile))
  try:
    dSrcTable = GetTable(sSrcFile, uSrcFileStartLine, uSrcFileTrimLines)
    dSelectors = LoadSelectors(sSelectorsFile)
    dRes = SelectTable(dSrcTable, dSelectors)
  except:
    Log("GetSelectTable/Info: Unable to load table from {0}, using selectors from {1}.".format(sSrcFile, sSelectorsFile))
    raise
  del dSrcTable
  del dSelectors
  Log("GetSelectTable/Info: Process completed.")
  return dRes

def SaveTable(ldTable, sDestFile):
  global sDelimiter, sQuote, sEncoding
  Log("SaveTable/Info: Saving table to {0}...".format(sDestFile))
  try:
    oDestFile = open(sDestFile, "w", newline = "", encoding=sEncoding)
  except IOError:
    Log("SaveTable/Error: Unable to open {0}.".format(sDestFile))
    raise
  try:
    oHeaderWriter = csv.writer(oDestFile, delimiter = sDelimiter, quotechar = None, skipinitialspace = True)
    oWriter = csv.DictWriter(oDestFile, fieldnames = ldTable[0], delimiter = sDelimiter, quotechar = sQuote, quoting = csv.QUOTE_NONNUMERIC, doublequote = True, escapechar = "\\", skipinitialspace = True)
    oHeaderWriter.writerow(ldTable[0])
    for dRow in ldTable[1]:
      oWriter.writerow(dRow)
  except:
    oDestFile.close()
    Log("SaveTable/Error: Invalid format.")
    raise
  oDestFile.close()
  del oWriter
  del oHeaderWriter
  Log("SaveTable/Info: Process completed.")
  return True

def JoinTables(ldDestTable, ldSrcTable):
  Log("JoinTables/Info: Joining tables.")
  if ldDestTable[0] != ldSrcTable[0]:
    Log("JoinTables/Error: Tables have different keys.")
    raise
  for dElem in ldSrcTable[1]:
    ldDestTable[1].append(dElem)
  Log("JoinTables/Info: Process completed.")
  return True

def LoadDict(sSrcFile, sKey):
  Log("LoadDict/Info: Loading from {0}...".format(sSrcFile))
  try:
    dSrcTable = LoadTable(sSrcFile)
  except:
    Log("LoadDict/Error: Unable to read {0}.".format(sSrcFile))
    raise
  if not sKey in dSrcTable[0]:
    Log("LoadDict/Error: Invalid format while reading {0}.".format(sSrcFile))
    raise Error
  dRes = dict()
  for dLine in dSrcTable[1]:
    dNewDict = dict()
    for sField in dSrcTable[0]:
      if sField == sKey:
        continue
      else:
        dNewDict[sField] = dLine[sField]
    dRes[dLine[sKey]] = dNewDict
  del dSrcTable
  Log("LoadDict/Info: Process completed.")
  return dRes

def LoadDictList(sSrcFile, sKey):
  Log("LoadDictList/Info: Loading from {0}...".format(sSrcFile))
  try:
    dSrcTable = LoadTable(sSrcFile)
  except:
    Log("LoadDictList/Error: Unable to read {0}.".format(sSrcFile))
    raise
  if not sKey in dSrcTable[0]:
    Log("LoadDictList/Error: Invalid format while reading {0}.".format(sSrcFile))
    raise Error
  dRes = dict()
  for dLine in dSrcTable[1]:
    dNewDict = dict()
    for sField in dSrcTable[0]:
      if sField == sKey:
        continue
      else:
        dNewDict[sField] = dLine[sField]
    if dLine[sKey] in dRes:
      dRes[dLine[sKey]].append(dNewDict)
    else:
      dRes[dLine[sKey]] = [dNewDict]
  del dSrcTable
  Log("LoadDictList/Info: Process completed.")
  return dRes

def LoadDictSimple(sSrcFile, sKey, sField):
  Log("LoadDictSimple/Info: Loading from {0}...".format(sSrcFile))
  try:
    dSrcTable = LoadTable(sSrcFile)
  except:
    Log("LoadDictSimple/Error: Unable to read {0}.".format(sSrcFile))
    raise
  if not sKey in dSrcTable[0]:
    Log("LoadDictSimple/Error: Invalid format while reading {0}.".format(sSrcFile))
    raise Error
  if not sField in dSrcTable[0]:
    Log("LoadDictSimple/Error: Invalid format while reading {0}.".format(sSrcFile))
    raise Error
  dRes = dict()
  for dLine in dSrcTable[1]:
    dRes[dLine[sKey]] = dLine[sField]
  del dSrcTable
  Log("LoadDictSimple/Info: Process completed.")
  return dRes

def LoadDictSimpleList(sSrcFile, sKey, sField):
  Log("LoadDictSimpleList/Info: Loading from {0}...".format(sSrcFile))
  try:
    dSrcTable = LoadTable(sSrcFile)
  except:
    Log("LoadDictSimpleList/Error: Unable to read {0}.".format(sSrcFile))
    raise
  if not sKey in dSrcTable[0]:
    Log("LoadDictSimpleList/Error: Invalid format while reading {0}.".format(sSrcFile))
    raise Error
  if not sField in dSrcTable[0]:
    Log("LoadDictSimpleList/Error: Invalid format while reading {0}.".format(sSrcFile))
    raise Error
  dRes = dict()
  for dLine in dSrcTable[1]:
    if dLine[sKey] not in dRes:
      dRes[dLine[sKey]] = [dLine[sField]]
    else:
      dRes[dLine[sKey]].append(dLine[sField])
  del dSrcTable
  Log("LoadDictSimpleList/Info: Process completed.")
  return dRes

def GetTableFromDict(dSrc, sKeyField):
  dNewTable = []
  dNewFieldList = []
  bFieldListSet = False
  try:
    for sKey, dItem in dSrc.items():
      dNewDict = dict()
      if not bFieldListSet:
        dNewFieldList = list(dItem.keys())
        bFieldListSet = True
      dNewDict[sKeyField] = sKey
      for sKey2 in dNewFieldList:
        dNewDict[sKey2] = dItem[sKey2]
      dNewTable.append(dNewDict)
    dNewFieldList.insert(0,sKeyField)
  except:
    Log("GetTableFromDict/Error: Invalid format.")
    raise
  return (dNewFieldList, dNewTable)

def GetTableFromDictEx(dSrc, sKeyField, dFieldList):
  dNewTable = []
  dNewFieldList = []
  bFieldListSet = False
  if dFieldList != None:
    dNewFieldList = dFieldList
    bFieldListSet = True
  try:
    for sKey, dItem in dSrc.items():
      dNewDict = dict()
      if not bFieldListSet:
        dNewFieldList = list(dItem.keys())
        bFieldListSet = True
      if sKeyField != None:
        dNewDict[sKeyField] = sKey
      for sKey2 in dNewFieldList:
        dNewDict[sKey2] = dItem[sKey2]
      dNewTable.append(dNewDict)
    if sKeyField != None:
      dNewFieldList.insert(0,sKeyField)
  except:
    Log("GetTableFromDictEx/Error: Invalid format.")
    raise
  return (dNewFieldList, dNewTable)

def GetTableFromDictList(dSrc, sKeyField):
  dNewTable = []
  dNewFieldList = []
  bFieldListSet = False
  try:
    for sKey, dItem in dSrc.items():
      for dRow in dItem:
        dNewDict = dict()
        dNewDict[sKeyField] = sKey
        if not bFieldListSet:
          dNewFieldList = list(dRow.keys())
          bFieldListSet = True
        for sKey2 in dNewFieldList:
          dNewDict[sKey2] = dRow[sKey2]
        dNewTable.append(dNewDict)
    dNewFieldList.insert(0,sKeyField)
  except:
    Log("GetTableFromDictList/Error: Invalid format.")
    raise
  return (dNewFieldList, dNewTable)

def GetTableFromDictListUniqueKey(dSrc, sKeyField):
  dNewTable = []
  dNewFieldList = []
  bFieldListSet = False
  try:
    for sKey, dItem in dSrc.items():
      uCount = 1
      for dRow in dItem:
        dNewDict = dict()
        dNewDict[sKeyField] = sKey + "-" + str(uCount)
        uCount = uCount + 1
        if not bFieldListSet:
          dNewFieldList = list(dRow.keys())
          bFieldListSet = True
        for sKey2 in dNewFieldList:
          dNewDict[sKey2] = dRow[sKey2]
        dNewTable.append(dNewDict)
    dNewFieldList.insert(0,sKeyField)
  except:
    Log("GetTableFromDictListUniqueKey/Error: Invalid format.")
    raise
  return (dNewFieldList, dNewTable)

def GetTableFromDictSimple(dSrc, sKeyField, sValueField):
  dNewTable = []
  try:
    for sKey, dItem in dSrc.items():
      dNewTable.append({sKeyField: sKey, sValueField: dItem})
  except:
    Log("GetTableFromDictSimple/Error: Invalid format.")
    raise
  return ([sKeyField, sValueField], dNewTable)

def GetTableFromDictSimpleList(dSrc, sKeyField, sValueField):
  dNewTable = []
  try:
    for sKey, dItem in dSrc.items():
      for sRow in dItem:
        dNewTable.append({sKeyField: sKey, sValueField: sRow})
  except:
    Log("GetTableFromDictSimpleList/Error: Invalid format.")
    raise
  return ([sKeyField, sValueField], dNewTable)

def OpenDB(sDBFile):
  Log("OpenDB/Info: Opening database file {0}...".format(sDBFile))
  try:
    oSQLConn = sqlite3.connect(sDBFile)
    oSQLConn.row_factory = sqlite3.Row
    oSQLCursor = oSQLConn.cursor()
  except:
    Log("OpenDB/Error: Unable to open database {0}.".format(sDBFile))
    raise
  Log("OpenDB/Info: Process completed.")
  return (oSQLConn, oSQLCursor)

def QueryDB(oSQLLink, sQuery):
  oSQLLink[1].execute(sQuery)

def CloseDB(oSQLLink):
  oSQLLink[0].commit()
  oSQLLink[0].close()
  return True;

def InsertDBTable(oSQLLink, ldTable, sTableName):
  ldNewTable = []
  Log("InsertDBTable/Info: Creating database table {0}...".format(sTableName))
  oSQLLink[1].execute("DROP TABLE IF EXISTS {0}".format(sTableName))
  try:
    for dRow in ldTable[1]:
      dNewRow = []
      for sKey in ldTable[0]:
        dNewRow.append(dRow[sKey])
      ldNewTable.append(tuple(dNewRow))
    sQuery1 = "CREATE TABLE {0} (".format(sTableName) + "\"" + "\", \"".join(ldTable[0]) + "\"" + ")"
    lsQuery2 = []
    for sKey in ldTable[0]:
      lsQuery2.append("?")
    sQuery2 = "INSERT INTO {0} VALUES (".format(sTableName) + ",".join(lsQuery2) + ")"
    oSQLLink[1].execute(sQuery1)
    oSQLLink[1].executemany(sQuery2, ldNewTable)
  except:
    Log("InsertDBTable/Error: An SQL error occurred while creating table {0}.".format(sTableName))
    del ldNewTable
    raise
  del ldNewTable
  Log("InsertDBTable/Info: Process completed.")
  return True

def InsertDBValues(oSQLLink, sTableName, ldTable, lsColumns):
  ldNewTable = []
  if lsColumns == None:
    lsKeys = ldTable[0]
  else:
    lsKeys = lsColumns
  Log("InsertDBValues/Info: Updating database table {0}...".format(sTableName))
  try:
    for dRow in ldTable[1]:
      dNewRow = []
      for sKey in lsKeys:
        dNewRow.append(dRow[sKey])
      ldNewTable.append(tuple(dNewRow))
    lsQuery = []
    for sKey in lsKeys:
      lsQuery.append("?")
    sQuery = "INSERT INTO {0} VALUES (".format(sTableName) + ",".join(lsQuery) + ")"
    oSQLLink[1].executemany(sQuery, ldNewTable)
  except:
    Log("InsertDBValues/Error: An SQL error occurred while updating table {0}.".format(sTableName))
    del ldNewTable
    raise
  del ldNewTable
  Log("InsertDBValues/Info: Process completed.")
  return True

def GetDBTable(oSQLLink, sTable, lsColumns):
  ldTable = []
  lsKeys = list(lsColumns)
  lsKeys1 = list(lsColumns)
  for sElem in enumerate(lsKeys1):
    if sElem[1].lower() == "group":
      lsKeys1[sElem[0]] = '"{0}"'.format(lsKeys1[sElem[0]])
  sQuery = """
    SELECT {0} FROM {1}
  """.format(",".join(lsKeys1), sTable)
  Log("GetDBTable/Info: Reading database table {0}...".format(sTable))
  try:
    for dResRow in oSQLLink[1].execute(sQuery):
      dRow = {}
      for sKey in lsKeys:
        dRow[sKey] = dResRow[sKey]
      ldTable.append(dRow)
  except:
    Log("GetDBTable/Error: An error occurred while reading table {0}.".format(sTable))
    del ldTable
    del lsKeys
    del lsKeys1
    raise
  Log("GetDBTable/Info: Process completed.")
  return (lsKeys, ldTable)

def ExportDBTable(oSQLLink, sTable, lsColumns, sDestFile):
  Log("ExportDBTable/Info: Exporting database table {0}...".format(sTable))
  try:
    ldTable = GetDBTable(oSQLLink[1], sTable, lsColumns)
  except:
    Log("ExportDBTable/Error: An error occurred while exporting table {0}.".format(sTable))
    raise
  try:
    SaveTable(ldTable, sDestFile)
  except:
    Log("ExportDBTable/Error: An error occurred while exporting table {0}.".format(sTable))
    del ldTable
    raise
  del ldTable
  Log("ExportDBTable/Info: Process completed.")
  return True

def TranslateTable(dTable, dTransTable):
  try:
    dTrans = {}
    for sField in dTable[0]:
      dTrans[sField] = []
    for dRow in dTransTable[1]:
      if dRow['field'] in dTrans:
        dTrans[dRow['field']].append((dRow['search'], dRow['subst']))
  except:
    Log("TranslateTable/Error: Invalid format for the source file {0}.".format(sTransFile))
    raise
  for dRow in dTable[1]:
    for sField in dTable[0]:
      for dTransEl in dTrans[sField]:
        dRow[sField] = re.sub(dTransEl[0], dTransEl[1], dRow[sField], flags=re.IGNORECASE)
  del dTrans
  return True

def TranslateTableFromFile(dTable, sTransFile):
  try:
    dTransTable = LoadTable(sTransFile)
  except:
    Log("TranslateTable/Error: Unable to read the source file {0}.".format(sTransFile))
    raise
  try:
    dTrans = {}
    for sField in dTable[0]:
      dTrans[sField] = []
    for dRow in dTransTable[1]:
      if dRow['field'] in dTrans:
        dTrans[dRow['field']].append((dRow['search'], dRow['subst']))
  except:
    Log("TranslateTable/Error: Invalid format for the source file {0}.".format(sTransFile))
    raise
  for dRow in dTable[1]:
    for sField in dTable[0]:
      for dTransEl in dTrans[sField]:
        dRow[sField] = re.sub(dTransEl[0], dTransEl[1], dRow[sField], flags=re.IGNORECASE)
  del dTrans
  return True

def CoalesceTable(dSrc, sKey, lsFields):
  dTemp = dict()
  dRes = list()
  lsResHeader = list()
  Log("CoalesceTable/Info: Coalescing table with key \"{0}\"...".format(sKey))
  try:
    lsResHeader.append(sKey)
    for sField in lsFields:
      lsResHeader.append(sField)
    for dRow in dSrc[1]:
      if dRow[sKey] not in dTemp:
        dTemp[dRow[sKey]] = dict()
        dTemp[dRow[sKey]][sKey] = dRow[sKey]
        for sField in lsFields:
          dTemp[dRow[sKey]][sField] = None
      for sField in lsFields:
        if dRow[sField] != None and dRow[sField] != "":
          dTemp[dRow[sKey]][sField] = dRow[sField]
    for sKey in sorted(dTemp.keys()):
      bOK = True
      for sField in lsFields:
        if dTemp[sKey][sField] == None:
          bOK = False
          break
      if bOK:
        dRes.append(dTemp[sKey])
  except:
    Log("CoalesceTable/Error: Invalid parameters.")
    raise
  dResTable = (lsResHeader, dRes)
  return dResTable

def SelectUniqueRows(dSrc, sKey):
  dRes = list()
  dKey = dict()
  print("SelectUniqueRows/Info: Selecting unique rows...")
  uIndex = 0
  for dRow in dSrc[1]:
    if dRow[sKey] not in dKey:
      dKey[dRow[sKey]] = uIndex
      dRes.append(dRow)
      uIndex = uIndex + 1
    else:
      dRes[dKey[dRow[sKey]]] = dRow
  return (dSrc[0], dRes)

def GetTableOrders(sSrcFile):
  global lsWhitespace, sSrcEncoding, sSrcQuote
  sSrcDelimiter = "|"
  Log("GetTableOrders/Info: Importing from {0}...".format(sSrcFile))
  try:
    oSrcFile = open(sSrcFile, "r", newline = "", encoding = sSrcEncoding)
  except IOError:
    Log("GetTableOrders/Error: Unable to open {0}.".format(sSrcFile))
    raise
  sSrcStream = oSrcFile.read()
  oSrcFile.close()
  Log("GetTableOrders/Info: Processing the file...")
  oMemFile = io.StringIO(sSrcStream, newline = "")
  oMemFile.seek(0)
  del sSrcStream
  lsSrcKeys = ["episode","datetime","record","worksample","name","birthday","gender","insurancenumber","other_1","other_2","other_3","other_4","other_5","other_6","other_7","sample","tests"]
  lsKeys = ["record","episode","name","birthday","gender","insurancenumber","datetime","sample","test"]
  oReaderSrc = csv.DictReader(oMemFile, fieldnames=lsSrcKeys, delimiter = sSrcDelimiter, quotechar = sSrcQuote, doublequote = True, escapechar = None, skipinitialspace = True)
  oMemFile.seek(0)
  ldTable = list()
  for dRow in oReaderSrc:
    sBirthday = datetime.datetime.strptime(dRow['birthday'], "%Y%m%d").strftime("%Y-%m-%d")
    sDateTime = datetime.datetime.strptime(dRow['datetime'], "%Y%m%d%H%M").strftime("%Y-%m-%d %H:%M")
    lsTests = dRow['tests'].split(',')
    for dTest in lsTests:
        ldTable.append(
          { 
            'record': dRow['record'],
            'episode': dRow['episode'],
            'name': CleanStrName(dRow['name']),
            'birthday': sBirthday,
            'gender': dRow['gender'],
            'insurancenumber': dRow['insurancenumber'],
            'datetime': sDateTime,
            'sample': dRow['sample'],
            'test': dTest
          }
        )
  del oReaderSrc
  oMemFile.close()
  Log("GetTableOrder/Info: Process completed.")
  return (lsKeys, ldTable)

# Main routine start

if C_VERBOSE:
  os.system("cls")
Log("")
Log("---")
Log("Ferramenta para importação de dados do LIS")
Log("Versão: 2020-08-23")
Log("Data e hora: " + datetime.datetime.now().strftime(sTimestampFormat))
Log("---")
