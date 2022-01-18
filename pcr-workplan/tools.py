# tools.py (utf-8)
# 
# Edited by: RR-DSE
# Timestamp: 22-01-18 20:42:17

# ------------
# Dependencies
# ------------

import io
import os
import sys
import csv
import datetime
import re
import csv

# ---------
# Constants
# ---------

sEncoding = "utf8"
sQuote = '"'
sDelimiter = ","

# Functions

def LoadTable(sSrcFile):
  global sDelimiter, sQuote, sEncoding
  try:
    oSrcFile = open(sSrcFile, "r", newline = "", encoding = sEncoding)
  except IOError:
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
    del oReaderSrc
    oMemFile.close()
    raise
  del oReaderSrc
  oMemFile.close()
  return (lsKeys, ldTable)

def LoadTableMem(oMemString):
  global sDelimiter, sQuote, sEncoding
  try:
    oMemString.seek(0)
  except IOError:
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
    del oReaderSrc
    oMemFile.close()
    raise
  del oReaderSrc
  oMemFile.close()
  return (lsKeys, ldTable)

def LoadTableEx(sSrcFile, sDelimiter, sQuote, sEncoding):
  try:
    oSrcFile = open(sSrcFile, "r", newline = "", encoding = sEncoding)
  except IOError:
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
    del oReaderSrc
    oMemFile.close()
    raise
  del oReaderSrc
  oMemFile.close()
  return (lsKeys, ldTable)

def LoadSelectors(sSrcFile):
  try:
    dSrcTable = LoadTable(sSrcFile)
  except:
    raise
  if not "source" in dSrcTable[0]:
    raise Error
  if not "field" in dSrcTable[0]:
    raise Error
  if not "type" in dSrcTable[0]:
    raise Error
  dRes = []
  for dRow in dSrcTable[1]:
    dRes.append((dRow["source"], dRow["field"], dRow["type"]))
  del dSrcTable
  return dRes
  
def SelectTable(ldTable, ldSelectors):
  ldRes = []
  lsNewKeys = [] 
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
    del ldRes
    del lsNewKeys
    raise
  return (lsNewKeys, ldRes)

def SaveTable(ldTable, sDestFile):
  global sDelimiter, sQuote, sEncoding
  try:
    oDestFile = open(sDestFile, "w", newline = "", encoding=sEncoding)
  except IOError:
    raise
  try:
    oHeaderWriter = csv.writer(oDestFile, delimiter = sDelimiter, quotechar = None, skipinitialspace = True)
    oWriter = csv.DictWriter(oDestFile, fieldnames = ldTable[0], delimiter = sDelimiter, quotechar = sQuote, quoting = csv.QUOTE_NONNUMERIC, doublequote = True, escapechar = "\\", skipinitialspace = True)
    oHeaderWriter.writerow(ldTable[0])
    for dRow in ldTable[1]:
      oWriter.writerow(dRow)
  except:
    oDestFile.close()
    raise
  oDestFile.close()
  del oWriter
  del oHeaderWriter
  return True

def JoinTables(ldDestTable, ldSrcTable):
  if ldDestTable[0] != ldSrcTable[0]:
    raise
  for dElem in ldSrcTable[1]:
    ldDestTable[1].append(dElem)
  return True

def LoadDict(sSrcFile, sKey):
  try:
    dSrcTable = LoadTable(sSrcFile)
  except:
    raise
  if not sKey in dSrcTable[0]:
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
  return dRes

def LoadDictList(sSrcFile, sKey):
  try:
    dSrcTable = LoadTable(sSrcFile)
  except:
    raise
  if not sKey in dSrcTable[0]:
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
  return dRes

def LoadDictSimple(sSrcFile, sKey, sField):
  try:
    dSrcTable = LoadTable(sSrcFile)
  except:
    raise
  if not sKey in dSrcTable[0]:
    raise Error
  if not sField in dSrcTable[0]:
    raise Error
  dRes = dict()
  for dLine in dSrcTable[1]:
    dRes[dLine[sKey]] = dLine[sField]
  del dSrcTable
  return dRes

def LoadDictSimpleList(sSrcFile, sKey, sField):
  try:
    dSrcTable = LoadTable(sSrcFile)
  except:
    raise
  if not sKey in dSrcTable[0]:
    raise Error
  if not sField in dSrcTable[0]:
    raise Error
  dRes = dict()
  for dLine in dSrcTable[1]:
    if dLine[sKey] not in dRes:
      dRes[dLine[sKey]] = [dLine[sField]]
    else:
      dRes[dLine[sKey]].append(dLine[sField])
  del dSrcTable
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
    raise
  return (dNewFieldList, dNewTable)

def GetTableFromDictSimple(dSrc, sKeyField, sValueField):
  dNewTable = []
  try:
    for sKey, dItem in dSrc.items():
      dNewTable.append({sKeyField: sKey, sValueField: dItem})
  except:
    raise
  return ([sKeyField, sValueField], dNewTable)

def GetTableFromDictSimpleList(dSrc, sKeyField, sValueField):
  dNewTable = []
  try:
    for sKey, dItem in dSrc.items():
      for sRow in dItem:
        dNewTable.append({sKeyField: sKey, sValueField: sRow})
  except:
    raise
  return ([sKeyField, sValueField], dNewTable)

def CoalesceTable(dSrc, sKey, lsFields):
  dTemp = dict()
  dRes = list()
  lsResHeader = list()
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
