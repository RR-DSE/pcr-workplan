# report.py (utf-8)
# 
# Edited by: RR-DSE
# Timestamp: 22-01-18 20:26:50

import fpdf
import tools
import os
import io
import sys
import datetime
import xlrd
import glob
import csv
import math
import matplotlib.pyplot as pyplot

# Constants

sConfigFolder = "config"
sTemporaryFolder = "temp"

# Global variables

sDateTimeFormat = "%Y-%m-%d %H:%M:%S"

fPlotTextHeight = 0.05
fPlotTextVStep = 0.02
  
oReport = None
dConfigReport = None
dConfigOptions = None
dDateTime = None

fBaselineMin = 3
fBaselineMax = 10

dProtocols = dict()

# Configuration functions

def LoadConfigReport():
  global dConfigReport, sConfigFolder
  try:
    dConfigReport = tools.LoadDictSimple("{}\\report.csv".format(sConfigFolder), "parameter", "value")
  except Exception as dError:
    print("Erro: Não foi possível carregar o ficheiro de configuração dos relatórios PDF.")
    print("Mensagem de erro:\n" + str(dError))
    raise
  return

def LoadConfigOptions():
  global sConfigFolder, dConfigOptions
  try:
    dConfigOptions = tools.LoadDictSimple("{}\\options.csv".format(sConfigFolder), "parameter", "value")
  except Exception as dError:
    print("Erro: Não foi possível carregar o ficheiro de configuração para opções.")
    print("Mensagem de erro:\n" + str(dError))
    raise
  return

# Auxiliary functions

def Min(lfList):
  fRes = lfList[0]
  for fValue in lfList:
    if fValue < fRes:
      fRes = fValue
  return fRes

def Max(lfList):
  fRes = lfList[0]
  for fValue in lfList:
    if fValue > fRes:
      fRes = fValue
  return fRes

def Mean(lfList):
  fN = float(len(lfList))
  fRes = 0.0
  for fValue in lfList:
    fRes = fRes + fValue / fN
  return fRes

def SDSample(lfList):
  fMean = Mean(lfList)
  fN = float(len(lfList) - 1)
  fRes = 0.0
  for fValue in lfList:
    fRes = fRes + (fValue - fMean)**2
  fRes = fRes / fN
  fRes = math.sqrt(fRes)
  return fRes

def BaselineInspection(lfList, fMax):
  global fBaselineMin, fBaselineMax
  for uIndex in range(fBaselineMin+1, fBaselineMax):
    fValue = lfList[uIndex] / fMax
    fMean = Mean(lfList[2:uIndex]) / fMax
    fSD1 = SDSample(lfList[2:uIndex])
    fSD2 = SDSample(lfList[2:uIndex + 1])
    fRes = fValue / fMean
  return None

def CheckThreshold(fThreshold, fMax):
  bRes = True
  if fThreshold / fMax > 0.10:
    bRes = False
  return bRes

def Threshold_1(lfList):
  global fBaselineMin, fBaselineMax
  fMean = Mean(lfList[fBaselineMin-1:fBaselineMax])
  fSD = SDSample(lfList[fBaselineMin-1:fBaselineMax])
  fRes = fMean + 10 * fSD
  return fRes

def Threshold_2(lfList):
  fMean = Mean(lfList)
  fSD = SDSample(lfList)
  fRes = fMean + 10 * fSD
  return fRes

def Ct_1(lfList):
  fT = Threshold_1(lfList)
  uRes = None
  uIndex = 1
  uCurrRes = 0
  uValCount = 0
  for fValue in lfList:
    if fValue > fT:
      if uValCount == 0:
        uCurrRes = uIndex
      uValCount = uValCount + 1
    else:
      uValCount = 0
    if uValCount >= 3:
      break
    uIndex = uIndex + 1
  if uValCount > 0:
    uRes = uCurrRes
  return uRes

def Ct_2(lfList, fT):
  uRes = None
  uIndex = 1
  uCurrRes = 0
  uValCount = 0
  for fValue in lfList:
    if fValue > fT:
      if uValCount == 0:
        uCurrRes = uIndex
      uValCount = uValCount + 1
    else:
      uValCount = 0
    if uValCount >= 3:
      break
    uIndex = uIndex + 1
  if uValCount > 0:
    uRes = uCurrRes
  return uRes

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
          if "types_amplicons" not in dProtocols[dRow['protocol']]:
            dProtocols[dRow['protocol']]['types_amplicons'] = dict()
          if dRow['type'] not in dProtocols[dRow['protocol']]['types_amplicons']:
            dProtocols[dRow['protocol']]['types_amplicons'][dRow['type']] = dict()
          dProtocols[dRow['protocol']]['types_amplicons'][dRow['type']][int(dRow['id'])] = dRow['ref']
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
        if dRow['parameter'].lower() == 'workplan_extra':
          dProtocols[dRow['protocol']]['workplan_extra'] = dRow['title'] if dRow['title'] != None else ""
  except Exception as dError:
    print("Erro: Não foi possível carregar pelo menos um dos ficheiros de configuração para protocolos.")
    print("Mensagem de erro:\n" + str(dError))
    raise
  return

def GetTable(sFile):
  dTable = tools.LoadTable(sFile)[1]
  return dTable

def GetDet1Raw(sFile):
  global fBaselineMin, fBaselineMax
  oSrcFile = open(sFile, "r", newline = "")
  sSrcStream = oSrcFile.read()
  oSrcFile.close()
  oMemFile = io.StringIO(newline = "")
  oMemFile.write(sSrcStream)
  del sSrcStream
  oMemFile.seek(0)
  oReaderSrcKeys = csv.reader(oMemFile, delimiter = "\t", quotechar = "\"", skipinitialspace = True)
  lsKeys = next(oReaderSrcKeys)
  del oReaderSrcKeys
  oMemFile.seek(0)
  oMemFile.readline()
  oReaderSrc = csv.DictReader(oMemFile, fieldnames = lsKeys, delimiter = "\t", quotechar = "\"", quoting = csv.QUOTE_NONE, doublequote = True, escapechar = "\\", skipinitialspace = True)
  ldTable = list(oReaderSrc)
  for dRow in ldTable:
    for sKey, sValue in dRow.items():
      if isinstance(sValue, str) and (sValue.lower() == "null" or sValue == ""):
        dRow[sKey] = None
  del oReaderSrc
  oMemFile.close()
  dData = dict()
  for dRow in ldTable:
    for sKey, sValue in dRow.items():
      if sKey != None and sKey.strip().lower() != "step" and sKey.strip().lower() != "cycle" and sKey.strip().lower() != "channel" and sKey.strip().lower() != "temperature":
        if sKey not in dData:
          dData[sKey] = dict()
        if dRow["Channel"] not in dData[sKey]:
          dData[sKey][dRow["Channel"]] = list()
        dData[sKey][dRow["Channel"]].append(float(sValue.replace(" ", "").replace(",",".")))
  for sWell in dData:
    for sChannel in dData[sWell]:
      #fBase = dData[sWell][sChannel][2]
      fBase = Min(dData[sWell][sChannel][fBaselineMin-1:fBaselineMax])
      uIndex = 0
      for fValue in dData[sWell][sChannel]:
        dData[sWell][sChannel][uIndex] = fValue - fBase
        uIndex = uIndex + 1
  return dData

def GetDet1XLSX(sFile):
  oWB = xlrd.open_workbook(sFile)
  oSheet = oWB.sheet_by_name("Plate")
  lsCol1Values = oSheet.col_values(0)
  lsCol1Types = oSheet.col_types(0)
  uRow = 0
  dData = dict()
  for sCol1Value in lsCol1Values:
    if uRow == 0:
      uRow = uRow + 1
      continue
    if lsCol1Types[uRow] == 1 and sCol1Value != "":
      sWell = sCol1Value.split("-")[0][0] + str(int(sCol1Value.split("-")[0][1:]))
      sChannel = str(int(sCol1Value.split("-")[1].replace("Ch", "")))
      sChannelTitle = oSheet.cell_value(uRow, 1) if oSheet.cell_value(uRow, 1) != "[none]" else None
      sAmplicon = oSheet.cell_value(uRow, 3) if str(oSheet.cell_value(uRow, 3)) != "" else None
      sSample = oSheet.cell_value(uRow, 8) if str(oSheet.cell_value(uRow, 8)) != "" else None
      if sWell and sChannel and sChannelTitle and sAmplicon and sSample:
        if sSample not in dData:
          dData[sSample] = dict()
        if sAmplicon not in dData[sSample]:
          dData[sSample][sAmplicon] = dict()
        dData[sSample][sAmplicon]['well'] = sWell
        dData[sSample][sAmplicon]['channel'] = sChannel
        dData[sSample][sAmplicon]['channel_title'] = sChannelTitle
    uRow = uRow + 1
  oWB.release_resources()
  return dData

# PDF functions

def PlotPCR(
  dPlate,
  dThreshold,
  sSample,
  sAmplicon,
  sFigureFile = None,
  sAmpliconTitle = None,
  sChannelTitle = None,
  uFigure = 1,
  dYAxis = None):
  global fPlotTextHeight, fPlotTextVStep
  lfValues = dPlate[sSample][sAmplicon]['data']
  sChannel = dPlate[sSample][sAmplicon]['channel_title']
  sWell = dPlate[sSample][sAmplicon]['well']
  if dThreshold:
    fThreshold = dThreshold[sChannel][sAmplicon]['value']
  else:
    fThreshold = Threshold_1(lfValues)
  if dYAxis:
    fYMin = dYAxis[sAmplicon][sChannel]['y_axis'][0]
    fYMax = dYAxis[sAmplicon][sChannel]['y_axis'][1]
  else:
    fYMin = Min(lfValues)
    fYMax = Max(lfValues)
  if fThreshold > fYMax:
    fYMax = fThreshold
  fYMin = fYMin - ((fYMax - fYMin) * 0.05)
  if fYMin < 0.0:
    fYMin = 0.0
  fYMax = fYMax + ((fYMax - fYMin) * 0.05)
  fYDelta = fYMax - fYMin
  uCt = Ct_2(lfValues, fThreshold)
  if uCt != None:
    fCt = float(uCt)
  else:
    fCt = None
  lfX = list(range(1, len(lfValues) + 1))
  lfXLabels = list()
  uLabel = 0
  for uTick in lfX:
    if uLabel == 0:
      lfXLabels.append(str(uTick))
      uLabel = 1
    else:
      lfXLabels.append("")
      uLabel = 0
  list(range(1, len(lfValues) + 1, 2))
  fXMax = len(lfValues)
  bThreshold = CheckThreshold(fThreshold, fYMax)
  if bThreshold and fCt and Max(lfValues) / fYMax < 0.50:
    bThreshold = False
  pyplot.figure(num = uFigure, figsize = (6, 3), tight_layout=True)
  pyplot.clf()
  pyplot.plot(lfX, lfValues, "-", color="k", linewidth = 1)
  pyplot.axis("on")
  pyplot.xlim(1, fXMax)
  pyplot.ylim(fYMin, fYMax)
  pyplot.xticks(lfX, lfXLabels, color="k", fontfamily="sans-serif", fontsize="x-small", fontstretch="normal", fontweight="normal")
  pyplot.yticks(color="k", fontfamily="sans-serif", fontsize="x-small", fontstretch="normal", fontweight="normal")
  if bThreshold:
    pyplot.hlines(fThreshold, 1, fXMax, color="k", linestyle = "--", linewidth = 1, alpha = 0.5)
    if fCt:
      pyplot.vlines(fCt, fYMin, fYMax, color="k", linestyle = ":", linewidth = 1, alpha = 0.5)
  else:
    f1 = math.floor(math.log10(fYMax))
    f2 = math.floor(fYMax / 10**f1) * 10**f1
    f3 = f2 / 20.0
    lfYTicks = list()
    fCurr = 0.0
    while fCurr <= fYMax:
      fCurr = fCurr + f3
      lfYTicks.append(fCurr)
    pyplot.yticks(ticks = lfYTicks)
    pyplot.tick_params(axis="y", grid_linestyle = ":", grid_color="gray", grid_alpha = 0.5)
    pyplot.grid(True, axis="y")
    #pyplot.rc("grid", linestyle="--", color="gray")
  sActualAmplicon = sAmplicon
  if sAmpliconTitle and sAmpliconTitle != "":
    sActualAmplicon = sAmpliconTitle
  sActualChannel = sChannel
  if sChannelTitle and sChannelTitle != "":
    sActualChannel = sChannelTitle
  fXTextPos = 2
  fTextHeight = fPlotTextHeight * fYDelta
  fTextVStep = fPlotTextVStep * fYDelta
  fYTextPos = fYMax - fYDelta * 0.07 - fTextHeight
  pyplot.text(fXTextPos, fYTextPos, sActualAmplicon, fontfamily="sans-serif", fontsize="small", fontstretch="normal", fontweight="bold")
  fYTextPos = fYTextPos - fTextHeight - fTextVStep * 2
  pyplot.text(fXTextPos, fYTextPos, "Canal: " + sActualChannel, fontfamily="sans-serif", fontsize="x-small", fontstretch="normal", fontweight="normal")
  fYTextPos = fYTextPos - fTextHeight - fTextVStep
  pyplot.text(fXTextPos, fYTextPos, "Poço: " + sWell, fontfamily="sans-serif", fontsize="x-small", fontstretch="normal", fontweight="normal")
  fYTextPos = fYTextPos - fTextHeight - fTextVStep
  pyplot.text(fXTextPos, fYTextPos, "Limiar: " + str(round(fThreshold, 1)), fontfamily="sans-serif", fontsize="x-small", fontstretch="normal", fontweight="normal")
  fYTextPos = fYTextPos - fTextHeight - fTextVStep * 2
  if fCt != None:
    sCt = str(round(fCt, 1))
  else:
    sCt = "NA"
  if not bThreshold:
    sCt = "Inspecionar"
  pyplot.text(fXTextPos, fYTextPos, "Ct: " + sCt, fontfamily="sans-serif", fontsize="small", fontstretch="normal", fontweight="bold")
  if sFigureFile:
    pyplot.savefig(sFigureFile)
  return

def StartReport():
  global oReport, dConfigReport, dDateTime
  oReport = fpdf.FPDF(
    orientation = dConfigReport['page_orientation'],
    unit = dConfigReport['page_unit'],
    format = dConfigReport['page_format'])
  oReport.set_margins(
    dConfigReport['page_margin_left'],
    dConfigReport['page_margin_top'],
    dConfigReport['page_margin_right'])
  dDateTime = datetime.datetime.now()
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

def DrawCell(sText, sStyle, fWidth, fHeight, sAlign, sBorder, bFill = False, uAdvance = 0, fPosX = 0.0):
  global oReport, dConfigReport
  if not fHeight:
    fActualHeight = dConfigReport["font_height_" + sStyle.lower()]
  else:
    fActualHeight = fHeight
  MovePos(dConfigReport["page_section_advance"] * float(uAdvance) + fPosX, 0.0)
  SetFont(sStyle)
  oReport.cell(fWidth, h = fActualHeight, txt = sText if sText else "", border = sBorder.upper(), ln = 1, align = sAlign, fill = bFill)
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

# Main routine

print("")
print("A carregar configuracoes...")

try:
  LoadConfigOptions()
  LoadConfigReport()
  LoadProtocols()
except:
  sys.exit()

if dConfigOptions['experiments_folder'] == None:
  dConfigOptions['experiments_folder'] = ""
if dConfigOptions['worklist_folder'] == None:
  dConfigOptions['worklist_folder'] = ""

lsFolders = glob.glob(dConfigOptions['experiments_folder']+"\\*")
lsExperiments = list()
for sFolder in lsFolders:
  lsExperiments.append(sFolder.replace(dConfigOptions['experiments_folder']+"\\", "").lower())

uArgLen = len(sys.argv) - 1
if uArgLen > 0:
  sExperimentFile = sys.argv[1]
else:
  sExperimentFile = None

bCancel = False

if uArgLen == 0:
  bProcess = True
  while bProcess:
    print("")
    sExperimentFile = input("Indique o ficheiro da corrida (\"C\" para cancelar): \n  ").strip().lower()
    if sExperimentFile == "c":
      bProcess = False
      bCancel = True
    elif sExperimentFile == "":
      continue
    else:
      if sExperimentFile not in lsExperiments:
        print("")
        print("Erro: Ficheiro nao existente")
        continue
      else:
        bProcess = False
else:
  if sExperimentFile not in lsExperiments:
    print("")
    print("Erro: Ficheiro nao existente")
    bCancel = True

if bCancel:
  sys.exit()

print("")
print("A carregar dados...")

try:
  dSamples = GetTable(dConfigOptions['experiments_folder'] + "\\{0}\\{0}_samples.csv".format(sExperimentFile))
  dPCRInfo = GetTable(dConfigOptions['experiments_folder'] + "\\{0}\\{0}_pcr.csv".format(sExperimentFile))
  sProtocol = dPCRInfo[0]['protocol']
  dProtocol = dProtocols[sProtocol]
  if "baseline_min" in dProtocol and dProtocol['baseline_min'] > 0:
    fBaselineMin = dProtocol['baseline_min']
  if "baseline_max" in dProtocol and dProtocol['baseline_max'] > 0:
    fBaselineMax = dProtocol['baseline_max']
  dPlate = GetDet1XLSX(dConfigOptions['experiments_folder'] + "\\{0}\\{0}.xlsx".format(sExperimentFile))
  dRaw = GetDet1Raw(dConfigOptions['experiments_folder'] + "\\{0}\\{0}.txt".format(sExperimentFile))
  dThreshold = dict()
  for sSample in dPlate:
    for sAmplicon in dPlate[sSample]:
      dPlate[sSample][sAmplicon]['data'] = dRaw[dPlate[sSample][sAmplicon]['well']][dPlate[sSample][sAmplicon]['channel']]
      sChannelTitle = dPlate[sSample][sAmplicon]['channel_title']
      if sChannelTitle not in dThreshold:
        dThreshold[sChannelTitle] = dict()
      if sAmplicon not in dThreshold[sChannelTitle]:
        dThreshold[sChannelTitle][sAmplicon] = dict()
        dThreshold[sChannelTitle][sAmplicon]['data'] = list()
      dThreshold[sChannelTitle][sAmplicon]['data'].extend(dPlate[sSample][sAmplicon]['data'][fBaselineMin-1:fBaselineMax])
      dThreshold[sChannelTitle][sAmplicon]['value'] = Threshold_2(dThreshold[sChannelTitle][sAmplicon]['data'])
  dYAxis = dict()
  for sSample, dSample in dPlate.items():
    for sAmplicon, dAmplicon in dSample.items():
      if sAmplicon not in dYAxis:
        dYAxis[sAmplicon] = dict()
      if dAmplicon['channel_title'] not in dYAxis[sAmplicon]:
        dYAxis[sAmplicon][dAmplicon['channel_title']] = dict()
        dYAxis[sAmplicon][dAmplicon['channel_title']]['data'] = list()
      dYAxis[sAmplicon][dAmplicon['channel_title']]['data'].extend(dAmplicon['data'])
      dYAxis[sAmplicon][dAmplicon['channel_title']]['y_axis'] = (Min(dYAxis[sAmplicon][dAmplicon['channel_title']]['data']), Max(dYAxis[sAmplicon][dAmplicon['channel_title']]['data']))
except Exception as dError:
  print("")
  print("Erro: Nao foi possivel processar os dados. Verifique os ficheiros da corrida.")
  print("Mensagem de erro:\n" + str(dError))
  print("")
  sys.exit()

print("")

# Report generation

StartReport()

dExtras = dict()
for dSample in dSamples:
  if "extra_reactions" in dProtocol:
    luIDs = list(dProtocol['extra_reactions'].keys())
    luIDs.sort()
    for uID in luIDs:
      sExtraType = dProtocol['extra_reactions'][uID]['type']
      sExtraTitle = dProtocol['extra_reactions'][uID]['title']
      uSampleFactor = dProtocol['extra_reactions'][uID]['sample_factor']
      uSampleCount = len(dSamples) + 1
      if sExtraType not in dExtras:
        dExtras[sExtraType] = list()
      uExtraCount = len(dExtras[sExtraType])
      if uSampleFactor == 0:
        if uExtraCount == 0:
          dExtras[sExtraType].append(sExtraTitle)
      elif float(uSampleCount) / float(uSampleFactor) > float(uExtraCount):
        dExtras[sExtraType].append(sExtraTitle + "_" + str(uExtraCount + 1))

uImageIndex = 1

for sExtraType, lsExtraIDs in dExtras.items():
  uProcessIndex = 0
  uProcessCount = len(lsExtraIDs)
  for sExtraID in lsExtraIDs:
    uProcessIndex = uProcessIndex + 1
    print("")
    print("A gerar relatorio para: {} ({} de {} do tipo {})...".format(sExtraID, str(uProcessIndex), str(uProcessCount), dProtocol['types'][sExtraType]['title']))
    AddPage()
    SetPos(0, 0)
    WriteLine("#INSTITUTION# | #DEPARTMENT#", "header_1")
    WriteLine("Resultados de protocolo qPCR / RT-qPCR", "title_1")
    InfoLine("Protocolo", dProtocol['title'], 1, 0)
    InfoLine("Data e hora", dDateTime.strftime(sDateTimeFormat), 1, 0)
    InfoLine("Ficheiro", sExperimentFile, 1, 0)
    DrawRule(dConfigReport['page_body_width'], None)
    MovePos(0.0, dConfigReport['page_header_sep'])
    fYPos = GetPos()[1]
    fXDisp = 0
    fXPos = dConfigReport['report_info_x_correction']
    SetPos(fXPos + fXDisp, fYPos)
    DrawCell("ID: ", "small_1", dConfigReport['report_info_caption_width'], dConfigReport['report_info_height'], "R", "")
    fXDisp = fXDisp + dConfigReport['report_info_caption_width']
    SetPos(fXPos + fXDisp, fYPos)
    DrawCell(sExtraID, "large_1_bold", dConfigReport['report_info_text_2_width'], dConfigReport['report_info_height'], "L", "")
    fXDisp = fXDisp + dConfigReport['report_info_text_2_width']
    SetPos(fXPos + fXDisp, fYPos)
    DrawCell("Tipo: ", "small_1", dConfigReport['report_info_caption_width'], dConfigReport['report_info_height'], "R", "")
    fXDisp = fXDisp + dConfigReport['report_info_caption_width']
    SetPos(fXPos + fXDisp, fYPos)
    DrawCell(dProtocol['types'][sExtraType]['title'], "large_1", dConfigReport['report_info_text_2_width'], dConfigReport['report_info_height'], "L", "")
    fYPos = fYPos + dConfigReport['report_info_height'];
    fXDisp = 0
    SetPos(fXPos + fXDisp, fYPos)
    SetPos(0, GetPos()[1] + dConfigReport['pcr_info_plot_sep'])
    luAmplicons = list(dProtocol['types_amplicons'][sExtraType].keys())
    luAmplicons.sort()
    uColumn = 0
    for uAmplicon in luAmplicons:
      sAmpliconRef = dProtocol['amplicons'][dProtocol['types_amplicons'][sExtraType][uAmplicon]]['ref']
      sAmpliconTitle = dProtocol['amplicons'][dProtocol['types_amplicons'][sExtraType][uAmplicon]]['title']
      sFigureFile = "{}\\pcr_{}.png".format(sTemporaryFolder, str(uImageIndex))
      PlotPCR(dPlate, None, sExtraID, sAmpliconRef, sFigureFile, sAmpliconTitle, dYAxis = dYAxis)
      Image(sFigureFile, None, None, 0, dConfigReport['pcr_plot_height'])
      if uColumn == 0:
        MovePos(dConfigReport['pcr_plot_h_disp'], 0)
        uColumn = 1
      else:
        SetPos(0, GetPos()[1] + dConfigReport['pcr_plot_height'] + dConfigReport['pcr_plot_v_sep'])
        uColumn = 0
      uImageIndex = uImageIndex + 1

uProcessIndex = 0
uProcessCount = len(dSamples)

for dSample in dSamples:
  uProcessIndex = uProcessIndex + 1
  print("")
  print("A gerar relatorio para: {} ({} de {} amostras)...".format(dSample['sample'], str(uProcessIndex), str(uProcessCount)))
  AddPage()
  SetPos(0, 0)
  WriteLine("#INSTITUTION# | #DEPARTMENT#", "header_1")
  WriteLine("Resultados de protocolo qPCR / RT-qPCR", "title_1")
  InfoLine("Protocolo", dProtocol['title'], 1, 0)
  InfoLine("Data e hora", dDateTime.strftime(sDateTimeFormat), 1, 0)
  InfoLine("Ficheiro", sExperimentFile, 1, 0)
  DrawRule(dConfigReport['page_body_width'], None)
  MovePos(0.0, dConfigReport['page_header_sep'])
  fYPos = GetPos()[1]
  fXDisp = 0
  fXPos = dConfigReport['report_info_x_correction']
  SetPos(fXPos + fXDisp, fYPos)
  DrawCell("Amostra: ", "small_1", dConfigReport['report_info_caption_width'], dConfigReport['report_info_height'], "R", "")
  fXDisp = fXDisp + dConfigReport['report_info_caption_width']
  SetPos(fXPos + fXDisp, fYPos)
  DrawCell(dSample['sample'], "large_1_bold", dConfigReport['report_info_text_2_width'], dConfigReport['report_info_height'], "L", "")
  fXDisp = fXDisp + dConfigReport['report_info_text_2_width']
  SetPos(fXPos + fXDisp, fYPos)
  DrawCell("Data de colheita: ", "small_1", dConfigReport['report_info_caption_width'], dConfigReport['report_info_height'], "R", "")
  fXDisp = fXDisp + dConfigReport['report_info_caption_width']
  SetPos(fXPos + fXDisp, fYPos)
  DrawCell(dSample['sample_date'], "large_1", dConfigReport['report_info_text_2_width'], dConfigReport['report_info_height'], "L", "")
  fYPos = fYPos + dConfigReport['report_info_height'];
  fXDisp = 0
  SetPos(fXPos + fXDisp, fYPos)
  DrawCell("Nome: ", "small_1", dConfigReport['report_info_caption_width'], dConfigReport['report_info_height'], "R", "")
  fXDisp = fXDisp + dConfigReport['report_info_caption_width']
  SetPos(fXPos + fXDisp, fYPos)
  DrawCell(dSample['name'], "large_1", dConfigReport['report_info_text_width'], dConfigReport['report_info_height'], "L", "")
  fYPos = fYPos + dConfigReport['report_info_height'];
  fXDisp = 0
  SetPos(fXPos + fXDisp, fYPos)
  DrawCell("Data de nascimento: ", "small_1", dConfigReport['report_info_caption_width'], dConfigReport['report_info_height'], "R", "")
  fXDisp = fXDisp + dConfigReport['report_info_caption_width']
  SetPos(fXPos + fXDisp, fYPos)
  DrawCell(dSample['birthday'], "large_1", dConfigReport['report_info_text_width'], dConfigReport['report_info_height'], "L", "")
  fYPos = fYPos + dConfigReport['report_info_height'];
  fXDisp = 0
  SetPos(fXPos + fXDisp, fYPos)
  DrawCell("Serviço: ", "small_1", dConfigReport['report_info_caption_width'], dConfigReport['report_info_height'], "R", "")
  fXDisp = fXDisp + dConfigReport['report_info_caption_width']
  SetPos(fXPos + fXDisp, fYPos)
  DrawCell(dSample['department'], "large_1", dConfigReport['report_info_text_width'], dConfigReport['report_info_height'], "L", "")
  fYPos = fYPos + dConfigReport['report_info_height'];
  fXDisp = 0
  SetPos(fXPos + fXDisp, fYPos)
  SetPos(0, GetPos()[1] + dConfigReport['pcr_info_plot_sep'])
  luAmplicons = list(dProtocol['types_amplicons']['unknown'].keys())
  luAmplicons.sort()
  uColumn = 0
  for uAmplicon in luAmplicons:
    sAmpliconRef = dProtocol['amplicons'][dProtocol['types_amplicons']['unknown'][uAmplicon]]['ref']
    sAmpliconTitle = dProtocol['amplicons'][dProtocol['types_amplicons']['unknown'][uAmplicon]]['title'] 
    sFigureFile = "{}\\pcr_{}.png".format(sTemporaryFolder, str(uImageIndex))
    PlotPCR(dPlate, None, dSample['sample'], sAmpliconRef, sFigureFile, sAmpliconTitle, dYAxis = dYAxis)
    Image(sFigureFile, None, None, 0, dConfigReport['pcr_plot_height'])
    if uColumn == 0:
      MovePos(dConfigReport['pcr_plot_h_disp'], 0)
      uColumn = 1
    else:
      SetPos(0, GetPos()[1] + dConfigReport['pcr_plot_height'] + dConfigReport['pcr_plot_v_sep'])
      uColumn = 0
    uImageIndex = uImageIndex + 1

oReport.output(dConfigOptions['experiments_folder'] + "\\{0}\\{0}_report.pdf".format(sExperimentFile))
os.startfile(dConfigOptions['experiments_folder'] + "\\{0}\\{0}_report.pdf".format(sExperimentFile))
