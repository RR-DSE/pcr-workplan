# extra_protocol_3.py (utf-8)
# 
# Edited by: RR-DSE
# Timestamp: 22-01-18 20:08:24

import io
import datetime

dTransType = {'nc': "Negative", 'pc': "Positive", 'unknown': "Unknown"}
dTransColor = {'nc': "r=0,g=255,b=204", 'pc': "r=255,g=204,b=0", 'unknown': "r=0,g=204,b=255"}

def SearchPCRSlotWell(dPCRSlots, uRow, uColumn):
  tsRet = None
  for sType, dSlots in dPCRSlots.items():
    for uSlot in dSlots:
      sID = dSlots[uSlot]['id']
      for uWell in dSlots[uSlot]['wells']:
        if dSlots[uSlot]['wells'][uWell]['row'] == uRow and dSlots[uSlot]['wells'][uWell]['column'] == uColumn:
          tsRet = (sType, sID)
          break
  return tsRet

def Workplan(dData = None):
  oString = io.StringIO(newline = "")
  luKeys = list(dData['protocol']['filter_options'].keys())
  luKeys.sort()
  oString.write("Version=1.0.0\r\n")
  oString.write("Filter;Sample ID;Type;Assay;Lot No;Extraction Method;Extraction Lot;Extraction Material;Color;Note\r\n")
  for uRow in range(1, 9):
    for uColumn in range(1, 13):
      sType = "Empty"
      sID = ""
      sAssay = ""
      sLot = ""
      sExpirationDate = ""
      sColor = "r=220,g=220,b=220"
      tsSearch = SearchPCRSlotWell(dData['pcr_slots'], uRow, uColumn)
      if tsSearch:
        sType = dTransType[tsSearch[0]]
        sID = tsSearch[1]
        sAssay = "AssayCode"
        sLot = "{}  ({})".format(dData['protocol']['lot'], dData['protocol']['lot_expiration_date'])
        sColor = dTransColor[tsSearch[0]]
      for uKey in luKeys:
        sString = "{};{};{};{};{};;;;java.awt.Color[{}];;\r\n".format(dData['protocol']['filter_options'][uKey]['data'], sID, sType, sAssay, sLot, sColor)
        oString.write(sString)
  oString.seek(0)
  oFile = open("{}\\{}\\{}_det2_samples.csv".format(dData['experiments_folder'], dData['file'], dData['file']), "w", encoding = "utf-8", newline = "")
  oFile.write(oString.read())
  oFile.close()
  oString.close()

def GetPlateCode(dData = None):
  oFile = open("{}\\det2_currplate.txt".format(dData['config_folder']), "r+", encoding = "utf-8", newline = "")
  uID = int(oFile.readline()) + 1
  sID = "DET2{}".format(str(uID).zfill(6))
  oFile.seek(0)
  oFile.write("{}\r\n".format(str(uID).zfill(6)))
  oFile.truncate()
  oFile.close()
  return sID
