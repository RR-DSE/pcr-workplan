# extra_protocol_2.py (utf-8)
# 
# Edited by: RR-DSE
# Timestamp: 22-01-18 20:08:50

import io
import datetime

dTransType = {'nc': "Negative", 'pc': "Positive", 'unknown': "Unknown"}
dTransRow = {1: "A", 2: "B", 3: "C", 4: "D", 5: "E", 6: "F", 7: "G", 8: "H"}

def SearchPCRSlotWell(dPCRSlots, uRow, uColumn):
  tsRet = None
  for sType, dSlots in dPCRSlots.items():
    for uSlot in dSlots:
      sID = dSlots[uSlot]['id']
      for uWell in dSlots[uSlot]['wells']:
        if dSlots[uSlot]['wells'][uWell]['row'] == uRow and dSlots[uSlot]['wells'][uWell]['column'] == uColumn:
          tsRet = (sType, sID, uWell)
          break
  return tsRet

def Workplan(dData = None):
  lsChannels = list()
  for sType, dType in dData['protocol']['mixes'].items():
    for uWell, dWell in dType.items():
      if 'amplicons' in dWell:
        for sAmplicon, dAmplicon in dWell['amplicons'].items():
          if dAmplicon['channel_title'] not in lsChannels:
            lsChannels.append(dAmplicon['channel_title'])
  for sChannel in lsChannels:
    oString = io.StringIO(newline = "")
    oString.write("Row,Column,*Target Name,*Sample Name,*Biological Group\r\n")
    for uRow in range(1, 9):
      for uColumn in range(1, 13):
        sID = ""
        sTarget = ""
        bFound = False
        tsSearch = SearchPCRSlotWell(dData['pcr_slots'], uRow, uColumn)
        if tsSearch:
          if 'amplicons' in dData['protocol']['mixes'][tsSearch[0]][tsSearch[2]]:
            for sAmplicon, dAmplicon in dData['protocol']['mixes'][tsSearch[0]][tsSearch[2]]['amplicons'].items():
              if dAmplicon['channel_title'] == sChannel:
                bFound = True
                sID = tsSearch[1]
                sTarget = dData['protocol']['amplicons'][sAmplicon]['abbreviation']
                sString = "{},{},{},{},\r\n".format(dTransRow[uRow], str(uColumn), sTarget, sID)
                oString.write(sString)
    oString.seek(0)
    oFile = open("{}\\{}\\{}_det2_{}_samples.csv".format(dData['experiments_folder'], dData['file'], dData['file'], sChannel), "w", encoding = "utf-8", newline = "")
    oFile.write(oString.read())
    oFile.close()
    oString.close()

def GetPlateCode(dData = None):
  return None
