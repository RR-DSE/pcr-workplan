# extra_protocol_1.py (utf-8)
# 
# Edited by: RR-DSE
# Timestamp: 22-01-18 19:57:33

import io
import datetime
import shutil

def Workplan(dData = None):
  shutil.copyfile("{}\\report.cmd".format(dData['commands_folder']), "{}\\{}\\report.cmd".format(dData['experiments_folder'], dData['file']))

def GetPlateCode(dData = None):
  return None
