@ECHO OFF
SET TOOLSDIR=\"pcr_workplan"
SET PATH=%PATH%;\Python
MODE 80,30
COLOR 70
TITLE Windows Python Environment
PROMPT Path: $P$_$G
SET CURRDIR=%~p0
SET CURRDIR=%CURRDIR:~0,-1%
FOR %%f IN (%CURRDIR%) DO SET CURRFOLDER=%%~nxf
CD %TOOLSDIR%
CLS
ECHO #INSTITUTION# ^| #DEPARTMENT# ^| PCR
ECHO Rotina para criacao de relatorio
ECHO --------------------------------------------------------------
ECHO.
ECHO A criar relatorio para experiencia %CURRFOLDER%...
PYTHON det1_report.py %CURRFOLDER%
CD %CURRDIR%
ECHO.
PAUSE
