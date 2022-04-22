@ECHO OFF
SETLOCAL
set dirname=%~dp0
java -ea -cp "%dirname%\lib\*;" com.ka.spreadsheet.diff.SpreadSheetDiffer %*
ENDLOCAL
