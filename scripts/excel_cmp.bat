@ECHO OFF
SETLOCAL
set dirname=%~dp0
java -ea -cp "%dirname%\dist\*;" com.ka.spreadsheet.diff.SpreadSheetDiffer %*
ENDLOCAL