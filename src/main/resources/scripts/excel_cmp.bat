@ECHO OFF
SETLOCAL
set dirname=%~dp0
java -ea -Xmx512m -cp "%dirname%\dist\*;" com.ka.spreadsheet.diff.SpreadSheetDiffer %*
ENDLOCAL