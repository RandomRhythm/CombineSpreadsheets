'Combine columns of speadsheets based on matching column.
'Requires Microsoft Excel. If blank entries exist in matching columns then sort by that column so empty entries are last.
'may take a while for large spreadsheets. Haven't looked into optimization but fastest option would be to not use Excel.

'Copyright (c) 2018 Ryan Boyle randomrhythm@rhythmengineering.com.

'This program is free software: you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation, either version 3 of the License, or
'(at your option) any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with this program.  If not, see <http://www.gnu.org/licenses/>.

Const forwriting = 241
Const ForAppending = 8
Const ForReading = 1
Dim intTabCounter
Dim strMainSSKey
Dim intLastRowEntry
Dim intLastImportRow
Dim boolCaseSensitive
Dim boolShowSecondExcel

'config section
strMainSSKey = ""
strMainMatchKey = ""
boolCaseSensitive = False
boolShowSecondExcel = True 'Set to true to show second Excel window (be sure not to close it out until script has finished)
'end config section

Set WshShell = CreateObject("WScript.Shell")
Set objEnv = WshShell.Environment("Process")



if objenv("PROMPT") <> "" and WScript.Arguments.count < 1 then
  wscript.echo "No parameters passed. Pass quoted column names and file paths:" & vbcrlf & "CombineSpreadsheets.vbs " & chr(34) & "Spreadsheet Path 1" & chr(34) & " " & Chr(34) & "Column Name 1" & Chr(34) & " " & chr(34) & "Spreadsheet Path 2" & chr(34) & " " & Chr(34) & "Column Name 2" & Chr(34) & vbcrlf & vbcrlf & "Example:" & vbcrlf & "CombineSpreadsheets.vbs " & chr(34) & "c:\spreadsheets\ss1.xlsx" & chr(34) & " " & Chr(34) & "MD5" & Chr(34) & " " & chr(34) & "c:\spreadsheets\ss2.xlsx" & chr(34) & " " & Chr(34) & "MD5 Hash" & Chr(34)
  wscript.quit
end if
Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
if WScript.Arguments.Count = 0 then

elseif WScript.Arguments(0) <> "" then
  for each argument in WScript.Arguments
	if OpenFilePath1 = "" and objFSO.fileexists(argument) then 
		OpenFilePath1 = argument
	elseif strMainSSKey = "" then
		strMainSSKey = argument
	elseif OpenFilePath2 = "" and objFSO.fileexists(argument) then 
		OpenFilePath2 = argument
	elseif strMainMatchKey = "" and OpenFilePath2 = "" then
		wscript.echo "Problem accessing file path:" & strMainSSKey & ". Confirm the file exists. Script will now exit"
		wscript.quit(3)
	elseif strMainMatchKey = "" then
		strMainMatchKey = argument
	end if
  next
end if
'msgbox OpenFilePath1  & "|" & strMainSSKey & "|" & OpenFilePath2 & "|" & strMainMatchKey
if OpenFilePath1 = "" and strMainSSKey = "" and OpenFilePath2 = "" and strMainMatchKey = "" then
	wscript.echo "No parameters provided/processed. The script will now prompt for the required parameters"
elseif OpenFilePath2 = "" then
	wscript.echo "Secondary spreadsheet path not identified. Script will now exit"
	wscript.quit(4)
elseif strMainMatchKey = "" and strMainSSKey <> "" then
	strMainMatchKey = strMainSSKey
end if

'set inital values

intTabCounter = 1
intWriteRowCounter = 1



CurrentDirectory = GetFilePath(wscript.ScriptFullName)



if OpenFilePath1 = "" or objFSO.fileexists(OpenFilePath1) = False then
	if OpenFilePath1 <> "" and objFSO.fileexists(OpenFilePath1) = False then
		wscript.echo "file does not exist:" & OpenFilePath1 & ". Please open the main spreadsheet"
	else
		wscript.echo "Please open the main spreadsheet"
	end if
	OpenFilePath1 = SelectFile( )
end if




Dim DictKeyLocation: Set DictKeyLocation = CreateObject("Scripting.Dictionary")'

if objFSO.fileexists(OpenFilePath1) = False then
	wscript.echo "file does not exist:" & OpenFilePath1 & ". The script will now exit."
	wscript.quit(5)
end if
Set objExcel = CreateObject("Excel.Application")
OpenFilePath1 = OpenFilePath1
Set objWorkbook = objExcel.Workbooks.Open _
    (OpenFilePath1)
    objExcel.Visible = True

if strMainSSKey = "" then
	strMainSSKey = inputbox("Please type the exact text for the column header name you want to use for combining the spreadsheets")
end if
	
mycolumncounter = 1
int_Main_Location = -1
Do Until objExcel.Cells(1,mycolumncounter).Value = ""
  if cStr(objExcel.Cells(1,mycolumncounter).Value) = cStr(strMainSSKey) then int_Main_Location = mycolumncounter 'Match key
  mycolumncounter = mycolumncounter +1
loop
intLastRowEntry = mycolumncounter
if int_Main_Location = -1 then 
  wscript.echo "Problem parsing header. Make sure the column header text matches. The provided text was " & CHr(34) & strMainSSKey & Chr(34) & ". Script will now exit"
  wscript.quit
end if

if OpenFilePath2 = "" then
	wscript.echo "Please open the import spreadsheet"
	OpenFilePath2 = SelectFile( )
end if
if objFSO.fileexists(OpenFilePath2) = False then
	wscript.echo "file does not exist:" & OpenFilePath2 & ". The script will now exit."
	wscript.quit(6)
end if

Set objExcel2 = CreateObject("Excel.Application")
Set objWorkbook2 = objExcel2.Workbooks.Open _
    (OpenFilePath2)
    objExcel2.Visible = boolShowSecondExcel 'Hide this to not confuse or allow accidental closure

if strMainMatchKey = "" then
	strMainMatchKey = inputbox("Please type the exact text for the column header name you want to use for combining the recently selected spreadsheet")
end if
secondcolumncounter = 1
int_MainMatch_Location = -1
Do Until objExcel2.Cells(1,secondcolumncounter).Value = ""
  if objExcel2.Cells(1,secondcolumncounter).Value = strMainMatchKey then int_MainMatch_Location = secondcolumncounter 'Match key
  secondcolumncounter = secondcolumncounter +1
loop
intLastImportRow = secondcolumncounter -1
if int_MainMatch_Location = -1 then 
  wscript.echo "Problem parsing match header. Make sure the column header text matches. The provided column name was " & Chr(34) & strMainMatchKey & Chr(34) & ". Script will now exit"
  objWorkbook2.Close
  objExcel2.Quit
  wscript.quit(10)
end if

'mark where match keys are located
intRowCounter = 2
Do Until objExcel2.Cells(intRowCounter,int_MainMatch_Location).Value = "" 'loop till you hit null value (end of rows)

	strTmpMatchKey = objExcel2.Cells(intRowCounter,int_MainMatch_Location).Value
	if boolCaseSensitive = False then strTmpMatchKey = lcase(strTmpMatchKey)
	if DictKeyLocation.exists(strTmpMatchKey) = false then DictKeyLocation.add strTmpMatchKey, intRowCounter
	intRowCounter = intRowCounter + 1
loop

'add additional items to header row
inttmpLastRowEntry = intLastRowEntry
'msgbox DictKeyLocation.item(strTmpMatchKey)
for secondcolumncounter =1 to intLastImportRow
  strSStempValue = objExcel2.Cells(1,secondcolumncounter).Value
   objExcel.Cells(1,inttmpLastRowEntry).Value = strSStempValue
  inttmpLastRowEntry = inttmpLastRowEntry +1
next

intHitCount = 0
intMissCount = 0
intRowCounter = 2
Do Until objExcel.Cells(intRowCounter,int_Main_Location).Value = "" 'loop till you hit null value (end of rows)

  strTmpMatchKey = objExcel.Cells(intRowCounter,int_Main_Location).Value
  if boolCaseSensitive = False then strTmpMatchKey = lcase(strTmpMatchKey)
  if DictKeyLocation.exists(strTmpMatchKey) = true then
	intHitCount = intHitCount +1
    inttmpLastRowEntry = intLastRowEntry
    'msgbox DictKeyLocation.item(strTmpMatchKey)
    for secondcolumncounter =1 to intLastImportRow
      strSStempValue = objExcel2.Cells(DictKeyLocation.item(strTmpMatchKey),secondcolumncounter).Value
       objExcel.Cells(intRowCounter,inttmpLastRowEntry).Value = strSStempValue
      inttmpLastRowEntry = inttmpLastRowEntry +1
    next
  else
	intMissCount = intMissCount + 1
	logdata CurrentDirectory & "\missed.log", date & " " & time & " unmatched entry:" & chr(34) & strTmpMatchKey & Chr(34), False
  end if
  intRowCounter = intRowCounter +1
loop
objWorkbook2.Close
objExcel2.Quit
wscript.echo "finished combining spreadsheets. " & DictKeyLocation.count & " entries were read. " & intHitCount & " entries were combined. " & intMissCount & " entries could not be matched."

function LogData(TextFileName, TextToWrite,EchoOn)
Dim strTmpFilName1
Dim strTmpFilName2


Set fsoLogData = CreateObject("Scripting.FileSystemObject")
if EchoOn = True then wscript.echo TextToWrite
  If fsoLogData.fileexists(TextFileName) = False Then
      'Creates a replacement text file 
      fsoLogData.CreateTextFile TextFileName, True
  End If
on error resume next
Set WriteTextFile = fsoLogData.OpenTextFile(TextFileName,ForAppending, False)
if err.number <> 0 then
  msgbox "Error writting to " & TextFileName & " perhaps the file is locked?"
  err.number = 0
  Set WriteTextFile = fsoLogData.OpenTextFile(TextFileName,ForAppending, False)
  if err.number <> 0 then exit function
end if

on error goto 0
WriteTextFile.WriteLine TextToWrite
WriteTextFile.Close
Set fsoLogData = Nothing
End Function

Function GetFilePath (ByVal FilePathName)
found = False

Z = 1

Do While found = False and Z < Len((FilePathName))

 Z = Z + 1

         If InStr(Right((FilePathName), Z), "\") <> 0 And found = False Then
          mytempdata = Left(FilePathName, Len(FilePathName) - Z)
          
             GetFilePath = mytempdata

             found = True

        End If      

Loop

end Function

Function GetData(contents, ByVal EndOfStringChar, ByVal MatchString)
MatchStringLength = Len(MatchString)
x= 0

do while x < len(contents) - (MatchStringLength +1)

  x = x + 1
  if Mid(contents, x, MatchStringLength) = MatchString then
    'Gets server name for section
    for y = 1 to len(contents) -x
      if instr(Mid(contents, x + MatchStringLength, y),EndOfStringChar) = 0 then
          TempData = Mid(contents, x + MatchStringLength, y)
        else
          exit do  
      end if
    next
  end if
loop
GetData = TempData
end Function






Function SelectFile( )
    ' File Browser via HTA
    ' Author:   Rudi Degrande, modifications by Denis St-Pierre and Rob van der Woude
    ' Features: Works in Windows Vista and up (Should also work in XP).
    '           Fairly fast.
    '           All native code/controls (No 3rd party DLL/ XP DLL).
    ' Caveats:  Cannot define default starting folder.
    '           Uses last folder used with MSHTA.EXE stored in Binary in [HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32].
    '           Dialog title says "Choose file to upload".
    ' Source:   http://social.technet.microsoft.com/Forums/scriptcenter/en-US/a3b358e8-15&?lig;-4ba3-bca5-ec349df65ef6

    Dim objExec, strMSHTA, wshShell

    SelectFile = ""

    ' For use in HTAs as well as "plain" VBScript:
    strMSHTA = "mshta.exe ""about:" & "<" & "input type=file id=FILE>" _
             & "<" & "script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
             & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);" & "<" & "/script>"""
    ' For use in "plain" VBScript only:
    ' strMSHTA = "mshta.exe ""about:<input type=file id=FILE>" _
    '          & "<script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
    '          & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>"""

    Set wshShell = CreateObject( "WScript.Shell" )
    Set objExec = wshShell.Exec( strMSHTA )

    SelectFile = objExec.StdOut.ReadLine( )

    Set objExec = Nothing
    Set wshShell = Nothing
End Function



Function isIPaddress(strIPaddress)
DIm arrayTmpquad
Dim boolReturn_isIP
boolReturn_isIP = True
if instr(strIPaddress,".") then
  arrayTmpquad = split(strIPaddress,".")
  for each item in arrayTmpquad
    if isnumeric(item) = false then boolReturn_isIP = false
  next
else
  boolReturn_isIP = false
end if
isIPaddress = boolReturn_isIP
END FUNCTION

Function ReturnHostFromHeader(strtmpline)
if instr(strtmpline, "Host: ") then
  strtmpline = getdata(strtmpline, ":", "Host: ")
  if right(strtmpline,6) = "Accept" then strtmpline = left(strtmpline,len(strtmpline)-6)
  if right(strtmpline,10) = "Connection" then strtmpline = left(strtmpline,len(strtmpline)-10)
  ReturnHostFromHeader = strtmpline
end if
End Function

  


Function RemoveDups(strRMdupsData, strSplitChar)
Dim ArrayRemoveDups
Dim strReturnRemoveDups
if instr(strRMdupsData, strSplitChar) then
  ArrayRemoveDups = split(strRMdupsData, strSplitChar)
  Dim dicTmpRemoveDuplicate: Set dicTmpRemoveDuplicate = CreateObject("Scripting.Dictionary")
  for xRP = 0 to ubound(ArrayRemoveDups)
    if not dicTmpRemoveDuplicate.Exists(ArrayRemoveDups(xRP)) then _
    dicTmpRemoveDuplicate.Add ArrayRemoveDups(xRP), dicTmpRemoveDuplicate.Count
  next
  For Each Item In dicTmpRemoveDuplicate
    if strReturnRemoveDups =  "" Then
      strReturnRemoveDups = Item
    else
      strReturnRemoveDups = strReturnRemoveDups & strSplitChar & Item
    end if
  next
    RemoveDups = strReturnRemoveDups 

else
  RemoveDups = strRMdupsData
end if



End Function  


Function IsPrivateIP(strIP)
Dim boolReturnIsPrivIp
Dim ArrayOctet
boolReturnIsPrivIp = False
if isIPaddress(strIP) = False then
  IsPrivateIP = False
  exit function
end if
if left(strIP,3) = "10." then
  boolReturnIsPrivIp = True
elseif left(strIP,4) = "172." then
  ArrayOctet = split(strIP,".")
  if ArrayOctet(1) >15 and ArrayOctet(1) < 32 then
    boolReturnIsPrivIp = True
  end if
elseif left(strIP,7) = "192.168" then
  boolReturnIsPrivIp = True
end if
IsPrivateIP = boolReturnIsPrivIp
End Function




Sub Write_Spreadsheet_line(strSSrow)
Dim intColumnCounter
if instr(strSSrow,"|") then
  strSSrow = split(strSSrow, "|")
  for intColumnCounter = 1 to ubound(strSSrow) + 1
    objExcel.Cells(intWriteRowCounter, intColumnCounter).Value = strSSrow(intColumnCounter -1)
  next
else
    objExcel.Cells(intWriteRowCounter, 1).Value = strSSrow
end if
intWriteRowCounter = intWriteRowCounter + 1
end sub

Sub Add_Workbook_Worksheet(strWorksheetName)
Set objWorkbook = objExcel.Worksheets(objExcel.Worksheets.count)
objWorkbook.Activate

objExcel.ActiveWorkbook.Worksheets.Add
intWriteRowCounter = 1
Set objSheet1 = objExcel.Worksheets(objExcel.Worksheets(objExcel.Worksheets.count -1).name)
    Set objSheet2 = objExcel.Worksheets(objExcel.Worksheets(objExcel.Worksheets.count).name)
    objSheet2.Move objSheet1


objExcel.Worksheets(objExcel.Worksheets.count).Name = strWorksheetName
Set objWorkbook = objExcel.Worksheets(objExcel.Worksheets.count)
objWorkbook.Activate


end sub

Sub Move_next_Workbook_Worksheet(strWorksheetName)
intTabCounter = intTabCounter + 1
if objExcel.Worksheets.count < intTabCounter then
  Add_Workbook_Worksheet(strWorksheetName)
else
  Set objWorkbook = objExcel.Worksheets(intTabCounter)
  objWorkbook.Activate
  if strWorksheetName <> "" then objExcel.Worksheets(intTabCounter).Name = strWorksheetName
  intWriteRowCounter = 1
end if
end sub


Function ReturnPairedListfromDict(tmpDictionary)
Dim strTmpDictList
For Each Item In tmpDictionary
  if strTmpDictList = "" then 
  
    strTmpDictList = tmpDictionary.Item(Item) & " - " & Item
  else
    strTmpDictList = strTmpDictList & ", " & tmpDictionary.Item(Item) & " - " & Item
  end if

next

ReturnPairedListfromDict = strTmpDictList
End Function

Function ReturnListfromDict(tmpDictionary)
Dim strTmpDictList
For Each Item In tmpDictionary
  if strTmpDictList = "" then 
  
    strTmpDictList =  Item
  else
    strTmpDictList = strTmpDictList & vbCrLf & Item
  end if

next

ReturnListfromDict = strTmpDictList
End Function


