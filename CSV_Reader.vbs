'CSV Reader for VBScript 

'Copyright (c) 2019 Ryan Boyle randomrhythm@rhythmengineering.com.

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

'constants for handling file access
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateTrue = -1
Const TristateFalse = 0

Dim DictHeader: Set DictHeader = CreateObject("Scripting.Dictionary") 'Maping between header text and integer column location. Header text is populated from a file. 
Dim strUniqueColumn 'Used to store the text label for the unique column
Dim intUniqueLocation 'Unique column numeric location
Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim dictUniqueLoc: Set dictUniqueLoc = CreateObject("Scripting.Dictionary")'unique value location (example: MD5 hash)
Dim ArraySpreadSheetData()'used to hold the csv file in memory
dim tmpArrayPointer() 'temporary location pointer 
Dim boolCaseSensitive 'force case sensitive matching
Dim strDelimiter 'This is the delimiter character 

'Configurations for CSV parsing
boolCaseSensitive = True 'Setting to false forces everything to lowercase
strDelimiter =  "," 'delimiter character. Use VbTab for tab separated or "," for comma separated
strUniqueColumn = "Name" 'exact text match of column header that contains the unique data to track
intHrowAbort = 6 'Number of rows in to abort if header has not been identified.

'set path string
CurrentDirectory = GetFilePath(wscript.ScriptFullName)


'Example of returning a CSV cell value - ReturnSpreadSheetItem(CSV string, column number)
msgbox ReturnSpreadSheetItem("black, gray, white, red, blue, green, yellow, orange, brown, purple, pink", 2) 'column location starts at zero

'example using mixed quotes csv to return cell value - Water,"Hydrogen hydroxide (HH or HOH), hydrogen oxide, dihydrogen monoxide (DHMO)", Liquid
msgbox ReturnSpreadSheetItem("Water," & chr(34) & "Hydrogen hydroxide (HH or HOH), hydrogen oxide, dihydrogen monoxide (DHMO)" & chr(34) & ",Liquid", 1)

'example row quoted multi-line span
msgbox ReturnSpreadSheetItem("Water," & chr(34) & "Hydrogen hydroxide (HH or HOH), " & vbCrLf & "hydrogen oxide, " & vbCrLf & "dihydrogen monoxide (DHMO)" & chr(34) & ",Liquid", 1)

'example reading in CSV to array for cell value access
strCsvPath = CurrentDirectory & "\titanic.csv" 'Path to CSV file
LoadKeyValue CurrentDirectory & "\header.txt" , DictHeader 'load header value list into dict (these are the column names you want to work with)
redim preserve ArraySpreadSheetData(1)'array for loading CSV into memory

loadSpreadSheetData strCsvPath, False 'load CSV into array. Each row is loaded into ArraySpreadSheetData

'Example using header to column location dictionary DictHeader. Pulls Name value from row number 2
MsgBox ReturnSpreadSheetItem(ArraySpreadSheetData(2), DictHeader.item("Name"))

'Retrieve CSV row from unique value - "Mrs. James Joseph (Margaret Tobin) Brown"
tmpCSVline = ArraySpreadSheetData(dictUniqueLoc.Item("Mrs. James Joseph (Margaret Tobin) Brown"))
age = ReturnSpreadSheetItem(tmpCSVline, DictHeader.item("Age")) 'Get cell value from CSV row
MsgBox "Molly Brown's age was " & age


sub loadSpreadSheetData(strSpreadSheetFpath, boolUnicode) 'This is an example of loading a CSV into memory for use
boolHeaderMatch = False
If strSpreadSheetFpath = "" then
  wscript.echo "Please open the csv"
  OpenFilePath1 = SelectFile( )
else
OpenFilePath1 = strSpreadSheetFpath
end if
if objFSO.fileexists(OpenFilePath1) then
  if boolUnicode = True then
    ANSIorUnicode = TristateTrue
  else
    ANSIorUnicode = TristateFalse
  end if
  Set objFile = objFSO.OpenTextFile(OpenFilePath1, ForReading, false, ANSIorUnicode)
  intCSVRowLocation = 0
  boolSuppressNoHash = False
  Do While Not objFile.AtEndOfStream
    if not objFile.AtEndOfStream then 'read file
        On Error Resume Next
        intCSVRowLocation = intCSVRowLocation + 1
        strLineData = objFile.ReadLine 
        on Error GoTo 0
        if BoolHeaderLocSet =True then
			if instrrev(strLineData, strDelimiter) + 1 = chr(34) then 'grab the rest of the row if it contains a quoted return character
			  do while right(strLineData,1) <> Chr(34) and Not objFile.AtEndOfStream
				if right(strLineData,1) <> Chr(34) then
				  strLineData = strLineData & objFile.ReadLine
				end if
			  loop
			end if
          
        end if

        ArraySpreadSheetData(intCSVRowLocation) = strLineData
        redim preserve ArraySpreadSheetData(intCSVRowLocation +1)
		if intCSVRowLocation > intHrowAbort and BoolHeaderLocSet = False and boolUnicode = False then 'failing to load data 
			redim ArraySpreadSheetData(1) 'clear array. X # of rows in and no header means something went wrong
			loadSpreadSheetData strSpreadSheetFpath, true 'try reading as unicode
			exit sub
		end if
		
		
	If boolHeaderMatch = False Then 	
		If boolCaseSensitive = False Then
			strHeadCompare = LCase(strLineData)
		Else	
			strHeadCompare = strLineData
		End If	
		
		for each headerText in DictHeader 'loop through header items 
      if instr(strHeadCompare, headerText) > 0 then 'match header text
        boolHeaderMatch = True 'header row found
        exit for
      end if
		Next
	End If	
		
      if BoolHeaderLocSet = False and boolHeaderMatch = True then
        if instr(strLineData, "Image Path") > 0 and instr(strLineData,	"MD5") > 0 and instr(strLineData, "Entry Location") > 0 then 'autoruns
          boolSuppressNoHash = True
        end if
          'header row

          SetHeaderLocations strLineData
          BoolHeaderLocSet = True
		  'msgbox "header location set"
        elseIf BoolHeaderLocSet = True then
          if instr(strLineData, strDelimiter) then
            strUniqueVal = ReturnSpreadSheetItem(strLineData, intUniqueLocation)
            if strUniqueVal <> "" then
              If boolCaseSensitive = False Then strUniqueVal = lcase(strUniqueVal) 'needs to be lower case for comparison
              if dictUniqueLoc.exists(strUniqueVal) = false then
                dictUniqueLoc.add strUniqueVal, intCSVRowLocation
                if boolSpreadSheetDebug = true then msgbox "unique -" & intUniqueLocation & "|" & intCSVRowLocation & "|" & strUniqueVal
              end if
            else
              if boolSuppressNoHash = False then Msgbox "Could not process line in SpreadSheet: " & strLineData
            end if
          else
            Msgbox "no commas-" & strLineData
          end if
        end if
    end if
  loop
  objFile.close

else'file does not exist

end if
end sub


Function returnCellLocation(strQuotedLine, cellNumber) 'needed to support mixed quoted non-quoted csv
dim StrReturnCellL
  strTmpHArray = split(strQuotedLine, strDelimiter)
  redim tmpArrayPointer(ubound(strTmpHArray))
  boolQuoted = False
  intArrayCount = 0
  for cellCount = 0 to ubound(strTmpHArray)
	if boolQuoted = False then 
		tmpArrayPointer(intArrayCount) = cellCount
		if cellNumber = intArrayCount then StrReturnCellL = cellCount
		intArrayCount = intArrayCount + 1 
	end if

	if instr(strTmpHArray(cellCount),chr(34)) > 0 then 
		if boolQuoted = False and left(strTmpHArray(cellCount), 1) = chr(34) and right(strTmpHArray(cellCount),1) = chr(34) then
			boolQuoted = False
		elseif boolQuoted = True and right(strTmpHArray(cellCount), 1) = chr(34) then 
			boolQuoted = False
		elseif boolQuoted = False and left(strTmpHArray(cellCount), 1) = chr(34) then
			boolQuoted = True
		else
			'ignore quotes that aren't at the begening or end 
		end if
	end if
  next
returnCellLocation = StrReturnCellL  
end Function


Function ReturnSpreadSheetItem(strCSVrow, intColumnLocation) 'pass this function the csv row and which column you want to get the value
Dim strSpreadSheetItem

intArrayPointer = returnCellLocation(strCSVrow, intColumnLocation)
if instr(strCSVrow, strDelimiter) > 0 Then
	strTmpHArray = split(strCSVrow, strDelimiter)
	if ubound(tmpArrayPointer) >= intColumnLocation and cint(intColumnLocation) > -1 then
		if ubound(tmpArrayPointer) = intArrayPointer then
			strSpreadSheetItem = replace(strTmpHArray(intArrayPointer), Chr(34), "")
		elseif (tmpArrayPointer(intColumnLocation) +1 <> tmpArrayPointer(intColumnLocation +1)) then
			strSpreadSheetItem = ""
			for itemCount = 0 to tmpArrayPointer(intColumnLocation +1) - (tmpArrayPointer(intColumnLocation) +1)
				strSpreadSheetItem = AppendValues(strSpreadSheetItem, replace(strTmpHArray(intArrayPointer + itemCount), Chr(34), ""), strDelimiter)
			next
		else
			strSpreadSheetItem = replace(strTmpHArray(intArrayPointer), Chr(34), "")
		end if
	
	else
		msgbox "SpreadSheet array mismatch:strCSVrow=" & strCSVrow & "&intArrayPointer=" & intArrayPointer  & "&ubound(tmpArrayPointer)=" & ubound(tmpArrayPointer)
		if cint(intArrayPointer) > -1 AND cint(intArrayPointer) <= ubound(strTmpHArray) then
			strSpreadSheetItem = replace(strTmpHArray(tmpArrayPointer(intArrayPointer)), Chr(34), "")
		end if
	end if

end if
ReturnSpreadSheetItem = strSpreadSheetItem
End Function


Sub SetHeaderLocations(StrHeaderText) 'sets the integer location for the header text
if instr(StrHeaderText, strDelimiter) or instr(StrHeaderText, vbtab) then
  if instr(StrHeaderText, strDelimiter) then 
    strTmpHArray = split(StrHeaderText, strDelimiter)
  else
    MsgBox "missing delimiter. Script will now exit"
    WScript.Quit (4)
  end if
  for inthArrayLoc = 0 to ubound(strTmpHArray)
    strCellData = ReturnSpreadSheetItem(StrHeaderText, inthArrayLoc)
  If boolCaseSensitive = False Then
  	strCellData = LCase(strCellData)
  End If
    If strUniqueColumn = strCellData Then
    	intUniqueLocation = inthArrayLoc
    End If  
    for each headerText in DictHeader 'loop through header items
      if strCellData = headerText then 'match header text
        DictHeader.item(headerText) = inthArrayLoc
        exit for
      end if
    next 
  next
else
  Msgbox "error parsing header"
end if
end sub

Sub LoadKeyValue(strListPath, dictToLoad)
if objFSO.fileexists(strListPath) then
  Set objFile = objFSO.OpenTextFile(strListPath)
  Do While Not objFile.AtEndOfStream
    if not objFile.AtEndOfStream then 'read file
        On Error Resume Next
        strData = objFile.ReadLine
        If boolCaseSensitive = False Then
          if dictToLoad.exists(lcase(strData)) = False then dictToLoad.add lcase(strData), ""
		Else
		  if dictToLoad.exists(strData) = False then dictToLoad.add strData, ""
		End if	
        on error goto 0
    end if
  loop
end if
end sub



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


Function AppendValues(strAggregate,strAppend, strSepChar)
    if strAggregate = "" then
      strAggregate = strAppend
    else
      strAggregate = strAggregate & strSepChar & strAppend
    end if
AppendValues = strAggregate
end Function