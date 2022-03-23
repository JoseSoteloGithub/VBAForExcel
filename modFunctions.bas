Option Explicit

Function funAddColumnHeader(ByVal columnHeader As String, ByVal headerRow As Long)
    ' Adds column header to the right of all other columns
    Dim getPosition As Range

    Set getPosition = modFunctions.getRangeWhole(columnHeader, ActiveSheet) 'ActiveSheet.Cells.Find(what:=columnHeader, after:=Range("A1"), lookat:=xlWhole, SearchDirection:=xlNext)
    
    If getPosition Is Nothing Then
    
        ActiveSheet.Cells(headerRow, Columns.Count).End(xlToLeft).Offset(, 1) = columnHeader
        
        Set getPosition = modFunctions.getRangeWhole(columnHeader, ActiveSheet) 'Set getPosition = ActiveSheet.Cells.Find(what:=columnHeader, after:=Range("A1"), lookat:=xlWhole, SearchDirection:=xlNext)
        
        getPosition.Offset(, -1).Copy
        
        getPosition.PasteSpecial Paste:=xlPasteFormats
        
        Application.CutCopyMode = False
    
    End If

End Function

Function funLoopThroughRows(targetSheet As Worksheet, columnLetter As String, rowNumber As Long) As Range
    'Iterate through visible rows
    
    Dim thisRange As Range

    Set thisRange = targetSheet.Range(columnLetter & rowNumber)

    If thisRange.EntireRow.Hidden = False Then
    
        If Not IsError(thisRange) Then
    
'            If IsNumeric(thisRange) Then
'
                If thisRange = "" Then
                
                    Set funLoopThroughRows = thisRange
                
                End If
'
'            End If
            
        End If
        
    End If

End Function

Sub subLaunchChrome(Optional navigatePath As String)

    ' Launches Chrome and navigates to passed path
            
    If navigatePath = Empty Then
    
        navigatePath = ""
    
    End If

    Dim chromePath As String
    
    chromePath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    
    Shell (chromePath & " -url " & navigatePath)

End Sub

Function funGetWorksheetByRange(findString As String, targetWorkbook As Workbook) As Worksheet

    ' 
            
    Dim currentSheet As Worksheet

    For Each currentSheet In targetWorkbook.Worksheets
    
        Dim getPosition As Range
        
        Set getPosition = modFunctions.getRangeWhole(findString, currentSheet)
    
        If Not getPosition Is Nothing Then
        
            Set funGetWorksheetByRange = currentSheet
        
            Exit For
        
        End If
    
    Next currentSheet

End Function

Function funOpenNewestFile(strFolder As String, fileQualifyer As String, Optional relativeToWhatDate As Date)

    ' Open the newest file in a path
                
    If relativeToWhatDate = "12:00:00 AM" Then
    
        relativeToWhatDate = Date
    
    End If

    If InStr(strFolder, "/") > 0 Then

        If Right(strFolder, 1) <> "/" Then
        
            strFolder = strFolder & "/"
        
        End If

    ElseIf InStr(strFolder, "\") > 0 Then

        If Right(strFolder, 1) <> "\" Then
        
            strFolder = strFolder & "\"
        
        End If
        
    End If

    'Need to validate newest pricing file is open
    Dim strFileName As String
    
    Dim strFileSpec As String: strFileSpec = strFolder & "*.x*"
    Dim fileList() As String
    Dim intFoundFiles As Integer
    strFileName = Dir(strFileSpec)
    Do While Len(strFileName) > 0
        ReDim Preserve fileList(intFoundFiles)
        fileList(intFoundFiles) = strFileName
        intFoundFiles = intFoundFiles + 1
        strFileName = Dir
    Loop
        
    Dim iFileList As Long
        
    For iFileList = 0 To UBound(fileList)
    
        Dim currentFileDateString As String
    
        currentFileDateString = Left(fileList(iFileList), InStr(fileList(iFileList), " ") - 1)
        
        currentFileDateString = Replace(currentFileDateString, ".", "/")
        
        Dim currentFileDate As Date
        
        currentFileDate = CDate(currentFileDateString)
        
        Dim fileDateMax As Date
        
        If currentFileDateString > fileDateMax And _
        InStr(UCase(fileList(iFileList)), UCase(fileQualifyer)) > 0 And _
        currentFileDate <= relativeToWhatDate Then
        
            Dim madeItToVerification As Boolean
            
            madeItToVerification = True
        
            fileDateMax = currentFileDateString
            
            Dim fileDateMaxIndex As Long
            
            fileDateMaxIndex = iFileList
        
        End If
    
    Next iFileList

    If madeItToVerification = True Then

        'Below loops through each open workbook to see if it's the workbook to open
    
        Dim currentWeeklyLCRWorkbook As Workbook
    
        For Each currentWeeklyLCRWorkbook In Workbooks
        
            If InStr(UCase(currentWeeklyLCRWorkbook.Name), UCase(fileList(fileDateMaxIndex))) > 0 Then
            
                Exit For
            
            End If
            
        Next currentWeeklyLCRWorkbook
        
        If currentWeeklyLCRWorkbook Is Nothing Then
            
            SetAttr strFolder & fileList(fileDateMaxIndex), vbNormal
            
            Set currentWeeklyLCRWorkbook = Workbooks.Open(strFolder & fileList(fileDateMaxIndex))
        
            If currentWeeklyLCRWorkbook.ReadOnly = True Then
        
                currentWeeklyLCRWorkbook.ChangeFileAccess Mode:=xlReadWrite
        
            End If
        
        End If
    
        Set funOpenNewestFile = currentWeeklyLCRWorkbook

    Else
    
        Set funOpenNewestFile = Nothing

    End If

End Function            

Sub funApplyAutofilterToSheet(sheetThatNeedsAFilter As Worksheet, headerRow As Long, Optional lastRow As Long)

    ' Applies Autofilter to worksheet
                    
    If Not sheetThatNeedsAFilter.AutoFilter Is Nothing Then
    
        sheetThatNeedsAFilter.AutoFilterMode = False
    
    End If
    
    Dim lastColumnLetter As String
    
    lastColumnLetter = modFunctions.getLastColumnLetter(sheetThatNeedsAFilter, headerRow)
    
    If lastRow = 0 Then
    
        lastRow = modFunctions.getLastRow(sheetThatNeedsAFilter)
    
    End If
    
    sheetThatNeedsAFilter.Range("A" & headerRow & ":" & lastColumnLetter & lastRow).AutoFilter
    
End Sub                

Function returnToWindow(windowIndex As Long)

    ' If window changes, this can return to the previous window on passed index 
                    
    ActiveWindow.ActivatePrevious

    Do Until ActiveWindow.WindowNumber = windowIndex

        ActiveWindow.ActivateNext
    
    Loop

End Function                

Function getColumnLetter(findString As String, currentSheet As Worksheet, Optional headerRow As Long)

    ' Returns column letter by header string
                    
    Dim getPosition As Range

    If headerRow = 0 Then

        Set getPosition = currentSheet.Cells.Find(What:=findString, After:=currentSheet.Range("A1"), LookIn:=xlFormulas, LookAt:=xlWhole, SearchDirection:=xlNext)
    
    Else
    
        Set getPosition = currentSheet.Rows(headerRow).Find(What:=findString, After:=currentSheet.Range("A" & headerRow), LookAt:=xlWhole, SearchDirection:=xlNext)
    
    End If
    
    If Not getPosition Is Nothing Then
    
        getColumnLetter = Split(getPosition.Address, "$")(1)
    
    Else
    
        getColumnLetter = "Not Found"
    
    End If
    
End Function
                        
Function getColumnNumber(findString As String, currentSheet As Worksheet, Optional headerRow As Long)

    ' Returns column index by pass header string
                            
    Dim getPosition As Range

    If headerRow = 0 Then

        Set getPosition = currentSheet.Cells.Find(What:=findString, After:=currentSheet.Range("A1"), LookAt:=xlWhole, SearchDirection:=xlNext)
    
    Else
    
        Set getPosition = currentSheet.Rows(headerRow).Find(What:=findString, After:=currentSheet.Range("A" & headerRow), LookAt:=xlWhole, SearchDirection:=xlNext)
    
    End If
    
    If Not getPosition Is Nothing Then
    
        getColumnNumber = getPosition.Column
    
    Else
    
        getColumnNumber = "Not Found"
    
    End If
    
End Function
                                
Function getColumnLetterPart(findString As String, currentSheet As Worksheet)

    ' Returns column letter by partial header string match
                                    
    Dim getPosition As Range

    Set getPosition = currentSheet.Cells.Find(What:=findString, After:=currentSheet.Range("A1"), LookAt:=xlPart, SearchDirection:=xlNext)
    
    If Not getPosition Is Nothing Then
    
        getColumnLetterPart = Split(getPosition.Address, "$")(1)
        
    Else
    
        getColumnLetterPart = "Not Found"
    
    End If
    
End Function

Function getRangePart(findString As String, currentSheet As Worksheet, Optional withinRange As Range, Optional afterRange As Range) As Range

    Dim getPosition As Range

    If withinRange Is Nothing And afterRange Is Nothing Then

        Set getPosition = currentSheet.Cells.Find(What:=findString, After:=currentSheet.Range("A1"), LookAt:=xlPart, SearchDirection:=xlNext)
    
    ElseIf Not withinRange Is Nothing And afterRange Is Nothing Then

        Set getPosition = currentSheet.Range(withinRange.Address).Find(What:=findString, After:=currentSheet.Range("A" & withinRange.Row), LookAt:=xlPart, SearchDirection:=xlNext)
    
    ElseIf Not withinRange Is Nothing And Not afterRange Is Nothing Then
    
        Set getPosition = currentSheet.Range(withinRange.Address).Find(What:=findString, After:=currentSheet.Range(afterRange.Address), LookAt:=xlPart, SearchDirection:=xlNext)
    
    ElseIf withinRange Is Nothing And Not afterRange Is Nothing Then
    
        Set getPosition = currentSheet.Cells.Find(What:=findString, After:=currentSheet.Range(afterRange.Address), LookAt:=xlPart, SearchDirection:=xlNext)
    
    Else
    
        Set getPosition = currentSheet.Range(withinRange.Address).Find(What:=findString, After:=currentSheet.Range("A" & withinRange.Row), LookAt:=xlPart, SearchDirection:=xlNext)
    
    End If
    
    Set getRangePart = getPosition
    
End Function

Function getRangeWhole(findString As String, currentSheet As Worksheet, Optional withinRange As Range, Optional afterRange As Range) As Range

    Dim getPosition As Range

    If Len(findString) <= 255 Then

        If withinRange Is Nothing And afterRange Is Nothing Then

            Set getPosition = currentSheet.Cells.Find(What:=findString, After:=currentSheet.Range("A1"), LookAt:=xlWhole, SearchDirection:=xlNext)
    
        ElseIf Not withinRange Is Nothing And afterRange Is Nothing Then
    
            Set getPosition = currentSheet.Range(withinRange.Address).Find(What:=findString, After:=currentSheet.Range("A" & withinRange.Row), LookAt:=xlWhole, SearchDirection:=xlNext)
    
        ElseIf Not withinRange Is Nothing And Not afterRange Is Nothing Then
        
            Set getPosition = currentSheet.Range(withinRange.Address).Find(What:=findString, After:=currentSheet.Range(afterRange.Address), LookAt:=xlWhole, SearchDirection:=xlNext)
        
        ElseIf withinRange Is Nothing And Not afterRange Is Nothing Then
        
            Set getPosition = currentSheet.Cells.Find(What:=findString, After:=currentSheet.Range(afterRange.Address), LookAt:=xlWhole, SearchDirection:=xlNext)
        
        End If
    
    End If
    
    Set getRangeWhole = getPosition
    
End Function

Function getColumnHeaderRow(findString As String, currentSheet As Worksheet)

    If findString = "" Then

        getColumnHeaderRow = 1

    Else

        Dim getPosition As Range
    
        Set getPosition = currentSheet.Cells.Find(What:=findString, After:=currentSheet.Range("A1"), LookAt:=xlWhole, SearchDirection:=xlNext)
        
        If getPosition Is Nothing Then
        
            getColumnHeaderRow = -1
        
        Else
        
            getColumnHeaderRow = Split(getPosition.Address, "$")(2)

        End If

    End If

End Function                                
                
Function getLastRow(currentSheet As Worksheet)

    Dim getPosition As Range

    Set getPosition = currentSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)

    If getPosition Is Nothing Then

        getLastRow = 0

    Else

        Dim splitAddress As Range
    
        Set splitAddress = getPosition
    
        ''' xlCellTypeLastCell
    
        If currentSheet.ProtectContents = False Then
    
            Set getPosition = currentSheet.Cells.SpecialCells(xlCellTypeLastCell)
        
            If getPosition.Row > splitAddress.Row Then
            
                If Not IsEmpty(getPosition) Then
                
                    Set splitAddress = getPosition
                
                End If
                
            End If
    
        End If
    
        ''''''
    
        ''' UsedRange
        
        Set getPosition = currentSheet.UsedRange
    
        Dim usedRangeAsRange As Range
        
        If UBound(Split(getPosition.Address, "$")) = 4 Then
        
            If IsNumeric(Split(getPosition.Address, "$")(4)) Then
        
                Set usedRangeAsRange = currentSheet.Range("$A$" & Split(getPosition.Address, "$")(4))
        
            Dim usedRangeLastRow As Long
        
                usedRangeLastRow = usedRangeAsRange.Row
            
            Else
            
                usedRangeLastRow = 0
            
            End If
        
        ElseIf UBound(Split(getPosition.Address, "$")) = 2 Then
        
            If IsNumeric(Split(getPosition.Address, "$")(2)) Then
        
                Set usedRangeAsRange = currentSheet.Range("$A$" & Split(getPosition.Address, "$")(2))
        
                usedRangeLastRow = usedRangeAsRange.Row
            
            Else
            
                usedRangeLastRow = 0
            
            End If
        
        Else
        
            If IsNumeric(Split(getPosition.Address, ":")(1)) Then
        
                Set usedRangeAsRange = currentSheet.Range(Split(getPosition.Address, ":")(1))
            
                usedRangeLastRow = usedRangeAsRange.Row
        
            Else
            
                usedRangeLastRow = 0
            
            End If
        
        End If
        
        If usedRangeLastRow > splitAddress.Row Then
        
            Set splitAddress = usedRangeAsRange
            
        End If
    
        ''''''
        
        getLastRow = splitAddress.Row

    End If

End Function

                                                                                                                        
Function getLastColumnLetter(currentSheet As Worksheet, Optional headerRow As Long)

    ''' Find All
    
    Dim getPosition As Range
    
    Set getPosition = currentSheet.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, After:=currentSheet.Cells(Rows.Count, Columns.Count))

    Dim splitAddress As Range

    If getPosition Is Nothing Then
    
        Set getPosition = currentSheet.Range("A1")
    
    End If

    Set splitAddress = getPosition

    ''''''

    ''' xlCellTypeLastCell

    If currentSheet.ProtectContents = False Then

        Set getPosition = currentSheet.Cells.SpecialCells(xlCellTypeLastCell)
    
        If getPosition.Column > splitAddress.Column Then
        
            If Not IsEmpty(getPosition) Then
            
                Set splitAddress = getPosition
            
            End If
            
        End If

    End If

    ''''''
    
    Set getPosition = currentSheet.UsedRange

    If Not getPosition Is Nothing Then

        If InStr(getPosition.Address, ":") > 0 Then

            Set getPosition = currentSheet.Range(Split(getPosition.Address, ":")(1))
        
            If getPosition.Column > splitAddress.Column Then
            
                Set splitAddress = getPosition
            
            End If
    
        Else
        
            If getPosition.Column > splitAddress.Column Then
            
                Set splitAddress = getPosition
            
            End If
        
        End If
    
    End If

    ''' with HeaderRow

    If headerRow > 0 Then

        Set getPosition = currentSheet.Cells(headerRow, Columns.Count).End(xlToLeft)

        If getPosition.Column > splitAddress.Column Then
        
            Set splitAddress = getPosition
        
        End If
        
    End If

    ''''''
    
    getLastColumnLetter = Split(splitAddress.Address, "$")(1)

End Function

Function getLastColumnNumber(currentSheet As Worksheet, Optional headerRow As Long)

    ''' Find All
    
    Dim getPosition As Range
    
    Set getPosition = currentSheet.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)

    Dim splitAddress As Range

    Set splitAddress = getPosition

    ''''''

    ''' xlCellTypeLastCell

    If currentSheet.ProtectContents = False Then

        Set getPosition = currentSheet.Cells.SpecialCells(xlCellTypeLastCell)
    
        If getPosition.Column > splitAddress.Column Then
        
            If Not IsEmpty(getPosition) Then
            
                Set splitAddress = getPosition
            
            End If
            
        End If

    End If

    ''''''
    
    Set getPosition = currentSheet.UsedRange

    If Not getPosition Is Nothing Then

        If InStr(getPosition.Address, ":") > 0 Then

            Dim testForAlpha As Boolean

            testForAlpha = modFunctions.isAlpha(Left(Replace(Split(getPosition.Address, ":")(1), "$", ""), 1))

            If testForAlpha = True Then

                Set getPosition = currentSheet.Range(Split(getPosition.Address, ":")(1))
            
                If getPosition.Column > splitAddress.Column Then
                
                    Set splitAddress = getPosition
                
                End If
    
            End If
    
        Else
        
            If getPosition.Column > splitAddress.Column Then
            
                Set splitAddress = getPosition
            
            End If
        
        End If
    
    End If

    ''' with HeaderRow

    If headerRow <> 0 Then

        Set getPosition = currentSheet.Cells(headerRow, Columns.Count).End(xlToLeft)

        If getPosition.Column > splitAddress.Column Then
        
            Set splitAddress = getPosition
        
        End If
        
    End If

    ''''''

    getLastColumnNumber = splitAddress.Column

End Function

Function isAlpha(str As String) As Boolean
 
    Dim Flag As Boolean
    
    Flag = True
     
    Dim i As Long
     
    For i = 1 To Len(Trim(str))
     
        If Asc(Mid(Trim(str), i, 1)) > 64 And Asc(Mid(Trim(str), i, 1)) < 90 Or _
        Asc(Mid(Trim(str), i, 1)) > 96 And Asc(Mid(Trim(str), i, 1)) < 123 Then
     
        Else
            Flag = False
        End If
    
    Next i
     
    If Flag = True Then
        
        isAlpha = True
    
    Else
        
        isAlpha = False
    
    End If
 
End Function                                                                                                

Function getColumnLetterCreateIfMissing(lookForString As String, headerRow As Long, currentSheet As Worksheet)

    Dim getPosition As Range
    
    Set getPosition = modFunctions.getRangeWhole(lookForString, currentSheet)
    
    If getPosition Is Nothing Then

        Dim priceNotificationLastColumnLetter As String
        
        priceNotificationLastColumnLetter = modFunctions.getLastColumnLetter(currentSheet, headerRow)
    
        If priceNotificationLastColumnLetter = "A" And currentSheet.Range(priceNotificationLastColumnLetter & headerRow) = "" Then
        
            currentSheet.Range(priceNotificationLastColumnLetter & headerRow).Value = lookForString
        
        Else
        
            currentSheet.Range(priceNotificationLastColumnLetter & headerRow).Offset(, 1).Value = lookForString
        
        End If
    
    End If

    getColumnLetterCreateIfMissing = modFunctions.getColumnLetter(lookForString, currentSheet)

End Function

Function getWorkbook(nameOfWorkbookToGet As String) As Workbook
    
    If InStr(nameOfWorkbookToGet, """") > 0 Then
    
        nameOfWorkbookToGet = Replace(nameOfWorkbookToGet, """", "")
    
    End If

    Dim workBookToGet As Workbook

    For Each workBookToGet In Workbooks
        
        If Not workBookToGet Is ThisWorkbook Then
        
            ''' This checks to make sure that there is a workbook that has the string "Price Agreement Analyzer" in the name
        
            If InStr(workBookToGet.FullName, nameOfWorkbookToGet) > 0 Then
            
                Exit For
            
            End If
            
        End If
    
    Next workBookToGet
    
    Set getWorkbook = workBookToGet

End Function

Function getWorksheet(nameOfWorksheetToGet As String, Optional withinWorkbook As Workbook) As Worksheet

    If withinWorkbook Is Nothing Then
    
        Set withinWorkbook = ActiveWorkbook
    
    End If

    Dim workSheetToGet As Worksheet

    For Each workSheetToGet In withinWorkbook.Worksheets
        
        If InStr(workSheetToGet.Name, nameOfWorksheetToGet) > 0 Then
        
            Exit For
        
        End If
    
    Next workSheetToGet
    
    If Not workSheetToGet Is Nothing Then
    
        Set getWorksheet = workSheetToGet
    
    Else
    
        Set getWorksheet = Nothing
    
    End If

End Function

Function indexOf(searchString As String, arrayToSearchIn As Variant) As Long

    Dim foundSomething As Boolean
    
    foundSomething = False

    Dim iSearch As Long

    For iSearch = 0 To UBound(arrayToSearchIn)
    
        If searchString = arrayToSearchIn(iSearch) Then
            
            indexOf = iSearch
            
            foundSomething = True
            
            Exit For
        
        End If
    
    Next iSearch

    If foundSomething = False Then

        indexOf = -1
    
    End If

'    indexOf = Application.Match(searchString, arrayToSearchIn, False) - 1
    
    If IsError(indexOf) Then
    
        indexOf = -1
    
    End If

End Function

Sub subOpenFolderLocation(strDirectory As String)

    ''' Open folder destination
    
    Dim pID As Variant
    Dim Sh As Object
'    Dim strDirectory As String
    Dim w As Object
    Dim thisPath As String
    
'    strDirectory = "\\DATA.SHAMROCKFOODS.COM\UDRIVE\" & Environ("UserName") & "\Desktop"
    
    strDirectory = UCase(strDirectory)
    
    Set Sh = CreateObject("shell.application")
    
    For Each w In Sh.Windows
        
        If w <> "Internet Explorer" Then
        
            thisPath = w.document.Folder.self.Path
            
            thisPath = UCase(thisPath)

            If strDirectory = thisPath Then
    
                'if already open, bring it front
                w.Visible = False
                w.Visible = True
                Exit For
    
            End If
        
        End If
        
    Next
    
    If w Is Nothing Then
        
        'if you get here, the folder isn't open so open it
        pID = Shell("explorer.exe " & strDirectory, vbNormalFocus)
    
    End If
    
End Sub


Sub subUnhideALL()

    ActiveSheet.Cells.EntireRow.Hidden = False
    
    ActiveSheet.Cells.EntireColumn.Hidden = False
    
    If (ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode) Or ActiveSheet.FilterMode Then
        
        ActiveSheet.ShowAllData
    
    End If

End Sub

Sub subVisibleToHyperlink(Optional startRange As Range)

    If startRange Is Nothing Then
    
        Set startRange = ActiveCell
    
    End If

    Application.Calculation = xlCalculationManual
    
    Application.ScreenUpdating = False

    Dim lastRow As Long
    
    Dim i As Long
    
    lastRow = modFunctions.getLastRow(ActiveSheet)
    
    For i = startRange.Row To lastRow
    
        ActiveSheet.Cells(i, startRange.Column).Hyperlinks.Delete
    
        If ActiveSheet.Cells(i, startRange.Column).EntireRow.Hidden = False Then
        
            If Not IsError(ActiveSheet.Cells(i, startRange.Column)) Then
                
                If ActiveSheet.Cells(i, startRange.Column) <> "" Then
                
                    ActiveSheet.Cells(i, startRange.Column).Formula = Replace(ActiveSheet.Cells(i, startRange.Column).Formula, "P:\", "\\data\crossdep$\")
                
                    If InStr(ActiveSheet.Cells(i, startRange.Column).Formula, "=HYPERLINK") = 0 Then
                    
                        If InStr(ActiveSheet.Cells(i, startRange.Column).Value, """") > 0 Then ActiveSheet.Cells(i, startRange.Column).Value = Replace(ActiveSheet.Cells(i, startRange.Column).Value, """", "")
                    
                        ActiveSheet.Cells(i, startRange.Column).Formula = "=HYPERLINK(""" & ActiveSheet.Cells(i, startRange.Column).Formula & """,""" & ActiveSheet.Cells(i, startRange.Column).Formula & """)"
                    
                    End If
                
                End If
        
            ElseIf IsError(ActiveSheet.Cells(i, startRange.Column)) Then
            
                Dim errorAsString As String
            
                Select Case ActiveSheet.Cells(i, startRange.Column)
            
                    Case CVErr(xlErrDiv0)
                        errorAsString = "#DIV/0!"
                    
                    Case CVErr(xlErrNA)
                        errorAsString = "#N/A"
                    
                    Case CVErr(xlErrName)
                        errorAsString = "#NAME?"
                    
                    Case CVErr(xlErrNull)
                        errorAsString = "#NULL!"
                    
                    Case CVErr(xlErrNum)
                        errorAsString = "#NUM!"
                    
                    Case CVErr(xlErrRef)
                        errorAsString = "#REF!"
                    
                    Case CVErr(xlErrValue)
                        errorAsString = "#VALUE!"
                    
                End Select
            
                ActiveSheet.Cells(i, startRange.Column) = errorAsString
                
            End If
        
        End If
        
    Next i

    Application.CutCopyMode = False

    Application.Calculation = xlCalculationAutomatic
    
    Application.ScreenUpdating = True

End Sub

Public Function CONCATENATEMULTIPLE(Ref As Range, Separator As String) As String
    
    Dim Cell As Range
    
    Dim Result As String
    
    For Each Cell In Ref
        
        Result = Result & Cell.Value & Separator
    
    Next Cell
    
    CONCATENATEMULTIPLE = Left(Result, Len(Result) - 1)
    
End Function
