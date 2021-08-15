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
