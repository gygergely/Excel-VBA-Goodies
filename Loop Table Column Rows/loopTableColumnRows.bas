Attribute VB_Name = "loopTableColumnRows"
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: loopTableRowsInAColumnRange
' Purpose: Loop through rows of a pre-defined table column
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Gergely Gyetvai
' Date: 05/01/2019
' ----------------------------------------------------------------
Sub loopTableRowsInAColumnRange()

    Dim rng As Range
    Dim cell As Range
    Dim i As Long

    Set rng = shTableRows.Range("tbl_grocery[country_code]")
    
    Debug.Print "RANGE APPROACH 1 - FOR EACH LOOP"
    Debug.Print "--------------------------------"
    
    For Each cell In rng
    
        If cell.Offset(0, 1).Value2 = "Pasta - Ravioli" Then
            Debug.Print cell.Value2 & vbTab & cell.Offset(0, 1) & vbTab & cell.Offset(0, 2)
        End If
        
    Next cell
    
    Debug.Print "--------------------------------"
    Debug.Print ""
    Debug.Print "RANGE APPROACH 2 - FOR LOOP"
    Debug.Print "--------------------------------"
    
    For i = 1 To rng.Rows.Count
    
        If rng.Cells(i, 1).Offset(0, 1).Value2 = "Pasta - Ravioli" Then
            Debug.Print rng.Cells(i, 1).Value2 & vbTab & rng.Cells(i, 1).Offset(0, 1) & vbTab & rng.Cells(i, 1).Offset(0, 2)
        End If
        
    Next i
    
    Debug.Print "--------------------------------"
    
End Sub

' ----------------------------------------------------------------
' Procedure Name: loopTableRowsInAColumnListObject
' Purpose: Loop through rows a pre-defined table column
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Gergely Gyetvai
' Date: 05/01/2019
' ----------------------------------------------------------------
Sub loopTableRowsInAColumnListObject()

    Dim tbl As ListObject
    Dim lrow As Range

    Set tbl = shTableRows.ListObjects("tbl_grocery")
    
    Debug.Print "LISTOBJECT APPROACH - FOR EACH LOOP"
    Debug.Print "-----------------------------------"

    For Each lrow In tbl.ListColumns("country_code").DataBodyRange.Rows
        
        If lrow.Offset(0, 1).Value2 = "Pasta - Ravioli" Then
            Debug.Print lrow.Value2 & vbTab & lrow.Offset(0, 1) & vbTab & lrow.Offset(0, 2)
        End If
    
    Next lrow
    
    Debug.Print "-----------------------------------"
    
End Sub
