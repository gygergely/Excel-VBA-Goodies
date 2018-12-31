Attribute VB_Name = "PrintDictionary"
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: printDictionaryToSheet
' Purpose: Create a dictionary with 100 elements and print out keys and values to a sheet
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Gergely Gyetvai
' Date: 30/12/2018
' ----------------------------------------------------------------
Sub printDictionaryToSheet()

    Dim baseDictionary As Object
    Dim sh As Worksheet
    Dim i As Long
    
    'Clear target ranges
    Call clearRanges
    
    'Create a dictionary
    Set baseDictionary = CreateObject("scripting.dictionary")
    
    For i = 1 To 100
        baseDictionary.Add Key:=i, Item:=f_generateRandomString(6, "alpha")
    Next i

    Set sh = ThisWorkbook.Worksheets("DictionaryPrint")
    
    'Print dictionary to rows
    sh.Cells(11, 3).Resize(1, baseDictionary.Count).Value2 = baseDictionary.Keys
    sh.Cells(12, 3).Resize(1, baseDictionary.Count).Value2 = baseDictionary.Items

    'Print dictionary to columns
    sh.Cells(16, 3).Resize(baseDictionary.Count, 1).Value2 = Application.Transpose(baseDictionary.Keys)
    sh.Cells(16, 4).Resize(baseDictionary.Count, 1).Value2 = Application.Transpose(baseDictionary.Items)

End Sub

' ----------------------------------------------------------------
' Procedure Name: clearRanges
' Purpose: Clear content of pre-defined ranges
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Gergely Gyetvai
' Date: 30/12/2018
' ----------------------------------------------------------------
Sub clearRanges()

    Dim sh As Worksheet
    
    'Clear target ranges
    Set sh = ThisWorkbook.Worksheets("DictionaryPrint")
    sh.Range("printDictRows").ClearContents
    sh.Range("printDictColumns").ClearContents
    
End Sub
