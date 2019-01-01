Attribute VB_Name = "NamesWithIdenticalRefersTo"
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: namesWithSameRefersTo
' Purpose: Finding names with identical RefersTo property and create a report
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Gergely Gyetvai
' Date: 31/12/2018
' ----------------------------------------------------------------
Sub namesWithSameRefersTo()

    Dim nm As Name
    Dim nameDictionary As Object, duplicateDictionary As Object
    Dim dictKey As String, duplDictKey As String
    Dim counter As Long
    
    'Check if ActiveWorkbook has any names
    If ActiveWorkbook.Names.Count > 0 Then
    
    'Creating dictionaries:
    '   - nameDictionary: holding unique RefersTos as key and name's name as value
    '   - duplicateDictionary: holding unique name's name as key and RefersTos as value
    
    Set nameDictionary = CreateObject("scripting.dictionary")
    Set duplicateDictionary = CreateObject("scripting.dictionary")
        
        'Loop through all the names in the ActiveWorkbook, adding names RefersTo property to nameDictionary, if it is
        'already existing add them to duplicateDictionary in order to find the 1st occurence check if the name's name is existing
        'in the duplicateDictionary go forward otherwise add it
        
        For Each nm In ActiveWorkbook.Names
        
            dictKey = CStr("'" & nm.RefersTo)
        
            If nameDictionary.exists(dictKey) Then
                counter = counter + 1
            
                duplicateDictionary.Add Key:=nm.Name, Item:=dictKey
                duplDictKey = nameDictionary(dictKey)
            
                If duplicateDictionary.exists(duplDictKey) = False Then
                    duplicateDictionary.Add Key:=duplDictKey, Item:=dictKey
                End If
            
            Else
                nameDictionary.Add Key:=dictKey, Item:=nm.Name
            End If
        
        Next nm
        
        'If there is any duplicated or multiplicated RefersTo create a report
        If counter > 0 Then
            Call createReport(duplicateDictionary)
        Else
            MsgBox "No multiplicated RefersTo found in Names.", vbInformation, "No mulitplicated RefersTo"
        End If

    Else

        MsgBox "There is no name in " & ActiveWorkbook.Name, vbInformation, "No Name"
    
    End If

End Sub

' ----------------------------------------------------------------
' Procedure Name: createReport
' Purpose: Create a formatted report
' Procedure Kind: Sub
' Procedure Access: Public
' Parameter inputDictionary (Object): a dictionary holding names with identical RefersTo property
' Author: Gergely Gyetvai
' Date: 31/12/2018
' ----------------------------------------------------------------
Sub createReport(inputDictionary As Object)

    Dim reportWb As Workbook
    Dim sh As Worksheet
    Dim rng As Range, sortRng As Range
    Dim reportHeader As String

    'Report header full name of the workbook
    reportHeader = ActiveWorkbook.FullName
    
    'Add new workbook
    Set reportWb = Workbooks.Add
    Set sh = reportWb.Worksheets(1)
    
    'Add column headers
    With sh.Cells(5, 1)
        .Offset(0, 0) = "Name's Name"
        .Offset(0, 1) = "Refers To"
    End With
    
    'Print out dictionary keys and items
    sh.Cells(6, 1).Resize(inputDictionary.Count, 1) = Application.Transpose(inputDictionary.keys)
    sh.Cells(6, 2).Resize(inputDictionary.Count, 1) = Application.Transpose(inputDictionary.items)
    
    'Sort results
    Set rng = sh.Cells(6, 2).Resize(inputDictionary.Count, 1)
    Set sortRng = sh.Cells(6, 1).Resize(inputDictionary.Count, 2)
    
    sh.Sort.SortFields.Clear
    sh.Sort.SortFields.Add2 Key:=rng, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With sh.Sort
        .SetRange sortRng
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    
    'Autofit Name & Refers To column
    sh.Columns(1).AutoFit
    sh.Columns(2).AutoFit
    
    'Add report header to the new workbook
    sh.Cells(1, 1) = "Multiplicated Refers Tos " & reportHeader

    'Add generate details
    sh.Cells(2, 1) = "Generated: " & Now() & " by " & Environ("UserName")
    
    'Format columns headers
    Set rng = sh.Range(Cells(5, 1).Address, Cells(5, 2).Address)
    
    With rng
        .Interior.Color = RGB(84, 130, 53)
        .Font.Color = RGB(255, 255, 255)
        .Font.Size = 9
        .Font.Bold = True
        .RowHeight = rng.RowHeight * 2.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    'Freeze pane
    sh.Cells(6, 1).Select
    ActiveWindow.FreezePanes = True
    
    'Format header
    Set rng = sh.Range(Cells(1, 1).Address, Cells(2, 1).Address)
    
    With rng
        .Font.Size = 11
        .Font.Color = RGB(55, 86, 35)
        .Font.Bold = True
    End With
    
    
End Sub
