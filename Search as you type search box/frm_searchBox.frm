VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_searchBox 
   Caption         =   "Search box"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7980
   OleObjectBlob   =   "frm_searchBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_searchBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'INITALIZE START
'------------------------------------------------
Private Sub UserForm_Initialize()

    'Load list during initalization
    Call loadList
    
    Me.tbox_srch_ID.SetFocus
    
End Sub
'INITALIZE END
'------------------------------------------------

'CLICK & DBLCLICK START
'------------------------------------------------

Private Sub cmd_add_Click()

    Call addItemByClick

End Sub

Private Sub lbox_ID_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Call addItemByClick
    
End Sub

Private Sub lbox_Word_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Call addItemByClick

End Sub

Private Sub cmb_cancel_Click()

    Unload Me
    
End Sub
'CLICK & DBLCLICK END
'------------------------------------------------

'CHANGE EVENTS START
'------------------------------------------------

Private Sub lbox_ID_Change()

    Me.lbox_Word.ListIndex = Me.lbox_ID.ListIndex
    Me.lbox_Word.TopIndex = Me.lbox_ID.TopIndex

End Sub

Private Sub lbox_Word_Change()

    Me.lbox_ID.ListIndex = Me.lbox_Word.ListIndex
    Me.lbox_ID.TopIndex = Me.lbox_Word.TopIndex

End Sub

Private Sub tbox_srch_ID_Keyup(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    Me.tbox_srch_Word.Value = ""
    Call loadList
    
End Sub
Private Sub tbox_srch_Word_Keyup(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Me.tbox_srch_ID.Value = ""
    Call loadList

End Sub
'CHANGE EVENTS END
'------------------------------------------------


' ----------------------------------------------------------------
' Procedure Name: loadList
' Purpose: Filter source list based on search terms
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Gergely Gyetvai
' ----------------------------------------------------------------
Sub loadList()

    Dim baseArray() As Variant
    Dim resultArray() As Variant
    Dim IDArray() As Variant
    Dim wordArray() As Variant
    Dim counter As Long, i As Long

    On Error Resume Next
    
    'Clear current list boxes
    Me.lbox_ID.Clear
    Me.lbox_Word.Clear
    
    'Assign source list to an array, unfiltered
    baseArray = WordList.Range("tbl_WordList")
    
    'Set a default value to filter match counter
    counter = 0
    
    'Iterate through the source list, if search term is found add item to result array
    For i = LBound(baseArray) To UBound(baseArray)
        If ((InStr(1, baseArray(i, 1), Me.tbox_srch_ID.Value, vbTextCompare) > 0 And Me.tbox_srch_Word.Value = "") Or _
            (InStr(1, baseArray(i, 2), Me.tbox_srch_Word.Value, vbTextCompare) > 0 And tbox_srch_ID.Value = "")) Then
            
            counter = counter + 1
            
            ReDim Preserve resultArray(1 To 2, 1 To counter)
            resultArray(1, counter) = baseArray(i, 1)
            resultArray(2, counter) = baseArray(i, 2)
        End If
    Next i
    
    'If there is at least one match, separate result array to two arrays and load them to the listboxes
    If counter > 0 Then
        ReDim IDArray(1 To UBound(resultArray, 2), 1 To 1)
        ReDim wordArray(1 To UBound(resultArray, 2), 1 To 1)
        
        For i = LBound(resultArray, 2) To UBound(resultArray, 2)
            IDArray(i, 1) = resultArray(1, i)
            wordArray(i, 1) = resultArray(2, i)
        Next i
        
        Me.lbox_ID.List = IDArray
        Me.lbox_Word.List = wordArray
        
    End If
    
    On Error GoTo 0
    
End Sub

' ----------------------------------------------------------------
' Procedure Name: addItemByClick
' Purpose: Simple add word and form hide call, called in click events
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Gergely Gyetvai
' ----------------------------------------------------------------
Sub addItemByClick()

    Call addWord
    Me.Hide

End Sub

' ----------------------------------------------------------------
' Procedure Name: addWord
' Purpose: Add selected item to a cell
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Gergely Gyetvai
' ----------------------------------------------------------------
Sub addWord()

    Dim selectedValue As String, selectedIndex As Long
    Dim cell As Range

    selectedValue = ""
    
    Set cell = ActiveCell
    
    If Me.lbox_ID.ListIndex > -1 Then
    
        selectedIndex = Me.lbox_ID.ListIndex
        selectedValue = Me.lbox_ID.List(selectedIndex)

        'Keeping the leading 0s / work around format you search are as text
        cell.Value = "'" & selectedValue
        
    End If
    
End Sub


