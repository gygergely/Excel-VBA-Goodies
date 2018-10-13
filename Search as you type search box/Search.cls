VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' ----------------------------------------------------------------
' Procedure Name: Worksheet_BeforeDoubleClick
' Purpose: Show the search box if user double click in 'searchArea'
' Procedure Kind: Sub
' Procedure Access: Private
' Author: Gergely Gyetvai
' ----------------------------------------------------------------
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

    If Not Intersect(Target, Range("searchArea")) Is Nothing Then
        
        Dim cell As Range
        
        Set cell = Target
        
        frm_searchBox.Show
        Unload frm_searchBox
        
        cell.Offset(1, 0).Activate
    
    End If

End Sub

