Attribute VB_Name = "LOOPOLEOBJ"
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: OLEObjectsLoop
' Purpose: Loop through OLE objects in the ActiveWorkbook and print their name to the immediate window
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Gergely Gyetvai
' Date: 11/11/2018
' ----------------------------------------------------------------
Sub OLEObjectsLoop()

    Dim oleObj As OLEObject
    Dim sh As Worksheet
    Dim counter As Long

    For Each sh In ActiveWorkbook.Worksheets
    
        If sh.OLEObjects.Count > 0 Then
                
            For Each oleObj In sh.OLEObjects
                Debug.Print oleObj.Name
            Next oleObj

        End If
        
    Next sh
    
End Sub
