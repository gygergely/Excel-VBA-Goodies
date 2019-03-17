Attribute VB_Name = "GetGAL"
Option Explicit

' ----------------------------------------------------------------
' Procedure Name: GetAllGALMembers
' Purpose: get global address list details from Outlook
' Procedure Kind: Sub
' Procedure Access: Public
' Author: Gergely Gyetvai
' Date: 17/03/2019
' ----------------------------------------------------------------
Sub GetAllGALMembers()
    
    Dim i As Long, lastRow As Long, hitCounter As Long
    Dim resultArray() As Variant
    Dim outlApp As Outlook.Application
    Dim outlNameSpace As Outlook.Namespace
    Dim outlGAL As Outlook.AddressList
    Dim outlEntry As Outlook.AddressEntries
    Dim outlMember As Outlook.AddressEntry
    
    'Set up Outlook
    Set outlApp = Outlook.Application
    Set outlNameSpace = outlApp.GetNamespace("MAPI")
    Set outlGAL = outlNameSpace.GetGlobalAddressList()
    
    hitCounter = 0

    'clear all current entries
    lastRow = GAL.Cells.Find("*", , , , xlByRows, xlPrevious).Row
    
    If lastRow > 9 Then
        GAL.Range(Cells(10, 1).Address, Cells(lastRow, 9).Address).Delete xlShiftUp
    End If
    
    'store address entries
    Set outlEntry = outlGAL.AddressEntries
    
    On Error Resume Next

    'loop through address entries and extract details
    For i = 1 To outlEntry.Count
        Application.StatusBar = "Export: &i & " \ " outlEntryCount"

        Set outlMember = outlEntry.Item(i)
        
        'check if address type is user type (not e.g distribution list)
        If outlMember.AddressEntryUserType = olExchangeUserAddressEntry Then
            
            hitCounter = hitCounter + 1
            
            ReDim Preserve resultArray(1 To 9, 1 To hitCounter)
            
            'add to array
            'display name
            resultArray(1, hitCounter) = outlMember.GetExchangeUser.Name
            'FirstName
            resultArray(2, hitCounter) = outlMember.GetExchangeUser.FirstName
            'LastName
            resultArray(3, hitCounter) = outlMember.GetExchangeUser.LastName
            'Phone
            resultArray(4, hitCounter) = outlMember.GetExchangeUser.BusinessTelephoneNumber
            'Email
            resultArray(5, hitCounter) = outlMember.GetExchangeUser.PrimarySmtpAddress
            'Title
            resultArray(6, hitCounter) = outlMember.GetExchangeUser.JobTitle
            'Department
            resultArray(7, hitCounter) = outlMember.GetExchangeUser.Department
            'Location
            resultArray(8, hitCounter) = outlMember.GetExchangeUser.OfficeLocation
            'City
            resultArray(9, hitCounter) = outlMember.GetExchangeUser.City
        
            
        End If
    Next i
    
    'print results to GAL Data Tab
    resultArray = f_transposeArray(resultArray)
    
    GAL.Cells(10, 1).Resize(UBound(resultArray, 1), UBound(resultArray, 2)).Value = resultArray
            
    'save the workbook with results
    ThisWorkbook.Save
    
    'update details
    GAL.Cells(5, 2) = Now()
    
    GAL.Cells(6, 2) = Application.UserName

    'clear the variables
    Set outlApp = Nothing
    Set outlNameSpace = Nothing
    Set outlGAL = Nothing

End Sub

' ----------------------------------------------------------------
' Procedure Name: f_transposeArray
' Purpose: transpose a 2D array
' Procedure Kind: Function
' Procedure Access: Public
' Parameter srcArray (Variant): array to transpose
' Return Type: Variant
' Author: Gergely Gyetvai
' Date: 17/03/2019
' ----------------------------------------------------------------
Function f_transposeArray(srcArray As Variant) As Variant

    Dim x As Long
    Dim y As Long
    Dim Xupper As Long
    Dim Yupper As Long
    Dim tempArray As Variant
    
    Xupper = UBound(srcArray, 2)
    Yupper = UBound(srcArray, 1)
    
    ReDim tempArray(1 To Xupper, 1 To Yupper)
    
    For x = 1 To Xupper
    
        For y = 1 To Yupper
            
            tempArray(x, y) = srcArray(y, x)
            
        Next y
        
    Next x
    
    f_transposeArray = tempArray
    
End Function
