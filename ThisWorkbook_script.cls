Const validMACAddress As String = "AA:BB:CC:11:22:33"
Const protectedSheetName As String = "Main"
Private locked As Boolean

'##### Events #####

Private Sub Workbook_Open()
    CheckAndUnlockAndShow
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    locked = True
    LockAndHide
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    LockAndHide
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    UnlockAndShow
End Sub


'##### Functions #####

Sub CheckAndUnlockAndShow()
    'Enviroment settings
    locked = True
    Application.DisplayFormulaBar = False
    'Sheet settings
    LockAndHide
    'Enable checks
    If validMACAddress = GetMyMACAddress() Then
        locked = False
        UnlockAndShow
    End If
End Sub

Sub UnlockAndShow()
    If locked = False Then
        Sheets(protectedSheetName).Visible = xlSheetVisible
        Sheets(protectedSheetName).Select
        Sheets("Empty").Delete
    End If
End Sub

Sub LockAndHide()
    If Application.Sheets.Count = 1 Then
        With ThisWorkbook
            .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Empty"
        End With
    End If
    Sheets(protectedSheetName).Visible = xlSheetVeryHidden
End Sub


'##### Utility #####

Function GetMyMACAddress() As String

    'Declaring the necessary variables.
    Dim strComputer     As String
    Dim objWMIService   As Object
    Dim colItems        As Object
    Dim objItem         As Object
    Dim myMACAddress    As String
    
    'Set the computer.
    strComputer = "."
    
    'The root\cimv2 namespace is used to access the Win32_NetworkAdapterConfiguration class.
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    
    'A select query is used to get a collection of network adapters that have the property IPEnabled equal to true.
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    
    'Loop through all the collection of adapters and return the MAC address of the first adapter that has a non-empty IP.
    For Each objItem In colItems
        If Not IsNull(objItem.IPAddress) Then myMACAddress = objItem.MACAddress
        Exit For
    Next
    
    'Return the IP string.
    GetMyMACAddress = myMACAddress

End Function
