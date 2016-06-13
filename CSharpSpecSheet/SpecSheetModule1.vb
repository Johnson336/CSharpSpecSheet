Dim FilePath As String
Dim iLoop As Variant
Dim missingcount As Integer
Dim PrevHDDQTY As Integer
Dim drive As Variant
Dim PCdrives As String
Dim network As Variant
Dim PCnetwork As String
Dim accessories As Variant
Dim PCaccessories As String
Dim ctrlValue As Variant
Dim strPass As String
Dim wSheet As Worksheet
Dim pipe As String
Dim sp As String
Dim listTitle As String
Dim portName As String
Private Declare Function FindWindow Lib "User32" _
Alias "FindWindowA" ( _
ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long

Private Declare Function GetWindowLong Lib "User32" _
Alias "GetWindowLongA" ( _
ByVal hwnd As Long, _
ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "User32" _
Alias "SetWindowLongA" (ByVal hwnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long

Private Declare Function DrawMenuBar Lib "User32" ( _
ByVal hwnd As Long) As Long

' start clipboard API calls
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) _
   As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) _
   As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
   ByVal dwBytes As Long) As Long
Declare Function CloseClipboard Lib "User32" () As Long
Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) _
   As Long
Declare Function EmptyClipboard Lib "User32" () As Long
Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
   ByVal lpString2 As Any) As Long
Declare Function SetClipboardData Lib "User32" (ByVal wFormat _
   As Long, ByVal hMem As Long) As Long

Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096
' end clipboard API calls

Function ClipBoard_SetData(MyString As String)
   Dim hGlobalMemory As Long, lpGlobalMemory As Long
   Dim hClipMemory As Long, X As Long

   ' Allocate moveable global memory.
   '-------------------------------------------
   hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)

   ' Lock the block to get a far pointer
   ' to this memory.
   lpGlobalMemory = GlobalLock(hGlobalMemory)

   ' Copy the string to this global memory.
   lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

   ' Unlock the memory.
   If GlobalUnlock(hGlobalMemory) <> 0 Then
      MsgBox "Could not unlock memory location. Copy aborted."
      GoTo OutOfHere2
   End If

   ' Open the Clipboard to copy data to.
   If OpenClipboard(0&) = 0 Then
      MsgBox "Could not open the Clipboard. Copy aborted."
      Exit Function
   End If

   ' Clear the Clipboard.
   X = EmptyClipboard()

   ' Copy the data to the Clipboard.
   hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere2:

   If CloseClipboard() = 0 Then
      MsgBox "Could not close Clipboard."
   End If

End Function

Sub RemoveCaption(objForm As Object)
     
    Dim lStyle          As Long
    Dim hMenu           As Long
    Dim mhWndForm       As Long
     
    If Val(Application.Version) < 9 Then
        mhWndForm = FindWindow("ThunderXFrame", objForm.Caption) 'XL97
    Else
        mhWndForm = FindWindow("ThunderDFrame", objForm.Caption) 'XL2000+
    End If
    lStyle = GetWindowLong(mhWndForm, -16)
    lStyle = lStyle And Not &HC00000
    SetWindowLong mhWndForm, -16, lStyle
    DrawMenuBar mhWndForm
    
End Sub
Private Sub FormAdmin()
    formAdminPassword.Show
End Sub
Private Sub CheckAdminPassword()

With formAdminPassword.txtAdminPassword
strPass = .Text
If strPass <> "TCGisgr8" Then
    MsgBox "Incorrect Password!", vbCritical, "Access Denied"
    strPass = ""
    .Text = ""
    .SetFocus
    Exit Sub
End If
End With

Sheets("Data").Select
Application.Visible = True
Unload formAdminPassword
Unload formSpecSheet
For Each wSheet In Worksheets
    wSheet.Unprotect Password:="TCGisgr8"
Next wSheet
End Sub
Sub openform()
    

For Each wSheet In Worksheets
    wSheet.Protect Password:="TCGisgr8", UserInterfaceOnly:=True
Next wSheet

Sheets("Sheet1").Select
Application.Visible = False
frmLoading.Show False
Application.Wait Now + TimeValue("00:00:01")
Run "UpdateRanges"
ActiveWorkbook.Save

formSpecSheet.Show False

'// Update dropdowns from ranges
With formSpecSheet
    .labelDate.Caption = Date
    .txtHDDQTY.Text = CStr(.spinHDDQTY.Value)
    .dropCondition.RowSource = "Data!rangeCondition"
    .dropBrand.RowSource = "Data!rangeBrand"
    .dropFormFactor.RowSource = "Data!rangeFormFactor"
    .txtCPUQTY.Text = CStr(.spinCPUQTY.Value)
    .txtCPUCores.Text = CStr(.spinCPUCores.Value)
    .txtCaddyQTY.Text = CStr(.spinCaddyQTY.Value)
    .dropCPUType.RowSource = "Data!rangeCPUType"
    .dropMemoryType.RowSource = "Data!rangeMemoryType"
    .dropMemorySize.RowSource = "Data!rangeMemorySize"
    .dropMemoryRating.RowSource = "Data!rangeMemoryRating"
    .dropMemorySpeed.RowSource = "Data!rangeMemorySpeed"
    .dropHDDType.RowSource = "Data!rangeHDDType"
    .dropHDDRPM.RowSource = "Data!rangeHDDRPM"
    .dropOpticalDrive.RowSource = "Data!rangeOpticalDrive"
    .dropVideo.RowSource = "Data!rangeVideo"
    .dropCOA.RowSource = "Data!rangeCOA"
    .dropDamage.RowSource = "Data!rangeDamage"

    For Each iLoop In .framePorts.Controls
        If TypeName(iLoop) = "TextBox" Then
            iLoop.Text = "0"
        End If
    Next iLoop
End With
PrevHDDQTY = 0
Unload frmLoading
Run "checkFG"
'populate tester field with username
formSpecSheet.txtTester.Value = Environ("USERNAME")

'check for autoupdater.bat
If Dir(ActiveWorkbook.Path & "\autoupdater.bat") <> "" Then
    formSpecSheet.btnUpdate.Enabled = True
    formSpecSheet.btnSync.Enabled = True
Else
    formSpecSheet.btnUpdate.Enabled = False
    formSpecSheet.btnSync.Enabled = False
End If


End Sub
Private Sub spinHDDQTYchange()
With formSpecSheet
    .txtHDDQTY.Text = CStr(.spinHDDQTY.Value)
    If .spinHDDQTY.Value = 0 Then
        .txtHDDSize.Enabled = False
        .txtHDDSize.Locked = True
        .txtHDDSize.Text = "N/A"
        .txtHDDSize.TabStop = False
        .txtHDDSerial.Enabled = False
        .txtHDDSerial.Locked = True
        .txtHDDSerial.Text = "N/A"
        .txtHDDSerial.TabStop = False
        .dropHDDType.Enabled = False
        .dropHDDType.Locked = True
        .dropHDDType.ListIndex = -1
        .dropHDDType.TabStop = False
        .dropHDDRPM.Enabled = False
        .dropHDDRPM.Locked = True
        .dropHDDRPM.ListIndex = -1
        .dropHDDRPM.TabStop = False
    Else
        If PrevHDDQTY = 0 Then
            .txtHDDSize.Enabled = True
            .txtHDDSize.Locked = False
            .txtHDDSize.Text = ""
            .txtHDDSize.TabStop = True
            .dropHDDType.Enabled = True
            .dropHDDType.Locked = False
            .dropHDDType.TabStop = True
            .dropHDDRPM.Enabled = True
            .dropHDDRPM.Locked = False
            .dropHDDRPM.TabStop = True
        End If
    End If
    If .spinHDDQTY.Value = 1 Then
        .txtHDDSerial.Enabled = True
        .txtHDDSerial.Locked = False
        .txtHDDSerial.Text = ""
        .txtHDDSerial.TabStop = True
    End If
    If .spinHDDQTY.Value > 1 Then
        .txtHDDSerial.Enabled = False
        .txtHDDSerial.Locked = True
        .txtHDDSerial.Text = "Various"
        .txtHDDSerial.TabStop = False
    End If
PrevHDDQTY = .spinHDDQTY.Value
End With
End Sub
Private Sub spinCPUQTYchange()
With formSpecSheet
    .txtCPUQTY.Text = CStr(.spinCPUQTY.Value)
End With
End Sub
Private Sub spinCPUCoreschange()
With formSpecSheet
    .txtCPUCores.Text = CStr(.spinCPUCores.Value)
End With
End Sub

Private Sub spinCaddyQTYchange()
With formSpecSheet
    .txtCaddyQTY.Text = CStr(.spinCaddyQTY.Value)
End With
End Sub
Private Sub UpdateRanges()
'// Update ranges on data sheet
Sheets("Data").Range(Sheets("Data").Range("A2"), Sheets("Data").Range("A2").End(xlDown)).Name = "rangeCondition"
Sheets("Data").Range(Sheets("Data").Range("B2"), Sheets("Data").Range("B2").End(xlDown)).Name = "rangeBrand"
Sheets("Data").Range(Sheets("Data").Range("C2"), Sheets("Data").Range("C2").End(xlDown)).Name = "rangeFormFactor"
Sheets("Data").Range(Sheets("Data").Range("D2"), Sheets("Data").Range("D2").End(xlDown)).Name = "rangeCPUType"
Sheets("Data").Range(Sheets("Data").Range("E2"), Sheets("Data").Range("E2").End(xlDown)).Name = "rangeMemoryType"
Sheets("Data").Range(Sheets("Data").Range("F2"), Sheets("Data").Range("F2").End(xlDown)).Name = "rangeMemorySize"
Sheets("Data").Range(Sheets("Data").Range("G2"), Sheets("Data").Range("G2").End(xlDown)).Name = "rangeMemoryRating"
Sheets("Data").Range(Sheets("Data").Range("H2"), Sheets("Data").Range("H2").End(xlDown)).Name = "rangeMemorySpeed"
Sheets("Data").Range(Sheets("Data").Range("I2"), Sheets("Data").Range("I2").End(xlDown)).Name = "rangeHDDType"
Sheets("Data").Range(Sheets("Data").Range("J2"), Sheets("Data").Range("J2").End(xlDown)).Name = "rangeHDDRPM"
Sheets("Data").Range(Sheets("Data").Range("K2"), Sheets("Data").Range("K2").End(xlDown)).Name = "rangeOpticalDrive"
Sheets("Data").Range(Sheets("Data").Range("L2"), Sheets("Data").Range("L2").End(xlDown)).Name = "rangeVideo"
Sheets("Data").Range(Sheets("Data").Range("M2"), Sheets("Data").Range("M2").End(xlDown)).Name = "rangeCOA"
Sheets("Data").Range(Sheets("Data").Range("N2"), Sheets("Data").Range("N2").End(xlDown)).Name = "rangeDamage"


End Sub

Function IsFileOpen(filename As String)
' This function checks to see if a file is open or not. If the file is
' already open, it returns True. If the file is not open, it returns
' False. Otherwise, a run-time error occurs because there is
' some other problem accessing the file.
    Dim filenum As Integer, errnum As Integer

    On Error Resume Next   ' Turn error checking off.
    filenum = FreeFile()   ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open filename For Input Lock Read As #filenum
    Close filenum          ' Close the file.
    errnum = Err           ' Save the error number that occurred.
    On Error GoTo 0        ' Turn error checking back on.

    ' Check to see which error occurred.
    Select Case errnum

        ' No error occurred.
        ' File is NOT already open by another user.
        Case 0
         IsFileOpen = False

        ' Error number for "Permission Denied."
        ' File is already opened by another user.
        Case 70
            IsFileOpen = True

        ' Another error occurred.
        Case Else
            Error errnum
    End Select

End Function
Sub SaveData()
FilePath = ActiveWorkbook.Path
ActiveWorkbook.Save
ActiveSheet.Copy
Application.ScreenUpdating = False
Application.DisplayAlerts = False

If Dir(ActiveWorkbook.Path & "\data", vbDirectory) = "" Then
    MkDir Path:=ActiveWorkbook.Path & "\data"
End If

    ActiveWorkbook.SaveAs FilePath & "\data\SpecSheetData.xlsx", FileFormat:=51

Application.DisplayAlerts = True
ActiveWorkbook.Close Savechanges:=False
Application.ScreenUpdating = True
End Sub
Sub PrintLabel()
FilePath = ActiveWorkbook.Path

'Shell FilePath & "\SpecSheet.bat", vbNormalFocus

Dim Shex As Object
Set Shex = CreateObject("Shell.Application")
Shex.Open (FilePath & "\SpecSheet.lbx")

Application.Wait DateAdd("s", 4, Now)
Application.SendKeys "%fp~", True
Application.Wait DateAdd("s", 3, Now)
Application.SendKeys "~", True
Application.Wait DateAdd("s", 1, Now)
Application.SendKeys "%{F4}", True
End Sub

Private Sub GenerateDescription()

'// Set up field names
Dim descriptionFields(1 To 28) As String
descriptionFields(1) = "Condition: "
descriptionFields(2) = "Brand: "
descriptionFields(3) = "Model: "
descriptionFields(4) = "Form Factor: "
descriptionFields(5) = "LCD Size: "
descriptionFields(6) = "QTY: "
descriptionFields(7) = "Cores: "
descriptionFields(8) = "Type: "
descriptionFields(9) = "Speed: "
descriptionFields(10) = "Bus Speed: "
descriptionFields(11) = "Type: "
descriptionFields(12) = "Size: "
descriptionFields(13) = "Rating: "
descriptionFields(14) = "Speed: "
descriptionFields(15) = "QTY: "
descriptionFields(16) = "Size: "
descriptionFields(17) = "Interface: "
descriptionFields(18) = "RPM: "
descriptionFields(19) = "Video: "
descriptionFields(20) = "Drives: "
descriptionFields(21) = "Network: "
descriptionFields(22) = "COA: "
descriptionFields(23) = "Installed: "
descriptionFields(24) = "Ports: "
descriptionFields(25) = "Accessories: "
descriptionFields(26) = "Notes: "
descriptionFields(27) = "Series: "
descriptionFields(28) = "CPU: "

Dim sheetFields(1 To 28) As String
sheetFields(1) = "C2"
sheetFields(2) = "D2"
sheetFields(3) = "F2"
sheetFields(4) = "G2"
sheetFields(5) = "AD2"
sheetFields(6) = "H2"
sheetFields(7) = "I2"
sheetFields(8) = "L2"
sheetFields(9) = "J2"
sheetFields(10) = "K2"
sheetFields(11) = "M2"
sheetFields(12) = "N2"
sheetFields(13) = "AE2"
sheetFields(14) = "O2"
sheetFields(15) = "P2"
sheetFields(16) = "Q2"
sheetFields(17) = "R2"
sheetFields(18) = "U2"
sheetFields(19) = "T2"
sheetFields(20) = "S2"
sheetFields(21) = "V2"
sheetFields(22) = "W2"
sheetFields(23) = "AF2"
sheetFields(24) = "AG2"
sheetFields(25) = "AC2"
sheetFields(26) = "X2"
sheetFields(27) = "AJ2" 'cpu series
sheetFields(28) = "AK2" 'cpu type + cpu name combined

'// Generate description
' Old description generation
'listDescription = "You are bidding on the following item:<ul>"
'For iLoop = 1 To UBound(descriptionFields)
'    listDescription = listDescription & "<li><b>" & descriptionFields(iLoop) & "</b>" & Range(sheetFields(iLoop)).Value & "</li>"
'Next iLoop
'listDescription = listDescription & "</ul>"
'Range("Y2").Value = listDescription

'// New Description Generation
listDescription = "You are bidding on the following item:<ul>"

' System Info
listDescription = listDescription & "<li><b>System Info</b><ul>"
  'brand
  listDescription = listDescription & "<li>" & descriptionFields(2) & Range(sheetFields(2)).Value & "</li>"
  'model
  listDescription = listDescription & "<li>" & descriptionFields(3) & Range(sheetFields(3)).Value & "</li>"
  'form factor
  listDescription = listDescription & "<li>" & descriptionFields(4) & Range(sheetFields(4)).Value & "</li>"
  'conditional LCD size
  If Range(sheetFields(5)).Value <> "N/A" Then
    'LCD Size
    listDescription = listDescription & "<li>" & descriptionFields(5) & Range(sheetFields(5)).Value & "</li>"
  End If
  'condition
  listDescription = listDescription & "<li>" & descriptionFields(1) & Range(sheetFields(1)).Value & "</li>"
listDescription = listDescription & "</ul></li>"
'CPU
listDescription = listDescription & "<li><b>CPU</b><ul>"
  'QTY
  listDescription = listDescription & "<li>" & descriptionFields(6) & Range(sheetFields(6)).Value & "</li>"
  'Type
  listDescription = listDescription & "<li>" & descriptionFields(8) & Range(sheetFields(8)).Value & "</li>"
  'Series
  listDescription = listDescription & "<li>" & descriptionFields(27) & Range(sheetFields(27)).Value & "</li>"
  'Cores
  listDescription = listDescription & "<li>" & descriptionFields(7) & Range(sheetFields(7)).Value & "</li>"
  'Speed
  listDescription = listDescription & "<li>" & descriptionFields(9) & Range(sheetFields(9)).Value & "</li>"
  'Bus
  listDescription = listDescription & "<li>" & descriptionFields(10) & Range(sheetFields(10)).Value & "</li>"
listDescription = listDescription & "</ul></li>"
'Memory
listDescription = listDescription & "<li><b>Memory</b><ul>"
  'Size
  listDescription = listDescription & "<li>" & descriptionFields(12) & Range(sheetFields(12)).Value & "</li>"
  'Type
  listDescription = listDescription & "<li>" & descriptionFields(11) & Range(sheetFields(11)).Value & "</li>"
  'Rating
  listDescription = listDescription & "<li>" & descriptionFields(13) & Range(sheetFields(13)).Value & "</li>"
  'Speed
  listDescription = listDescription & "<li>" & descriptionFields(14) & Range(sheetFields(14)).Value & "</li>"
listDescription = listDescription & "</ul></li>"
'Hard Drive
listDescription = listDescription & "<li><b>Hard Drive</b><ul>"
  'QTY
  listDescription = listDescription & "<li>" & descriptionFields(15) & Range(sheetFields(15)).Value & "</li>"
  If Range(sheetFields(15)).Value > 0 Then
    'Size
    listDescription = listDescription & "<li>" & descriptionFields(16) & Range(sheetFields(16)).Value & "</li>"
    'Interface
    listDescription = listDescription & "<li>" & descriptionFields(17) & Range(sheetFields(17)).Value & "</li>"
    'RPM
    listDescription = listDescription & "<li>" & descriptionFields(18) & Range(sheetFields(18)).Value & "</li>"
  End If
listDescription = listDescription & "</ul></li>"
'Operating System
listDescription = listDescription & "<li><b>Operating System</b><ul>"
  'COA
  listDescription = listDescription & "<li>" & descriptionFields(22) & Range(sheetFields(22)).Value & "</li>"
  'Installed
  listDescription = listDescription & "<li>" & descriptionFields(23) & Range(sheetFields(23)).Value & "</li>"
listDescription = listDescription & "</ul></li>"
'Other
listDescription = listDescription & "<li><b>Other</b><ul>"
  'Video
  listDescription = listDescription & "<li>" & descriptionFields(19) & Range(sheetFields(19)).Value & "</li>"
  'Optical Drives
  listDescription = listDescription & "<li>" & descriptionFields(20) & Range(sheetFields(20)).Value & "</li>"
  'Network
  listDescription = listDescription & "<li>" & descriptionFields(21) & Range(sheetFields(21)).Value & "</li>"
  'Ports
  listDescription = listDescription & "<li>" & descriptionFields(24) & Range(sheetFields(24)).Value & "</li>"
  'Accessories
  listDescription = listDescription & "<li>" & descriptionFields(25) & Range(sheetFields(25)).Value & "</li>"
  'Notes
  listDescription = listDescription & "<li>" & descriptionFields(26) & Range(sheetFields(26)).Value & "</li>"
listDescription = listDescription & "</ul></li>"
'End
listDescription = listDescription & "</ul>"
Range("Y2").Value = listDescription

End Sub

Sub GenerateTitle()
'// Generate title
pipe = " | "
sp = " "

'// Brand & Model
listTitle = Range("D2").Value & sp & Range("F2").Value & sp

'// LCD Size if Laptop
If Range("G2").Value = "Laptop" Or Range("G2").Value = "Tablet Laptop" Or Range("G2").Value = "All-in-One" Then
    listTitle = listTitle & Range("AD2").Value & sp
End If

'// Form Factor
listTitle = listTitle & Range("G2").Value

'// CPU QTY & Speed
If Range("H2") > 0 Then
    If Range("H2") > 1 Then
        listTitle = listTitle & pipe & Range("H2").Value & "x " & Range("J2").Value & sp & Range("L2").Value
    Else
        listTitle = listTitle & pipe & Range("J2").Value & sp & Range("L2").Value
    End If
End If

'// CPU Name
If formSpecSheet.txtCPUName.Text <> "" And formSpecSheet.txtCPUName.Text <> "N/A" Then
     listTitle = listTitle & sp & formSpecSheet.txtCPUName.Text
End If

'// Memory Size & Speed
If Range("N2").Value = "None" Then
    listTitle = listTitle & pipe & "No RAM"
Else
    listTitle = listTitle & pipe & Range("N2").Value & sp & Range("AE2").Value
End If

'// Hard Drives
If Range("P2") > 0 Then
    If Range("P2") > 1 Then
        listTitle = listTitle & pipe & Range("P2").Value & "x " & Range("Q2").Value
    Else
        listTitle = listTitle & pipe & Range("Q2").Value
    End If
End If

'// Optical Drives - leave off if title > 80 chars
'// EDIT: Always include optical drive
'If Range("AA2").Value <> "None" And Len(listTitle & Range("AA2").Value) < 81 Then
If Range("AA2").Value <> "None" Then
    listTitle = listTitle & pipe & Range("AA2").Value
End If

'// Title Length Check
'If Len(listTitle) > 80 Then
'    MsgBox "The auction title is over 80 characters. It will need to be edited manually in Kyozou.", vbInformation
'End If

Range("AB2").Value = listTitle

End Sub
Private Sub CheckForMissing()

'// Check for missing inputs
Run "ClearMissingBG"
missingcount = 0
With formSpecSheet
For Each iLoop In .Controls
    If TypeName(iLoop) = "TextBox" Or TypeName(iLoop) = "ComboBox" Then
        If iLoop.Name <> "txtNotes" Then
            If .txtHDDQTY.Text = 0 Then
                'If iLoop.Name = "dropHDDType" Or iLoop.Name = "dropHDDRPM" Or iLoop.Name = "txtHDDSerial" Then
                If iLoop.Name = "dropHDDType" Or iLoop.Name = "dropHDDRPM" Or iLoop.Name = "txtHDDSerial" Then
                    GoTo MissingLoopEnd
                End If
            End If
            If iLoop.Name = "txtCPUName" Or iLoop.Name = "dropDamage" Or iLoop.Name = "checkCPUHT" Or iLoop.Name = "checkCaddyNA" Then
                GoTo MissingLoopEnd
            End If
            
                Select Case TypeName(iLoop)
                    Case "TextBox"
                        ctrlValue = iLoop.Text
                    Case "ComboBox"
                        ctrlValue = iLoop.Value
                End Select
            If ctrlValue = "" Then
                iLoop.BackColor = &H80FFFF
                missingcount = missingcount + 1
            End If
        End If
    End If
MissingLoopEnd:
Next iLoop
End With
End Sub
Private Sub ClearMissingBG()
For Each iLoop In formSpecSheet.Controls
    If TypeName(iLoop) = "TextBox" Or TypeName(iLoop) = "ComboBox" Then
        iLoop.BorderStyle = 0
        iLoop.BackColor = &H80000005
    End If
Next iLoop
End Sub

Private Sub FillSheet()
Sheets("Sheet1").Select
With formSpecSheet
Range("A2").Value = .txtISPF.Text
Range("B2").Value = .labelDate.Caption
Range("C2").Value = .dropCondition.Value

If .dropBrand.Value = "Other:" Then
Range("D2").Value = .txtOtherBrand.Text
Else
Range("D2").Value = .dropBrand.Value
End If

Range("E2").Value = .txtSerial.Text
Range("F2").Value = .txtModel.Text
Range("G2").Value = .dropFormFactor.Value

'Select Case True
'    Case .option1CPU
'        Range("H2").Value = 1
'    Case .option2CPU
'        Range("H2").Value = 2
'    Case .option3CPU
'        Range("H2").Value = 3
'    Case .option4CPU
'        Range("H2").Value = 4
'End Select

'Select Case True
'    Case .option1CPUCore
'        Range("I2").Value = 1
'    Case .option2CPUCore
'        Range("I2").Value = 2
'    Case .option4CPUCore
'        Range("I2").Value = 4
'    Case .option6CPUCore
'        Range("I2").Value = 6
'End Select
Range("H2").Value = .txtCPUQTY.Text

If .checkCPUHT.Value = True Then
  Range("I2").Value = .txtCPUCores.Text & " w/ HT"
Else
    Range("I2").Value = .txtCPUCores.Text
End If

If IsNumeric(.txtCPUSpeed.Text) Then
    If CDbl(.txtCPUSpeed.Text) < 1 Then
        Range("J2").Value = (CDbl(.txtCPUSpeed.Text) * 1000) & "MHz"
    Else
        Range("J2").Value = .txtCPUSpeed.Text & "GHz"
    End If
End If

If .txtFSBSpeed.Text = "?" Then
    Range("K2").Value = "Unknown"
Else
    Range("K2").Value = .txtFSBSpeed.Text
End If

Range("L2").Value = .dropCPUType.Value
Range("M2").Value = .dropMemoryType.Value
Range("N2").Value = .dropMemorySize.Value
Range("O2").Value = .dropMemorySpeed.Value
Range("P2").Value = .txtHDDQTY.Text
If .txtHDDQTY.Value = 0 Then
    Range("Q2").Value = "N/A"
Else
    If IsNumeric(.txtHDDSize.Text) Then
        If CInt(.txtHDDSize.Text) >= 1000 Then
            Range("Q2").Value = Int(CInt(.txtHDDSize.Text) / 1000) & "TB"
        Else
            Range("Q2").Value = .txtHDDSize.Text & "gb"
        End If
    End If
End If
Range("R2").Value = .dropHDDType.Value

Range("AA2").Value = .dropOpticalDrive.Value

PCdrives = ""
If .dropOpticalDrive.Value <> "None" Then
    PCdrives = .dropOpticalDrive.Value
End If
For Each drive In .frameOtherDrives.Controls
      If TypeName(drive) = "CheckBox" And drive.Caption <> "None" And drive.Value = True Then
            If PCdrives <> "" Then
                PCdrives = PCdrives & ", "
            End If
            PCdrives = PCdrives & drive.Caption
      End If
Next drive
If PCdrives = "" Then
    PCdrives = "None"
End If
Range("S2").Value = PCdrives
Range("T2").Value = .dropVideo.Value
If .txtVideoModel.Text <> "N/A" Then
    Range("T2") = Range("T2") & " - " & .txtVideoModel.Text
End If
If .txtVRAM.Text <> "N/A" Then
    Range("T2") = Range("T2") & " " & .txtVRAM.Text & "mb"
End If

'If .dropHDDRPM.ListIndex = -1 Then
'    Range("U2").Value = "Unknown"
'Else
'    Range("U2").Value = .dropHDDRPM.Value
'End If
Range("U2").Value = .dropHDDRPM.Value

PCnetwork = ""
For Each network In .frameNetwork.Controls
    If TypeName(network) = "CheckBox" And network.Caption <> "None" And network.Value = True Then
        If PCnetwork <> "" Then
            PCnetwork = PCnetwork & ", "
        End If
    PCnetwork = PCnetwork & network.Caption
    End If
Next network
If PCnetwork = "" Then
    PCnetwork = "None"
End If
Range("V2").Value = PCnetwork
Range("W2").Value = .dropCOA.Value

'Serial Insertion Into Notes
Range("X2").Value = "S/N: " & .txtSerial.Text


'CPU Name Insertion into Notes
'this gets its own line now - skip it
'If .txtCPUName.Text <> "N/A" And .txtCPUName.Text <> "" Then
'    Range("X2").Value = Range("X2").Value & " | CPU: " & .txtCPUName.Text
'End If

'add regular notes
If .txtNotes.Text <> "" Then
    Range("X2").Value = Range("X2").Value & " | " & .txtNotes.Text
End If

'Caddy QTY insertion into notes
Dim caddystring As String
caddystring = ""
If .checkCaddyNA.Value = False Then
    If .txtCaddyQTY.Value = 0 Then
        caddystring = "No HDD caddies"
    ElseIf .txtCaddyQTY.Value = 1 Then
        caddystring = "Includes 1x HDD caddy"
    Else
        caddystring = "Includes " & .txtCaddyQTY.Value & "x HDD caddies"
    End If
    Range("X2").Value = Range("X2").Value & " | " & caddystring
End If

'Damage Insertion into Notes
If .dropDamage.Value <> "" And .dropDamage.Value <> "N/A" Then
    Range("X2").Value = Range("X2").Value & " | " & .dropDamage.Value
End If

    'use tester field instead of username
If .txtTester.Value <> "" Then
    Range("Z2").Value = LCase(.txtTester.Value)
Else
    'unless tester field is empty, then get username
    Range("Z2").Value = Environ("USERNAME")
End If

PCaccessories = ""
For Each accessories In .frameAccessories.Controls
    If TypeName(accessories) = "CheckBox" And accessories.Caption <> "None" And accessories.Value = True Then
        If PCaccessories <> "" Then
            PCaccessories = PCaccessories & ", "
        End If
    PCaccessories = PCaccessories & accessories.Caption
    End If
Next accessories
If PCaccessories = "" Then
    PCaccessories = "None"
End If
Range("AC2").Value = PCaccessories
If .txtLCDSize.Value <> "N/A" Then
    Range("AD2").Value = .txtLCDSize.Value & """"
Else
    Range("AD2").Value = .txtLCDSize.Value
End If
Range("AE2").Value = .dropMemoryRating.Value

Select Case True
    Case .optionOSNo
        Range("AF2").Value = "No"
    Case .optionOSYes
        Range("AF2").Value = "Yes"
End Select
Range("AG2").Value = ""
    For Each iLoop In .framePorts.Controls
        If TypeName(iLoop) = "TextBox" Then
            If iLoop.Text <> "0" Then
                portName = iLoop.Name
                If portName = "PS2" Then
                    portName = "PS/2"
                End If
                If portName = "SDCard" Then
                    portName = "SD Card"
                End If
                If portName = "SVideo" Then
                    portName = "S-Video"
                End If
                If Range("AG2").Value = "" Then
                    Range("AG2").Value = iLoop.Text & "x " & portName
                Else
                    Range("AG2").Value = Range("AG2").Value & ", " & iLoop.Text & "x " & portName
                End If
            End If
        End If
    Next iLoop

Range("AH2").Value = .txtHDDSerial.Text
Range("AI2").Value = .txtWeight.Text & "lb"
Range("AJ2").Value = .txtCPUName.Text
Range("AK2").Value = .dropCPUType.Value & " " & .txtCPUName.Text
End With

End Sub
Private Sub PrintLabelClick()
Run "CheckForMissing"
    If missingcount > 0 Then
        Exit Sub
    End If
    
'error checking
If IsFileOpen(ActiveWorkbook.Path & "\data\SpecSheetData.xlsx") Then
    MsgBox "Please close P-touch Editor before trying to print!"
    Exit Sub
End If

Run "FillSheet"
Run "GenerateDescription"
Run "GenerateTitle"
Run "SaveData"

'save a log of this unit into archive
Run "ArchiveUnit"

If formSpecSheet.checkMCF.Value = True Then
    Run "OutputMCF"
End If
If formSpecSheet.checkFG.Value = True Then
    Run "OutputFG"
End If

Run "PrintLabel"

'MsgBox Range("AB2").Value & vbNewLine & Len(Range("AB2").Value)
End Sub
Private Sub ClearForm()
Unload formSpecSheet
Run "openform"
End Sub
Private Sub LoadClick()

Dim fileLoadName As String
Dim str As String
Dim arr As Variant
Dim f As Object
Set f = Application.FileDialog(3)


'ensure directory exists
If Dir(ActiveWorkbook.Path & "\saves", vbDirectory) = "" Then
    MkDir Path:=ActiveWorkbook.Path & "\saves"
End If

ChDir (ActiveWorkbook.Path & "\saves")
f.AllowMultiSelect = False
f.Title = "Choose template to load"
f.Filters.Clear
f.Filters.Add "Comma Separated Value Files", "*.csv"
f.InitialView = msoFileDialogViewDetails
f.Show
If f.SelectedItems.Count > 0 Then
    fileLoadName = f.SelectedItems(1)
Else
    Exit Sub
End If

'MsgBox "files chosen: " & f.SelectedItems.Count & " : " & fileLoadName

Open fileLoadName For Input As #1
Line Input #1, str
Close #1

arr = Split(str, ",")

With formSpecSheet
'MsgBox "found " & UBound(arr) & " fields"
    If .lblVersion.Caption >= arr(0) Then
        If IsNumeric(arr(3)) Then
            .dropCondition.ListIndex = arr(3)
        Else
            str = arr(3)
            For i = 0 To .dropCondition.ListCount - 1
                If (.dropCondition.List(i) = str) Then
                    .dropCondition.ListIndex = i
                End If
            Next i
        End If
    'check for other brands correctly
    If IsNumeric(arr(4)) Then
        .dropBrand.ListIndex = arr(4)
    Else
        .dropBrand.ListIndex = 12
        .txtOtherBrand.Text = arr(4)
    End If
    .txtModel.Text = arr(6)
    
    'dropformfactor
    If IsNumeric(arr(7)) Then
        .dropFormFactor.ListIndex = arr(7)
    Else
        str = arr(7)
        For i = 0 To .dropFormFactor.ListCount - 1
            If (.dropFormFactor.List(i) = str) Then
                .dropFormFactor.ListIndex = i
            End If
        Next i
    End If
    
    .spinCPUQTY.Value = arr(8)
    .spinCPUCores.Value = arr(9)
    .txtCPUSpeed.Text = arr(10)
    .txtFSBSpeed.Text = arr(11)
    
    'dropcputype
    If IsNumeric(arr(12)) Then
        .dropCPUType.ListIndex = arr(12)
    Else
        str = arr(12)
        For i = 0 To .dropCPUType.ListCount - 1
            If (.dropCPUType.List(i) = str) Then
                .dropCPUType.ListIndex = i
            End If
        Next i
    End If
    
    'dropmemorytype
    If IsNumeric(arr(13)) Then
        .dropMemoryType.ListIndex = arr(13)
    Else
        str = arr(13)
        For i = 0 To .dropMemoryType.ListCount - 1
            If (.dropMemoryType.List(i) = str) Then
                .dropMemoryType.ListIndex = i
            End If
        Next i
    End If
    'dropmemorysize
    If IsNumeric(arr(14)) Then
        .dropMemorySize.ListIndex = arr(14)
    Else
        str = arr(14)
        For i = 0 To .dropMemorySize.ListCount - 1
            If (.dropMemorySize.List(i) = str) Then
                .dropMemorySize.ListIndex = i
            End If
        Next i
    End If
    'dropmemoryspeed
    If IsNumeric(arr(15)) Then
        .dropMemorySpeed.ListIndex = arr(15)
    Else
        str = arr(15)
        For i = 0 To .dropMemorySpeed.ListCount - 1
            If (.dropMemorySpeed.List(i) = str) Then
                .dropMemorySpeed.ListIndex = i
            End If
        Next i
    End If
    'dropopticaldrive
    If IsNumeric(arr(19)) Then
        .dropOpticalDrive.ListIndex = arr(19)
    Else
        str = arr(19)
        For i = 0 To .dropOpticalDrive.ListCount - 1
            If (.dropOpticalDrive.List(i) = str) Then
                .dropOpticalDrive.ListIndex = i
            End If
        Next i
    End If
    'dropvideo
    If IsNumeric(arr(20)) Then
        .dropVideo.ListIndex = arr(20)
    Else
        str = arr(20)
        For i = 0 To .dropVideo.ListCount - 1
            If (.dropVideo.List(i) = str) Then
                .dropVideo.ListIndex = i
            End If
        Next i
    End If
    'dropcoa
    If IsNumeric(arr(22)) Then
        .dropCOA.ListIndex = arr(22)
    Else
        str = arr(22)
        For i = 0 To .dropCOA.ListCount - 1
            If (.dropCOA.List(i) = str) Then
                .dropCOA.ListIndex = i
            End If
        Next i
    End If
    .txtNotes.Text = arr(23)
    .txtLCDSize.Text = arr(24)
    'dropmemoryrating
    If IsNumeric(arr(25)) Then
        .dropMemoryRating.ListIndex = arr(25)
    Else
        str = arr(25)
        For i = 0 To .dropMemoryRating.ListCount - 1
            If (.dropMemoryRating.List(i) = str) Then
                .dropMemoryRating.ListIndex = i
            End If
        Next i
    End If
    .USB.Text = arr(26)
    .Ethernet.Text = arr(27)
    .Modem.Text = arr(28)
    .VGA.Text = arr(29)
    .DVI.Text = arr(30)
    .SVideo.Text = arr(31)
    .PS2.Text = arr(32)
    .Audio.Text = arr(33)
    .eSATAp.Text = arr(34)
    .Serial.Text = arr(35)
    .Parallel.Text = arr(36)
    .PCMCIA.Text = arr(37)
    .SDCard.Text = arr(38)
    .Firewire.Text = arr(39)
    .eSATA.Text = arr(40)
    .HDMI.Text = arr(41)
    .SCSI.Text = arr(42)
    .DisplayPort.Text = arr(43)
    .txtCPUName.Text = arr(44)
    .checkCPUHT.Value = arr(45)
    
    Else
    MsgBox "Save file version (" & arr(0) & ") doesn't match program version (" & .lblVersion.Caption & "), unable to load."
    End If
    

End With
End Sub
Private Sub SaveClick()

Dim str As String
Dim IntialName As String
Dim brandStr As String
Dim brandSave As String
Dim fileSaveName As Variant
If formSpecSheet.dropBrand.Value = "Other:" Then
    brandStr = formSpecSheet.txtOtherBrand.Text
    brandSave = formSpecSheet.txtOtherBrand.Text
Else
    brandStr = formSpecSheet.dropBrand.Value
    'brandSave = formSpecSheet.dropBrand.ListIndex
    brandSave = formSpecSheet.dropBrand.Value
End If
InitialName = Replace(formSpecSheet.dropFormFactor.Value & " " & brandStr & " " & formSpecSheet.txtModel.Text & ".csv", " ", ".")

'ensure directory exists
If Dir(ActiveWorkbook.Path & "\saves", vbDirectory) = "" Then
    MkDir Path:=ActiveWorkbook.Path & "\saves"
End If
ChDir (ActiveWorkbook.Path & "\saves")

'set default save name and open file chooser
fileSaveName = Application.GetSaveAsFilename(InitialFileName:=InitialName, _
    fileFilter:="Comma Separated Value File (*.csv), *.csv")

'user clicked cancel, abort out
If fileSaveName = False Then
    Exit Sub
End If

With formSpecSheet

      
'str = .lblVersion.Caption & "," & .txtISPF.Text & "," & .labelDate.Caption & "," & .dropCondition.ListIndex & "," & brandSave _
'     & "," & .txtSerial.Text & "," & .txtModel.Text & "," & .dropFormFactor.ListIndex & "," & .spinCPUQTY.Value _
'      & "," & .spinCPUCores.Value & "," & .txtCPUSpeed.Text & "," & .txtFSBSpeed.Text & "," & .dropCPUType.ListIndex _
'       & "," & .dropMemoryType.ListIndex & "," & .dropMemorySize.ListIndex & "," & .dropMemorySpeed.ListIndex & "," & .spinHDDQTY.Value _
'        & "," & .txtHDDSize.Text & "," & .dropHDDType.ListIndex & "," & .dropOpticalDrive.ListIndex & "," & .dropVideo.ListIndex _
'         & "," & .dropHDDRPM.ListIndex & "," & .dropCOA.ListIndex & "," & .txtNotes.Text & "," & .txtLCDSize.Text _
'          & "," & .dropMemoryRating.ListIndex & "," & .USB.Text & "," & .Ethernet.Text & "," & .Modem.Text _
'           & "," & .VGA.Text & "," & .DVI.Text & "," & .SVideo.Text & "," & .PS2.Text & "," & .Audio.Text _
'            & "," & .eSATAp.Text & "," & .Serial.Text & "," & .Parallel.Text & "," & .PCMCIA.Text & "," & .SDCard.Text _
'             & "," & .Firewire.Text & "," & .eSATA.Text & "," & .HDMI.Text & "," & .SCSI.Text & "," & .DisplayPort.Text _
'              & "," & .txtCPUName.Text & "," & .checkCPUHT.Value

' switch from saving index values to real values of drop downs
str = .lblVersion.Caption & "," & .txtISPF.Text & "," & .labelDate.Caption & "," & .dropCondition.Value & "," & brandSave _
     & "," & .txtSerial.Text & "," & .txtModel.Text & "," & .dropFormFactor.Value & "," & .spinCPUQTY.Value _
      & "," & .spinCPUCores.Value & "," & .txtCPUSpeed.Text & "," & .txtFSBSpeed.Text & "," & .dropCPUType.Value _
       & "," & .dropMemoryType.Value & "," & .dropMemorySize.Value & "," & .dropMemorySpeed.Value & "," & .spinHDDQTY.Value _
        & "," & .txtHDDSize.Text & "," & .dropHDDType.Value & "," & .dropOpticalDrive.Value & "," & .dropVideo.Value _
         & "," & .dropHDDRPM.Value & "," & .dropCOA.Value & "," & Replace(.txtNotes.Text, ",", "") & "," & .txtLCDSize.Text _
          & "," & .dropMemoryRating.Value & "," & .USB.Text & "," & .Ethernet.Text & "," & .Modem.Text _
           & "," & .VGA.Text & "," & .DVI.Text & "," & .SVideo.Text & "," & .PS2.Text & "," & .Audio.Text _
            & "," & .eSATAp.Text & "," & .Serial.Text & "," & .Parallel.Text & "," & .PCMCIA.Text & "," & .SDCard.Text _
             & "," & .Firewire.Text & "," & .eSATA.Text & "," & .HDMI.Text & "," & .SCSI.Text & "," & .DisplayPort.Text _
              & "," & .txtCPUName.Text & "," & .checkCPUHT.Value


End With

' Store current form info into save file
Open fileSaveName For Output As #1
Print #1, str
Close #1

End Sub
Private Sub OutputFG()
'// Output FG template line information in easy to copy/paste format
Dim str As String
Dim ispf As String
Dim model As String
Dim Serial As String
Dim brand As String
Dim formfactor As String
Dim speed As String
Dim cpu As String
Dim cpunum As String
Dim cpuname As String
Dim ram As String
Dim hdd As String
Dim serverhdd As String
Dim optical As String
Dim coa As String
Dim notes As String
Dim powerStr As String

Sheets("Sheet1").Select
With formSpecSheet

FilePath = ActiveWorkbook.Path & "\data\FG.txt"

ispf = .txtISPF.Text
model = .txtModel.Text
Serial = .txtSerial.Text
If .dropBrand.Value = "Other:" Then
    brand = .txtOtherBrand.Text
Else
    brand = .dropBrand.Value
End If
formfactor = .dropFormFactor.Value
speed = .txtCPUSpeed.Text
cpunum = .spinCPUQTY.Value
If cpunum > 1 Then
    cpu = cpunum & "x " & .dropCPUType.Value
Else
    cpu = .dropCPUType.Value
End If
cpuname = .txtCPUName.Text

' ram fix
Dim stickTypes() As String
Dim stickCount() As Integer
Dim found As Boolean
ReDim stickTypes(0)
ReDim stickCount(0)
ram = ""

'traverse memorylistbox
For i = 0 To (.memoryListbox.ListCount - 1)
    found = False
    'check stored sticks
    For k = 0 To UBound(stickTypes)
        If stickTypes(k) = .memoryListbox.List(i) Then
            found = True
            stickCount(k) = stickCount(k) + 1
        End If
    Next k
    If found = False Then 'stick wasn't already stored in stickTypes, add it there
        ReDim Preserve stickTypes(UBound(stickTypes) + 1)
        ReDim Preserve stickCount(UBound(stickCount) + 1)
        stickTypes(UBound(stickTypes)) = .memoryListbox.List(i)
        stickCount(UBound(stickCount)) = 1
    End If
Next i

'// servers still get full stick types, others get truncated
If formfactor = "Server" Then
    For i = 1 To UBound(stickTypes)
        ram = ram & stickCount(i) & "x " & UCase(stickTypes(i))
        If i < UBound(stickTypes) Then
            ram = ram & ", "
        End If
    Next i
Else
    If UBound(stickTypes) > 1 Then '// more than one size of ram stick, truncate
        ram = UCase(.dropMemorySize.Value & " " & .dropMemoryType.Value & " " & .dropMemoryRating.Value)
    Else
        For i = 1 To UBound(stickTypes)
            ram = ram & stickCount(i) & "x " & UCase(stickTypes(i))
            If i < UBound(stickTypes) Then
                ram = ram & ", "
            End If
        Next i
    End If
End If




If .spinHDDQTY.Value > 1 Then
    hdd = .spinHDDQTY.Value & "x "
Else
    hdd = ""
End If
If IsNumeric(.txtHDDSize.Text) Then
        If CInt(.txtHDDSize.Text) >= 1000 Then
            hdd = hdd & Int(CInt(.txtHDDSize.Text) / 1000) & "TB " & .dropHDDType.Value
        Else
            hdd = hdd & .txtHDDSize.Text & "gb " & .dropHDDType.Value
        End If
Else
    hdd = "N/A"
End If

'// special sever hdd - # and size are no longer combined
serverhdd = ""
If IsNumeric(.txtHDDSize.Text) Then
        If CInt(.txtHDDSize.Text) >= 1000 Then
            serverhdd = serverhdd & Int(CInt(.txtHDDSize.Text) / 1000) & "TB " & .dropHDDType.Value
        Else
            serverhdd = serverhdd & .txtHDDSize.Text & "gb " & .dropHDDType.Value
        End If
Else
    serverhdd = "N/A"
End If

optical = .dropOpticalDrive.Value
coa = UCase(.dropCOA.Value)


notes = .txtNotes.Text
'// add special case laptop addons to notes
If formfactor = "Laptop" Or formfactor = "Tablet Laptop" Then
    '// check for extended
    If .checkAccessories4.Value = "True" Then
        notes = notes & " | Extended battery"
        '// check for no battery
    ElseIf .checkAccessories4.Value = "False" And .checkAccessories3.Value = "False" Then
        notes = notes & " | No battery"
    End If
    '// check for fingerprint
    If .checkAccessories5.Value = "True" Then
        notes = notes & " | Fingerprint reader"
    End If
    '// check for webcam
    If .checkAccessories6.Value = "True" Then
        notes = notes & " | Webcam"
    End If
End If

'Caddy QTY insertion into notes
Dim caddystring As String
caddystring = ""
If .checkCaddyNA.Value = False Then
    If .txtCaddyQTY.Value = 0 Then
        caddystring = "No HDD caddies"
    ElseIf .txtCaddyQTY.Value = 1 Then
        caddystring = "Includes 1x HDD caddy"
    Else
        caddystring = "Includes " & .txtCaddyQTY.Value & "x HDD caddies"
    End If
    notes = notes & " | " & caddystring
End If

'// put damage into end of notes
notes = notes & " | " & .dropDamage.Value

If .checkAccessories2.Value = True Then
    powerStr = "YES"
Else
    powerStr = "NO"
End If
'//

'// special case for servers due to spreadsheet formatting
If formfactor = "Server" Then
    str = ispf & vbTab & vbTab & vbTab & model & vbTab & Serial & vbTab & brand & vbTab & formfactor & vbTab & cpunum & vbTab & speed & vbTab & .dropCPUType.Value & vbTab & .txtCPUName.Value & vbTab & ram & vbTab & .spinHDDQTY.Value & vbTab & serverhdd & vbTab & optical & vbTab & coa & vbTab & notes & vbTab & "1"
ElseIf formfactor = "Laptop" Or formfactor = "Tablet Laptop" Then
    str = ispf & vbTab & model & vbTab & Serial & vbTab & brand & vbTab & speed & vbTab & cpu & vbTab & cpuname & vbTab & ram & vbTab & .txtHDDSize.Text & vbTab & .txtLCDSize.Text & vbTab & optical & vbTab & powerStr & vbTab & coa & vbTab & notes & vbTab & "1"
Else '// desktops primarily
    str = ispf & vbTab & model & vbTab & Serial & vbTab & brand & vbTab & formfactor & vbTab & speed & vbTab & cpu & vbTab & cpuname & vbTab & ram & vbTab & hdd & vbTab & optical & vbTab & powerStr & vbTab & coa & vbTab & notes & vbTab & "1"
End If

'// Export our info into the output file
Open FilePath For Output As #1
Print #1, str
Close #1

' Save our FG info to the clipboard
ClipBoard_SetData (str)

End With
End Sub

Private Sub OutputMCF()
'// Output title specs to text file for automatic MCF comparison & numbering

' full file data
Dim strData As String
'each line in the original text file
Dim strLine As String
'each line in array format for comparisons
Dim arrString() As String
'current pc spec
Dim spec As String
Dim found As Boolean
Dim firstChar As String
found = False
Dim cpuspeed As String
Dim cpunum As String
Dim brand As String
Dim lcd As String
lcd = ""
Dim thisLine As Boolean
Dim ram As String

Dim lineNum As Integer
lineNum = 0
cpunum = ""
thisLine = False

'build current pc spec
With formSpecSheet
If IsNumeric(.txtCPUSpeed.Text) Then
    If CDbl(.txtCPUSpeed.Text) < 1 Then
        cpuspeed = (CDbl(.txtCPUSpeed.Text) * 1000) & "MHz"
    Else
        cpuspeed = .txtCPUSpeed.Text & "GHz"
    End If
End If
If .spinCPUQTY.Value > 1 Then
    cpunum = .spinCPUQTY.Value & "x "
End If

If .dropBrand.Value = "Other:" Then
    brand = .txtOtherBrand.Text
Else
    brand = .dropBrand.Value
End If

If IsNumeric(.txtLCDSize.Text) Then
  lcd = .txtLCDSize.Text & """"
End If

'add screen size if it has one
spec = brand & " " & .txtModel.Text & " "
If lcd <> "" Then
spec = spec & lcd & " "
End If

'check ram correctly
'// Memory Size & Speed
If .dropMemorySize.Text = "None" Then
    ram = "No RAM"
Else
    ram = .dropMemorySize.Value & " " & .dropMemoryRating.Value
End If

spec = spec & .dropFormFactor.Value & " | " & cpunum & cpuspeed & " " & .dropCPUType.Value & " " & .txtCPUName.Text & " | " & ram

If .spinHDDQTY.Value > 0 Then
    spec = spec & " | "
    If .spinHDDQTY.Value > 1 Then
        spec = spec & .spinHDDQTY.Value & "x "
    End If
    'correctly check for terabyte drives
    If CInt(.txtHDDSize.Text) >= 1000 Then
        spec = spec & Int(CInt(.txtHDDSize.Text) / 1000) & "TB"
    Else
        spec = spec & .txtHDDSize.Text & "gb"
    End If
End If
spec = spec & " | " & .dropOpticalDrive.Value & " | " & .dropDamage.Value
End With

'open the original text file to read the lines
FilePath = ActiveWorkbook.Path & "\data\preMCF.txt"

If Dir(FilePath) <> "" Then ' only open for input if the file already exists

    Open FilePath For Input As #1
    'continue until the end of the file
    While EOF(1) = False
        'read the current line of text
        Line Input #1, strLine
        'replace tabs with spaces so Trim will work
        strLine = Replace(strLine, vbTab, "  ")
    
    
        'get first char of line
        firstChar = Left(Trim(strLine), 1)
    
        'if first char is an asterisk, remove it
        If firstChar = "*" Then
            strLine = Right(Trim(strLine), Len(strLine) - 1)
            firstChar = Left(Trim(strLine), 1)
        End If
        
        'check lines for a comment character after removing white space
        '// valid comment characters at beginning of line are dash (-) or single quote (')
        If firstChar = "-" Or firstChar = "'" Then
            'append the commented line with no formatting
            strData = strData + strLine + vbCrLf
    
        'not a comment line
        Else
            'split line into array for processing
            arrString = Strings.Split(strLine, " , ")
    
            'verify array length for proper comparison
            If UBound(arrString) >= 3 Then
                lineNum = lineNum + 1
                'test current specs against line read from file
                If arrString(3) = spec Then
                    found = True
                    thisLine = True
                    ' increment count
                    arrString(1) = CInt(arrString(1)) + 1
                    ' increment weight
                    arrString(2) = CInt(arrString(2)) + CInt(formSpecSheet.txtWeight.Text)
                End If
                'check if we just incremented this line and asterisk it for easy spotting
                If thisLine = True Then
                    thisLine = False
                    strData = strData & "*" & Strings.Join(arrString, " , ") & vbCrLf
                Else
                    'regular line, no asterisk
                    'rebuild the line from our split array and append to end of file
                    strData = strData & Strings.Join(arrString, " , ") & vbCrLf
                End If
            End If
        End If
    Wend
    Close #1
End If

'entry wasn't found, create a new line
If found = False Then
    strData = strData & "*#" & lineNum + 1 & " , 1 , " & formSpecSheet.txtWeight.Text & " , " & spec
End If



'reopen the file for output
Open FilePath For Output As #1
Print #1, strData
Close #1

End Sub

Private Sub checkFG()
With formSpecSheet

If .checkFG.Value = "False" Then
    .frameInstalledMemory.Visible = False
    .addMemory.Visible = False
    .removeMemory.Visible = False
    .frameMemorySpeed.Visible = True
    .frameMemoryType.Visible = True
    .Image18.Visible = True
Else
    .frameInstalledMemory.Visible = True
    .addMemory.Visible = True
    .removeMemory.Visible = True
    .frameMemorySpeed.Visible = False
    .frameMemoryType.Visible = False
    .Image18.Visible = False
End If

End With
End Sub

Private Sub ArchiveUnit()

''// store every unit tested to be later recalled by entering only serial number
Dim filename As String
Dim out As String

filename = UCase(formSpecSheet.txtSerial.Text) & ".csv"

If InStr(filename, "|") > 0 Or InStr(filename, "*") > 0 _
    Or InStr(filename, "<") > 0 Or InStr(filename, ":") > 0 Or InStr(filename, "/") > 0 Or InStr(filename, "\") > 0 Then
    MsgBox "Serial number contains invalid characters for archiving unit, please use only alpha-numeric characters."
    Exit Sub
End If



'ensure directory exists
If Dir(ActiveWorkbook.Path & "\archive", vbDirectory) = "" Then
    MkDir Path:=ActiveWorkbook.Path & "\archive"
End If
ChDir (ActiveWorkbook.Path & "\archive")


With formSpecSheet
    out = .txtISPF.Text & ", " & .labelDate.Caption & ", " & .dropCondition.Value & ", " & .dropBrand.Value _
     & ", " & .txtOtherBrand.Text & ", " & .txtSerial.Text & ", " & .txtModel.Text & ", " & .dropFormFactor.Value _
      & ", " & .spinCPUQTY.Value & ", " & .spinCPUCores.Value & ", " & .checkCPUHT.Value & ", " & .txtCPUSpeed.Text _
       & ", " & .dropCPUType.Value & ", " & .txtFSBSpeed.Text & ", " & .txtCPUName.Text & ", " & .dropMemorySize.Value _
        & ", " & .dropMemoryRating.Value & ", " & .dropMemoryType.Value & ", " & .dropMemorySpeed.Value & ", " & .txtWeight.Text _
         & ", " & .spinHDDQTY.Value & ", " & .txtHDDSize.Text & ", " & .dropHDDType.Value & ", " & .dropHDDRPM.Value _
          & ", " & .txtHDDSerial.Text & ", " & .dropVideo.Value & ", " & .txtVideoModel.Text & ", " & .txtVRAM.Text _
           & ", " & .dropOpticalDrive.Value & ", " & .CheckBox4.Value & ", " & .CheckBox1.Value & ", " & .CheckBox3.Value _
            & ", " & .txtLCDSize.Text & ", " & .CheckBox8.Value & ", " & .CheckBox5.Value & ", " & .CheckBox6.Value _
             & ", " & .CheckBox7.Value & ", " & .CheckBox9.Value & ", " & .dropCOA.Value & ", " & .optionOSNo.Value _
              & ", " & .optionOSYes.Value & ", " & Replace(.txtNotes.Text, ",", "") & ", " & .checkAccessories0.Value & ", " & .checkAccessories1.Value _
               & ", " & .checkAccessories2.Value & ", " & .checkAccessories3.Value & ", " & .checkAccessories4.Value _
                & ", " & .checkAccessories5.Value & ", " & .checkAccessories6.Value & ", " & .checkAccessories7.Value _
                 & ", " & .checkAccessories8.Value & ", " & .dropDamage.Value & ", " & .USB.Text & ", " & .Ethernet.Text _
                  & ", " & .Modem.Text & ", " & .VGA.Text & ", " & .DVI.Text & ", " & .SVideo.Text & ", " & .PS2.Text _
                   & ", " & .Audio.Text & ", " & .eSATAp.Text & ", " & .Serial.Text & ", " & .Parallel.Text & ", " & .PCMCIA.Text _
                    & ", " & .SDCard.Text & ", " & .Firewire.Text & ", " & .eSATA.Text & ", " & .HDMI.Text & ", " & .SCSI.Text _
                     & ", " & .DisplayPort.Text & ", " & .lblVersion.Caption & ", " & .txtTester.Text & ", " & .spinCaddyQTY.Value _
                     & ", " & .checkCaddyNA
    


End With

' Store current form info into save file
Open filename For Output As #1
Print #1, out
Close #1


End Sub
Sub getFromArchive(filename As String)

Dim inputStr As String
Dim arr As Variant
Dim str As String


Open filename For Input As #1
Line Input #1, inputStr
Close #1

arr = Split(inputStr, ", ")



With formSpecSheet
'MsgBox "found " & UBound(arr) & " fields"
    If .lblVersion.Caption >= arr(70) Then
        .txtISPF.Text = arr(0)
        
        'dropcondition2
        str = arr(2)
        For i = 0 To .dropCondition.ListCount - 1
            If (.dropCondition.List(i) = str) Then
                .dropCondition.ListIndex = i
            End If
        Next i
        
        'dropbrand3
        str = arr(3)
        For i = 0 To .dropBrand.ListCount - 1
            If (.dropBrand.List(i) = str) Then
                .dropBrand.ListIndex = i
            End If
        Next i
        
        .txtOtherBrand.Text = arr(4)
        
        .txtModel.Text = arr(6)
        
        'dropformfactor7
        str = arr(7)
        For i = 0 To .dropFormFactor.ListCount - 1
            If (.dropFormFactor.List(i) = str) Then
                .dropFormFactor.ListIndex = i
            End If
        Next i
        
        .spinCPUQTY.Value = arr(8)
        
        .spinCPUCores.Value = arr(9)
        
        .checkCPUHT.Value = arr(10)
        
        .txtCPUSpeed.Text = arr(11)
        
        'dropcputype12
        str = arr(12)
        For i = 0 To .dropCPUType.ListCount - 1
            If (.dropCPUType.List(i) = str) Then
                .dropCPUType.ListIndex = i
            End If
        Next i
        
        .txtFSBSpeed.Text = arr(13)
        
        .txtCPUName.Text = arr(14)
        
        'dropmemorysize15
        str = arr(15)
        For i = 0 To .dropMemorySize.ListCount - 1
            If (.dropMemorySize.List(i) = str) Then
                .dropMemorySize.ListIndex = i
            End If
        Next i
        
        'dropmemoryrating16
        str = arr(16)
        For i = 0 To .dropMemoryRating.ListCount - 1
            If (.dropMemoryRating.List(i) = str) Then
                .dropMemoryRating.ListIndex = i
            End If
        Next i
        
        'dropmemorytype17
        str = arr(17)
        For i = 0 To .dropMemoryType.ListCount - 1
            If (.dropMemoryType.List(i) = str) Then
                .dropMemoryType.ListIndex = i
            End If
        Next i
        
        'dropmemoryspeed18
        str = arr(18)
        For i = 0 To .dropMemorySpeed.ListCount - 1
            If (.dropMemorySpeed.List(i) = str) Then
                .dropMemorySpeed.ListIndex = i
            End If
        Next i
        
        .txtWeight.Text = arr(19)
        
        .spinHDDQTY.Value = arr(20)
        
        .txtHDDSize.Text = arr(21)
        
        'drophddtype22
        str = arr(22)
        For i = 0 To .dropHDDType.ListCount - 1
            If (.dropHDDType.List(i) = str) Then
                .dropHDDType.ListIndex = i
            End If
        Next i
        
        'drophddrpm23
        str = arr(23)
        For i = 0 To .dropHDDRPM.ListCount - 1
            If (.dropHDDRPM.List(i) = str) Then
                .dropHDDRPM.ListIndex = i
            End If
        Next i
        
        .txtHDDSerial.Text = arr(24)
        
        'dropvideo25
        str = arr(25)
        For i = 0 To .dropVideo.ListCount - 1
            If (.dropVideo.List(i) = str) Then
                .dropVideo.ListIndex = i
            End If
        Next i
        
        .txtVideoModel.Text = arr(26)
        
        .txtVRAM.Text = arr(27)
        
        'dropopticaldrive28
        str = arr(28)
        For i = 0 To .dropOpticalDrive.ListCount - 1
            If (.dropOpticalDrive.List(i) = str) Then
                .dropOpticalDrive.ListIndex = i
            End If
        Next i
        
        'none
        .CheckBox4.Value = arr(29)
        'fdd
        .CheckBox1.Value = arr(30)
        'tape
        .CheckBox3.Value = arr(31)
        
        .txtLCDSize.Text = arr(32)
        
        '8, 5, 6 , 7, 9
        .CheckBox8.Value = arr(33)
        .CheckBox5.Value = arr(34)
        .CheckBox6.Value = arr(35)
        .CheckBox7.Value = arr(36)
        .CheckBox9.Value = arr(37)
        
        'dropwindowscoa38
        str = arr(38)
        For i = 0 To .dropCOA.ListCount - 1
            If (.dropCOA.List(i) = str) Then
                .dropCOA.ListIndex = i
            End If
        Next i
        
        .optionOSNo.Value = arr(39)
        .optionOSYes.Value = arr(40)
        
        .txtNotes.Text = arr(41)
        
        .checkAccessories0.Value = arr(42)
        .checkAccessories1.Value = arr(43)
        .checkAccessories2.Value = arr(44)
        .checkAccessories3.Value = arr(45)
        .checkAccessories4.Value = arr(46)
        .checkAccessories5.Value = arr(47)
        .checkAccessories6.Value = arr(48)
        .checkAccessories7.Value = arr(49)
        .checkAccessories8.Value = arr(50)
        
        'dropdamage51
        str = arr(51)
        For i = 0 To .dropDamage.ListCount - 1
            If (.dropDamage.List(i) = str) Then
                .dropDamage.ListIndex = i
            End If
        Next i
        
        .USB.Text = arr(52)
        .Ethernet.Text = arr(53)
        .Modem.Text = arr(54)
        .VGA.Text = arr(55)
        .DVI.Text = arr(56)
        .SVideo.Text = arr(57)
        .PS2.Text = arr(58)
        .Audio.Text = arr(59)
        .eSATAp.Text = arr(60)
        .Serial.Text = arr(61)
        .Parallel.Text = arr(62)
        .PCMCIA.Text = arr(63)
        .SDCard.Text = arr(64)
        .Firewire.Text = arr(65)
        .eSATA.Text = arr(66)
        .HDMI.Text = arr(67)
        .SCSI.Text = arr(68)
        .DisplayPort.Text = arr(69)
        
        'set tester field if the archive file contains it
        If UBound(arr) > 70 Then
            .txtTester.Text = arr(71)
        End If
        
        'set caddy qty and checkbox
        If UBound(arr) > 72 Then
            .spinCaddyQTY.Value = arr(72)
            .checkCaddyNA.Value = arr(73)
        End If
        
        
    
    
    Else
    MsgBox "Save file version (" & arr(70) & ") doesn't match program version (" & .lblVersion.Caption & "), unable to load."
    End If
    

End With


End Sub

Sub importQR(desc As String)
''// WIP -- Will be used for populating spec sheet by scanning a printed label's barcode
    Dim str As String
    Dim arr(1 To 49) As String
    Dim start As Integer
    Dim finish As Integer
    Dim format As Boolean
    Dim havehdd As Boolean
    
    For i = 1 To UBound(arr)
        arr(i) = ""
    Next i
        
    
    If InStr(1, desc, "System Info", vbTextCompare) > 0 Then
        ' new label format
        format = True
    Else
        ' old label format
        format = False
    End If
    
    
    ' new label format
    If format = True Then
    
    finish = 1
    havehdd = False
    
    'brand
    start = InStr(finish, desc, "Brand: ", vbTextCompare)
    finish = InStr(start, desc, "</li>", vbTextCompare)
    str = Mid(desc, start, (finish - start))
    arr(4) = Split(str, ": ")(1)
    'model
    start = InStr(finish, desc, "Model: ", vbTextCompare)
    finish = InStr(start, desc, "</li>", vbTextCompare)
    str = Mid(desc, start, (finish - start))
    arr(6) = Split(str, ": ")(1)
    'form factor
    start = InStr(finish, desc, "Form Factor: ", vbTextCompare)
    finish = InStr(start, desc, "</li>", vbTextCompare)
    str = Mid(desc, start, (finish - start))
    arr(7) = Split(str, ": ")(1)
    'condition
    start = InStr(finish, desc, "Condition: ", vbTextCompare)
    finish = InStr(start, desc, "</li>", vbTextCompare)
    str = Mid(desc, start, (finish - start))
    arr(3) = Split(str, ": ")(1)
    'cpu qty
    start = InStr(finish, desc, "QTY: ", vbTextCompare)
    finish = InStr(start, desc, "</li>", vbTextCompare)
    str = Mid(desc, start, (finish - start))
    arr(8) = Split(str, ": ")(1)
    'cpu type
    start = InStr(finish, desc, "Type: ", vbTextCompare)
    finish = InStr(start, desc, "</li>", vbTextCompare)
    str = Mid(desc, start, (finish - start))
    arr(12) = Split(str, ": ")(1)
    'cpu series
    start = InStr(finish, desc, "Series: ", vbTextCompare)
    finish = InStr(start, desc, "</li>", vbTextCompare)
    str = Mid(desc, start, (finish - start))
    arr(44) = Split(str, ": ")(1)
    'cpu cores
    start = InStr(finish, desc, "Cores: ", vbTextCompare)
    finish = InStr(start, desc, "</li>", vbTextCompare)
    str = Mid(desc, start, (finish - start))
    If InStr(1, str, "w", vbTextCompare) > 0 Then
        str = Left(str, InStr(1, str, "w", vbTextCompare) - 2)
        arr(45) = True
    Else
        arr(45) = False
    End If
    arr(9) = Split(str, ": ")(1)
    'cpu speed
    start = InStr(finish, desc, "Speed: ", vbTextCompare)
    finish = InStr(start, desc, "</li>", vbTextCompare)
    str = Mid(desc, start, (finish - start))
    str = Left(str, InStr(1, str, "Hz", vbTextCompare) - 2)
    arr(10) = Split(str, ": ")(1)
    'bus speed
    start = InStr(finish, desc, "Bus Speed: ", vbTextCompare)
    finish = InStr(start, desc, "</li>", vbTextCompare)
    str = Mid(desc, start, (finish - start))
    arr(11) = Split(str, ": ")(1)
    'memory size
    start = InStr(finish, desc, "Size: ", vbTextCompare)
    finish = InStr(start, desc, "</li>", vbTextCompare)
    str = Mid(desc, start, (finish - start))
    arr(14) = Split(str, ": ")(1)
    'memory type
    start = InStr(finish, desc, "Type: ", vbTextCompare)
    finish = InStr(start, desc, "</li>", vbTextCompare)
    str = Mid(desc, start, (finish - start))
    arr(13) = Split(str, ": ")(1)
    'memory rating
    start = InStr(finish, desc, "Rating: ", vbTextCompare)
    finish = InStr(start, desc, "</li>", vbTextCompare)
    str = Mid(desc, start, (finish - start))
    arr(25) = Split(str, ": ")(1)
    'memory speed
    start = InStr(finish, desc, "Speed: ", vbTextCompare)
    finish = InStr(start, desc, "</li>", vbTextCompare)
    str = Mid(desc, start, (finish - start))
    arr(15) = Split(str, ": ")(1)
    'hdd qty
    start = InStr(finish, desc, "QTY: ", vbTextCompare)
    finish = InStr(start, desc, "</li>", vbTextCompare)
    str = Mid(desc, start, (finish - start))
    arr(46) = Split(str, ": ")(1)
    If arr(46) > 0 Then 'we have HDDs
    
    havehdd = True
    'hdd size
    start = InStr(finish, desc, "Size: ", vbTextCompare)
    finish = InStr(start, desc, "</li>", vbTextCompare)
    str = Mid(desc, start, (finish - start))
    str = Left(str, Len(str) - 2) ' chop out gb and TB
    arr(47) = Split(str, ": ")(1)
    
    'hdd interface
    start = InStr(finish, desc, "Interface: ", vbTextCompare)
    finish = InStr(start, desc, "</li>", vbTextCompare)
    str = Mid(desc, start, (finish - start))
    arr(48) = Split(str, ": ")(1)
    
    'hdd rpm
    start = InStr(finish, desc, "RPM: ", vbTextCompare)
    finish = InStr(start, desc, "</li>", vbTextCompare)
    str = Mid(desc, start, (finish - start))
    arr(49) = Split(str, ": ")(1)
    
    End If
    
    
    'optional
    'video
    'start = InStr(finish, desc, "Video: ", vbTextCompare)
    'finish = InStr(start, desc, "</li>", vbTextCompare)
    'str = Mid(desc, start, (finish - start))
    'drives
    'start = InStr(finish, desc, "Drives: ", vbTextCompare)
    'finish = InStr(start, desc, "</li>", vbTextCompare)
    'str = Mid(desc, start, (finish - start))
    
    'network
    'start = InStr(finish, desc, "Network: ", vbTextCompare)
    'finish = InStr(start, desc, "</li>", vbTextCompare)
    'str = Mid(desc, start, (finish - start))
    
    'ports
    'start = InStr(finish, desc, "Ports: ", vbTextCompare)
    'finish = InStr(start, desc, "</li>", vbTextCompare)
    'str = Mid(desc, start, (finish - start))
    
    'accessories
    'start = InStr(finish, desc, "Accessories: ", vbTextCompare)
    'finish = InStr(start, desc, "</li>", vbTextCompare)
    'str = Mid(desc, start, (finish - start))
    
    'notes
    start = InStr(finish, desc, "Notes: ", vbTextCompare)
    finish = InStr(start, desc, "</li>", vbTextCompare)
    str = Mid(desc, start, (finish - start))
    
    Else 'Old label format
    ' don't support old labels yet
    Exit Sub
    
    End If
    
' begin filling in sheet with filled array info
With formSpecSheet
'MsgBox "found " & UBound(arr) & " fields"
        If IsNumeric(arr(3)) Then
            .dropCondition.ListIndex = arr(3)
        Else
            str = arr(3)
            For i = 0 To .dropCondition.ListCount - 1
                If (.dropCondition.List(i) = str) Then
                    .dropCondition.ListIndex = i
                End If
            Next i
        End If
    'check for other brands correctly
    If IsNumeric(arr(4)) Then
        .dropBrand.ListIndex = arr(4)
    Else
        .dropBrand.ListIndex = 12
        .txtOtherBrand.Text = arr(4)
    End If
    .txtModel.Text = arr(6)
    
    'dropformfactor
    If IsNumeric(arr(7)) Then
        .dropFormFactor.ListIndex = arr(7)
    Else
        str = arr(7)
        For i = 0 To .dropFormFactor.ListCount - 1
            If (.dropFormFactor.List(i) = str) Then
                .dropFormFactor.ListIndex = i
            End If
        Next i
    End If
    
    .spinCPUQTY.Value = arr(8)
    .spinCPUCores.Value = arr(9)
    .txtCPUSpeed.Text = arr(10)
    .txtFSBSpeed.Text = arr(11)
    
    'dropcputype
    If IsNumeric(arr(12)) Then
        .dropCPUType.ListIndex = arr(12)
    Else
        str = arr(12)
        For i = 0 To .dropCPUType.ListCount - 1
            If (.dropCPUType.List(i) = str) Then
                .dropCPUType.ListIndex = i
            End If
        Next i
    End If
    
    'dropmemorytype
    If IsNumeric(arr(13)) Then
        .dropMemoryType.ListIndex = arr(13)
    Else
        str = arr(13)
        For i = 0 To .dropMemoryType.ListCount - 1
            If (.dropMemoryType.List(i) = str) Then
                .dropMemoryType.ListIndex = i
            End If
        Next i
    End If
    'dropmemorysize
    If IsNumeric(arr(14)) Then
        .dropMemorySize.ListIndex = arr(14)
    Else
        str = arr(14)
        For i = 0 To .dropMemorySize.ListCount - 1
            If (.dropMemorySize.List(i) = str) Then
                .dropMemorySize.ListIndex = i
            End If
        Next i
    End If
    'dropmemoryspeed
    If IsNumeric(arr(15)) Then
        .dropMemorySpeed.ListIndex = arr(15)
    Else
        str = arr(15)
        For i = 0 To .dropMemorySpeed.ListCount - 1
            If (.dropMemorySpeed.List(i) = str) Then
                .dropMemorySpeed.ListIndex = i
            End If
        Next i
    End If
    'dropopticaldrive
    If IsNumeric(arr(19)) Then
        .dropOpticalDrive.ListIndex = arr(19)
    Else
        str = arr(19)
        For i = 0 To .dropOpticalDrive.ListCount - 1
            If (.dropOpticalDrive.List(i) = str) Then
                .dropOpticalDrive.ListIndex = i
            End If
        Next i
    End If
    'dropvideo
    If IsNumeric(arr(20)) Then
        .dropVideo.ListIndex = arr(20)
    Else
        str = arr(20)
        For i = 0 To .dropVideo.ListCount - 1
            If (.dropVideo.List(i) = str) Then
                .dropVideo.ListIndex = i
            End If
        Next i
    End If
    'dropcoa
    If IsNumeric(arr(22)) Then
        .dropCOA.ListIndex = arr(22)
    Else
        str = arr(22)
        For i = 0 To .dropCOA.ListCount - 1
            If (.dropCOA.List(i) = str) Then
                .dropCOA.ListIndex = i
            End If
        Next i
    End If
    .txtNotes.Text = arr(23)
    .txtLCDSize.Text = arr(24)
    'dropmemoryrating
    If IsNumeric(arr(25)) Then
        .dropMemoryRating.ListIndex = arr(25)
    Else
        str = arr(25)
        For i = 0 To .dropMemoryRating.ListCount - 1
            If (.dropMemoryRating.List(i) = str) Then
                .dropMemoryRating.ListIndex = i
            End If
        Next i
    End If
    .USB.Text = arr(26)
    .Ethernet.Text = arr(27)
    .Modem.Text = arr(28)
    .VGA.Text = arr(29)
    .DVI.Text = arr(30)
    .SVideo.Text = arr(31)
    .PS2.Text = arr(32)
    .Audio.Text = arr(33)
    .eSATAp.Text = arr(34)
    .Serial.Text = arr(35)
    .Parallel.Text = arr(36)
    .PCMCIA.Text = arr(37)
    .SDCard.Text = arr(38)
    .Firewire.Text = arr(39)
    .eSATA.Text = arr(40)
    .HDMI.Text = arr(41)
    .SCSI.Text = arr(42)
    .DisplayPort.Text = arr(43)
    .txtCPUName.Text = arr(44)
    .checkCPUHT.Value = arr(45)
    .spinHDDQTY.Value = arr(46)
    
    If havehdd = True Then
    .txtHDDSize.Text = arr(47)
    
        'drophddtype
    If IsNumeric(arr(48)) Then
        .dropHDDType.ListIndex = arr(48)
    Else
        str = arr(48)
        For i = 0 To .dropHDDType.ListCount - 1
            If (.dropHDDType.List(i) = str) Then
                .dropHDDType.ListIndex = i
            End If
        Next i
    End If
    
        'drophddrpm
    str = arr(49)
    For i = 0 To .dropHDDRPM.ListCount - 1
        If (.dropHDDRPM.List(i) = str) Then
            .dropHDDRPM.ListIndex = i
        End If
    Next i
    
    End If
    
    
    
End With
    
End Sub
