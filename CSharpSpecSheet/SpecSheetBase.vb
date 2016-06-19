Private Sub addMemory_Click()
If formSpecSheet.dropMemorySize.Value <> "" And formSpecSheet.dropMemorySize.Value <> "None" And formSpecSheet.dropMemoryRating.Value <> "" And formSpecSheet.dropMemoryRating.Value <> "N/A" Then
formSpecSheet.memoryListbox.AddItem (formSpecSheet.dropMemorySize.Value & " " & formSpecSheet.dropMemoryType.Value & " " & formSpecSheet.dropMemoryRating.Value)
End If
End Sub

Private Sub btnLoad_Click()
Run "LoadClick"
End Sub

Private Sub btnSave_Click()
Run "SaveClick"
End Sub

Private Sub btnSync_Click()
'open the original text file to read the lines
FilePath = ActiveWorkbook.Path & "\autoupdater.bat"

If Dir(FilePath) <> "" Then ' only run updater if it exists in directory
    ChDir (ActiveWorkbook.Path)
    Call Shell(FilePath, vbNormalFocus)
End If
End Sub

Private Sub btnUpdate_Click()
Dim oServ As Object
Dim cProc As Variant
Dim oProc As Object

'open the original text file to read the lines
FilePath = ActiveWorkbook.Path & "\autoupdater.bat"

If Dir(FilePath) <> "" Then ' only run updater if it exists in directory
    ChDir (ActiveWorkbook.Path)
    Call Shell(FilePath, vbNormalFocus)
    ' kill excel process so it can update correctly
    Set oServ = GetObject("winmgmts:")
    Set cProc = oServ.ExecQuery("Select * from Win32_Process")

    For Each oProc In cProc

        'Rename EXCEL.EXE in the line below with the process that you need to Terminate.
        'NOTE: It is 'case sensitive

        If oProc.Name = "EXCEL.EXE" Then
        errReturnCode = oProc.Terminate()
        End If
    Next
End If
End Sub

Private Sub buttonQRInput_Click()
    formQRInput.Show
End Sub

Private Sub checkAccessories0_Click()
Select Case checkAccessories0.Value
Case False
If checkAccessories1.Value = False And checkAccessories2.Value = False And checkAccessories3.Value = False And checkAccessories4.Value = False And checkAccessories5.Value = False And checkAccessories6.Value = False And checkAccessories7.Value = False And checkAccessories8.Value = False Then
checkAccessories0.Value = True
Else
checkAccessories0.Value = False
End If
Case True
checkAccessories0.Value = True
checkAccessories1.Value = False
checkAccessories2.Value = False
checkAccessories3.Value = False
checkAccessories4.Value = False
checkAccessories5.Value = False
checkAccessories6.Value = False
checkAccessories7.Value = False
checkAccessories8.Value = False
End Select
End Sub
Private Sub checkAccessories1_Click()
If checkAccessories1.Value = False And checkAccessories2.Value = False And checkAccessories3.Value = False And checkAccessories4.Value = False And checkAccessories5.Value = False And checkAccessories6.Value = False And checkAccessories7.Value = False And checkAccessories8.Value = False Then
checkAccessories0.Value = True
Else
checkAccessories0.Value = False
End If
End Sub
Private Sub checkAccessories2_Click()
If checkAccessories1.Value = False And checkAccessories2.Value = False And checkAccessories3.Value = False And checkAccessories4.Value = False And checkAccessories5.Value = False And checkAccessories6.Value = False And checkAccessories7.Value = False And checkAccessories8.Value = False Then
checkAccessories0.Value = True
Else
checkAccessories0.Value = False
End If
End Sub
Private Sub checkAccessories3_Click()
If checkAccessories1.Value = False And checkAccessories2.Value = False And checkAccessories3.Value = False And checkAccessories4.Value = False And checkAccessories5.Value = False And checkAccessories6.Value = False And checkAccessories7.Value = False And checkAccessories8.Value = False Then
checkAccessories0.Value = True
Else
checkAccessories0.Value = False
End If
End Sub
Private Sub checkAccessories4_Click()
If checkAccessories1.Value = False And checkAccessories2.Value = False And checkAccessories3.Value = False And checkAccessories4.Value = False And checkAccessories5.Value = False And checkAccessories6.Value = False And checkAccessories7.Value = False And checkAccessories8.Value = False Then
checkAccessories0.Value = True
Else
checkAccessories0.Value = False
End If
End Sub
Private Sub checkAccessories5_Click()
If checkAccessories1.Value = False And checkAccessories2.Value = False And checkAccessories3.Value = False And checkAccessories4.Value = False And checkAccessories5.Value = False And checkAccessories6.Value = False And checkAccessories7.Value = False And checkAccessories8.Value = False Then
checkAccessories0.Value = True
Else
checkAccessories0.Value = False
End If
End Sub
Private Sub checkAccessories6_Click()
If checkAccessories1.Value = False And checkAccessories2.Value = False And checkAccessories3.Value = False And checkAccessories4.Value = False And checkAccessories5.Value = False And checkAccessories6.Value = False And checkAccessories7.Value = False And checkAccessories8.Value = False Then
checkAccessories0.Value = True
Else
checkAccessories0.Value = False
End If
End Sub
Private Sub checkAccessories7_Click()
If checkAccessories1.Value = False And checkAccessories2.Value = False And checkAccessories3.Value = False And checkAccessories4.Value = False And checkAccessories5.Value = False And checkAccessories6.Value = False And checkAccessories7.Value = False And checkAccessories8.Value = False Then
checkAccessories0.Value = True
Else
checkAccessories0.Value = False
End If
End Sub
Private Sub checkAccessories8_Click()
If checkAccessories1.Value = False And checkAccessories2.Value = False And checkAccessories3.Value = False And checkAccessories4.Value = False And checkAccessories5.Value = False And checkAccessories6.Value = False And checkAccessories7.Value = False And checkAccessories8.Value = False Then
checkAccessories0.Value = True
Else
checkAccessories0.Value = False
End If
End Sub

Private Sub checkFG_Click()
Run "checkFG"

End Sub

Private Sub dropFormFactor_Change()
If dropFormFactor.Text = "Laptop" Or dropFormFactor.Text = "Tablet" Or dropFormFactor.Text = "All-in-One" Or dropFormFactor.Text = "Tablet Laptop" Then
    txtLCDSize.Enabled = True
    txtLCDSize.Locked = False
    txtLCDSize.Text = ""
    txtLCDSize.TabStop = True
Else
    txtLCDSize.Enabled = False
    txtLCDSize.Locked = True
    txtLCDSize.Text = "N/A"
    txtLCDSize.TabStop = False
End If
End Sub

Private Sub dropMemoryRating_Change()
'update memory speed automatically according to memory rating
dropMemorySpeed.ListIndex = dropMemoryRating.ListIndex
'update memory type based on memory rating
If InStr(dropMemoryRating.Value, "PC-") > 0 Then
    dropMemoryType.ListIndex = 0
ElseIf InStr(dropMemoryRating.Value, "PC2-") > 0 Then
    dropMemoryType.ListIndex = 1
ElseIf InStr(dropMemoryRating.Value, "PC3-") > 0 Then
    dropMemoryType.ListIndex = 2
ElseIf InStr(dropMemoryRating.Value, "PC4-") > 0 Then
    dropMemoryType.ListIndex = 3
Else
    dropMemoryType.ListIndex = 6
End If
End Sub


Private Sub removeMemory_Click()
For intCount = formSpecSheet.memoryListbox.ListCount - 1 To 0 Step -1
    If formSpecSheet.memoryListbox.Selected(intCount) Then formSpecSheet.memoryListbox.RemoveItem (intCount)
Next intCount
End Sub

Private Sub spinCaddyQTY_Change()
Run "spinCaddyQTYchange"
End Sub

Private Sub spinCPUCores_Change()
Run "spinCPUCoreschange"
End Sub

Private Sub spinCPUQTY_Change()
Run "spinCPUQTYchange"
End Sub

Private Sub txtCPUQTY_Change()

End Sub
Private Sub txtSerial_Enter()
With formSpecSheet.txtSerial
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub
Private Sub txtSerial_Change()
'find and load log file if it exists
If InStr(formSpecSheet.txtSerial.Text, "|") > 0 Or InStr(formSpecSheet.txtSerial.Text, "*") > 0 _
    Or InStr(formSpecSheet.txtSerial.Text, "<") > 0 Or InStr(formSpecSheet.txtSerial.Text, ":") > 0 Then
    Exit Sub
End If
If formSpecSheet.txtSerial.Text <> "" And Dir(ActiveWorkbook.Path & "\archive\" & UCase(formSpecSheet.txtSerial.Text) & ".csv") <> "" Then
    formSpecSheet.txtSerial.BackColor = RGB(153, 255, 153)
    getFromArchive (ActiveWorkbook.Path & "\archive\" & UCase(formSpecSheet.txtSerial.Text) & ".csv")
Else
    formSpecSheet.txtSerial.BackColor = RGB(255, 255, 255)
End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

If Application.Visible = False And CloseMode <> 1 Then
    Application.Quit
End If
End Sub

Private Sub btnAdmin_Click()
    Run "FormAdmin"
End Sub

Private Sub btnClearForm_Click()
Run "clearForm"
End Sub

Private Sub btnGenerate_Click()
Run "PrintLabelClick"
End Sub

Private Sub CheckBox1_Click()
If CheckBox1.Value = False And CheckBox3.Value = False Then
CheckBox4.Value = True
Else
CheckBox4.Value = False
End If
End Sub


Private Sub CheckBox3_Click()
If CheckBox1.Value = False And CheckBox3.Value = False Then
CheckBox4.Value = True
Else
CheckBox4.Value = False
End If
End Sub

Private Sub CheckBox4_Click()
Select Case CheckBox4.Value
Case False
If CheckBox1.Value = False And CheckBox3.Value = False Then
CheckBox4.Value = True
Else
CheckBox4.Value = False
End If
Case True
CheckBox1.Value = False
CheckBox3.Value = False
CheckBox4.Value = True
End Select
End Sub

Private Sub CheckBox5_Click()
If CheckBox5.Value = False And CheckBox6.Value = False And CheckBox7.Value And CheckBox9.Value = False = False Then
CheckBox8.Value = True
Else
CheckBox8.Value = False
End If
End Sub

Private Sub CheckBox6_Click()
If CheckBox5.Value = False And CheckBox6.Value = False And CheckBox7.Value = False And CheckBox9.Value = False Then
CheckBox8.Value = True
Else
CheckBox8.Value = False
End If
End Sub

Private Sub CheckBox7_Click()
If CheckBox5.Value = False And CheckBox6.Value = False And CheckBox7.Value = False And CheckBox9.Value = False Then
CheckBox8.Value = True
Else
CheckBox8.Value = False
End If
End Sub
Private Sub CheckBox9_Click()
If CheckBox5.Value = False And CheckBox6.Value = False And CheckBox7.Value = False And CheckBox9.Value = False Then
CheckBox8.Value = True
Else
CheckBox8.Value = False
End If
End Sub

Private Sub CheckBox8_Click()
Select Case CheckBox8.Value
Case False
If CheckBox5.Value = False And CheckBox6.Value = False And CheckBox7.Value = False And CheckBox9.Value = False Then
CheckBox8.Value = True
Else
CheckBox8.Value = False
End If
Case True
CheckBox5.Value = False
CheckBox6.Value = False
CheckBox7.Value = False
CheckBox9.Value = False
CheckBox8.Value = True
End Select
End Sub

Private Sub dropBrand_Change()
If dropBrand.Text = "Other:" Then
    txtOtherBrand.Enabled = True
    txtOtherBrand.Locked = False
    txtOtherBrand.Text = ""
    txtOtherBrand.TabStop = True
Else
    txtOtherBrand.Enabled = False
    txtOtherBrand.Locked = True
    txtOtherBrand.Text = "N/A"
    txtOtherBrand.TabStop = False
End If
End Sub


Private Sub dropVideo_Change()
If dropVideo.Text <> "None" And dropVideo.Text <> "Onboard" Then
    txtVideoModel.Enabled = True
    txtVideoModel.Locked = False
    txtVideoModel.Text = ""
    txtVideoModel.TabStop = True
    txtVRAM.Enabled = True
    txtVRAM.Locked = False
    txtVRAM.Text = ""
    txtVRAM.TabStop = True
Else
    txtVideoModel.Enabled = False
    txtVideoModel.Locked = True
    txtVideoModel.Text = "N/A"
    txtVideoModel.TabStop = False
    txtVRAM.Enabled = False
    txtVRAM.Locked = True
    txtVRAM.Text = "N/A"
    txtVRAM.TabStop = False
End If
End Sub
Private Sub dropCPUType_Change()
If dropCPUType.Text <> "N/A" Then
    txtCPUName.Enabled = True
    txtCPUName.Locked = False
    txtCPUName.Text = ""
    txtCPUName.TabStop = True
Else
    txtCPUName.Enabled = False
    txtCPUName.Locked = True
    txtCPUName.Text = "N/A"
    txtCPUName.TabStop = False
End If
End Sub

Private Sub labelDate_Click()
LDate = Date
labelDate.Caption = LDate
End Sub

Private Sub spinHDDQTY_Change()
Run "spinHDDQTYchange"
End Sub
