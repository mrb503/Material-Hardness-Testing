Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()

Sheet1.ComboBox1.Value = ""
Range("C17:C18").Value = ""
SpinButton1.Value = 0
SpinButton2.Value = 0
SpinButton3.Value = 0
Range("G7:G9").Value = ""



End Sub

Private Sub CommandButton2_Click()

If Sheet1.ComboBox1.Value = "Knoop" Then
If Range("G7").Value = "" Or Range("G9").Value = "" Then
MsgBox "Error! Missing one or more required values!"
Else
h = 14.2 * ((Range("G7")) / (Range("G9") ^ 2))
Range("C17").Value = h
Range("C18").Value = "HK"
End If

ElseIf Sheet1.ComboBox1.Value = "Vickers" Then
If Range("G7").Value = "" Or Range("G8").Value = "" Then
MsgBox "Error! Missing one or more required values!"
Else
h = (1.854 * Range("G7")) / (Range("G8") ^ 2)
Range("C17").Value = h
Range("C18").Value = "HV"
End If

ElseIf Sheet1.ComboBox1.Value = "Brinell" Then
If Range("G7").Value = "" Or Range("G8").Value = "" Or Range("G9").Value = "" Then
MsgBox "Error! Missing one or more required values!"
Else
pi = Application.WorksheetFunction.pi()
num = 2 * Range("G7").Value
denom = pi * Range("G9").Value * (Range("G9").Value - (Sqr((Range("G9").Value ^ 2) - (Range("G8").Value ^ 2))))
h = num / denom
Range("C17").Value = h
Range("C18").Value = "BHN"
End If

ElseIf Sheet1.ComboBox1.Value = "" Then
MsgBox "Error! No hardness test selected"

End If


End Sub

Private Sub SpinButton1_Change()

Range("G7").Value = SpinButton1.Value
SpinButton1.Min = 0
SpinButton1.SmallChange = 1

End Sub

Private Sub SpinButton2_Change()

Range("G8").Value = SpinButton2.Value
SpinButton2.Min = 0
SpinButton2.SmallChange = 1

End Sub

Private Sub SpinButton3_Change()

Range("G9").Value = SpinButton3.Value
SpinButton3.Min = 0
SpinButton3.SmallChange = 1

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
