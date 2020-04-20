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
h = 14.2 * ((Range("G7")) / (Range("G8") ^ 2))
Range("C17").Value = h
Range("C18").Value = "HK"

ElseIf Sheet1.ComboBox1.Value = "Vickers" Then
h = (1.854 * Range("G7")) / (Range("G8") ^ 2)
Range("C17").Value = h
Range("C18").Value = "HV"

ElseIf Sheet1.ComboBox1.Value = "Brinell" Then
pi = Application.WorksheetFunction.pi()
num = 2 * Range("G7")
denom = pi * Range("G9") * (Range("G9") - Sqr((Range("G9") ^ 2) - (Range("G8") ^ 2)))
h = num / denom
Range("C17").Value = h
Range("C18").Value = "BHN"

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
