Private Sub CommandButton1_Click()
'заполняет элементы массива рандомными положительными значениями от 0 до 100'
For i = 1 To 30
Cells(1, i) = Int((100 * Rnd) + 0)
Next i
End Sub

Private Sub CommandButton2_Click()
'находит и выводит произведение элементов массива, которые являются четными и не оканчиваются на 0'
p = 1
For i = 1 To 30
If Cells(1, i) Mod 2 = 0 And Cells(1, i) Mod 10 <> 0 Then
p = p * Cells(1, i)
End If
Next i
MsgBox (p)
End Sub

Private Sub CommandButton3_Click()
'закрывает форму'
UserForm1.Hide
End Sub