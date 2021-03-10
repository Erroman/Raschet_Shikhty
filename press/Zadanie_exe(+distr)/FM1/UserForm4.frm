VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "РАСЧЕТ"
   ClientHeight    =   9240
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   11604
   OleObjectBlob   =   "UserForm4.dsx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Shlp, Sin As String
Private M, MemComp(4, 100), MemCompSel(4, 17) As Variant
Private r, X As Boolean
Private NumCmp, BispD(12), BispRW1(15), BispRW2(15) As Integer
Private Ps, Ns, Mp(17), Ms(17), Nd(17), Mps, Mss, Mpz, Msz, NumClc, NO, WO As Double

Private Sub CommandButton3_Click() 'ВЫХОД
UserForm4.Hide
End Sub
Private Sub CommandButton2_Click() 'БД КОМПОНЕНТОВ
X = True
UserForm5.Show
End Sub
Private Sub InP2()
Sin = InputBox(Shlp, "ВВОД ДАННЫХ", M)
 If IsNumeric(Sin) = True Then M = Sin
r = False
End Sub


Private Sub CommandButton5_Click() 'Переход к таблице задания
Dim gg, NNN, i, j, k, kr, krn As Integer
gg = 0
For i = 3 To FG5.Rows - 1
 If Val(FG5.TextMatrix(i, 4)) = 0 Or Val(FG5.TextMatrix(i, 3)) = 0 Then gg = 1
Next i
  If gg = 0 And FG5.Rows > 3 Then
 UserForm2.Label3.Caption = "Расчет № " & NumClc
UserForm3.FG1.Cols = 17
NNN = Ns + 3
UserForm2.Label4.Caption = NNN - 1
UserForm3.FG1.Rows = NNN
UserForm3.FG1.ColWidth(0) = 330
UserForm3.FG1.RowHeight(0) = 360
UserForm3.FG1.RowHeight(1) = 360
UserForm3.FG1.RowHeight(NNN - 1) = 360
UserForm3.FG1.Height = 140
UserForm3.FG1.Left = 6
UserForm3.FG1.Top = 20
UserForm3.FG1.Width = 610
UserForm3.FG2.Rows = 3

' Номера строк таблицы
    For i = 2 To UserForm3.FG1.Rows - 2
        UserForm3.FG1.TextMatrix(i, 0) = i - 1
        UserForm3.FG1.RowHeight(i) = 360
    Next i
        UserForm3.FG1.TextMatrix(1, 0) = "НО"
        UserForm3.FG1.TextMatrix(UserForm3.FG1.Rows - 1, 0) = "ВО"
        
' Номера столбцов таблицы
    For i = 1 To UserForm3.FG1.Cols - 1
        If i < 11 Then UserForm3.FG1.TextMatrix(0, i) = i
        'If i > 9 Then UserForm3.FG1.TextMatrix(0, i) = i
        UserForm3.FG1.ColWidth(i) = 710
        UserForm3.FG2.ColWidth(i) = 710
       If i = 16 Then UserForm3.FG1.ColWidth(i) = 860: UserForm3.FG2.ColWidth(i) = 860
    Next i
        UserForm3.FG1.TextMatrix(0, 16) = "Вес П"
        
   ' Обнуление области таблицы
For i = 1 To 16
  For j = 1 To NNN - 1
   UserForm3.FG1.TextMatrix(j, i) = Format(0, "#0.000")
  Next j
Next i
     
    ' Загрузка данных
 
  For j = 2 To NNN - 2
   kr = 10: k1 = 0: k2 = 0
For i = 3 To FG5.Rows - 1
    k = Val(FG5.TextMatrix(i, 4))
      If k < 11 Then
         UserForm3.FG1.TextMatrix(j, k) = Format(FG5.TextMatrix(i, 2), "#0.000")
      End If
      If k = 11 Then
          kr = k + k1 + Mid(FG5.TextMatrix(i, 4), 4, 1) - 1
          UserForm3.FG1.TextMatrix(j, kr) = Format(FG5.TextMatrix(i, 2), "#0.000")
          UserForm3.FG1.TextMatrix(0, kr) = FG5.TextMatrix(i, 4)
          k2 = k2 + 1
      End If
      If k = 12 Then
          kr = k + k2 + Mid(FG5.TextMatrix(i, 4), 4, 1) - 2
          UserForm3.FG1.TextMatrix(j, kr) = Format(FG5.TextMatrix(i, 2), "#0.000")
          UserForm3.FG1.TextMatrix(0, kr) = FG5.TextMatrix(i, 4)
          k1 = k1 + 1
      End If
      
Next i
  Next j

 '--
 kr = 10: k1 = 0: k2 = 0
For i = 3 To FG5.Rows - 1
    k = Val(FG5.TextMatrix(i, 4))
      If k < 11 Then
         UserForm3.FG2.TextMatrix(2, k) = FG5.TextMatrix(i, 6)
      End If
      If k = 11 Then
          kr = k + k1 + Mid(FG5.TextMatrix(i, 4), 4, 1) - 1
          UserForm3.FG2.TextMatrix(2, kr) = FG5.TextMatrix(i, 6)
          k2 = k2 + 1
      End If
      If k = 12 Then
          kr = k + k2 + Mid(FG5.TextMatrix(i, 4), 4, 1) - 2
          UserForm3.FG2.TextMatrix(2, kr) = FG5.TextMatrix(i, 6)
          k1 = k1 + 1
      End If
'Stop
Next i
'--


   If FG5.TextMatrix(1, 3) > 0 Then
    UserForm3.FG1.TextMatrix(1, FG5.TextMatrix(1, 4)) = Format(FG5.TextMatrix(1, 3), "#0.000")
   End If
   If FG5.TextMatrix(2, 3) > 0 Then
    UserForm3.FG1.TextMatrix(NNN - 1, FG5.TextMatrix(2, 4)) = Format(FG5.TextMatrix(2, 3), "#0.000")
   End If
 
  X = True
  UserForm3.Show
  
Else
Style = vbYes + vbCritical + vbDefaultButto2
        Msg = "ВВЕДЕНЫ НЕ ВСЕ ДАННЫЕ"
        Title = "СООБЩЕНИЯ": Help = "DEMO.HLP": Ctxt = 1000
        Responce = MsgBox(Msg, Style, Title, Help, Ctxt)
    End If

End Sub


Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
   If r = False Then M = TextBox1.Text: r = True
If KeyCode = 13 Then
        Shlp = "ВВЕДИТЕ НОМЕР РАСЧЕТА:"
            Call InP2
        TextBox1.Text = M
        NumClc = Val(TextBox1.Text)
  TextBox2.Enabled = True
  TextBox2.SetFocus
  TextBox2.SelStart = 0
  TextBox2.SelLength = 6
 End If
End Sub
Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If r = False Then M = TextBox2.Text: r = True
If KeyCode = 13 Then
        Shlp = "ВВЕДИТЕ ВЕС НИЖНЕГО ОБЕДНЕНИЯ:"
            Call InP2
        TextBox2.Text = M
        NO = Val(TextBox2.Text)
 If Label8.Visible = True Then
    Msz = NO + Mpz * Ns + WO
    Label8.Caption = Format(Msz, "#0.000")
 End If
  TextBox3.Enabled = True
  TextBox3.SetFocus
  TextBox3.SelStart = 0
  TextBox3.SelLength = 10
  FG5.TextMatrix(1, 3) = Format(NO, "#0.000")
 End If
End Sub
Private Sub TextBox3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If r = False Then M = TextBox3.Text: r = True
If KeyCode = 13 Then
        Shlp = "ВВЕДИТЕ ВЕС ПОРЦИИ:"
            Call InP2
        TextBox3.Text = Format(M, "#0.000")
        Mpz = TextBox3.Text
 If Label8.Visible = True Then
    Msz = NO + Mpz * Ns + WO
    Label8.Caption = Format(Msz, "#0.000")
 End If
    TextBox4.Enabled = True
  TextBox4.SetFocus
  TextBox4.SelStart = 0
  TextBox4.SelLength = 6
 End If
 FG5.TextMatrix(0, 2) = "            ВЕС        НА   " & Format(Mpz, "#0.00") & "  [кг]"
End Sub
Private Sub TextBox4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If r = False Then M = TextBox4.Text: r = True
If KeyCode = 13 Then
        Shlp = "ВВЕДИТЕ ЧИСЛО ССЫПОК:"
            Call InP2
         TextBox4.Text = M
        Ns = Val(TextBox4.Text)
 If FG5.Visible = True Then Call crFg5
 If Label8.Visible = True Then
    Msz = NO + Mpz * Ns + WO
    Label8.Caption = Format(Msz, "#0.000")
 End If
  TextBox5.Enabled = True
  TextBox5.SetFocus
  TextBox5.SelStart = 0
  TextBox5.SelLength = 6
 End If
End Sub

Sub crFg5()
Dim i As Integer
Mps = 0: Mss = 0
 For i = 1 To FG5.Rows - 1
    Mps = Mps + FG5.TextMatrix(i, 2)
    If i > 2 Then FG5.TextMatrix(i, 3) = Format(Ns * FG5.TextMatrix(i, 2), "#0.000")
    Mss = Mss + FG5.TextMatrix(i, 3)
 Next i
    Label1.Caption = "Суммарный вес:           " _
      & Format(Mps, "#0.000") & "              " _
        & Format(Mss, "#0.000")
End Sub
Private Sub TextBox5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If r = False Then M = TextBox5.Text: r = True
If KeyCode = 13 Then
        Shlp = "ВВЕДИТЕ ВЕС ВЕРХНЕГО ОБЕДНЕНИЯ:"
            Call InP2
        TextBox5.Text = M
        WO = Val(TextBox5.Text)
 If Label8.Visible = True Then
    Msz = NO + Mpz * Ns + WO
    Label8.Caption = Format(Msz, "#0.000")
 End If
    Label8.Visible = True
    Msz = NO + Mpz * Ns + WO
    Label8.Caption = Format(Msz, "#0.000")
    Msz = Label8.Caption
    FG5.TextMatrix(2, 3) = Format(WO, "#0.000")
    If FG5.Visible = False Then
        CommandButton1.Enabled = True
        CommandButton1.Visible = True
        CommandButton1.SetFocus
                            Else
        FG5.SetFocus
    End If
    
 End If
End Sub
Sub PrAddFG()
Dim i, j As Integer
If FG5.Rows < 18 Then
     FG5.Rows = FG5.Rows + 1
  For i = FG5.Row + 1 To FG5.Rows - 2
     j = FG5.Rows + FG5.Row - i
     FG5.TextMatrix(j, 1) = FG5.TextMatrix(j - 1, 1)
     FG5.TextMatrix(j, 2) = FG5.TextMatrix(j - 1, 2)
     FG5.TextMatrix(j, 3) = FG5.TextMatrix(j - 1, 3)
     FG5.TextMatrix(j, 4) = FG5.TextMatrix(j - 1, 4)
   Next i
     FG5.TextMatrix(FG5.Row + 1, 0) = FG5.Rows - 1
     FG5.TextMatrix(FG5.Row + 1, 1) = FG5.TextMatrix(FG5.Row, 1)
     FG5.Row = FG5.Rows - 1
     FG5.Col = 1
     FG5.CellBackColor = &H8000000F
    'FG5.Row = FG5.Rows - 1
    'FG5.Col = 1
    'FG5.CellBackColor = &H80000003
   For j = 1 To FG5.Rows - 1
     FG5.TextMatrix(j, 0) = j
   Next j
 FG5.TextMatrix(FG5.Rows - 1, 2) = Format(0, "0.000")
 FG5.TextMatrix(FG5.Rows - 1, 3) = Format(0, "0.000")
 FG5.TextMatrix(FG5.Rows - 1, 4) = "выбрать "
End If
 Call RTbl1
 Call CalcFG5
End Sub

Private Sub FG5_KeyPress(KeyAscii As Integer) 'Редактирование
 If KeyAscii = 13 Then
    If FG5.Col = 2 And FG5.Row > 2 Then Call PrInTblCl2
    If FG5.Col = 3 And FG5.Row > 2 Then Call PrInTblCl3
    If FG5.Col = 4 Then
         If (FG5.Row = 1 And FG5.TextMatrix(1, 3) > 0) Or _
            (FG5.Row = 2 And FG5.TextMatrix(2, 3) > 0) Or _
                FG5.Row > 2 Then Call PrInTblCl4
         End If
    End If
 'End If
End Sub
Sub PrInTblCl2()
Dim Msg, Style, Title, Help, Ctxt, Responce
    Dim Pmaxi, IDi As Integer
    Dim M As Variant
    Dim Pin, Rm As Double
    Dim S As String
        
      Pmaxi = MemCompSel(3, FG5.Row)
      IDi = MemCompSel(2, FG5.Row)
      
    If IsNumeric(FG5.TextMatrix(FG5.Row, FG5.Col)) = True Then _
     Rm = FG5.TextMatrix(FG5.Row, FG5.Col)
 M = Rm
 
 S = InputBox("ВВЕДИТЕ ВЕС МАТЕРИАЛА :", "ВВОД ДАННЫХ", M)
    If IsNumeric(S) = True Then
        Pin = S
       If Val(S) > Pmaxi Then
            Style = vbYes + vbCritical + vbDefaultButto2
        Msg = "ПРЕВЫШЕНО МАКСИМАЛЬНОЕ ЗНАЧЕНИЕ ВЕСА: Распределите на 2 дозатора"
        Title = "СООБЩЕНИЯ": Help = "DEMO.HLP": Ctxt = 1000
        Responce = MsgBox(Msg, Style, Title, Help, Ctxt)
       Else
          FG5.TextMatrix(FG5.Row, FG5.Col) = Format(S, "#0.000")
        If S > 0 Then FG5.CellBackColor = &HFFFF80             'метка Табл.1
        If S = 0 Then FG5.CellBackColor = &HE0E0E0
          FG5.TextMatrix(FG5.Row, 3) = Format(FG5.TextMatrix(FG5.Row, 2) * Ns, "#0.000")
          FG5.Col = 3
        If S > 0 Then FG5.CellBackColor = &HFFFF80
        If S = 0 Then FG5.CellBackColor = &HE0E0E0
          Call CalcFG5
          FG5.Col = 2
          FG5.SetFocus
       End If
    Else
      Style = vbYes + vbCritical + vbDefaultButto2
        Msg = "ОШИБКА ФОРМАТА ВВОДА"
        Title = "СООБЩЕНИЯ": Help = "DEMO.HLP": Ctxt = 1000
        Responce = MsgBox(Msg, Style, Title, Help, Ctxt)
    End If
End Sub

Sub PrInTblCl3()
Dim Msg, Style, Title, Help, Ctxt, Responce
Dim Pmaxi, IDi As Integer
    Dim M As Variant
    Dim Pin, Rm As Double
    Dim S As String
        
      Pmaxi = MemCompSel(3, FG5.Row) * Ns
      IDi = MemCompSel(2, FG5.Row)
      
    If IsNumeric(FG5.TextMatrix(FG5.Row, FG5.Col)) = True Then _
     Rm = FG5.TextMatrix(FG5.Row, FG5.Col)
 M = Rm
 S = InputBox("ВВЕДИТЕ ВЕС МАТЕРИАЛА :", "ВВОД ДАННЫХ", M)
    If IsNumeric(S) = True Then
        Pin = S
       If Val(S) > Pmaxi Then
            Style = vbYes + vbCritical + vbDefaultButto2
        Msg = "ПРЕВЫШЕНО МАКСИМАЛЬНОЕ ЗНАЧЕНИЕ ВЕСА: Распределите на 2 дозатора"
        Title = "СООБЩЕНИЯ": Help = "DEMO.HLP": Ctxt = 1000
        Responce = MsgBox(Msg, Style, Title, Help, Ctxt)
       Else
          FG5.TextMatrix(FG5.Row, FG5.Col) = Format(S, "#0.000")
        If S > 0 Then FG5.CellBackColor = &HFFFF80             'метка Табл.1
        If S = 0 Then FG5.CellBackColor = &HE0E0E0
          FG5.TextMatrix(FG5.Row, 2) = Format(FG5.TextMatrix(FG5.Row, 3) / Ns, "#0.000")
          FG5.Col = 2
        If S > 0 Then FG5.CellBackColor = &HFFFF80
        If S = 0 Then FG5.CellBackColor = &HE0E0E0
          FG5.Col = 3
          FG5.SetFocus
          Call CalcFG5
       End If
    Else
      Style = vbYes + vbCritical + vbDefaultButto2
        Msg = "ОШИБКА ФОРМАТА ВВОДА"
        Title = "СООБЩЕНИЯ": Help = "DEMO.HLP": Ctxt = 1000
        Responce = MsgBox(Msg, Style, Title, Help, Ctxt)
    End If
End Sub

Sub CalcFG5()
Dim i As Integer
Mps = 0: Mss = 0
 For i = 1 To FG5.Rows - 1
    Mps = Mps + FG5.TextMatrix(i, 2)
    Mss = Mss + FG5.TextMatrix(i, 3)
 Next i
    Label1.Caption = "Суммарный вес:           " _
      & Format(Mps, "#0.000") & "              " _
        & Format(Mss, "#0.000")
End Sub
Sub PrInTblCl4()      'Распределение материалов по дозаторам
Dim M, IDi As Integer
Dim S1 As String
  IDi = FG5.TextMatrix(FG5.Row, FG5.Col + 1)
Select Case IDi
     Case 1
If BispD(1) = 0 And BispD(2) = 0 Then S1 = "дозаторы №1 и №2": M = 1
If BispD(1) = 0 And BispD(2) = 1 Then S1 = "дозатор №1 ": M = 1
If BispD(1) = 1 And BispD(2) = 0 Then S1 = "дозатор №2 ": M = 2
If BispD(1) = 1 And BispD(2) = 1 Then S1 = "НЕТ СВОБОДНЫХ ДОЗАТОРОВ": M = 0
     Case 2
If BispD(3) = 0 And BispD(4) = 0 Then S1 = "дозаторы №3 и №4": M = 3
If BispD(3) = 0 And BispD(4) = 1 Then S1 = "дозатор №3 ": M = 3
If BispD(3) = 1 And BispD(4) = 0 Then S1 = "дозатор №4 ": M = 4
If BispD(3) = 1 And BispD(4) = 1 Then S1 = "НЕТ СВОБОДНЫХ ДОЗАТОРОВ": M = 0
     Case 3
If BispD(5) = 0 And BispD(4) = 0 And BispD(6) = 0 Then S1 = "дозаторы №4,№5 и №6": M = 5
If BispD(5) = 0 And BispD(4) = 1 And BispD(6) = 0 Then S1 = "дозаторы №5 и №6": M = 5
If BispD(5) = 1 And BispD(4) = 1 And BispD(6) = 0 Then S1 = "дозатор №6": M = 6
If BispD(5) = 1 And BispD(4) = 0 And BispD(6) = 0 Then S1 = "дозаторы №4 и №6": M = 6
If BispD(5) = 1 And BispD(4) = 0 And BispD(6) = 1 Then S1 = "дозатор №4": M = 4
If BispD(5) = 0 And BispD(4) = 1 And BispD(6) = 1 Then S1 = "дозатор №5": M = 5
If BispD(5) = 0 And BispD(4) = 0 And BispD(6) = 1 Then S1 = "дозаторы №4 и №5": M = 5
If BispD(5) = 1 And BispD(4) = 1 And BispD(6) = 1 Then S1 = "НЕТ СВОБОДНЫХ ДОЗАТОРОВ": M = 0
    Case 4
If BispD(5) = 0 And BispD(6) = 0 Then S1 = "дозаторы №5 и №6": M = 5
If BispD(5) = 0 And BispD(6) = 1 Then S1 = "дозатор №5 ": M = 5
If BispD(5) = 1 And BispD(6) = 0 Then S1 = "дозатор №6 ": M = 6
If BispD(5) = 1 And BispD(6) = 1 Then S1 = "НЕТ СВОБОДНЫХ ДОЗАТОРОВ": M = 0
    Case 5
   S1 = "дозатор №7": M = 7
    Case 6
If BispD(8) = 0 And BispD(9) = 0 And BispD(10) = 0 Then S1 = "дозаторы №8,№9 и №10": M = 8
If BispD(8) = 0 And BispD(9) = 1 And BispD(10) = 0 Then S1 = "дозаторы №8 и №10": M = 8
If BispD(8) = 1 And BispD(9) = 1 And BispD(10) = 0 Then S1 = "дозатор №10": M = 10
If BispD(8) = 1 And BispD(9) = 0 And BispD(10) = 0 Then S1 = "дозаторы №9 и №10": M = 9
If BispD(8) = 1 And BispD(9) = 0 And BispD(10) = 1 Then S1 = "дозатор №9": M = 9
If BispD(8) = 0 And BispD(9) = 1 And BispD(10) = 1 Then S1 = "дозатор №8": M = 8
If BispD(8) = 0 And BispD(9) = 0 And BispD(10) = 1 Then S1 = "дозаторы №8 и №9": M = 8
If BispD(8) = 1 And BispD(9) = 1 And BispD(10) = 1 Then S1 = "НЕТ СВОБОДНЫХ ДОЗАТОРОВ": M = 0
     Case 7
If BispD(7) = 0 And BispD(8) = 0 And BispD(9) = 0 And BispD(10) = 0 Then S1 = "дозаторы №7, №8,№9 и №10": M = 8
If BispD(7) = 0 And BispD(8) = 0 And BispD(9) = 1 And BispD(10) = 0 Then S1 = "дозаторы №7, №8 и №10": M = 8
If BispD(7) = 0 And BispD(8) = 1 And BispD(9) = 1 And BispD(10) = 0 Then S1 = "дозаторы №7, №10": M = 10
If BispD(7) = 0 And BispD(8) = 1 And BispD(9) = 0 And BispD(10) = 0 Then S1 = "дозаторы №7, №9 и №10": M = 9
If BispD(7) = 0 And BispD(8) = 1 And BispD(9) = 0 And BispD(10) = 1 Then S1 = "дозаторы №7, №9": M = 9
If BispD(7) = 0 And BispD(8) = 0 And BispD(9) = 1 And BispD(10) = 1 Then S1 = "дозаторы №7, №8": M = 8
If BispD(7) = 0 And BispD(8) = 0 And BispD(9) = 0 And BispD(10) = 1 Then S1 = "дозаторы №7, №8 и №9": M = 8
If BispD(7) = 0 And BispD(8) = 1 And BispD(9) = 1 And BispD(10) = 1 Then S1 = "дозатор №7": M = 7
If BispD(7) = 1 And BispD(8) = 0 And BispD(9) = 0 And BispD(10) = 0 Then S1 = "дозаторы №8,№9 и №10": M = 8
If BispD(7) = 1 And BispD(8) = 0 And BispD(9) = 1 And BispD(10) = 0 Then S1 = "дозаторы №8 и №10": M = 8
If BispD(7) = 1 And BispD(8) = 1 And BispD(9) = 1 And BispD(10) = 0 Then S1 = "дозатор №10": M = 10
If BispD(7) = 1 And BispD(8) = 1 And BispD(9) = 0 And BispD(10) = 0 Then S1 = "дозаторы №9 и №10": M = 9
If BispD(7) = 1 And BispD(8) = 1 And BispD(9) = 0 And BispD(10) = 1 Then S1 = "дозатор №9": M = 9
If BispD(7) = 1 And BispD(8) = 0 And BispD(9) = 1 And BispD(10) = 1 Then S1 = "дозатор №8": M = 8
If BispD(7) = 1 And BispD(8) = 0 And BispD(9) = 0 And BispD(10) = 1 Then S1 = "дозаторы №8 и №9": M = 8
If BispD(7) = 1 And BispD(8) = 1 And BispD(9) = 1 And BispD(10) = 1 Then S1 = "НЕТ СВОБОДНЫХ ДОЗАТОРОВ": M = 0
    Case 8
   S1 = "дозаторы №11, №12: m = 11"
    Case 9
   S1 = "дозаторы №9, №10, №11 и №12": M = 11
  End Select
   S = InputBox(S1, "ВЫБОР ДОЗАТОРА", M)
     If IsNumeric(S) = True Then
       If BispD(Val(S)) = 0 Then
           If Val(S) > 0 And Val(S) < 13 Then
   BispD(Val(FG5.TextMatrix(FG5.Row, 4))) = 0
   FG5.TextMatrix(FG5.Row, FG5.Col) = Format(S, "#0")
   FG5.CellBackColor = &HFFFF80
    If FG5.Row > 2 And Val(S) < 11 Then BispD(Val(S)) = 1
    End If
      If Val(S) = 11 Then
         BispRW1(FG5.Row - 2) = 1
         BispRW2(FG5.Row - 2) = 0
      End If
        If Val(S) = 12 Then
            BispRW2(FG5.Row - 2) = 1
            BispRW1(FG5.Row - 2) = 0
        End If
      
           Else
Msg = "УКАЗАННЫЙ ДОЗАТОР УЖЕ ИСПОЛЬЗУЕТСЯ"
Style = vbYes + vbCritical + vbDefaultButto2
Title = "СООБЩЕНИЯ": Help = "DEMO.HLP": Ctxt = 1000
Responce = MsgBox(Msg, Style, Title, Help, Ctxt)
      End If
     If Val(S) = 0 Then
   BispD(Val(FG5.TextMatrix(FG5.Row, 4))) = 0
   FG5.TextMatrix(FG5.Row, FG5.Col) = " Выбрать"
   FG5.CellBackColor = &H80000005
      End If
   End If
    j = 1: k = 1
  For i = 3 To FG5.Rows - 1
    If Val(FG5.TextMatrix(i, 4)) = 11 Then
        FG5.TextMatrix(i, 4) = "11-" & j
        j = j + 1
    End If
      If Val(FG5.TextMatrix(i, 4)) = 12 Then
          FG5.TextMatrix(i, 4) = "12-" & k
          k = k + 1
      End If
   Next i
End Sub

Sub RTbl1()
FG5.Height = 19 + FG5.Rows * 16
Label1.Move 130, 100 + FG5.Rows * 16
Label1.Caption = "Суммарный вес:                " & _
                Format(Mps, "#0.000") & _
"                      " & Format(Mss, "#0.000")
End Sub

Sub FormTabl()

'Ps = 140
'Ns = 26
 FG5.TextMatrix(0, 0) = "№№"
 FG5.TextMatrix(0, 1) = " КОМПОНЕНТ"
 FG5.TextMatrix(0, 2) = "            ВЕС        НА   " & Format(Mpz, "#0.00") & "  [кг]"
 FG5.TextMatrix(0, 3) = "             ВЕС              НА ЭЛЕКТРОД  [кг]"
 FG5.TextMatrix(0, 4) = "    ДОЗАТОР"
 
 
 'FG5.Rows = FG5.Rows + 1
 FG5.TextMatrix(1, 0) = ""
 FG5.TextMatrix(1, 1) = "Н.обеднение"
 FG5.TextMatrix(1, 2) = Format(0, "#0.000")
 FG5.TextMatrix(1, 3) = Format(NO, "#0.000")
 FG5.TextMatrix(1, 4) = "выбрать "
  If NO = 0 Then FG5.TextMatrix(1, 4) = ""
 FG5.TextMatrix(1, 5) = "1"
 FG5.TextMatrix(1, 6) = "90"
 FG5.Row = 1: FG5.Col = 1: FG5.CellBackColor = &HD0D0D0
 FG5.Row = 1: FG5.Col = 2: FG5.CellBackColor = &HD0D0D0
 FG5.Row = 1: FG5.Col = 3: FG5.CellBackColor = &HD0D0D0
 'FG5.TextMatrix(1, 5) = 1
 'FG5.Rows = FG5.Rows + 1
 FG5.TextMatrix(2, 0) = ""
 FG5.TextMatrix(2, 1) = "В.обеднение"
 FG5.TextMatrix(2, 2) = Format(0, "#0.000")
 FG5.TextMatrix(2, 3) = Format(WO, "#0.000")
 FG5.TextMatrix(2, 4) = "выбрать "
  If WO = 0 Then FG5.TextMatrix(2, 4) = ""
 FG5.TextMatrix(2, 5) = "1"
 FG5.TextMatrix(2, 6) = "90"
 'FG5.TextMatrix(2, 5) = 1
 
 FG5.Row = 2: FG5.Col = 1: FG5.CellBackColor = &HD0D0D0
 FG5.Row = 2: FG5.Col = 2: FG5.CellBackColor = &HD0D0D0
 FG5.Row = 2: FG5.Col = 3: FG5.CellBackColor = &HD0D0D0
 
 FG5.Visible = True
Call CalcFG5
 Call RTbl1
 
 
End Sub

Private Sub CommandButton1_Click() 'Выбор компонента
    Dim dbs As Database
    Dim qdf As QueryDef
    Dim rst  As DAO.Recordset
    Dim i, j, k As Integer
    
   ListBox1.ColumnCount = 4
   ListBox1.ColumnWidths = 90
   ListBox1.Clear
   ListBox2.ColumnCount = 4
   ListBox2.ColumnWidths = 90
   ListBox2.Clear

   S = "SELECT Material.Naim_mat,"
   S = S & " Material.Kod_marki,"
   S = S & " Material.Kod_doz,"
   S = S & " Material.Ves_max"
   S = S & " FROM Material"
 Set dbs = OpenDatabase("E:\DB\Baza.mdb")
 Set qdf = dbs.CreateQueryDef("")
  With qdf
           .SQL = S
       Set rst = .OpenRecordset()
    End With
   i = 0
    With rst
        Do While Not .EOF
        For j = 0 To 3
             MemComp(j, i) = .Fields(j).Value
       
        Next j
        
    If MemComp(1, i) = 1 Or MemComp(1, i) = 37 Or MemComp(1, i) = 38 Or MemComp(1, i) = 41 Then
       ListBox1.AddItem (MemComp(0, i))
       ListBox1.List(ListBox1.ListCount - 1, 1) = MemComp(1, i)
       ListBox1.List(ListBox1.ListCount - 1, 2) = MemComp(2, i)
       ListBox1.List(ListBox1.ListCount - 1, 3) = MemComp(3, i)
     Else
       ListBox2.AddItem (MemComp(0, i))
       ListBox2.List(ListBox2.ListCount - 1, 1) = MemComp(1, i)
       ListBox2.List(ListBox2.ListCount - 1, 2) = MemComp(2, i)
       ListBox2.List(ListBox2.ListCount - 1, 3) = MemComp(3, i)
     End If
       i = i + 1
              .MoveNext
             Loop
         .Close
    End With
dbs.Close

 ListBox1.Visible = True
 ListBox1.SetFocus
 ListBox1.ListIndex = 0
 ListBox2.Visible = True
 Call FormTabl
 CommandButton2.Visible = True
 CommandButton5.Visible = True
 Label1.Visible = True
 CommandButton1.Enabled = False
 
 End Sub

Private Sub FG5_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
Dim i, j, RwSelDel As Integer
    If KeyCode = 46 Then
     If FG5.Col = 1 Then
      If FG5.Rowsel > 2 Then
       If FG5.Rows > 4 Then
        RwSelDel = FG5.Rowsel
  BispD(Val(FG5.TextMatrix(FG5.Row, 4))) = 0
          FG5.RemoveItem (FG5.Rowsel)
                 For j = 3 To FG5.Rows - 1
                     FG5.TextMatrix(j, 0) = j - 2
                 Next j
                 For j = RwSelDel To FG5.Rows - 1
      MemCompSel(0, j) = j
      MemCompSel(1, j) = MemCompSel(1, j + 1)
      MemCompSel(2, j) = MemCompSel(2, j + 1)
      MemCompSel(3, j) = MemCompSel(3, j + 1)
                 Next j
                      Else
          FG5.Rows = 3
          For i = 0 To 4
             For j = 0 To 15
                MemCompSel(i, j) = 0
             Next j
          Next i
          For i = 0 To 10
        BispD(i) = 0
          Next i
      End If
       Call RTbl1
       Call CalcFG5
     End If
    End If
   End If
End Sub
Private Sub ListBox1_Click()
ListBox2.ListIndex = -1
ListBox2.Selected(0) = False
End Sub
Private Sub ListBox2_Click()
ListBox1.ListIndex = -1
ListBox1.Selected(0) = False
End Sub

Private Sub ListBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 13 Then Call LBtoFG
End Sub
Private Sub ListBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 13 Then Call LBtoFG
End Sub
Sub LBtoFG() 'Загрузка материалов в таблицу
If FG5.Rows < 18 Then
 FG5.Rows = FG5.Rows + 1
 FG5.TextMatrix(FG5.Rows - 1, 0) = FG5.Rows - 3
    If UserForm4.ActiveControl.TabIndex = 15 Then
       FG5.TextMatrix(FG5.Rows - 1, 1) = ListBox1.List(ListBox1.ListIndex)
      MemCompSel(0, FG5.Rows - 1) = FG5.Rows - 1
      MemCompSel(1, FG5.Rows - 1) = ListBox1.List(ListBox1.ListIndex, 1)
      MemCompSel(2, FG5.Rows - 1) = ListBox1.List(ListBox1.ListIndex, 2)
      MemCompSel(3, FG5.Rows - 1) = ListBox1.List(ListBox1.ListIndex, 3)
    End If
   If UserForm4.ActiveControl.TabIndex = 16 Then
       FG5.TextMatrix(FG5.Rows - 1, 1) = ListBox2.List(ListBox2.ListIndex)
      MemCompSel(0, FG5.Rows - 1) = FG5.Rows - 3
      MemCompSel(1, FG5.Rows - 1) = ListBox2.List(ListBox2.ListIndex, 1)
      MemCompSel(2, FG5.Rows - 1) = ListBox2.List(ListBox2.ListIndex, 2)
      MemCompSel(3, FG5.Rows - 1) = ListBox2.List(ListBox2.ListIndex, 3)
   End If
  FG5.Row = FG5.Rows - 1
  FG5.Col = 1
  FG5.CellBackColor = &H8000000F
 FG5.TextMatrix(FG5.Rows - 1, 2) = Format(0, "0.000")
 FG5.TextMatrix(FG5.Rows - 1, 3) = Format(0, "0.000")
 FG5.TextMatrix(FG5.Rows - 1, 4) = "выбрать "
 FG5.TextMatrix(FG5.Rows - 1, 5) = MemCompSel(2, FG5.Rows - 1)
 FG5.TextMatrix(FG5.Rows - 1, 6) = MemCompSel(1, FG5.Rows - 1)
End If
 FG5.Height = 19 + FG5.Rows * 16
 Call CalcFG5
 Label1.Move 130, 100 + FG5.Rows * 16
End Sub

Private Sub UserForm_Activate()
 If X = False Then

FG5.Cols = 7
FG5.Rows = 3
FG5.Height = 175
FG5.Left = 110
FG5.Top = 72
FG5.Width = 410


FG5.RowHeight(0) = 680
FG5.ColWidth(0) = 590
FG5.ColWidth(1) = 1600
FG5.ColWidth(2) = 1750
FG5.ColWidth(3) = 2300
FG5.ColWidth(4) = 1700
FG5.ColWidth(5) = 5
FG5.ColWidth(6) = 5

    FG5.Visible = False
    ListBox1.Visible = False
    ListBox2.Visible = False
    CommandButton2.Visible = False
    Label1.Visible = False
 
  TextBox1.SetFocus
  TextBox1.SelStart = 0
  TextBox1.SelLength = 6
  
  TextBox1.Height = 13.8
  TextBox1.Left = 278
  TextBox1.Top = 18
  TextBox1.Width = 48
  
  TextBox2.Height = 13.8
  TextBox2.Left = 222
  TextBox2.Top = 42
  TextBox2.Width = 24
  TextBox2.Enabled = False
  
  TextBox3.Height = 13.8
  TextBox3.Left = 258
  TextBox3.Top = 42
  TextBox3.Width = 48
  TextBox3.Enabled = False
  
  TextBox4.Height = 13.8
  TextBox4.Left = 312
  TextBox4.Top = 42
  TextBox4.Width = 24
  TextBox4.Enabled = False
  
  TextBox5.Height = 13.8
  TextBox5.Left = 342
  TextBox5.Top = 42
  TextBox5.Width = 24
  TextBox5.Enabled = False
  
  Label8.Height = 13.8
  Label8.Left = 378
  Label8.Top = 43
  Label8.Width = 54
  Label8.Visible = False
     
  Label2.Height = 12
  Label2.Left = 210
  Label2.Top = 19
  Label2.Width = 80
  Label2.Caption = "Расчет №:"
  
  Label3.Height = 12
  Label3.Left = 120
  Label3.Top = 43
  Label3.Width = 96
  Label3.Caption = "Вес электрода ="
  
  Label4.Height = 12
  Label4.Left = 246
  Label4.Top = 43
  Label4.Width = 6
  Label4.Caption = "+"
  
  
  Label5.Height = 12
  Label5.Left = 306
  Label5.Top = 43
  Label5.Width = 6
  Label5.Caption = "х"
  
  
  Label6.Height = 12
  Label6.Left = 336
  Label6.Top = 43
  Label6.Width = 6
  Label6.Caption = "+"
  
  Label7.Height = 12
  Label7.Left = 366
  Label7.Top = 43
  Label7.Width = 6
  Label7.Caption = "="
  
  CommandButton1.Enabled = False
  CommandButton1.Visible = False
  CommandButton5.Visible = False
  
 End If
End Sub






