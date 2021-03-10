VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "ВЫБОР ДОЗАТОРОВ"
   ClientHeight    =   8856
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   11568
   OleObjectBlob   =   "UserForm2.dsx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Shlp, Sin As String
Private M, MemComp(4, 100), MemCompSel(4, 17) As Variant
Private Mlb(5, 3), Moth(8, 5), Mlig(8, 5)
Private r, X As Boolean
Private NumCmp, BispD(16), BispRW1(15), BispRW2(15) As Integer
Private Ps, Ns, Mp(17), Ms(17), Nd(17), Mps, Mss, Mpz, Msz, NumClc, NO, WO As Double
Private Sub CommandButton1_Click() 'ВЫХОД
        For i = 0 To 16
    BispD(i) = 0
If i < FG5.Rows - 1 Then FG5.Row = i + 1: FG5.Col = 4: FG5.CellBackColor = &H80000005
        Next i
UserForm2.Hide
End Sub
Private Sub CommandButton2_Click() 'Переход к таблице задания
Dim gg, NNN, i, j, k, kr, krn As Integer
gg = 0
For i = 3 To FG5.Rows - 1
 If Val(FG5.TextMatrix(i, 4)) = 0 Or Val(FG5.TextMatrix(i, 3)) = 0 Then gg = 1
Next i
  If gg = 0 And FG5.Rows > 3 Then
UserForm3.FG1.Cols = 17
NNN = Val(Label4.Caption) + 1
UserForm3.FG1.Rows = NNN
UserForm3.FG2.Rows = 3
UserForm3.FG1.ColWidth(0) = 330
UserForm3.FG1.RowHeight(0) = 360
UserForm3.FG1.RowHeight(1) = 360
UserForm3.FG1.RowHeight(NNN - 1) = 360
UserForm3.FG1.Height = 140
UserForm3.FG1.Left = 6
UserForm3.FG1.Top = 20
UserForm3.FG1.Width = 610
UserForm3.Label2.Caption = Label5.Caption
UserForm3.Label3.Caption = Label6.Caption
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

   If FG5.TextMatrix(1, 3) > 0 Then
    UserForm3.FG1.TextMatrix(1, FG5.TextMatrix(1, 4)) = Format(FG5.TextMatrix(1, 3), "#0.000")
   End If
   If FG5.TextMatrix(2, 3) > 0 Then
    UserForm3.FG1.TextMatrix(NNN - 1, FG5.TextMatrix(2, 4)) = Format(FG5.TextMatrix(2, 3), "#0.000")
   End If
   
   
  'Загрузка материалов
  
 kr = 10: k1 = 0: k2 = 0
For i = 3 To FG5.Rows - 1
    k = Val(FG5.TextMatrix(i, 4))
      If k < 11 Then
         UserForm3.FG2.TextMatrix(1, k) = FG5.TextMatrix(i, 1)
      End If
      If k = 11 Then
          kr = k + k1 + Mid(FG5.TextMatrix(i, 4), 4, 1) - 1
          UserForm3.FG2.TextMatrix(1, kr) = FG5.TextMatrix(i, 1)
          k2 = k2 + 1
      End If
      If k = 12 Then
          kr = k + k2 + Mid(FG5.TextMatrix(i, 4), 4, 1) - 2
          UserForm3.FG2.TextMatrix(1, kr) = FG5.TextMatrix(i, 1)
          k1 = k1 + 1
      End If

Next i

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
  X = True
  UserForm3.Show
  
Else
Style = vbYes + vbCritical + vbDefaultButto2
        Msg = "ВВЕДЕНЫ НЕ ВСЕ ДАННЫЕ"
        Title = "СООБЩЕНИЯ": Help = "DEMO.HLP": Ctxt = 1000
        Responce = MsgBox(Msg, Style, Title, Help, Ctxt)
    End If

End Sub

Private Sub CommandButton3_Click()

UserForm2.FG5.TextMatrix(3, 2) = UserForm2.FG5.TextMatrix(3, 2) / 2
UserForm2.FG5.TextMatrix(4, 2) = UserForm2.FG5.TextMatrix(3, 2)

UserForm2.FG5.TextMatrix(3, 3) = UserForm2.FG5.TextMatrix(3, 3) / 2
UserForm2.FG5.TextMatrix(4, 3) = UserForm2.FG5.TextMatrix(3, 3)
UserForm2.FG5.TextMatrix(4, 4) = "выбрать "
FG5.Row = 4: FG5.Col = 0: FG5.CellBackColor = &HD0D0D0
FG5.Row = 4: FG5.Col = 1: FG5.CellBackColor = &HD0D0D0
FG5.Row = 4: FG5.Col = 2: FG5.CellBackColor = &HD0D0D0
FG5.Row = 4: FG5.Col = 3: FG5.CellBackColor = &HD0D0D0

CommandButton3.Visible = False

End Sub

Private Sub FG5_KeyPress(KeyAscii As Integer) 'Редактирование
 If KeyAscii = 13 Then
   If FG5.Col = 4 Then
     If (FG5.Row = 1 And FG5.TextMatrix(1, 3) > 0) Or _
            (FG5.Row = 2 And FG5.TextMatrix(2, 3) > 0) Or _
                FG5.Row > 2 Then Call PrInTblCl4
    End If
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
  Label1.Move 55, 125 + FG5.Rows * 16
  Label1.Caption = "Суммарный вес:              " & _
                Format(Mps, "#0.000") & _
                  "                   " & Format(Mss, "#0.000")
End Sub
Private Sub FG5_KeyDown(KeyCode As Integer, ByVal Shift As Integer) 'Удаление
Dim i, j, RwSelDel As Integer
    If KeyCode = 46 Then
     If FG5.TextMatrix(FG5.Row, 3) = 0 And FG5.Row > 2 Then
      If FG5.Col = 1 Then
       If FG5.Rows > 2 Then
        RwSelDel = FG5.Rowsel
        BispD(Val(FG5.TextMatrix(FG5.Row, 4))) = 0
        FG5.RemoveItem (FG5.Rowsel)
           For j = 3 To FG5.Rows - 1
              FG5.TextMatrix(j, 0) = j - 2
           Next j
       Else
        FG5.Rows = 1
       For i = 0 To 18
        BispD(i) = 0
       Next i
      End If
       Call CalcFG5
       Call RTbl1
      End If
     End If
    End If
End Sub

Private Sub UserForm_Activate()

FG5.Cols = 7
FG5.Left = 50
FG5.Top = 102
FG5.Width = 410

FG5.RowHeight(0) = 650
FG5.ColWidth(0) = 550
FG5.ColWidth(1) = 1600
FG5.ColWidth(2) = 1750
FG5.ColWidth(3) = 2100
FG5.ColWidth(4) = 1700
FG5.ColWidth(5) = 5
FG5.ColWidth(6) = 5

 FG5.TextMatrix(0, 0) = "№№"
 FG5.TextMatrix(0, 1) = " КОМПОНЕНТ"
 FG5.TextMatrix(0, 3) = "         ВЕС             НА ЭЛЕКТРОД [кг]"
 FG5.TextMatrix(0, 4) = "    ДОЗАТОР"
 
  FG5.Visible = True
  Label1.Visible = True
    
    For i = 1 To FG5.Rows - 1
   If FG5.TextMatrix(i, 3) > 0 Then
 FG5.Row = i: FG5.Col = 1: FG5.CellBackColor = &HD0D0D0
 FG5.Row = i: FG5.Col = 2: FG5.CellBackColor = &HD0D0D0
 FG5.Row = i: FG5.Col = 3: FG5.CellBackColor = &HD0D0D0
    Else
 FG5.Row = i: FG5.Col = 1: FG5.CellBackColor = &HB0B0B0
 FG5.Row = i: FG5.Col = 2: FG5.CellBackColor = &HB0B0B0
 FG5.Row = i: FG5.Col = 3: FG5.CellBackColor = &HB0B0B0
    End If
    Next i

    
 If FG5.TextMatrix(1, 3) = 0 Then
    FG5.Row = 1: FG5.Col = 4: FG5.CellBackColor = &HD0D0D0
 End If
 If FG5.TextMatrix(2, 3) = 0 Then
    FG5.Row = 2: FG5.Col = 4: FG5.CellBackColor = &HD0D0D0
 End If
    For i = 0 To 15
        BispRW1(i) = 0: BispRW2(i) = 0
    Next i
    
 If UserForm2.FG5.TextMatrix(4, 3) > 0 Then CommandButton3.Visible = False
    
End Sub







