VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "КОРРЕКТИРОВКА РАСЧЕТА"
   ClientHeight    =   5724
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   12972
   OleObjectBlob   =   "UserForm3.dsx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private FM(20, 105), Fmm(4, 100) As Double
Private L, NNN, NumZ, NumId, NumZp, N_rst As Integer
Private F As Double


Private Sub CommandButton4_Click()
UserForm3.Hide ' закрыть форму 1
End Sub

Private Sub CommandButton1_Click() ' Загрузка таблицы
Dim k, kr, k1, k2, i, j As Integer
NNN = UserForm2.Label4.Caption + 1
FG1.Rows = NNN

 For i = 2 To FG1.Rows - 2
        FG1.TextMatrix(i, 0) = i - 1
        FG1.RowHeight(i) = 360
    Next i
        FG1.TextMatrix(1, 0) = "НО"
        FG1.TextMatrix(FG1.Rows - 1, 0) = "ВО"
    For i = 1 To FG1.Cols - 1
        If i < 11 Then FG1.TextMatrix(0, i) = i
        FG1.ColWidth(i) = 710
        FG2.ColWidth(i) = 710
       If i = 16 Then FG1.ColWidth(i) = 860: FG2.ColWidth(i) = 860
    Next i
       FG1.TextMatrix(0, 16) = "Вес П"
       
For i = 1 To 16
    For j = 1 To NNN - 1
       FG1.TextMatrix(j, i) = Format(0, "#0.000")
    Next j
        FG2.Row = 0
        FG2.Col = i
        FG2.CellBackColor = &HE0E0E0
Next i
   
For i = 0 To Fmm(0, 0)
        FG1.Row = Fmm(1, i)
        FG1.Col = Fmm(2, i)
  If FG1.Row <> 0 And FG1.Col <> 0 Then FG1.CellBackColor = &H80000009
            Fmm(0, i) = 0
            Fmm(1, i) = 0
            Fmm(2, i) = 0
            Fmm(3, i) = 0
            Fmm(0, 0) = i
Next i
For j = 2 To NNN - 2
     kr = 10: k1 = 0: k2 = 0
 For i = 3 To UserForm2.FG5.Rows - 1
    k = Val(UserForm2.FG5.TextMatrix(i, 4))
     If k < 11 Then
       FG1.TextMatrix(j, k) = Format(UserForm2.FG5.TextMatrix(i, 2), "#0.000")
     End If
     If k = 11 Then
          kr = k + k1 + Mid(UserForm2.FG5.TextMatrix(i, 4), 4, 1) - 1
          FG1.TextMatrix(j, kr) = Format(UserForm2.FG5.TextMatrix(i, 2), "#0.000")
          FG1.TextMatrix(0, kr) = UserForm2.FG5.TextMatrix(i, 4)
          k2 = k2 + 1
     End If
     If k = 12 Then
          kr = k + k2 + Mid(UserForm2.FG5.TextMatrix(i, 4), 4, 1) - 2
          FG1.TextMatrix(j, kr) = Format(UserForm2.FG5.TextMatrix(i, 2), "#0.000")
          FG1.TextMatrix(0, kr) = UserForm2.FG5.TextMatrix(i, 4)
          k1 = k1 + 1
      End If
 Next i
        Next j
 
    If UserForm2.FG5.TextMatrix(1, 3) > 0 Then
    FG1.TextMatrix(1, UserForm2.FG5.TextMatrix(1, 4)) = Format(UserForm2.FG5.TextMatrix(1, 3), "#0.000")
   End If
   If UserForm2.FG5.TextMatrix(2, 3) > 0 Then
    FG1.TextMatrix(NNN - 1, UserForm2.FG5.TextMatrix(2, 4)) = Format(UserForm2.FG5.TextMatrix(2, 3), "#0.000")
   End If
Call Calk1
End Sub

Private Sub CommandButton5_Click() 'Загрузка задания в БД

Dim dbs As Database
    Dim qdf As QueryDef
    Dim rst  As DAO.Recordset
    Dim sSQL As String
    Dim i, j, k As Integer
    Dim Mnid(1000, 2), Nid As Integer
    
  M = N_rst
NumZ = Val(InputBox("ВВЕДИТЕ НОМЕР ЗАДАНИЯ:", "ВВОД ДАННЫХ", M, 4000, 4000))
  If IsNumeric(NumZ) = True Then
  
    'Поиск идентификаторов для выбранного номера задания
    
 S = "SELECT Zadanie_all.N_id, Zadanie_all.Kod FROM Zadanie_all WHERE Zadanie_all.N_zad =" & NumZ
 Set dbs = OpenDatabase("E:\DB\Baza.mdb")
 Set qdf = dbs.CreateQueryDef("")
  With qdf
           .SQL = S
       Set rst = .OpenRecordset()
    End With
  NumId_Max = 0
  Nid = 0
    With rst
        Do While Not .EOF
         NumId = .Fields(0)
          If NumId > NumId_Max Then NumId_Max = NumId
          Mnid(Nid, 0) = NumId
          Mnid(Nid, 1) = .Fields(1)
         Nid = Nid + 1
       .MoveNext
         Loop
         .Close
    End With
  dbs.Close
      
     ' Проверка наличия в БД номера с указанным идентификатором
  M = NumId_Max + 1
NumId = Val(InputBox("ВВЕДИТЕ ИДЕНТИФИКАТОР ЗАДАНИЯ:", "ВВОД ДАННЫХ", M, 4000, 4000))
  If IsNumeric(NumId) = True Then
   j = 0
    For i = 0 To Nid - 1
      Select Case NumId
       Case Mnid(i, 0)
        j = 1
        Kod = Mnid(i, 1)
      End Select
   Next i
   
   'Удаление из базы записи с сущ. идентификатором
  If j = 1 Then
    Msg = "УКАЗАННЫЙ ИДЕНТИФИКАТОР УЖЕ СУЩЕСТВУЕТ"
    Style = vbYesNo + vbCritical + vbDefaultButto2
    Title = "СООБЩЕНИЯ": Help = "DEMO.HLP": Ctxt = 1000
    Responce = MsgBox(Msg, Style, Title, Help, Ctxt)
    If vbYes = 6 Then
     'If Response = vbYes Then
        Set dbs = OpenDatabase("D:\DB\Baza.mdb")
        S = "DELETE FROM Zadanie_all WHERE Zadanie_all.Kod =" & Kod
  dbs.Execute S
        S = "DELETE FROM Zadanie_ss WHERE Zadanie_ss.Kod =" & Kod
  dbs.Execute S
  dbs.Close
    End If
  End If
   
  Set dbs = OpenDatabase("E:\DB\Baza.mdb")
  sSQL = "SELECT*FROM Zadanie_all"
  Set rst = dbs.OpenRecordset(sSQL)
  
          rst.AddNew
            rst.Fields(1) = NumZ
            rst.Fields(2) = NumId
            rst.Fields(3) = Date
            rst.Fields(4) = Time
            rst.Fields(5) = 1
         If Label3.Caption = 0 Then
             rst.Fields(6) = 0
         Else
            rst.Fields(6) = Label3.Caption
         End If
            rst.Fields(7) = FG1.Rows - 1
      
      For i = 1 To 15
            rst.Fields(7 + i) = FG2.TextMatrix(2, i) & ""
            If rst.Fields(7 + i) = "" Then rst.Fields(7 + i) = 0
            'FG2.TextMatrix(2, i) & ""
            'Else
            'rst.Fields(7 + i) = 0
      'Stop
       Next i
      
      
      NumZp = rst.Fields(0)
      
         rst.Update
    rst.Close

  sSQL = "SELECT*FROM Zadanie_ss"
  Set rst = dbs.OpenRecordset(sSQL)
  
    For i = 1 To FG1.Rows - 1
      'For j = 1 To 15
        'If FG2.TextMatrix(0, j) > 0 Then
          rst.AddNew
             rst.Fields(0) = NumZp
             rst.Fields(1) = FG1.TextMatrix(i, 0)
             rst.Fields(2) = FG1.TextMatrix(i, 1)
             rst.Fields(3) = FG1.TextMatrix(i, 2)
             rst.Fields(4) = FG1.TextMatrix(i, 3)
             rst.Fields(5) = FG1.TextMatrix(i, 4)
             rst.Fields(6) = FG1.TextMatrix(i, 5)
             rst.Fields(7) = FG1.TextMatrix(i, 6)
             rst.Fields(8) = FG1.TextMatrix(i, 7)
             rst.Fields(9) = FG1.TextMatrix(i, 8)
             rst.Fields(10) = FG1.TextMatrix(i, 9)
             rst.Fields(11) = FG1.TextMatrix(i, 10)
             rst.Fields(12) = FG1.TextMatrix(i, 11)
             rst.Fields(13) = FG1.TextMatrix(i, 12)
             rst.Fields(14) = FG1.TextMatrix(i, 13)
             rst.Fields(15) = FG1.TextMatrix(i, 14)
             rst.Fields(16) = FG1.TextMatrix(i, 15)
             'rst.Fields(17) = FG1.TextMatrix(i, 16)
          rst.Update
        'End If
      'Next j
    Next i
    
 rst.Close
 dbs.Close

End If
    End If
End Sub
Sub PrNid()

End Sub



Private Sub UserForm_Activate() 'Задание форм таблиц при открытии окна
Dim i As Integer

Label1.Caption = UserForm2.Label3.Caption
N_rst = Val(Mid(Label1.Caption, 10, 5))
'TextBox5.Text = N_rst

FG1.Cols = 17
FG2.Cols = 17
NNN = UserForm2.Label4.Caption
FG1.Rows = NNN + 1
FG2.Rows = 3
FG1.ColWidth(0) = 330
FG2.ColWidth(0) = 330
FG2.RowHeight(0) = 360
FG1.Height = 140
FG1.Left = 6
FG1.Top = 20
FG1.Width = 610
FG2.Height = 18
FG2.Left = 6
FG2.Top = 160
FG2.Width = 600
    For i = 2 To FG1.Rows - 2
        FG1.TextMatrix(i, 0) = i - 1
        FG1.RowHeight(i) = 360
    Next i
 FG1.RowHeight(1) = 360
 FG1.RowHeight(FG1.Rows - 1) = 360
    For i = 1 To FG1.Cols - 1
        If i < 11 Then FG1.TextMatrix(0, i) = i
        FG1.ColWidth(i) = 710
        FG2.ColWidth(i) = 710
       If i = 16 Then FG1.ColWidth(i) = 860: FG2.ColWidth(i) = 860
    Next i
 Fmm(0, 0) = 0
    
  Sc = "35;30;60;60"
    
   Call Calk1
End Sub


Private Sub InP1()  'Редактирование таблицы

Dim Rm As Double
Dim S, M As String
Dim Msg, Style, Title, Help, Ctxt, Responce

If IsNumeric(FG1.TextMatrix(FG1.Row, FG1.Col)) = True Then Rm = FG1.TextMatrix(FG1.Row, FG1.Col)
M = "0"
S = InputBox("РЕДАКТИРОВАТЬ:", "ВВОД ДАННЫХ", M)

If IsNumeric(S) = True And Val(S) < 90 Then
    FG1.TextMatrix(FG1.Row, FG1.Col) = Format(S, "#0.000")
    FG1.CellBackColor = &HFFFF80                                   'метка Табл.1
      FG2.Row = 0: FG2.Col = FG1.Col: FG2.CellBackColor = &HFFFF80 'метка Табл.2
        Fmm(0, 0) = Fmm(0, 0) + 1: i = Fmm(0, 0): Fmm(0, i) = i
        Fmm(1, i) = FG1.Row: Fmm(2, i) = FG1.Col: Fmm(3, i) = Rm
                                        Else
    FG1.TextMatrix(FG1.Row, FG1.Col) = Format(Rm, "#0.000")
            If Val(S) >= 90 Then
            Msg = "ПРЕВЫШЕНО МАКСИМАЛЬНОЕ ЗНАЧЕНИЕ"
                        Else
            Msg = "ОШИБКА ФОРМАТА ВВОДА"
        End If
    Style = vbYes + vbCritical + vbDefaultButto2
    Title = "СООБЩЕНИЯ": Help = "DEMO.HLP": Ctxt = 1000
    If S <> "" Then Responce = MsgBox(Msg, Style, Title, Help, Ctxt)
End If
 Call Calk1
End Sub
Private Sub Calk1() 'Расчет итоговых сумм
Dim i, j, k As Integer
Dim SumI(16) As Double
             'по столбцам
    SumI(16) = 0
For i = 1 To 15
    SumI(i) = 0
  For j = 1 To FG1.Rows - 1
        If IsNumeric(FG1.TextMatrix(j, i)) = True Then SumI(i) = SumI(i) + FG1.TextMatrix(j, i)
  Next j
    FG2.TextMatrix(0, i) = Format(SumI(i), "#0.000")
    SumI(16) = SumI(16) + SumI(i)
Next i
    FG2.TextMatrix(0, 16) = Format(SumI(16), "#0.000")
             'по строкам
For i = 1 To FG1.Rows - 1
    SumI(0) = 0
  For j = 1 To 15
    If IsNumeric(FG1.TextMatrix(i, j)) = True Then SumI(0) = SumI(0) + FG1.TextMatrix(i, j)
  Next j
    FG1.TextMatrix(i, 16) = Format(SumI(0), "#0.000")
 Next i
End Sub
Private Sub MtOf()
Dim i, j, k As Integer
    If Fmm(0, 0) = 0 Then
        For i = 1 To 16
            FG2.Row = 0
            FG2.Col = i
            FG2.CellBackColor = &HE0E0E0      'снятие метки Табл.2
        Next i
                     Else
        For i = 1 To 16                    'перебор столбцов
          k = 0
            For j = 1 To Fmm(0, 0)         'перебор архива коррекций
          If Fmm(2, j) = i Then k = 1
            Next j
          If k = 0 Then
            FG2.Row = 0
            FG2.Col = i
            FG2.CellBackColor = &HE0E0E0  'снятие метки Табл.2
          End If
        Next i
      End If
End Sub

Private Sub CommandButton2_Click()  'Шаг назад
            'снятие метки с Табл.1
i = Fmm(0, 0)
    If i > 0 And i < 100 Then
        FG1.Row = Fmm(1, i)
        FG1.Col = Fmm(2, i)
        FG1.TextMatrix(FG1.Row, FG1.Col) = Format(Fmm(3, i), "#0.000")
        FG1.CellBackColor = &H80000009
            Fmm(0, i) = 0
            Fmm(1, i) = 0
            Fmm(2, i) = 0
            Fmm(3, i) = 0
            Fmm(0, 0) = i - 1
        End If
   Call MtOf
   Call Calk1
End Sub

Private Sub FG1_DblClick() 'Редактирование
 Call InP1
End Sub

Private Sub FG1_KeyPress(KeyAscii As Integer) 'Редактирование
 If KeyAscii = 13 Then
Call InP1
 End If
End Sub





