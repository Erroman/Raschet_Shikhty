VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "ВЫБОР ЗАДАНИЯ"
   ClientHeight    =   9468
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   14040
   OleObjectBlob   =   "UserForm1.dsx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private M1(15, 100), M2(15, 100), M3(15, 100), SelRs(100, 12) As Variant
Private S, Sh, Sp, Num As String
Private X As Boolean
Private Ms(100, 15), Mlb(5, 3), Moth(8, 5), Mlig(8, 5)
Private NumSel, NNl, NNo, NNg, NNd, Nz, Ne, i As Integer

Private Sub CommandButton2_Click()
Call LBtoFG
End Sub

Private Sub CommandButton3_Click()
X = False
UserForm1.Hide
End Sub
Private Sub UserForm_Activate()
If X = False Then
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
ListBox1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
CommandButton2.Visible = False
CommandButton1.ControlTipText = "НАЖМИТЕ НА КНОПКУ ДЛЯ ВЫБОРА НОМЕРА РАСЧЕТА"
End If
End Sub

Private Sub CommandButton1_Click() 'Выбор задания (поиск выбранного )
Dim CntRs As Integer
Dim Msg, M, Style, Title, Help, Ctxt, Response
ListBox1.Clear
M = ""
Num = Val(InputBox("ВВЕДИТЕ НОМЕР РАСЧЕТА:", "ВВОД ДАННЫХ", M, 4000, 4000))
  If IsNumeric(Num) = True Then
     Num = Val(Num)
    S = "SELECT RASCHET_res.Data_ras,"
      S = S & " RASCHET_res.ras_id,"
      S = S & " RASCHET_res.Ves_el,"
      S = S & " RASCHET_res.K_porz,"
      S = S & " RASCHET_res.Ves_porz,"
      S = S & " RASCHET_res.Gub_ves_ras,"
      S = S & " RASCHET_res.Obedn_n,"
      S = S & " RASCHET_res.Obedn_v,"
      S = S & " RASCHET_res.Splav,"
      S = S & " RASCHET_res.kod_pres,"
      S = S & " RASCHET_res.diam,"
      S = S & " RASCHET_res.Kod_spl"
      S = S & " FROM RASCHET_res"
      S = S & " WHERE n_ras ="
      S = S & Val(Num)
     Nz = 12
 id = 1
 Label4.Caption = id
 Call CQD1
    CntRs = Ne - 1 'Число найденых расчетов в БД АРМ
    
If CntRs > 0 Then
 For i = 1 To CntRs
  ListBox1.AddItem (Ms(i, 0) & "       " & Ms(i, 1))
    SelRs(i, 0) = Val(Num)
    SelRs(i, 1) = Ms(i, 1): SelRs(i, 2) = Ms(i, 0): SelRs(i, 6) = Ms(i, 2)
    SelRs(i, 7) = Ms(i, 3): SelRs(i, 8) = Ms(i, 4): SelRs(i, 9) = Ms(i, 5)
    SelRs(i, 10) = Ms(i, 6): SelRs(i, 11) = Ms(i, 7): SelRs(i, 3) = Ms(i, 8)
    SelRs(i, 4) = Ms(i, 9): SelRs(i, 5) = Ms(i, 10): SelRs(i, 12) = Ms(i, 11)
 
 Next i
 
 Label1.Caption = "       Дата       Номер ID"
 Label2.Caption = "Расчет № " & Num
 Label3.Caption = "Найдено " & CntRs & " записей"
 ListBox1.Visible = True
 Label1.Visible = True
 Label2.Visible = True
 Label3.Visible = True
 ListBox1.SetFocus
 ListBox1.ListIndex = 0
Else
 ListBox1.Visible = False
 Label1.Visible = False
 Label2.Visible = False
 Label3.Visible = False
 
    Msg = "Расчет № " & Num & " в БД не найден."
    Style = vbYesNo + vbCritical
    Title = "Расчет № " & Num & " в БД не найден."
    Response = MsgBox(Msg, Style, Title, Help, Ctxt)
  'End If
  'End If
  End If
 End If
End Sub
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Call ListRes
End Sub
Private Sub ListBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call ListRes
End Sub
Sub CQD()
 Dim dbs As Database
   'Dim qdf As QueryDef
   'Dim rst As Recordset
    Dim i, j, k As Integer
  
  Set dbs = OpenDatabase("E:\DB\Baza.mdb")
  Set qdf = dbs.CreateQueryDef("")
    With qdf
           .SQL = S
       Set rst = .OpenRecordset()
    End With
   i = 1
    With rst
        Do While Not .EOF
        For k = 0 To Nz - 1
             Ms(i, k) = .Fields(k).Value
        Next k
   i = i + 1
            .MoveNext
             Loop
         .Close
    End With
dbs.Close
    Ne = i
Exit Sub
Err_Handler:
        com = "Remark Iai?aaeeuii caaaiu ia?aiao?u iienea " & Now
        gCommand.Execute (com)
        MyString = "Iai?aaeeuii caaaiu ia?aiao?u iienea " & Now
        com = "Invoke " & "Find.LstBx" & ".Clear()"
        gCommand.Execute (com)
        com = "Invoke " & "Find.LstBx" & ".additem(" & Chr(34) & MyString & Chr(34) & ")"
        gCommand.Execute (com)
        MsgBox Err.Description, vbExclamation
   
End Sub
Sub ListRes()
'Dim MyFont As StdFont
Dim i, j As Integer

For i = 0 To 8
    For j = 0 To 5
        Moth(i, j) = ""
        Mlig(i, j) = ""
     Next j
 Next i

ListBox2.ColumnCount = 4
Rows = 6
ListBox2.ColumnWidths = "170;110;150;60"
ListBox2.Clear
NumSel = ListBox1.ListIndex

    Mlb(0, 0) = "Номер расчета:"
    Mlb(0, 1) = SelRs(NumSel + 1, 0)
    Mlb(1, 0) = "Идентификатор расчета:"
    Mlb(1, 1) = Format(SelRs(NumSel + 1, 1))
    Mlb(2, 0) = "Дата:"
    Mlb(2, 1) = Format(SelRs(NumSel + 1, 2))
    Mlb(3, 0) = "Сплав:"
    Mlb(3, 1) = Format(SelRs(NumSel + 1, 3))
    UserForm2.Label5 = SelRs(NumSel + 1, 3)
    UserForm2.Label6 = SelRs(NumSel + 1, 12)
    Mlb(4, 0) = "Пресс:"
    Mlb(4, 1) = Format(SelRs(NumSel + 1, 4))
    Mlb(5, 0) = "Диаметр:"
    Mlb(5, 1) = Format(SelRs(NumSel + 1, 5))
    Mlb(0, 2) = "Вес электрода, кг:"
    Mlb(0, 3) = Format(SelRs(NumSel + 1, 6))
    Mlb(1, 2) = "Количество порций:"
    Mlb(1, 3) = Format(SelRs(NumSel + 1, 7))
    Mlb(2, 2) = "Вес порции, кг:"
    Mlb(2, 3) = Format(SelRs(NumSel + 1, 8))
    Mlb(3, 2) = "Вес губки(на 10 кг), кг:"
    Mlb(3, 3) = Format(SelRs(NumSel + 1, 9))
    Mlb(4, 2) = "Нижнее обеднение, кг:"
    Mlb(4, 3) = Format(SelRs(NumSel + 1, 10))
    Mlb(5, 2) = "Верхнее обеднение, кг:"
    Mlb(5, 3) = Format(SelRs(NumSel + 1, 11))

ListBox2.List() = Mlb()
      'Загрузка данных по отходам
 Num = Mlb(1, 1)
    Nz = 5
  S = "SELECT RASCHET_oth3.prizn_oth,"
     S = S & " RASCHET_oth3.pnaim,"
     S = S & " RASCHET_oth3.kol_oth,"
     S = S & " RASCHET_oth3.kod_doz,"
     S = S & " RASCHET_oth3.Ves_max FROM RASCHET_oth3"
     S = S & " WHERE ras_id ="
     S = S & Num
'Call CQD
'If Label4.Caption = 1 Then Call CQD1 Else CQD
 Call CQD1
   For i = 1 To Ne - 1
    Moth(i, 0) = Format(Ms(i, 0))
    Moth(i, 1) = Format(Ms(i, 1))
    Moth(i, 2) = Format(Ms(i, 2))
    Moth(i, 3) = Format(Ms(i, 2) * Mlb(2, 3) / 10, "#0.000")
    Moth(i, 4) = Format(Ms(i, 3))
    Moth(i, 5) = Format(Ms(i, 4))
 Next i
   NNo = Ne - 1
     'Загрузка данных по лигатуре
 Num = Mlb(1, 1)
    Nz = 5
  S = "SELECT RASCHET_lig3.kod_marki,"
     S = S & " RASCHET_lig3.marka,"
     S = S & " RASCHET_lig3.kol_lig,"
     S = S & " RASCHET_lig3.kod_doz,"
     S = S & " RASCHET_lig3.Ves_max FROM RASCHET_lig3"
     S = S & " WHERE ras_id ="
     S = S & Num
'Call CQD
'If Label4.Caption = 1 Then Call CQD1 Else CQD
 Call CQD1
  For i = 1 To Ne - 1
    Mlig(i, 0) = Format(Ms(i, 0))
    Mlig(i, 1) = Format(Ms(i, 1))
    Mlig(i, 2) = Format(Ms(i, 2))
    Mlig(i, 3) = Format(Ms(i, 2) * Mlb(2, 3) / 10, "#0.000")
    Mlig(i, 4) = Format(Ms(i, 3))
    Mlig(i, 5) = Format(Ms(i, 4))
  
  Next i
   NNl = Ne - 1
  
  ListBox3.Clear
  ListBox3.ColumnCount = 6
  ListBox3.ColumnWidths = "30;120;80;80;60;60"
  ListBox3.List() = Moth()
  ListBox3.List(0, 0) = "ID"
  ListBox3.List(0, 1) = "Материал"
  ListBox3.List(0, 2) = "Вес на 10 кг"
  ListBox3.List(0, 3) = "Вес порции"
  ListBox3.List(0, 4) = "KD"
  ListBox3.List(0, 5) = "Max вес"
  
  ListBox4.Clear
  ListBox4.ColumnCount = 6
  ListBox4.ColumnWidths = "30;120;80;80;60;60"
  ListBox4.List() = Mlig()
  ListBox4.List(0, 0) = "ID"
  ListBox4.List(0, 1) = "Материал"
  ListBox4.List(0, 2) = "Вес на 10 кг"
  ListBox4.List(0, 3) = "Вес порции"
  ListBox4.List(0, 4) = "KD"
  ListBox4.List(0, 5) = "Max вес"
  
  Frame2.Visible = True
  Frame3.Visible = True
  Frame4.Visible = True
  CommandButton2.Visible = True
 
End Sub
Sub LBtoFG() 'Загрузка материалов в таблицу
 UserForm2.FG5.Cols = 7
 NNg = 2
 NNd = 1
   If Mlb(3, 3) * Mlb(2, 3) / 10 > 80 Then NNd = 2
  UserForm2.FG5.Rows = NNg + NNo + NNl + 3
  UserForm2.FG5.TextMatrix(1, 0) = ""
  UserForm2.FG5.TextMatrix(1, 1) = "Н.обеднение"
  UserForm2.FG5.TextMatrix(1, 2) = Format(0, "#0.000")
  UserForm2.FG5.TextMatrix(1, 3) = Format(Mlb(4, 3), "#0.000")
  UserForm2.FG5.TextMatrix(1, 4) = "выбрать "
  If Mlb(4, 3) = 0 Then UserForm2.FG5.TextMatrix(1, 4) = ""
  UserForm2.FG5.TextMatrix(1, 5) = 1
  UserForm2.FG5.TextMatrix(2, 0) = ""
  UserForm2.FG5.TextMatrix(2, 1) = "В.обеднение"
  UserForm2.FG5.TextMatrix(2, 2) = Format(0, "#0.000")
  UserForm2.FG5.TextMatrix(2, 3) = Format(Mlb(5, 3), "#0.000")
  UserForm2.FG5.TextMatrix(2, 4) = "выбрать "
  If Mlb(5, 3) = 0 Then UserForm2.FG5.TextMatrix(2, 4) = ""
  UserForm2.FG5.TextMatrix(2, 5) = 1
   
   UserForm2.FG5.TextMatrix(3, 0) = 1
   UserForm2.FG5.TextMatrix(3, 1) = "Губка"
   UserForm2.FG5.TextMatrix(3, 2) = Format(Mlb(3, 3) * Mlb(2, 3) / (NNd * 10), "#0.000")
   UserForm2.FG5.TextMatrix(3, 3) = Format(Mlb(3, 3) * Mlb(2, 3) * Mlb(1, 3) / (NNd * 10), "#0.000")
   UserForm2.FG5.TextMatrix(3, 4) = "выбрать "
   UserForm2.FG5.TextMatrix(3, 5) = 1
   UserForm2.FG5.TextMatrix(3, 6) = "1"
   
   UserForm2.FG5.TextMatrix(4, 0) = "2"
   UserForm2.FG5.TextMatrix(4, 1) = "Губка"
   If NNd = 2 Then
      UserForm2.FG5.TextMatrix(4, 2) = Format(Mlb(3, 3) * Mlb(2, 3) / (NNd * 10), "#0.000")
      UserForm2.FG5.TextMatrix(4, 3) = Format(Mlb(3, 3) * Mlb(2, 3) * Mlb(1, 3) / (NNg * 10), "#0.000")
      UserForm2.FG5.TextMatrix(4, 4) = "выбрать "
   Else
      UserForm2.FG5.TextMatrix(4, 2) = 0
      UserForm2.FG5.TextMatrix(4, 3) = 0
      UserForm2.FG5.TextMatrix(4, 4) = ""
   End If
   UserForm2.FG5.TextMatrix(4, 5) = 1
   UserForm2.FG5.TextMatrix(4, 6) = "1"
   
 For i = 1 To NNo + NNl
    If i < NNo + 1 Then
  UserForm2.FG5.TextMatrix(i + NNg + 2, 0) = i + NNg + 2
  UserForm2.FG5.TextMatrix(i + NNg + 2, 1) = Moth(i, 1)
  UserForm2.FG5.TextMatrix(i + NNg + 2, 2) = Format(Moth(i, 3), "#0.000")
  UserForm2.FG5.TextMatrix(i + NNg + 2, 3) = Format(Moth(i, 3) * Mlb(1, 3), "#0.000")
  UserForm2.FG5.TextMatrix(i + NNg + 2, 4) = "выбрать "
  UserForm2.FG5.TextMatrix(i + NNg + 2, 5) = Moth(i, 4)
  UserForm2.FG5.TextMatrix(i + NNg + 2, 6) = Moth(i, 0)
    Else
  UserForm2.FG5.TextMatrix(i + NNg + 2, 0) = i + NNg + 2
  UserForm2.FG5.TextMatrix(i + NNg + 2, 1) = Mlig(i - NNo, 1)
  UserForm2.FG5.TextMatrix(i + NNg + 2, 2) = Format(Mlig(i - NNo, 3), "#0.000")
  UserForm2.FG5.TextMatrix(i + NNg + 2, 3) = Format(Mlig(i - NNo, 3) * Mlb(1, 3), "#0.000")
  UserForm2.FG5.TextMatrix(i + NNg + 2, 4) = "выбрать "
  UserForm2.FG5.TextMatrix(i + NNg + 2, 5) = Mlig(i - NNo, 4)
  UserForm2.FG5.TextMatrix(i + NNg + 2, 6) = Mlig(i - NNo, 0)
    End If
 'Stop
Next i
   
UserForm2.FG5.TextMatrix(0, 2) = "            ВЕС        НА   " _
    & Format(Mlb(2, 3), "#0.00") & "  [кг]"

 UserForm2.FG5.Height = 19 + UserForm2.FG5.Rows * 16
        
      Mps = 0: Mss = 0
 For i = 5 To UserForm2.FG5.Rows - 1
      Mps = Mps + UserForm2.FG5.TextMatrix(i, 2)
      Mss = Mss + UserForm2.FG5.TextMatrix(i, 3)
 
 Next i
      Mss = Mss + Mlb(4, 3) + Mlb(5, 3)
      
    UserForm2.Label1.Caption = "Суммарный вес:             " _
      & Format(Mps, "#0.000") & "                 " _
        & Format(Mss, "#0.000")
    UserForm2.Label2.Caption = "Вес электрода: " _
      & Format(Mss, "#0.0") & " = " _
       & Val(Mlb(4, 3)) & " + " _
        & Format(Mps, "#0.0") & " x " _
         & Val(Mlb(1, 3)) & " + " & Val(Mlb(5, 3))
         
    UserForm2.Label3.Caption = "Расчет №  " & Val(Mlb(0, 1))
    UserForm2.Label1.Move 55, 125 + (NNo + NNl + NNg + 3) * 16
    UserForm2.FG5.Visible = True
    UserForm2.Label1.Visible = True
    UserForm2.Label4.Caption = Val(Mlb(1, 3) + 2)
    'UserForm3.TextBox1.Text = Val(Mlb(0, 1))
    X = True
    
    UserForm2.Show

End Sub

Sub CQD1()
 Dim dbs As Database
   'Dim qdf As QueryDef
   'Dim rst As Recordset
    Dim i, j, k As Integer
  
  Set dbs = OpenDatabase("E:\DB\SyBase.mdb")
  Set qdf = dbs.CreateQueryDef("")
    With qdf
           .SQL = S
       Set rst = .OpenRecordset()
    End With
   i = 1
    With rst
        Do While Not .EOF
        For k = 0 To Nz - 1
             Ms(i, k) = .Fields(k).Value
        Next k
   i = i + 1
            .MoveNext
             Loop
         .Close
    End With
dbs.Close
    Ne = i
Exit Sub
Err_Handler:
        com = "Remark Iai?aaeeuii caaaiu ia?aiao?u iienea " & Now
        gCommand.Execute (com)
        MyString = "Iai?aaeeuii caaaiu ia?aiao?u iienea " & Now
        com = "Invoke " & "Find.LstBx" & ".Clear()"
        gCommand.Execute (com)
        com = "Invoke " & "Find.LstBx" & ".additem(" & Chr(34) & MyString & Chr(34) & ")"
        gCommand.Execute (com)
        MsgBox Err.Description, vbExclamation
   
End Sub


