VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisProject"
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Mb(100, 15), Mlb(5, 3), Moth(8, 5), Mlig(8, 5), SelRs(100, 12)
Private NumSel, NNl, NNo, NNg, Nz, Ne As Integer
Private S As String

Public Sub RR1()
Dim t As Tag
Set t = gTagDb.GetTag("R1")
gProject.TagDb.GetTag("R2") = Val(t)
End Sub
Public Sub RR2()
Dim Tag As Tag
'Set Tag = gTagDb.CreateTag("U1", analog) '   (FullName, [TagType])
 Dim userTag As Tag
  On Error GoTo ErrHandler

    Set userTag = gTagDb.CreateTag("Test\Analog")
    userTag.Description = "This is a test analog tag"
    userTag.MaximumValue = 200
    userTag.MinimumValue = -200
    userTag.ScaleFactor = 0.1

    userTag.Offset = 100
    'userTag.Units = "psi"
    userTag.WriteConfiguration

    'Release resources
    Set userTag = Nothing
ErrHandler:

    If Err Then
    'Display the error to the user
    MsgBox Err.Description
    Exit Sub
    End If
End Sub
Public Sub GMT() 'Генерация массива тегов
Dim userTag As Tag
Dim S As String

    For i = 1 To 100
        For j = 1 To 15
  On Error GoTo ErrHandler
        S = "d" & j & "\Z" & i
    Set userTag = gTagDb.CreateTag(S)
    userTag.MaximumValue = 200
    userTag.MinimumValue = 0
    userTag.ScaleFactor = 1
    userTag.WriteConfiguration

    Set userTag = Nothing
    
    Next j
   Next i
   
ErrHandler:

    If Err Then
    'Display the error to the user
    MsgBox Err.Description
    Exit Sub
    End If
End Sub
Public Sub RR4() ' Просмотр тегов
Dim t1, t2, t3 As Tag
Dim t4 As Double
Set t1 = gTagDb.GetTag("Num_Doz")
Set t2 = gTagDb.GetTag("N_ss")
     S = "d" & t1 & "\Z" & t2
Set t3 = gTagDb.GetTag(S)
t4 = Format(t3, "#0.000")
gProject.TagDb.GetTag("UU") = t4
End Sub

Public Sub RR5() ' Заполнение тегов
Dim i, j As Integer
    For i = 1 To 15
      For j = 1 To 50
     S = "d" & i & "\Z" & j
gProject.TagDb.GetTag(S) = i * j / 100
      Next j
    Next i
End Sub

Public Sub DMT() ' Очистка тегов
Dim i, j As Integer
    For i = 1 To 15
      For j = 1 To 100
     S = "d" & i & "\Z" & j
gProject.TagDb.GetTag(S) = 0
      Next j
    Next i
End Sub
Public Sub RZD()

Dim Ms(100, 3)
Dim dbs As Database
Dim qdf As QueryDef
Dim rst  As DAO.Recordset
Dim i, j, k As Integer
gProject.TagDb.GetTag("SL\Ind_R") = 1
Call DMT

NumZ = gTagDb.GetTag("SL\Num_Zad")
IdZ = gTagDb.GetTag("SL\Num_ID")


 
   S = "SELECT ZadS.kod,"
   S = S & " ZadS.Data,"
   S = S & " ZadS.Time,"
   S = S & " ZadS.N_ss"
   S = S & " FROM ZadS"
   S = S & " WHERE ZadS.N_zad ="
   S = S & Val(NumZ)
   S = S & " and ZadS.N_id ="
   S = S & Val(IdZ)
     
 Set dbs = OpenDatabase("C:\DB\Baza.mdb")
 Set qdf = dbs.CreateQueryDef("")
  With qdf
           .SQL = S
       Set rst = .OpenRecordset()
    End With
   
    With rst
        Do While Not .EOF
        Kod = .Fields(0).Value
        Nss = .Fields(3).Value
      .MoveNext
       Loop
        .Close
    End With
 
   S = "SELECT ZadM.c_ss, ZadM.num_doz,ZadM.wes FROM ZadM" & _
        " WHERE ZadM" & ".Kod =" & Val(Kod)
 
    Set qdf = dbs.CreateQueryDef("")
  With qdf
           .SQL = S
       Set rst = .OpenRecordset()
    End With
   i = 1: ndd = 10
    With rst
      Do While Not .EOF
       If IsNumeric(.Fields(1).Value) = False Then
            ndd = ndd + 1
            Nd = ndd
       Else: Nd = .Fields(1).Value: ndd = 10
        j = Val(.Fields(0).Value) + 1
        If .Fields(0).Value = "НО" Then j = 1
        If .Fields(0).Value = "ВО" Then j = Nss
       End If
    'FG1.TextMatrix(j, Val(Nd)) = Format(.Fields(2).Value, "#0.000")
    'gProject.TagDb.GetTag("D" & Val(Nd) & "\Z" & j) = Format(.Fields(2).Value, "#0.000")
    gProject.TagDb.GetTag("D" & Val(Nd) & "\Z" & j) = .Fields(2).Value
      i = i + 1
      .MoveNext
        Loop
         .Close
    End With
  
dbs.Close
gProject.TagDb.GetTag("SL\Ind_R") = 0
End Sub
Public Sub Rst1()

 UserForm1.Show
End Sub
Public Sub Rst2()

 UserForm4.Show
End Sub
Public Sub Rst3()

 UserForm5.Show
End Sub



