VERSION 5.00
Begin VB.Form frmClass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TSClass!!!"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   Icon            =   "frmClass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "GO!"
      Height          =   345
      Left            =   7830
      TabIndex        =   1
      Top             =   90
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Text            =   "uid=sa;pwd=;driver={SQL Server};server=garfield;pwd=;database=gestionjudicial"
      Top             =   90
      Width           =   6915
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      Height          =   6075
      Left            =   90
      TabIndex        =   6
      Top             =   630
      Width           =   8265
      Begin VB.ListBox List2 
         Height          =   1635
         Left            =   4170
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   330
         Width           =   4095
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "Save File"
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   1980
         Width           =   4065
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1350
         TabIndex        =   2
         Text            =   "c:\windows\temp\"
         Top             =   0
         Width           =   2745
      End
      Begin VB.ListBox List1 
         Height          =   1620
         Left            =   0
         TabIndex        =   3
         Top             =   330
         Width           =   4095
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Create Class"
         Height          =   375
         Left            =   4170
         TabIndex        =   4
         Top             =   1980
         Width           =   4095
      End
      Begin VB.TextBox Text2 
         Height          =   3675
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   2400
         Width           =   8265
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "UNIQUE KEY (FOR WHERE CLAUSES)"
         Height          =   195
         Left            =   4170
         TabIndex        =   11
         Top             =   60
         Width           =   2895
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Target Dir:"
         Height          =   195
         Left            =   30
         TabIndex        =   8
         Top             =   60
         Width           =   750
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Connect:"
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   150
      Width           =   645
   End
End
Attribute VB_Name = "frmClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public cn As rdoConnection
Public en As rdoEnvironment
Public rs As rdoResultset

Private Sub Command1_Click()
On Error GoTo HELL
rdoEnvironments(0).CursorDriver = rdUseOdbc
Set en = rdoEnvironments(0)
Set cn = en.OpenConnection(dsName:="", Prompt:=rdDriverNoPrompt, Connect:=Text3.Text)
On Error Resume Next
For Counter = 0 To cn.rdoTables.Count - 1
    List1.AddItem cn.rdoTables(Counter).Name
Next
List1.ListIndex = 0
Text3.Enabled = False
Command1.Enabled = False
fra.Enabled = True
List1.SetFocus
Exit Sub
HELL:
    MsgBox (Err.Description)
End Sub

Private Sub cmd_Click()
Dim Enter As String, Buffer As String, Counter As Integer, ListFields As String, TableName As String, ListEdit As String, ListVariables As String, ListUpdate As String, ListaWhere As String
Enter = Chr(13) + Chr(10)
TableName = cn.rdoTables(List1.ListIndex).Name
Buffer = ""
Buffer = Buffer + "VERSION 1.0 CLASS" + Enter
Buffer = Buffer + "BEGIN" + Enter
Buffer = Buffer + "  MultiUse = -1  'True" + Enter
Buffer = Buffer + "End" + Enter
Buffer = Buffer + "Attribute VB_Name = " + Chr(34) + TableName + Chr(34) + Enter
Buffer = Buffer + "Attribute VB_GlobalNameSpace = False" + Enter
Buffer = Buffer + "Attribute VB_Creatable = True" + Enter
Buffer = Buffer + "Attribute VB_PredeclaredId = False" + Enter
Buffer = Buffer + "Attribute VB_Exposed = False" + Enter
Buffer = Buffer + "Attribute VB_Ext_KEY = ""SavedWithClassBuilder"" ,""Yes""" + Enter
Buffer = Buffer + "Attribute VB_Ext_KEY = ""Top_Level"" ,""Yes""" + Enter
Buffer = Buffer + "'CLASE GENERADA POR TSCLASS!!!" + Enter
Buffer = Buffer + "'Tabla:" + TableName + Enter
Buffer = Buffer + Enter
Buffer = Buffer + "Public ErrMessage as String" + Enter
Buffer = Buffer + Enter
Buffer = Buffer + "Public Arr" + TableName + " as Variant" + Enter
Buffer = Buffer + Enter
For Counter = 0 To cn.rdoTables(List1.ListIndex).rdoColumns.Count - 1
    Buffer = Buffer + "Public c" + cn.rdoTables(List1.ListIndex).rdoColumns(Counter).Name + " as String" + Enter
Next
Buffer = Buffer + Enter
ListFields = ""
ListEdit = ""
ListVariables = ""
ListUpdate = ""
ListaWhere = ""
For Counter = 0 To cn.rdoTables(List1.ListIndex).rdoColumns.Count - 1
    ListVariables = ListVariables + "c" + cn.rdoTables(List1.ListIndex).rdoColumns(Counter).Name + " as String, "
    ListFields = ListFields + cn.rdoTables(List1.ListIndex).rdoColumns(Counter).Name + ", "
    ListEdit = ListEdit + "'" + Chr(34) + "+ c" + cn.rdoTables(List1.ListIndex).rdoColumns(Counter).Name + " + " + Chr(34) + "' , "
    ListUpdate = ListUpdate + cn.rdoTables(List1.ListIndex).rdoColumns(Counter).Name + "='" + Chr(34) + "+ c" + cn.rdoTables(List1.ListIndex).rdoColumns(Counter).Name + " + " + Chr(34) + "' , "
    If UniqueKey(cn.rdoTables(List1.ListIndex).rdoColumns(Counter).Name) Then
        ListaWhere = ListaWhere + cn.rdoTables(List1.ListIndex).rdoColumns(Counter).Name + "='" + Chr(34) + "+ c" + cn.rdoTables(List1.ListIndex).rdoColumns(Counter).Name + " + " + Chr(34) + "' , "
    End If
Next
ListVariables = Mid(ListVariables, 1, Len(ListVariables) - 2)
ListFields = Mid(ListFields, 1, Len(ListFields) - 2)
ListEdit = Mid(ListEdit, 1, Len(ListEdit) - 2)
ListUpdate = Mid(ListUpdate, 1, Len(ListUpdate) - 2)
ListaWhere = Mid(ListaWhere, 1, Len(ListaWhere) - 2)
Buffer = Buffer + "'ADD RECORD" + Enter
Buffer = Buffer + "Public Function AddRecord( cn As rdoConnection, " + ListVariables + ") as Boolean" + Enter
Buffer = Buffer + "Dim rs as rdoResultSet, Chain as String" + Enter
Buffer = Buffer + "AddRecord = True" + Enter
Buffer = Buffer + "On error goto HELL" + Enter
Buffer = Buffer + "Chain = ""Insert Into " + TableName + "(" + ListFields + ") Values (" + ListEdit + ")" + Chr(34) + Enter
Buffer = Buffer + "Set rs = cn.OpenResultset(Chain, rdConcurReadOnly, rdOpenStatic, rdExecDirect)" + Enter
Buffer = Buffer + "SIGUE:" + Enter
Buffer = Buffer + "On Error GoTo 0" + Enter
Buffer = Buffer + "Exit Function" + Enter
Buffer = Buffer + "HELL:" + Enter
Buffer = Buffer + "    ErrMessage = Err.Description" + Enter
Buffer = Buffer + "    AddRecord = False" + Enter
Buffer = Buffer + "    GoTo SIGUE" + Enter
Buffer = Buffer + "End Function" + Enter
Buffer = Buffer + Enter
Buffer = Buffer + "'EDIT RECORD" + Enter
Buffer = Buffer + "Public Function EditRecord( cn As rdoConnection, " + ListVariables + ") as Boolean" + Enter
Buffer = Buffer + "Dim rs as rdoResultSet, Chain as String" + Enter
Buffer = Buffer + "EditRecord = True" + Enter
Buffer = Buffer + "On error goto HELL" + Enter
Buffer = Buffer + "Chain = ""Update " + TableName + " set " + ListUpdate + " where " + StrTran(ListaWhere, ",", " and ") + Enter
Buffer = Buffer + "Set rs = cn.OpenResultset(Chain, rdConcurReadOnly, rdOpenStatic, rdExecDirect)" + Enter
Buffer = Buffer + "SIGUE:" + Enter
Buffer = Buffer + "On Error GoTo 0" + Enter
Buffer = Buffer + "Exit Function" + Enter
Buffer = Buffer + "HELL:" + Enter
Buffer = Buffer + "    ErrMessage = Err.Description" + Enter
Buffer = Buffer + "    EditRecord = False" + Enter
Buffer = Buffer + "    GoTo SIGUE" + Enter
Buffer = Buffer + "End Function" + Enter
Buffer = Buffer + Enter
Buffer = Buffer + "'DeleteRecord UN REGISTRO" + Enter
Buffer = Buffer + "Public Function DeleteRecord( cn As rdoConnection, " + ListVariables + ") as Boolean" + Enter
Buffer = Buffer + "Dim rs as rdoResultSet, Chain as String" + Enter
Buffer = Buffer + "DeleteRecord= True" + Enter
Buffer = Buffer + "On error goto HELL" + Enter
Buffer = Buffer + "Chain = ""delete from " + TableName + " where " + StrTran(ListaWhere, ",", " and ") + Enter
Buffer = Buffer + "Set rs = cn.OpenResultset(Chain, rdConcurReadOnly, rdOpenStatic, rdExecDirect)" + Enter
Buffer = Buffer + "SIGUE:" + Enter
Buffer = Buffer + "On Error GoTo 0" + Enter
Buffer = Buffer + "Exit Function" + Enter
Buffer = Buffer + "HELL:" + Enter
Buffer = Buffer + "    ErrMessage = Err.Description" + Enter
Buffer = Buffer + "    DeleteRecord= False" + Enter
Buffer = Buffer + "    GoTo SIGUE" + Enter
Buffer = Buffer + "End Function" + Enter
Buffer = Buffer + Enter
Buffer = Buffer + "'CHARGE RECORDS" + Enter
Buffer = Buffer + "Public Function ChargeRecords( cn As rdoConnection, " + ListVariables + ") as Boolean" + Enter
Buffer = Buffer + "Dim rs as rdoResultSet, Chain as String" + Enter
Buffer = Buffer + "ChargeRecords= True" + Enter
Buffer = Buffer + "On error goto HELL" + Enter
Buffer = Buffer + "Chain = ""select " + ListFields + " from " + TableName + " where " + StrTran(ListaWhere, ",", " and ") + Enter
Buffer = Buffer + "Set rs = cn.OpenResultset(Chain, rdConcurReadOnly, rdOpenStatic, rdExecDirect)" + Enter
Buffer = Buffer + "Arr" + TableName + "= CargaArray(rs)" + Enter
Buffer = Buffer + "SIGUE:" + Enter
Buffer = Buffer + "On Error GoTo 0" + Enter
Buffer = Buffer + "Exit Function" + Enter
Buffer = Buffer + "HELL:" + Enter
Buffer = Buffer + "    ErrMessage = Err.Description" + Enter
Buffer = Buffer + "    ChargeRecords= False" + Enter
Buffer = Buffer + "    GoTo SIGUE" + Enter
Buffer = Buffer + "End Function" + Enter
Buffer = Buffer + Enter
Buffer = Buffer + "'FIND A RECORD" + Enter
Buffer = Buffer + "Public Function FindRecords( cn As rdoConnection, " + ListVariables + ") as Boolean" + Enter
Buffer = Buffer + "Dim rs as rdoResultSet, Chain as String" + Enter
Buffer = Buffer + "FindRecords= True" + Enter
Buffer = Buffer + "On error goto HELL" + Enter
Buffer = Buffer + "Chain = ""select " + ListFields + " from " + TableName + " where " + StrTran(ListaWhere, ",", " and ") + Enter
Buffer = Buffer + "Set rs = cn.OpenResultset(Chain, rdConcurReadOnly, rdOpenStatic, rdExecDirect)" + Enter
Buffer = Buffer + "If Not rs.EOF Then" + Enter
For Counter = 0 To cn.rdoTables(List1.ListIndex).rdoColumns.Count - 1
    Buffer = Buffer + "    c" + cn.rdoTables(List1.ListIndex).rdoColumns(Counter).Name + " = rs(" + Chr(34) + cn.rdoTables(List1.ListIndex).rdoColumns(Counter).Name + Chr(34) + ")" + Enter
Next
Buffer = Buffer + "Else" + Enter
Buffer = Buffer + "    FindRecords= False" + Enter
Buffer = Buffer + "End If" + Enter
Buffer = Buffer + "SIGUE:" + Enter
Buffer = Buffer + "On Error GoTo 0" + Enter
Buffer = Buffer + "Exit Function" + Enter
Buffer = Buffer + "HELL:" + Enter
Buffer = Buffer + "    ErrMessage = Err.Description" + Enter
Buffer = Buffer + "    FindRecords= False" + Enter
Buffer = Buffer + "    GoTo SIGUE" + Enter
Buffer = Buffer + "End Function" + Enter
Buffer = Buffer + Enter
Buffer = Buffer + "Private Function CargaArray(ByVal rs As rdoResultset) As Variant" + Enter
Buffer = Buffer + "Dim ejex As Integer" + Enter
Buffer = Buffer + "Dim Ejey As Integer" + Enter
Buffer = Buffer + "Dim Arr As Variant" + Enter
Buffer = Buffer + "ReDim Arr(rs.RowCount, rs.rdoColumns.Count)" + Enter
Buffer = Buffer + "ejex = 0" + Enter
Buffer = Buffer + "Ejey = 0" + Enter
Buffer = Buffer + "Do" + Enter
Buffer = Buffer + "    Do Until rs.EOF" + Enter
Buffer = Buffer + "        For Each Columna In rs.rdoColumns" + Enter
Buffer = Buffer + "            Arr(ejex, Ejey) = Columna.Value" + Enter
Buffer = Buffer + "            Ejey = Ejey + 1" + Enter
Buffer = Buffer + "        Next" + Enter
Buffer = Buffer + "        Ejey = 0" + Enter
Buffer = Buffer + "        ejex = ejex + 1" + Enter
Buffer = Buffer + "        rs.MoveNext" + Enter
Buffer = Buffer + "    Loop" + Enter
Buffer = Buffer + "Loop Until rs.MoreResults = False" + Enter
Buffer = Buffer + "CargaArray = Arr" + Enter
Buffer = Buffer + "End Function" + Enter
Buffer = Buffer + Enter
If chk.Value = 1 Then
    Open Text1.Text + TableName + ".cls" For Output As #1
    Print #1, Buffer
    Close #1
End If
Text2.Text = Buffer
End Sub

Private Sub List1_Click()
List2.Clear
For Counter = 0 To cn.rdoTables(List1.ListIndex).rdoColumns.Count - 1
    List2.AddItem cn.rdoTables(List1.ListIndex).rdoColumns(Counter).Name
Next
List2.Selected(0) = True
List2.ListIndex = 0
End Sub

Private Function UniqueKey(CampoActual) As Boolean
Dim Counter As Integer
UniqueKey = False
For Counter = 0 To List2.ListCount - 1
    If UCase(CampoActual) = UCase(List2.List(Counter)) Then
        If List2.Selected(Counter) Then
            UniqueKey = True
            Exit For
        End If
    End If
Next
End Function

Private Sub List2_Click()
Dim Counter As Integer, OK As Boolean
OK = False
For Counter = 0 To List2.ListCount - 1
    If List2.Selected(Counter) Then
        OK = True
        Exit For
    End If
Next
cmd.Enabled = OK
End Sub

Public Function StrTran(ByVal Chain As String, ByVal Inicial As String, ByVal Final As String) As String
Dim Counter As Double, Buffer As String
Counter = 0
Do While True
    Counter = Counter + 1
    If Counter > Len(Chain) Then
        Exit Do
    End If
    If Mid(Chain, Counter, Len(Inicial)) = Inicial Then
        Buffer = Mid(Chain, 1, Counter - 1) + Final + Mid(Chain, Counter + Len(Inicial))
        Chain = Buffer
    End If
Loop
StrTran = Chain
End Function

Public Function CargaArray(ByVal rs As rdoResultset) As Variant
Dim ejex As Integer
Dim Ejey As Integer
Dim Arr As Variant
ReDim Arr(rs.RowCount, rs.rdoColumns.Count)
ejex = 0
Ejey = 0
Do
    Do Until rs.EOF
        For Each Columna In rs.rdoColumns
            Arr(ejex, Ejey) = Columna.Value
            Ejey = Ejey + 1
        Next
        Ejey = 0
        ejex = ejex + 1
        rs.MoveNext
    Loop
Loop Until rs.MoreResults = False
CargaArray = Arr
End Function

