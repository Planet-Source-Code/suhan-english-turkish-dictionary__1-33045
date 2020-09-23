VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sözlük"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1320
      List            =   "Form1.frx":000A
      TabIndex        =   9
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Prev"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4800
      TabIndex        =   2
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   570
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   "Turkish"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "English"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database 'db'yi database olarak tanimla
Dim rs As Recordset 'rs'yi recordset olarak tanimla
Dim SQLString As String 'SQLString'i string olarak tanimla
'ONEMLI:
'Bir database field'daki degerin numerik olup olmadigini bilir.
'Eger arama yaptirdigin field string ise, string olarak tanimlamalisin.
'Numerik ise, SQL'i WHERE English=1 seklinde kullanabilirsin.
'Aksi takdirde bir hata mesaji gelir.

Public Sub Command1_Click()
On Local Error Resume Next 'her ihtimale karsi
Set db = OpenDatabase(App.Path & "\sozluk.mdb") 'database'i goster
SQLString = "select * from soz where " & Combo1.Text & " like '" & Text9.Text & "*'"
'combo1'de fieldler var. text9'da ise arama yaptiracagimiz kayit var. LIKE ve * ile text9 ile baslayan her kayiti aratiyoruz.
Set rs = db.OpenRecordset(SQLString) 'recordset'i SQLString de tanimladigim sekilde ayarliyorum.
rs.MoveLast
rs.MoveFirst
'Neden son kayita gidip tekrar ilk kayita geldigimizi sorma.
'Bunu yapmazsan kayit sayisi yanlis cikiyor.Heralde bir bug.
Text1.Text = rs.Fields("english")
Text2.Text = rs.Fields("turkish")
Text4.Text = rs.RecordCount
'Buraya field sayisi kadar textbox girebilirsin ya da listbox'a aktarabilirsin.
End Sub

Public Sub Command2_Click()
On Local Error Resume Next
If Not rs.EOF Then 'recordset'in sonuna kadar.
rs.MoveNext 'birsonraki kayiti goster.
Text1.Text = rs.Fields("english")
Text2.Text = rs.Fields("turkish")
End If
End Sub

Private Sub Command3_Click()
On Local Error Resume Next
If Not rs.BOF Then 'recorset'in basina kadar
rs.MovePrevious 'bironceki kayiti goster
Text1.Text = rs.Fields("english")
Text2.Text = rs.Fields("turkish")
End If
End Sub

Public Sub Form_Load()
On Local Error Resume Next
Set db = OpenDatabase(App.Path & "\sozluk.mdb")
SQLString = "select * from soz"
'form ilk acildiginda tablodaki tüm field'larin tüm kayitlarini recordset'e aktariyorum.
Set rs = db.OpenRecordset(SQLString)
Text1.Text = rs.Fields("english")
Text2.Text = rs.Fields("turkish")
rs.MoveLast
rs.MoveFirst
Text4.Text = rs.RecordCount
Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
rs.Close 'exit tusu koyacaksan cikarken kapatmayi unutma
db.Close
End Sub

