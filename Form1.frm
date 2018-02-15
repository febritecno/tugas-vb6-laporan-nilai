VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tugas No 5 (VB)"
   ClientHeight    =   6870
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   6675
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   375
      Left            =   5505
      TabIndex        =   12
      Top             =   8280
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   661
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nama"
         Caption         =   "nama"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "uts"
         Caption         =   "uts"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "uas"
         Caption         =   "uas"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "tugas"
         Caption         =   "tugas"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "angka"
         Caption         =   "angka"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "huruf"
         Caption         =   "huruf"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3465.071
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   3465.071
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3720
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   3120
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2520
      Width           =   4695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4200
      Top             =   8280
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=nilai.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=nilai.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "mahasiswa"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1560
      Width           =   4695
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TAMBAH DATA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1680
      TabIndex        =   11
      Top             =   360
      Width           =   3540
   End
   Begin VB.Label Label6 
      Caption         =   $"Form1.frx":0015
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   6360
      Width           =   6375
   End
   Begin VB.Label Label5 
      Caption         =   "Nama : Febrian Dwi Putra (FTI-B)/3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   5760
      Width           =   6375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "TUGAS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "UAS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "UTS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Index           =   0
      Begin VB.Menu cls 
         Caption         =   "Exit"
         Index           =   2
      End
   End
   Begin VB.Menu ctk 
      Caption         =   "Cetak Data"
      Index           =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hasil As Double
Function bersih()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text1.SetFocus
End Function

Private Sub cls_Click(Index As Integer)
Close
End
End Sub

Private Sub Command1_Click()
db
If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" And Text4.Text = "" Then
    MsgBox "Kolom tidak anda isi!", vbCritical, "Informasi!"
Else
        Adodc1.Recordset.AddNew
        Adodc1.Recordset.Fields("nama") = Text1.Text
        Adodc1.Recordset.Fields("uts") = Text2.Text
        Adodc1.Recordset.Fields("uas") = Text3.Text
        Adodc1.Recordset.Fields("tugas") = Text4.Text
        hasil = (Int(Text2.Text) * 40 / 100) + (Int(Text3.Text) * 50 / 100) + (Int(Text4.Text) * 10 / 100)
        Adodc1.Recordset.Fields("angka") = hasil
        cek (hasil)
        Adodc1.Recordset.Update
MsgBox "Data Tersimpan!", vbInformation, "Informasi!"
bersih
End If
End Sub

Private Sub ctk_Click(Index As Integer)
db
N.Open "select * from mahasiswa", Conn
If Not N.EOF Then
Set DataReport1.DataSource = N
DataReport1.Show
Else
MsgBox "Data Sedang Kosong, Laporan Tidak Bisa Di Cetak!", vbCritical, "Informasi"
End If
End Sub

Private Function cek(x)
If x >= 90 Then
Adodc1.Recordset!huruf = "A"
ElseIf x >= 70 Then
Adodc1.Recordset!huruf = "B"
ElseIf x >= 50 Then
Adodc1.Recordset!huruf = "C"
ElseIf x < 50 Then
Adodc1.Recordset!huruf = "E"
End If
End Function

Private Function uts(x)
hasil = x * 40 / 100
x = hasil
End Function
Private Function uas(x)
hasil = x * 50 / 100
x = hasil
End Function
Private Function tugas(x)
hasil = x * 10 / 100
x = hasil
End Function
