VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   12030
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "formPegawai.frx":0000
      Height          =   1935
      Left            =   1680
      TabIndex        =   15
      Top             =   3840
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   3413
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   2280
      Top             =   6360
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   1085
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
      Connect         =   $"formPegawai.frx":0015
      OLEDBString     =   $"formPegawai.frx":00AE
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Pegawai"
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   9360
      TabIndex        =   9
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   6960
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   255
      Left            =   5520
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Insert"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   3960
      TabIndex        =   4
      Top             =   2640
      Width           =   3975
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   3960
      TabIndex        =   3
      Top             =   2280
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   3960
      TabIndex        =   2
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   3960
      TabIndex        =   1
      Top             =   1560
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   3960
      TabIndex        =   0
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label5 
      Caption         =   "No. HP"
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Alamat"
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Jabatan"
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Nama"
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "ID Pegawai"
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub textClear()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
End Sub


Private Sub cmdAdd_Click()
    Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields("id") = Text1.Text
    Adodc1.Recordset.Fields("nama") = Text2.Text
    Adodc1.Recordset.Fields("alamat") = Text3.Text
    Adodc1.Recordset.Fields("jabatan") = Text4.Text
    Adodc1.Recordset.Fields("nohp") = Text5.Text
    textClear
End Sub

Private Sub cmdDelete_Click()
    Adodc1.Recordset.Delete
    Adodc1.Recordset.MoveNext
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    DataGrid1.Refresh
End Sub

Private Sub cmdUpdate_Click()
    Adodc1.Recordset.Update
    Adodc1.Recordset!id = Text1.Text
    Adodc1.Recordset!nama = Text2.Text
    Adodc1.Recordset!jabatan = Text3.Text
    Adodc1.Recordset!alamat = Text4.Text
    Adodc1.Recordset!nohp = Text5.Text
    textClear
End Sub

