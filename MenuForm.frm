VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form MenuForm 
   Caption         =   "To-Do List"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc AdoDataTask 
      Height          =   330
      Left            =   2160
      Top             =   2400
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\development\VB6\ToDoList\ToDoList.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\development\VB6\ToDoList\ToDoList.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM Tarefas"
      Caption         =   "DataTask"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "MenuForm.frx":0000
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1720
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
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
            LCID            =   1046
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
            LCID            =   1046
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
   Begin VB.ListBox List1 
      Height          =   1035
      ItemData        =   "MenuForm.frx":001A
      Left            =   3480
      List            =   "MenuForm.frx":001C
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   360
      Left            =   3240
      TabIndex        =   2
      Top             =   3960
      Width           =   990
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   360
      Left            =   1920
      TabIndex        =   1
      Top             =   3960
      Width           =   990
   End
   Begin VB.Label lblBemVindo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To-Do List"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1845
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "MenuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
    If Not rs.EOF Then
        rs.Delete adAffectCurrent
        LoadDataGrid
    End If
End Sub


Private Sub cmdExit_Click()
    Unload MenuForm
    End
End Sub

Private Sub DataGrid1_dblClick()
    Dim selectRow As Long
    Dim selectCol As Long
    
    selectRow = DataGrid1.Row
    selectCol = DataGrid1.Col
    
    If selectRow > 0 And selectCol > 0 Then
        List1.AddItem DataGrid1.TextMatrix(selectRow, selectCol)
    End If
End Sub

Private Sub Form_Load()
    Call Connect
    LoadDataGrid
End Sub

