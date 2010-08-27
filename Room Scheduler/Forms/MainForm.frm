VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10380
   ClientLeft      =   90
   ClientTop       =   2265
   ClientWidth     =   17310
   BeginProperty Font 
      Name            =   "Bernard MT Condensed"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "MainForm.frx":0000
   ScaleHeight     =   10380
   ScaleWidth      =   17310
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flex 
      Height          =   1815
      Left            =   4440
      TabIndex        =   6
      Top             =   1440
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3201
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   375
      Left            =   3000
      Top             =   9480
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=STI"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "STI"
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "amirah@1"
      RecordSource    =   $"MainForm.frx":3F5FC
      Caption         =   "Adodc5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bernard MT Condensed"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid5 
      Bindings        =   "MainForm.frx":3F73C
      Height          =   6015
      Left            =   5040
      TabIndex        =   5
      Top             =   2880
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10610
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Day"
         Caption         =   "Day"
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
      BeginProperty Column02 
         DataField       =   "Start Time"
         Caption         =   "Start Time"
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
      BeginProperty Column03 
         DataField       =   "End Time"
         Caption         =   "End Time"
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
      BeginProperty Column04 
         DataField       =   "Room"
         Caption         =   "Room"
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
      BeginProperty Column05 
         DataField       =   "Subject"
         Caption         =   "Subject"
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
      BeginProperty Column06 
         DataField       =   "Section"
         Caption         =   "Section"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   3000
      Top             =   9120
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=STI"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "STI"
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "amirah@1"
      RecordSource    =   $"MainForm.frx":3F751
      Caption         =   "Adodc4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bernard MT Condensed"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid4 
      Bindings        =   "MainForm.frx":3F7F6
      Height          =   6015
      Left            =   5040
      TabIndex        =   4
      Top             =   2880
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10610
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "Section Name"
         Caption         =   "Section Name"
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
      BeginProperty Column02 
         DataField       =   "User"
         Caption         =   "User"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   2880
      Top             =   8760
      Visible         =   0   'False
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
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=STI"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "STI"
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "amirah@1"
      RecordSource    =   "SELECT id as ID, last_name as 'Last Name', first_name as 'First Name', middle_name as 'Middle Name', type as Type FROM users"
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bernard MT Condensed"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "MainForm.frx":3F80B
      Height          =   6015
      Left            =   5040
      TabIndex        =   3
      Top             =   2880
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10610
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "Last Name"
         Caption         =   "Last Name"
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
      BeginProperty Column02 
         DataField       =   "First Name"
         Caption         =   "First Name"
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
      BeginProperty Column03 
         DataField       =   "Middle Name"
         Caption         =   "Middle Name"
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
      BeginProperty Column04 
         DataField       =   "Type"
         Caption         =   "Type"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1140.095
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   2880
      Top             =   8280
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=STI"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "STI"
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "amirah@1"
      RecordSource    =   "SELECT id as ID, name as 'Subject Name' FROM subjects"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bernard MT Condensed"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "MainForm.frx":3F820
      Height          =   6015
      Left            =   5040
      TabIndex        =   2
      Top             =   2880
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10610
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "Subject Name"
         Caption         =   "Subject Name"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2760
      Top             =   7920
      Visible         =   0   'False
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
      CommandType     =   1
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "amirah@1"
      RecordSource    =   "SELECT id as ID, name as 'Subject Name' FROM rooms"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bernard MT Condensed"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "MainForm.frx":3F835
      Height          =   6015
      Left            =   5040
      TabIndex        =   0
      Top             =   2880
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10610
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "Subject Name"
         Caption         =   "Subject Name"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Schedules"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   675
      Left            =   5160
      TabIndex        =   1
      Top             =   1995
      Width           =   2415
   End
   Begin VB.Image Image13 
      Height          =   825
      Left            =   240
      Picture         =   "MainForm.frx":3F84A
      Top             =   9360
      Width           =   2190
   End
   Begin VB.Image Image12 
      Height          =   1245
      Left            =   0
      Picture         =   "MainForm.frx":4106E
      Top             =   7080
      Width           =   4455
   End
   Begin VB.Image Image11 
      Height          =   735
      Left            =   15960
      Picture         =   "MainForm.frx":44725
      Top             =   1920
      Width           =   765
   End
   Begin VB.Image Image10 
      Height          =   735
      Left            =   11280
      Picture         =   "MainForm.frx":44F14
      Top             =   1920
      Width           =   4635
   End
   Begin VB.Image Image9 
      Height          =   735
      Left            =   10080
      Picture         =   "MainForm.frx":4528C
      Top             =   1905
      Width           =   615
   End
   Begin VB.Image Image8 
      Height          =   735
      Left            =   9360
      Picture         =   "MainForm.frx":4595C
      Top             =   1900
      Width           =   600
   End
   Begin VB.Image Image7 
      Height          =   690
      Left            =   8640
      Picture         =   "MainForm.frx":460C8
      Top             =   1920
      Width           =   630
   End
   Begin VB.Image Image6 
      Height          =   1305
      Left            =   0
      Picture         =   "MainForm.frx":4677C
      Top             =   6120
      Width           =   4485
   End
   Begin VB.Image Image5 
      Height          =   1050
      Left            =   0
      Picture         =   "MainForm.frx":4A018
      Top             =   5160
      Width           =   4485
   End
   Begin VB.Image Image4 
      Height          =   990
      Left            =   0
      Picture         =   "MainForm.frx":4D2D2
      Top             =   4200
      Width           =   4485
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1320
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   299
      ImageHeight     =   68
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   22
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":501BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":555C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":5A9C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":5FDC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":651C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":6A371
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":6F51B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":74B75
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":7A1CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":80C15
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":8765B
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":8DBF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":94187
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":965A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":989C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":995FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":9A237
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":9AE31
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":9BA2B
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":9C6E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":9D3A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":9E1ED
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image3 
      Height          =   1020
      Left            =   0
      Picture         =   "MainForm.frx":9F033
      Top             =   3240
      Width           =   4485
   End
   Begin VB.Image Image2 
      Height          =   1020
      Left            =   0
      Picture         =   "MainForm.frx":A1E03
      Top             =   2280
      Width           =   4485
   End
   Begin VB.Image Image1 
      Height          =   15000
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function buttonsOut()
    Set Image2.Picture = ImageList1.ListImages(1).Picture
    Set Image3.Picture = ImageList1.ListImages(3).Picture
    Set Image4.Picture = ImageList1.ListImages(5).Picture
    Set Image5.Picture = ImageList1.ListImages(7).Picture
    Set Image6.Picture = ImageList1.ListImages(9).Picture
    Set Image7.Picture = ImageList1.ListImages(15).Picture
    Set Image8.Picture = ImageList1.ListImages(17).Picture
    Set Image9.Picture = ImageList1.ListImages(19).Picture
    Set Image11.Picture = ImageList1.ListImages(21).Picture
    Set Image12.Picture = ImageList1.ListImages(11).Picture
    Set Image13.Picture = ImageList1.ListImages(13).Picture
End Function

Private Function deployTables()
    Adodc1.ConnectionString = db
    'Adodc1.RecordSource = "SELECT id as ID, name as 'Subject Name' FROM rooms"
    Dim rs As New ADODB.Recordset
    Dim conn As New ADODB.Connection
    conn = db
    conn.Open
    rs.Open "SELECT * FROM rooms", conn
    
    flex.FormatString = "ID | Name"
    
    Do Until rs.EOF
        flex.Row = 1
        flex.AddItem
        flex.AddItem rs(1), 1
    
    Loop
    DataGrid1.DataSource
    
    
End Function

Private Sub Form_Load()
    deployTables
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
End Sub
Private Sub Image12_Click()
    Label1.Caption = "Reports"
    
    DataGrid1.Visible = False
    DataGrid2.Visible = False
    DataGrid3.Visible = False
    DataGrid4.Visible = False
    DataGrid5.Visible = False
End Sub

Private Sub Image13_Click()
    If MsgBox("Are you sure you wish to exit") = MsgBoxResult.No Then
        E.Cancel = True
    End If
End Sub

Private Sub Image2_Click()
    Label1.Caption = "Rooms"
    
    DataGrid1.Visible = True
    
    DataGrid2.Visible = False
    DataGrid3.Visible = False
    DataGrid4.Visible = False
    DataGrid5.Visible = False

End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set Image2.Picture = ImageList1.ListImages(2).Picture
End Sub

Private Sub Image3_Click()
    Label1.Caption = "Subjects"
    
    DataGrid2.Visible = True
    
    DataGrid1.Visible = False
    DataGrid3.Visible = False
    DataGrid4.Visible = False
    DataGrid5.Visible = False

End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set Image3.Picture = ImageList1.ListImages(4).Picture
End Sub

Private Sub Image4_Click()
    Label1.Caption = "Users"
    
    DataGrid3.Visible = True
    
    DataGrid2.Visible = False
    DataGrid1.Visible = False
    DataGrid4.Visible = False
    DataGrid5.Visible = False

End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set Image4.Picture = ImageList1.ListImages(6).Picture
End Sub

Private Sub Image5_Click()
    Label1.Caption = "Sections"
    
    DataGrid4.Visible = True
    
    DataGrid2.Visible = False
    DataGrid3.Visible = False
    DataGrid1.Visible = False
    DataGrid5.Visible = False
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set Image5.Picture = ImageList1.ListImages(8).Picture
End Sub

Private Sub Image6_Click()
    Label1.Caption = "Schedules"
    
    DataGrid5.Visible = True
    
    DataGrid2.Visible = False
    DataGrid3.Visible = False
    DataGrid4.Visible = False
    DataGrid1.Visible = False
    
End Sub

Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set Image6.Picture = ImageList1.ListImages(10).Picture
End Sub

Private Sub Image12_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set Image12.Picture = ImageList1.ListImages(12).Picture
End Sub
Private Sub Image13_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set Image13.Picture = ImageList1.ListImages(14).Picture
End Sub

Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set Image7.Picture = ImageList1.ListImages(16).Picture
End Sub

Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set Image8.Picture = ImageList1.ListImages(18).Picture
End Sub
Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set Image9.Picture = ImageList1.ListImages(20).Picture
End Sub
Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set Image11.Picture = ImageList1.ListImages(22).Picture
End Sub

