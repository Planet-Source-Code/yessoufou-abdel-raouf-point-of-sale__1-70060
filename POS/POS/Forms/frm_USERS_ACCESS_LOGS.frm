VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_USERS_ACCESS_LOGS 
   Caption         =   "USERS ACCES LOGS"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin OCX.b8Container b8Container1 
      Height          =   705
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   1244
      BackColor       =   16185592
      Begin MSComCtl2.DTPicker DTPFrom 
         Height          =   375
         Left            =   1410
         TabIndex        =   1
         Top             =   210
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   58654721
         CurrentDate     =   39262
      End
      Begin MSComCtl2.DTPicker DTPTo 
         Height          =   375
         Left            =   4200
         TabIndex        =   2
         Top             =   210
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   58654721
         CurrentDate     =   39262
      End
      Begin lvButton.lvButtons_H cmdLoad 
         Height          =   495
         Left            =   6360
         TabIndex        =   3
         Top             =   120
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   873
         Caption         =   "&Load"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgSize         =   32
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_USERS_ACCESS_LOGS.frx":0000
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3720
         TabIndex        =   5
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   210
         TabIndex        =   4
         Top             =   240
         Width           =   1035
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer 
      Height          =   10275
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   15255
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frm_USERS_ACCESS_LOGS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub sub_LOAD_REPORT(Optional lngID As Long, Optional DTPFrom As Date, Optional DTPTo As Date)
On Error GoTo errHandler

    Dim rec As New ADODB.Recordset
    Dim myApp As New CRAXDRT.Application
    Dim myAppReport As CRAXDRT.Report
        
    Set rec = cls_USERS_ACCESS_LOG_Obj.fn_LOAD_USERS_ACCESS_LOGS(lngID, DTPFrom, DTPTo)
    
    Set myAppReport = myApp.OpenReport(App.Path & "\Reports\Users_Access_logs_Rep.rpt")
    Call myAppReport.Database.SetDataSource(rec, 3, 1)
    myAppReport.FormulaFields.GetItemByName("CompanyName").Text = "'" & strCompanyName & "'"
    CRViewer.ReportSource = myAppReport
    CRViewer.ViewReport

    Exit Sub
    
EXITPROCEDURE:
    Exit Sub
    
errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
    GoTo EXITPROCEDURE
End Sub

Private Sub cmdLoad_Click()
    Call sub_LOAD_REPORT(1, DTPFrom.Value, DTPTo.Value)
End Sub

Private Sub Form_Load()
    DTPFrom.Value = Date
    DTPTo.Value = Date
    Call sub_LOAD_REPORT(0)
End Sub

