VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frm_EMPLOYEES_REP 
   Caption         =   "EMPLOYEES REPORT"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer 
      Height          =   11025
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
Attribute VB_Name = "frm_EMPLOYEES_REP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub sub_LOAD_REPORT(lngID As Long)
On Error GoTo errHandler

    Dim rec As New ADODB.Recordset
    Dim myApp As New CRAXDRT.Application
    Dim myAppReport As CRAXDRT.Report
        
    Set rec = cls_EMPLOYEES_Obj.fn_LOAD_EMPLOYEES(lngID)
    
    Set myAppReport = myApp.OpenReport(App.Path & "\Reports\Employees_Rep.rpt")
    Call myAppReport.Database.SetDataSource(rec, 3, 1)
    myAppReport.FormulaFields.GetItemByName("CompanyName").Text = "'" & strCompanyName & "'"
    CRViewer.ReportSource = myAppReport
    CRViewer.ViewReport


    
EXITPROCEDURE:
    Exit Sub

errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
    GoTo EXITPROCEDURE
End Sub

Private Sub Form_Load()
'    Call Mdl_FUNCTIONS.sub_FORM_SIZE(Me)
    Call sub_LOAD_REPORT(0)
End Sub



