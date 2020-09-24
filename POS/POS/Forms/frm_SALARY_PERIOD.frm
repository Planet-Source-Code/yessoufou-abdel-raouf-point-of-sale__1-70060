VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_SALARY_PERIOD 
   Caption         =   "SALARY PERIOD"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8580
   ScaleWidth      =   12060
   Begin OCX.b8Container ContainerList 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   13996
      BackColor       =   16185592
      Begin OCX.b8Container b8Container4 
         Height          =   7725
         Left            =   3480
         TabIndex        =   1
         Top             =   120
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   13626
         BackColor       =   16185592
         Begin VB.Frame fraDetails 
            BackColor       =   &H00F6F8F8&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   7545
            Left            =   90
            TabIndex        =   2
            Top             =   60
            Width           =   8175
            Begin MSComCtl2.DTPicker DTPEndDate 
               Height          =   345
               Left            =   1410
               TabIndex        =   13
               Top             =   570
               Width           =   2955
               _ExtentX        =   5212
               _ExtentY        =   609
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
               Format          =   20643841
               CurrentDate     =   39275
            End
            Begin MSComCtl2.DTPicker DTPStartDate 
               Height          =   345
               Left            =   1410
               TabIndex        =   14
               Top             =   90
               Width           =   2955
               _ExtentX        =   5212
               _ExtentY        =   609
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
               Format          =   20643841
               CurrentDate     =   39275
            End
            Begin VB.Label Label1 
               BackColor       =   &H00F6F8F8&
               Caption         =   "Start Date"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   120
               TabIndex        =   16
               Top             =   150
               Width           =   1215
            End
            Begin VB.Label Label2 
               BackColor       =   &H00F6F8F8&
               Caption         =   "End Date"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   120
               TabIndex        =   15
               Top             =   630
               Width           =   1215
            End
         End
      End
      Begin OCX.b8Container b8Container3 
         Height          =   7725
         Left            =   60
         TabIndex        =   3
         Top             =   120
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   13626
         BackColor       =   16185592
         Begin VB.ListBox lst 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7050
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   480
            Width           =   3075
         End
         Begin OCX.b8SideTab b8SideTab2 
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   150
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   661
            Caption         =   "Salary Periods"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Tahoma"
            FontSize        =   8.25
            ForeColor       =   16777215
         End
      End
   End
   Begin OCX.b8Container b8Container2 
      Height          =   675
      Left            =   0
      TabIndex        =   6
      Top             =   7950
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   1191
      BackColor       =   16185592
      Begin lvButton.lvButtons_H cmdAddNew 
         Height          =   495
         Left            =   180
         TabIndex        =   7
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Caption         =   "&Add New"
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
         Image           =   "frm_SALARY_PERIOD.frx":0000
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_SALARY_PERIOD.frx":0515
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   8196
         TabIndex        =   8
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Caption         =   "&Cancel"
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
         Image           =   "frm_SALARY_PERIOD.frx":082F
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_SALARY_PERIOD.frx":0D9E
      End
      Begin lvButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   2184
         TabIndex        =   9
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Caption         =   "&Edit"
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
         Image           =   "frm_SALARY_PERIOD.frx":10B8
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_SALARY_PERIOD.frx":1435
      End
      Begin lvButton.lvButtons_H cmdClose 
         Height          =   495
         Left            =   10200
         TabIndex        =   10
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Caption         =   "Cl&ose"
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
         Image           =   "frm_SALARY_PERIOD.frx":174F
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_SALARY_PERIOD.frx":1CDF
      End
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   6192
         TabIndex        =   11
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Caption         =   "&Save"
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
         Image           =   "frm_SALARY_PERIOD.frx":1FF9
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_SALARY_PERIOD.frx":2259
      End
      Begin lvButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   4188
         TabIndex        =   12
         Top             =   90
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Caption         =   "&Delete"
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
         Image           =   "frm_SALARY_PERIOD.frx":2573
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_SALARY_PERIOD.frx":2A86
      End
   End
End
Attribute VB_Name = "frm_SALARY_PERIOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngID As Long
Dim blnEdit As Boolean
Dim blnAdd As Boolean

Private Sub cmdDelete_Click()
On Error GoTo errHandler
    If MsgBox("Are you sure you want to Delete Salary Period? ", vbQuestion + vbYesNo, Title) = vbNo Then Exit Sub

    With cls_SALARY_PERIOD_Obj
        Call .fn_DELETE_SALARY_PERIOD(lngID)
        Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
        Call sub_LOAD_SALARY_PERIOD
    End With

EXITPROCEDURE:
    Exit Sub

errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
    GoTo EXITPROCEDURE
End Sub

Private Sub cmdEdit_Click()
    blnAdd = False
    blnEdit = True
    Call Mdl_FUNCTIONS.fn_UNSET_CONTROL_COLOR(Me, "All")
    Call subEnableCtl
End Sub

Private Sub cmdSave_Click()
On Error GoTo errHandler

    With cls_SALARY_PERIOD_Obj
        .StartDate = DTPStartDate.Value
        .EndDate = DTPEndDate.Value
        
        If blnAdd = True And blnEdit = False Then
            Call .fn_SAVE_SALARY_PERIOD
            Else
                Call .fn_UPDATE_SALARY_PERIOD(lngID)
        End If
    End With

    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
    Call subDisableCtl
    Call sub_LOAD_SALARY_PERIOD

EXITPROCEDURE:
    Exit Sub

errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
    GoTo EXITPROCEDURE
End Sub


Private Sub Form_Load()
    Call Mdl_FUNCTIONS.sub_FORM_SIZE(Me)
    Move 0, 0
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
    Call sub_LOAD_SALARY_PERIOD

    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdAddNew, cmdAddNew.hwnd, "Add new customer details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdEdit, cmdEdit.hwnd, "Edit existing customer details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdDelete, cmdDelete.hwnd, "Delete existing customer details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdSave, cmdSave.hwnd, "Save customer details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCancel, cmdCancel.hwnd, "Cancel the transaction.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdClose, cmdClose.hwnd, "Close the window.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
   
   
End Sub


Private Sub sub_LOAD_SALARY_PERIOD()

    Dim rec As New ADODB.Recordset
    
    Set rec = cls_SALARY_PERIOD_Obj.fN_LOAD_SALARY_PERIOD(0)
    lst.Clear
    If rec.AbsolutePosition <> -1 Then
        Do While Not rec.EOF
            lst.AddItem "From " & rec!StartDate & " To " & rec!EndDate
            lst.ItemData(lst.NewIndex) = rec!SalaryPeriodID
            rec.MoveNext
        Loop
    End If

End Sub

Private Sub lst_Click()
    lngID = lst.ItemData(lst.ListIndex)
    Call sub_LOAD_SALARY_PERIOD_DETAILS(lngID)
End Sub

Private Sub sub_LOAD_SALARY_PERIOD_DETAILS(lngID As Long)

    Dim rec As New ADODB.Recordset
    
    Set rec = cls_SALARY_PERIOD_Obj.fN_LOAD_SALARY_PERIOD(lngID)
    If rec.AbsolutePosition <> -1 Then
        DTPStartDate.Value = rec!StartDate
        DTPEndDate.Value = rec!EndDate
    End If

End Sub

Private Sub cmdaddNew_Click()

    blnAdd = True
    blnEdit = False
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call Mdl_FUNCTIONS.fn_UNSET_CONTROL_COLOR(Me, "All")
    Call subEnableCtl
    
End Sub

Private Sub cmdCancel_Click()
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
    Call subDisableCtl
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub subDisableCtl()

    cmdAddNew.Enabled = True
    cmdClose.Enabled = True

    cmdEdit.Enabled = False
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    cmdCancel.Enabled = False

    fraDetails.Enabled = False
    lst.Enabled = True
End Sub

Private Sub subEnableCtl()

    cmdAddNew.Enabled = False
    cmdClose.Enabled = False

    cmdEdit.Enabled = False
    cmdSave.Enabled = True
    cmdDelete.Enabled = False
    cmdCancel.Enabled = True
    fraDetails.Enabled = True

    lst.Enabled = False
End Sub

Private Sub lst_DblClick()
    Call cls_SALARIES_Obj.fn_CHECK_IF_PAID(frm_SALARIES.lngSalaryID, lngID)
    If blnCheckIfSalaryPaid = True Then
        MsgBox "This employee has already been paid", vbExclamation, Title
        Exit Sub
    End If
    frm_SALARIES.lngSalaryPeriodID = lngID
    Unload Me
End Sub
