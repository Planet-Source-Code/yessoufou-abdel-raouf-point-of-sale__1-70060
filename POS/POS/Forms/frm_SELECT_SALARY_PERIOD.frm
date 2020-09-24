VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_SELECT_SALARY_PERIOD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SALARY PERIOD"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin OCX.b8Container b8Container1 
      Height          =   855
      Left            =   0
      TabIndex        =   9
      Top             =   4470
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   1508
      BackColor       =   16185592
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   3720
         TabIndex        =   10
         Top             =   210
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
         Image           =   "frm_SELECT_SALARY_PERIOD.frx":0000
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_SELECT_SALARY_PERIOD.frx":056F
      End
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   1980
         TabIndex        =   11
         Top             =   210
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
         Image           =   "frm_SELECT_SALARY_PERIOD.frx":0889
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_SELECT_SALARY_PERIOD.frx":0AE9
      End
   End
   Begin OCX.b8Container b8Container4 
      Height          =   4425
      Left            =   3360
      TabIndex        =   0
      Top             =   0
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   7805
      BackColor       =   16185592
      Begin VB.Frame fraDetails 
         BackColor       =   &H00F6F8F8&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   975
         Left            =   90
         TabIndex        =   1
         Top             =   60
         Width           =   3615
         Begin MSComCtl2.DTPicker DTPEndDate 
            Height          =   345
            Left            =   1140
            TabIndex        =   2
            Top             =   570
            Width           =   2175
            _ExtentX        =   3836
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
            Format          =   51118081
            CurrentDate     =   39275
         End
         Begin MSComCtl2.DTPicker DTPStartDate 
            Height          =   345
            Left            =   1140
            TabIndex        =   3
            Top             =   90
            Width           =   2175
            _ExtentX        =   3836
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
            Format          =   51118081
            CurrentDate     =   39275
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
            TabIndex        =   5
            Top             =   630
            Width           =   1215
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
            TabIndex        =   4
            Top             =   150
            Width           =   1215
         End
      End
   End
   Begin OCX.b8Container b8Container3 
      Height          =   4425
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   7805
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
         Height          =   3735
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   3075
      End
      Begin OCX.b8SideTab b8SideTab2 
         Height          =   375
         Left            =   120
         TabIndex        =   8
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
Attribute VB_Name = "frm_SELECT_SALARY_PERIOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngID As Long
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    frm_SALARIES.lngSalaryPeriodID = lngID
    Unload Me
End Sub

Private Sub Form_Load()
'    Mdl_FUNCTIONS.sub_CENTER_FORM (Me)
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
    Call sub_LOAD_SALARY_PERIOD
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
