VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_ADD_LEAVES 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ADD/EDIT LEAVES"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin OCX.b8Container b8Container1 
      Height          =   3315
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5847
      BackColor       =   16185592
      Begin VB.Frame fraEmployeeDetails 
         BackColor       =   &H00F6F8F8&
         BorderStyle     =   0  'None
         Height          =   3045
         Left            =   120
         TabIndex        =   6
         Top             =   90
         Width           =   5295
         Begin VB.ComboBox cboLeaveType 
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
            Height          =   315
            Left            =   1740
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   120
            Width           =   3405
         End
         Begin VB.TextBox txtNotes 
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
            Height          =   945
            Left            =   1740
            TabIndex        =   4
            Top             =   1950
            Width           =   3405
         End
         Begin MSComCtl2.DTPicker DTApprovedDate 
            Height          =   375
            Left            =   1740
            TabIndex        =   1
            Top             =   570
            Width           =   3405
            _ExtentX        =   6006
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
            Format          =   58916865
            CurrentDate     =   39100
         End
         Begin MSComCtl2.DTPicker DTStartDate 
            Height          =   375
            Left            =   1740
            TabIndex        =   2
            Top             =   1020
            Width           =   3405
            _ExtentX        =   6006
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
            Format          =   58916865
            CurrentDate     =   39100
         End
         Begin MSComCtl2.DTPicker DTEndDate 
            Height          =   375
            Left            =   1740
            TabIndex        =   3
            Top             =   1470
            Width           =   3405
            _ExtentX        =   6006
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
            Format          =   58916865
            CurrentDate     =   39100
         End
         Begin MSComDlg.CommonDialog PictureDlg 
            Left            =   7680
            Top             =   3360
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Leave Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   11
            Top             =   120
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Approved Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   10
            Top             =   570
            Width           =   1290
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Left            =   150
            TabIndex        =   9
            Top             =   1020
            Width           =   885
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Left            =   150
            TabIndex        =   8
            Top             =   1470
            Width           =   810
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Notes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   7
            Top             =   1950
            Width           =   510
         End
      End
   End
   Begin OCX.b8Container b8Container6 
      Height          =   675
      Left            =   0
      TabIndex        =   12
      Top             =   3330
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1191
      BackColor       =   16185592
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   4005
         TabIndex        =   13
         Top             =   90
         Width           =   1275
         _ExtentX        =   2249
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
         Image           =   "frm_ADD_LEAVES.frx":0000
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_ADD_LEAVES.frx":056F
      End
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   2460
         TabIndex        =   14
         Top             =   90
         Width           =   1275
         _ExtentX        =   2249
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
         Image           =   "frm_ADD_LEAVES.frx":0889
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_ADD_LEAVES.frx":0AE9
      End
   End
End
Attribute VB_Name = "frm_ADD_LEAVES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngLeaveID As Long
Private Sub sub_LOAD_LEAVES(Optional cbo As ComboBox)

    Dim rec As New ADODB.Recordset
    Set rec = cls_REFERENCES_Obj.fn_LOAD_LEAVES
    
    cbo.Clear
    
    Do While Not rec.EOF
        cbo.AddItem rec!LeaveName
        cbo.ItemData(cbo.NewIndex) = rec!LeaveID
        rec.MoveNext
    Loop

End Sub

Private Sub cboLeaveType_Click()
    If cboLeaveType.ListIndex = -1 Then Exit Sub
    lngLeaveID = cboLeaveType.ItemData(cboLeaveType.ListIndex)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    
    If Mdl_FUNCTIONS.fn_REQUIRE_COMBO_FIELD(cboLeaveType, "Kindly select the leave type.") Then Exit Sub
    If DTStartDate.Value > DTEndDate.Value Then
        MsgBox "The start date should not be greater than today", vbExclamation, Title
        DTStartDate.SetFocus
        Exit Sub
    End If

    With cls_LEAVES_Obj
        .LeaveID = lngLeaveID
        .ApprovedDate = DTApprovedDate.Value
        .StartDate = DTStartDate.Value
        .EndDate = DTEndDate.Value
        .Notes = Trim(txtNotes.Text)
        .EmployeeID = frm_LEAVES.lngID
        If frm_LEAVES.blnLeaveAdd = True Then
            Call .fn_SAVE_LEAVES_RECORDS
        Else
            Call .fn_UPDATE_LEAVES_RECORDS(frm_LEAVES.lngLeaveDetailsID)
        End If
        
    End With
    MsgBox "Leave details saved successfully", vbInformation, Title
    frm_LEAVES.sub_LOAD_LEAVES_DETAILS (frm_LEAVES.lngID)
    Unload Me
End Sub

Private Sub Form_Load()
    Call Mdl_FUNCTIONS.sub_CENTER_FORM(Me)
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call sub_LOAD_LEAVES(cboLeaveType)
End Sub

Public Sub sub_LOAD_LEAVES_DETAILS(lngEmployeeID As Long, lngLeaveDetailsID As Long)

    Dim rec As New ADODB.Recordset
    
    Set rec = cls_LEAVES_Obj.fn_LOAD_LEAVES_DETAILS(lngEmployeeID, lngLeaveDetailsID)
      
    If rec.AbsolutePosition <> -1 Then
    
        lngLeaveID = rec!LeaveID
        cboLeaveType.ListIndex = fn_GET_LIST_INDEX(cboLeaveType, rec!LeaveID)
        DTApprovedDate.Value = rec!ApprovedDate
        DTStartDate.Value = rec!StartDate
        DTEndDate.Value = rec!EndDate
        txtNotes.Text = rec!Notes

    End If

    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frm_LEAVES.blnLeaveAdd = False
    frm_LEAVES.blnLeaveModify = False
End Sub
