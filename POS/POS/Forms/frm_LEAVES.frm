VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_LEAVES 
   BackColor       =   &H00FFFFFF&
   Caption         =   "EMPLOYEE LEAVES"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   12180
   Begin OCX.b8Container ContainerList 
      Height          =   7845
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   13838
      BackColor       =   16185592
      Begin OCX.b8Container b8Container4 
         Height          =   7665
         Left            =   3090
         TabIndex        =   1
         Top             =   90
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   13520
         BackColor       =   16185592
         Begin MSComctlLib.ListView lvw 
            Height          =   7365
            Left            =   150
            TabIndex        =   10
            Top             =   150
            Width           =   8505
            _ExtentX        =   15002
            _ExtentY        =   12991
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Leave Type"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Approved Date"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Start Date"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "End Date"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Notes"
               Object.Width           =   5292
            EndProperty
         End
      End
      Begin OCX.b8Container b8Container3 
         Height          =   7665
         Left            =   60
         TabIndex        =   2
         Top             =   90
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   13520
         BackColor       =   16185592
         Begin VB.ListBox lstEmployees 
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
            TabIndex        =   3
            Top             =   480
            Width           =   2775
         End
         Begin OCX.b8SideTab b8SideTab2 
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   150
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            Caption         =   "Employees"
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
      TabIndex        =   5
      Top             =   7860
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   1191
      BackColor       =   16185592
      Begin lvButton.lvButtons_H cmdAddNew 
         Height          =   495
         Left            =   180
         TabIndex        =   6
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
         Image           =   "frm_LEAVES.frx":0000
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_LEAVES.frx":0515
      End
      Begin lvButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   3520
         TabIndex        =   7
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
         Image           =   "frm_LEAVES.frx":082F
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_LEAVES.frx":0BAC
      End
      Begin lvButton.lvButtons_H cmdClose 
         Height          =   495
         Left            =   10200
         TabIndex        =   8
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
         Image           =   "frm_LEAVES.frx":0EC6
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_LEAVES.frx":1456
      End
      Begin lvButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   6860
         TabIndex        =   9
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
         Image           =   "frm_LEAVES.frx":1770
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_LEAVES.frx":1C83
      End
   End
End
Attribute VB_Name = "frm_LEAVES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lngID As Long

Dim strPictureName As String
Public blnLeaveModify As Boolean
Public blnLeaveAdd As Boolean
Dim lngSalaryID As Long
Public lngLeaveDetailsID As Long

Private Sub cmdClose_Click()
    Unload Me
End Sub



Private Sub cmdaddNew_Click()
        
    blnLeaveAdd = True
    blnLeaveModify = False
    
    frm_ADD_LEAVES.Show 1
    
End Sub


Private Sub cmdDelete_Click()
    If MsgBox("Are you sure you want to delete this leave details?", 4 + 32, Title) = vbYes Then
        Call cls_LEAVES_Obj.fn_DELETE_LEAVES(lngID, lngLeaveDetailsID)
        lvw.ListItems.Remove lvw.SelectedItem.Index
        Call sub_LOAD_LEAVES_DETAILS(lngID)
    End If
    
End Sub

Private Sub cmdEdit_Click()
    lngLeaveDetailsID = 0
    If lvw.ListItems.Count = 0 Then Exit Sub
    blnLeaveAdd = False
    blnLeaveModify = True
    lngLeaveDetailsID = lvw.SelectedItem.Text
    With frm_ADD_LEAVES
        Call .sub_LOAD_LEAVES_DETAILS(lngID, lngLeaveDetailsID)
        .Show 1
    End With
End Sub



Private Sub Form_Load()
    Call Mdl_FUNCTIONS.sub_FORM_SIZE(Me)
    Move 0, 0
    Call fn_SET_CONTROL_COLOR(Me, "All")
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call sub_LOAD_EMPLOYEES(lstEmployees)
    
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdAddNew, cmdAddNew.hwnd, "Add new leave details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdEdit, cmdEdit.hwnd, "Edit existing leave details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdDelete, cmdDelete.hwnd, "Delete existing leave details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdClose, cmdClose.hwnd, "Close the window.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor

End Sub

Private Sub sub_LOAD_EMPLOYEES(Optional lst As ListBox)

    Dim rec As New ADODB.Recordset
    Set rec = cls_EMPLOYEES_Obj.fn_LOAD_EMPLOYEES
    
    lst.Clear
    
    Do While Not rec.EOF
        lst.AddItem rec!FirstName & " " & rec!LastName
        lst.ItemData(lst.NewIndex) = rec!EmployeeID
        rec.MoveNext
    Loop

End Sub


Private Sub Form_Unload(Cancel As Integer)
    blnLeaveAdd = False
    blnLeaveModify = False
End Sub

Private Sub lstEmployees_Click()
    lngID = lstEmployees.ItemData(lstEmployees.ListIndex)
    Call sub_LOAD_LEAVES_DETAILS(lngID)
End Sub

Public Sub sub_LOAD_LEAVES_DETAILS(lngEmployeeID As Long)

    Dim rec As New ADODB.Recordset
    
    Set rec = cls_LEAVES_Obj.fn_LOAD_LEAVES(lngEmployeeID)
    lvw.ListItems.Clear
    If rec.AbsolutePosition <> -1 Then
        Do While Not rec.EOF
            Set lstItem = lvw.ListItems.Add(, , rec!LeaveDetailsID)
                lstItem.ListSubItems.Add , , Trim(rec!LeaveName)
                lstItem.ListSubItems.Add , , Trim(rec!ApprovedDate)
                lstItem.ListSubItems.Add , , Trim(rec!StartDate)
                lstItem.ListSubItems.Add , , Trim(rec!EndDate)
                lstItem.ListSubItems.Add , , Trim(rec!Notes)
        rec.MoveNext
        Loop
    End If
    
End Sub

Private Sub lvw_Click()
    lngLeaveDetailsID = lvw.SelectedItem.Text
End Sub
