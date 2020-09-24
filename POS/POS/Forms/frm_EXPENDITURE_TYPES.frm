VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_EXPENDITURE_TYPES 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TYPE OF EXPENDITURE"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin OCX.b8Container b8Container1 
      Height          =   4875
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   8599
      BackColor       =   16185592
      Begin OCX.b8Container b8Container2 
         Height          =   3945
         Left            =   90
         TabIndex        =   1
         Top             =   120
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   6959
         BackColor       =   16185592
         Begin OCX.b8Container b8Container5 
            Height          =   3555
            Left            =   3540
            TabIndex        =   4
            Top             =   90
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   6271
            BackColor       =   16185592
            Begin VB.Frame fraDetails 
               Appearance      =   0  'Flat
               BackColor       =   &H00F6F8F8&
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   585
               Left            =   1350
               TabIndex        =   5
               Top             =   180
               Width           =   3495
               Begin VB.TextBox txtXpenditure 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   150
                  TabIndex        =   6
                  Top             =   60
                  Width           =   3255
               End
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Expenditure"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   150
               TabIndex        =   7
               Top             =   240
               Width           =   1275
            End
         End
         Begin OCX.b8Container b8Container4 
            Height          =   3555
            Left            =   120
            TabIndex        =   2
            Top             =   90
            Width           =   3345
            _ExtentX        =   5900
            _ExtentY        =   6271
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
               Height          =   3150
               Left            =   150
               TabIndex        =   3
               Top             =   240
               Width           =   3045
            End
         End
      End
      Begin OCX.b8Container b8Container6 
         Height          =   675
         Left            =   90
         TabIndex        =   8
         Top             =   4110
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   1191
         BackColor       =   16185592
         Begin lvButton.lvButtons_H cmdAddNew 
            Height          =   495
            Left            =   150
            TabIndex        =   9
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
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
            Image           =   "frm_EXPENDITURE_TYPES.frx":0000
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EXPENDITURE_TYPES.frx":0515
         End
         Begin lvButton.lvButtons_H cmdCancel 
            Height          =   495
            Left            =   5625
            TabIndex        =   10
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
            Image           =   "frm_EXPENDITURE_TYPES.frx":082F
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EXPENDITURE_TYPES.frx":0D9E
         End
         Begin lvButton.lvButtons_H cmdEdit 
            Height          =   495
            Left            =   1515
            TabIndex        =   11
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
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
            Image           =   "frm_EXPENDITURE_TYPES.frx":10B8
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EXPENDITURE_TYPES.frx":1435
         End
         Begin lvButton.lvButtons_H cmdClose 
            Height          =   495
            Left            =   6990
            TabIndex        =   12
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
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
            Image           =   "frm_EXPENDITURE_TYPES.frx":174F
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EXPENDITURE_TYPES.frx":1CDF
         End
         Begin lvButton.lvButtons_H cmdSave 
            Height          =   495
            Left            =   4260
            TabIndex        =   13
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
            Image           =   "frm_EXPENDITURE_TYPES.frx":1FF9
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EXPENDITURE_TYPES.frx":2259
         End
         Begin lvButton.lvButtons_H cmdDelete 
            Height          =   495
            Left            =   2880
            TabIndex        =   14
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
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
            Image           =   "frm_EXPENDITURE_TYPES.frx":2573
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EXPENDITURE_TYPES.frx":2A86
         End
      End
   End
End
Attribute VB_Name = "frm_EXPENDITURE_TYPES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnAdd As Boolean
Dim blnEdit As Boolean

Dim lngID As Long

Private Sub cmdCancel_Click()
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
    Call subDisableCtl
    Call sub_LOAD_EXPENDITURE_TYPES
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    blnAdd = False
    blnEdit = True
    Call Mdl_FUNCTIONS.fn_UNSET_CONTROL_COLOR(Me, "All")
    Call subEnableCtl
    txtXpenditure.SetFocus
End Sub

Private Sub cmdaddNew_Click()
    blnAdd = True
    blnEdit = False
    
    Call Mdl_FUNCTIONS.fn_UNSET_CONTROL_COLOR(Me, "All")
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call subEnableCtl
    txtXpenditure.SetFocus
End Sub

Private Sub cmdSave_Click()
'On Error GoTo errHandler
    Dim ctr As Long
    
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtXpenditure, "Please enter the expenditure type") = True Then Exit Sub
    
    With cls_REFERENCES_Obj
        .ExpenditureName = txtXpenditure.Text
        
        If blnAdd = True And blnEdit = False Then
            Call .fn_SAVE_EXPENDITURE
            Else
                Call .fn_UPDATE_EXPENDITURE(lngID)
        End If
        
    End With
    
    
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
    Call subDisableCtl
    Call sub_LOAD_EXPENDITURE_TYPES
    
    Unload Me
    
'Exit Sub
'errHandler:
'    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    Call Mdl_FUNCTIONS.sub_CENTER_FORM(Me)
    Call sub_LOAD_EXPENDITURE_TYPES
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
    Call subDisableCtl
    
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdAddNew, cmdAddNew.hwnd, "Add new type of expenditure details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdEdit, cmdEdit.hwnd, "Edit existing type of expenditure details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdDelete, cmdDelete.hwnd, "Delete existing type of expenditure details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdSave, cmdSave.hwnd, "Save type of expenditure details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCancel, cmdCancel.hwnd, "Cancel the transaction.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdClose, cmdClose.hwnd, "Close the window.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    
End Sub

Private Sub sub_LOAD_EXPENDITURE_TYPES()

    Dim rec As New ADODB.Recordset
    Set rec = cls_REFERENCES_Obj.fn_LOAD_EXPENDITURE_TYPE(0)
    
    lst.Clear
    
    Do While Not rec.EOF
        lst.AddItem rec!ExpenditureName
        lst.ItemData(lst.NewIndex) = rec!ExpenditureTypeID
        rec.MoveNext
    Loop

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

Private Sub Form_Unload(Cancel As Integer)

    Call frm_EXPENDITURES.sub_LOAD_EXPENDITURE_TYPES(frm_EXPENDITURES.cboTypeOfXpenditure)
    
End Sub

Private Sub lst_Click()
    lngID = lst.ItemData(lst.ListIndex)
    Call sub_LOAD_EXPENDITURE_DETAILS(lngID)
End Sub

Private Sub sub_LOAD_EXPENDITURE_DETAILS(lngID As Long)

    Dim rec As New ADODB.Recordset
    Set rec = cls_REFERENCES_Obj.fn_LOAD_EXPENDITURE_TYPE(lngID)
    
    If rec.AbsolutePosition <> -1 Then
        txtXpenditure.Text = rec!ExpenditureName
        cmdEdit.Enabled = True
    End If
End Sub


