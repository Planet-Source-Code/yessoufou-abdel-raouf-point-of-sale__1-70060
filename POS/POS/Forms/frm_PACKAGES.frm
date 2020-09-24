VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_PACKAGES 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PACKAGES"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00F6F8F8&
      Height          =   2895
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   2385
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
         Height          =   2565
         Left            =   90
         TabIndex        =   9
         Top             =   180
         Width           =   2175
      End
   End
   Begin VB.Frame fraDetails 
      BackColor       =   &H00F6F8F8&
      Height          =   2895
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Width           =   5355
      Begin VB.TextBox txtPackageName 
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
         Height          =   345
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   2565
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Package Name"
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
         Left            =   150
         TabIndex        =   3
         Top             =   240
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F6F8F8&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   2910
      Width           =   7755
      Begin lvButton.lvButtons_H cmdAddNew 
         Height          =   495
         Left            =   90
         TabIndex        =   4
         Top             =   150
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
         Image           =   "frm_PACKAGES.frx":0000
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_PACKAGES.frx":0515
      End
      Begin lvButton.lvButtons_H cmdEdit 
         Height          =   495
         Left            =   1620
         TabIndex        =   5
         Top             =   150
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
         Image           =   "frm_PACKAGES.frx":082F
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_PACKAGES.frx":0BAC
      End
      Begin lvButton.lvButtons_H cmdSave 
         Height          =   495
         Left            =   4695
         TabIndex        =   6
         Top             =   150
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
         Image           =   "frm_PACKAGES.frx":0EC6
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_PACKAGES.frx":1126
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   495
         Left            =   6210
         TabIndex        =   7
         Top             =   150
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
         Image           =   "frm_PACKAGES.frx":1440
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_PACKAGES.frx":19AF
      End
      Begin lvButton.lvButtons_H cmdDelete 
         Height          =   495
         Left            =   3150
         TabIndex        =   10
         Top             =   150
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
         Image           =   "frm_PACKAGES.frx":1CC9
         ImgSize         =   24
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frm_PACKAGES.frx":21DC
      End
   End
End
Attribute VB_Name = "frm_PACKAGES"
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
    Call sub_LOAD_PACKAGES
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    blnAdd = False
    blnEdit = True
    Call Mdl_FUNCTIONS.fn_UNSET_CONTROL_COLOR(Me, "All")
    Call subEnableCtl
    txtPackageName.SetFocus
End Sub

Private Sub cmdaddNew_Click()
    blnAdd = True
    blnEdit = False
    
    Call Mdl_FUNCTIONS.fn_UNSET_CONTROL_COLOR(Me, "All")
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call subEnableCtl
    txtPackageName.SetFocus
End Sub

Private Sub cmdSave_Click()
On Error GoTo errHandler
    Dim ctr As Long
    
    If txtPackageName.Text = "" Then
        MsgBox "Please enter the aircraft type", vbExclamation, "Aircraft Type"
        txtPackageName.SetFocus
        Exit Sub
    End If
    
    With cls_REFERENCES_Obj
        .PackageName = txtPackageName.Text
        
        If blnAdd = True And blnEdit = False Then
            Call .fn_SAVE_PACKAGE
            Else
                Call .fn_UPDATE_PACKAGE(lngID)
        End If
        
    End With
    
    
    Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
    Call subDisableCtl
    Call sub_LOAD_PACKAGES
    
    Unload Me
    
Exit Sub
errHandler:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    Call Mdl_FUNCTIONS.sub_CENTER_FORM(Me)
    Call sub_LOAD_PACKAGES
    Call Mdl_FUNCTIONS.fn_SET_CONTROL_COLOR(Me, "All")
    Call subDisableCtl
    
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdAddNew, cmdAddNew.hwnd, "Add new package details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdEdit, cmdEdit.hwnd, "Edit existing package details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
'    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdDeleteCategory, cmdDeleteCategory.hwnd, "Delete existing category details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdSave, cmdSave.hwnd, "Save package details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCancel, cmdCancel.hwnd, "Cancel the transaction.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
'    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCloseCategory, cmdCloseCategory.hwnd, "Close the window.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
   
End Sub

Private Sub sub_LOAD_PACKAGES()

    Dim rec As New ADODB.Recordset
    Set rec = cls_REFERENCES_Obj.fn_LOAD_PACKAGES(0)
    
    lst.Clear
    
    If rec.AbsolutePosition <> -1 Then
        Do While Not rec.EOF
            lst.AddItem rec!PackageName
            lst.ItemData(lst.NewIndex) = rec!PackageID
            rec.MoveNext
        Loop
    End If
    
End Sub

Private Sub subDisableCtl()

    cmdAddNew.Enabled = True
'    cmdClose.Enabled = True
    
    cmdEdit.Enabled = False
    cmdSave.Enabled = False
'    cmdDelete.Enabled = False
    cmdCancel.Enabled = False
    
    fraDetails.Enabled = False
    lst.Enabled = True
    
End Sub

Private Sub subEnableCtl()

    cmdAddNew.Enabled = False
'    cmdClose.Enabled = False
    
    cmdEdit.Enabled = False
    cmdSave.Enabled = True
'    cmdDelete.Enabled = False
    cmdCancel.Enabled = True
    fraDetails.Enabled = True
    
    lst.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If blnProductPackage = True Then
        Call frm_PRODUCTS.sub_LOAD_PACKAGES
        ElseIf blnAddPackage = True Then
            Call frm_ADD_PACKAGES.sub_LOAD_PACKAGES
    End If
    
    blnProductPackage = False
    blnAddPackage = False
    
End Sub

Private Sub lst_Click()
    lngID = lst.ItemData(lst.ListIndex)
    Call sub_LOAD_PACKAGES_DETAILS(lngID)
End Sub

Private Sub sub_LOAD_PACKAGES_DETAILS(lngID As Long)

    Dim rec As New ADODB.Recordset
    Set rec = cls_REFERENCES_Obj.fn_LOAD_PACKAGES(lngID)
    
    If rec.AbsolutePosition <> -1 Then
        txtPackageName.Text = rec!PackageName
        cmdEdit.Enabled = True
    End If
End Sub

