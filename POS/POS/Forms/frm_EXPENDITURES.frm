VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_EXPENDITURES 
   BackColor       =   &H00FFFFFF&
   Caption         =   "XPENDITURE"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8355
   ScaleWidth      =   12150
   Begin OCX.b8Container b8Container4 
      Height          =   8235
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   14526
      BackColor       =   16185592
      Begin OCX.b8Container b8Container5 
         Height          =   4005
         Left            =   90
         TabIndex        =   8
         Top             =   4140
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   7064
         BackColor       =   16185592
         Begin MSComctlLib.ListView lvw 
            Height          =   2895
            Left            =   210
            TabIndex        =   9
            Top             =   180
            Width           =   11265
            _ExtentX        =   19870
            _ExtentY        =   5106
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
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Date"
               Object.Width           =   2207
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Type Of Expenditure"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Item(s)"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Unit Price"
               Object.Width           =   2294
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Unit/Qty"
               Object.Width           =   1765
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Total"
               Object.Width           =   1765
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Description"
               Object.Width           =   5998
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   795
            Left            =   210
            TabIndex        =   10
            Top             =   3060
            Width           =   11265
            Begin lvButton.lvButtons_H cmdTotal 
               Height          =   525
               Left            =   6150
               TabIndex        =   11
               Top             =   150
               Width           =   5115
               _ExtentX        =   9022
               _ExtentY        =   926
               CapAlign        =   1
               BackStyle       =   3
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFore           =   16711680
               cFHover         =   16711680
               cGradient       =   0
               Mode            =   0
               Value           =   0   'False
               cBack           =   16777215
               mPointer        =   99
               mIcon           =   "frm_EXPENDITURES.frx":0000
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Total"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   0
               TabIndex        =   12
               Top             =   240
               Width           =   1185
            End
         End
      End
      Begin OCX.b8Container b8Container2 
         Height          =   3345
         Left            =   90
         TabIndex        =   13
         Top             =   120
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   5900
         BackColor       =   16185592
         Begin OCX.b8Container b8Container6 
            Height          =   3045
            Left            =   210
            TabIndex        =   14
            Top             =   150
            Width           =   11295
            _ExtentX        =   19923
            _ExtentY        =   5371
            BackColor       =   16185592
            Begin VB.TextBox txtDescription 
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
               Height          =   1455
               Left            =   7470
               MaxLength       =   500
               MultiLine       =   -1  'True
               TabIndex        =   7
               Top             =   1170
               Width           =   3585
            End
            Begin VB.TextBox txtItem 
               Alignment       =   2  'Center
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
               Left            =   2100
               MaxLength       =   100
               TabIndex        =   2
               Top             =   1170
               Width           =   3585
            End
            Begin VB.TextBox txtUnitPrice 
               Alignment       =   1  'Right Justify
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
               Left            =   2100
               MaxLength       =   13
               TabIndex        =   3
               Top             =   1680
               Width           =   3585
            End
            Begin VB.TextBox txtQty 
               Alignment       =   2  'Center
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
               Left            =   2100
               MaxLength       =   10
               TabIndex        =   4
               Top             =   2190
               Width           =   3585
            End
            Begin VB.ComboBox cboTypeOfXpenditure 
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
               Left            =   2100
               Style           =   2  'Dropdown List
               TabIndex        =   0
               Top             =   180
               Width           =   6795
            End
            Begin VB.Frame Frame2 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   7290
               TabIndex        =   15
               Top             =   630
               Width           =   3885
               Begin VB.TextBox txtTotal 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Left            =   180
                  Locked          =   -1  'True
                  TabIndex        =   6
                  Top             =   30
                  Width           =   3585
               End
            End
            Begin MSComCtl2.DTPicker DTXpenditure 
               Height          =   375
               Left            =   2100
               TabIndex        =   1
               Top             =   690
               Width           =   3615
               _ExtentX        =   6376
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
               CustomFormat    =   "dd ddd - MMM - yyyy"
               Format          =   58392579
               CurrentDate     =   39109
            End
            Begin lvButton.lvButtons_H cboLoadCategories 
               Height          =   375
               Left            =   9000
               TabIndex        =   16
               Top             =   150
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   661
               Caption         =   "&Add New ..."
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
               Image           =   "frm_EXPENDITURES.frx":031A
               ImgSize         =   24
               cBack           =   -2147483633
               mPointer        =   99
               mIcon           =   "frm_EXPENDITURES.frx":082F
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Description"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   5910
               TabIndex        =   23
               Top             =   1170
               Width           =   1785
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Item(s)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   150
               TabIndex        =   22
               Top             =   1170
               Width           =   1785
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Total Amount"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   5880
               TabIndex        =   21
               Top             =   690
               Width           =   1785
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Unit/Quantity"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   150
               TabIndex        =   20
               Top             =   2190
               Width           =   1785
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Expenditure Date"
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
               Left            =   150
               TabIndex        =   19
               Top             =   660
               Width           =   1665
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Unit Price/ Amount"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   150
               TabIndex        =   18
               Top             =   1695
               Width           =   1785
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Type of Expenditure"
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
               Left            =   150
               TabIndex        =   17
               Top             =   165
               Width           =   1725
            End
         End
      End
      Begin OCX.b8Container b8Container3 
         Height          =   675
         Left            =   90
         TabIndex        =   24
         Top             =   3480
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   1191
         BackColor       =   16185592
         Begin lvButton.lvButtons_H cmdAddNew 
            Height          =   495
            Left            =   150
            TabIndex        =   25
            Top             =   90
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            Caption         =   "&Add To List"
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
            Image           =   "frm_EXPENDITURES.frx":0B49
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EXPENDITURES.frx":105E
         End
         Begin lvButton.lvButtons_H cmdCancel 
            Height          =   495
            Left            =   7902
            TabIndex        =   26
            Top             =   90
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            Caption         =   "&Clear List"
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
            Image           =   "frm_EXPENDITURES.frx":1378
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EXPENDITURES.frx":18E7
         End
         Begin lvButton.lvButtons_H cmdEdit 
            Height          =   495
            Left            =   2088
            TabIndex        =   27
            Top             =   90
            Width           =   1725
            _ExtentX        =   3043
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
            Image           =   "frm_EXPENDITURES.frx":1C01
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EXPENDITURES.frx":1F7E
         End
         Begin lvButton.lvButtons_H cmdClose 
            Height          =   495
            Left            =   9840
            TabIndex        =   28
            Top             =   90
            Width           =   1725
            _ExtentX        =   3043
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
            Image           =   "frm_EXPENDITURES.frx":2298
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EXPENDITURES.frx":2828
         End
         Begin lvButton.lvButtons_H cmdSave 
            Height          =   495
            Left            =   5964
            TabIndex        =   29
            Top             =   90
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            Caption         =   "&Save Expenditure"
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
            Image           =   "frm_EXPENDITURES.frx":2B42
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EXPENDITURES.frx":2DA2
         End
         Begin lvButton.lvButtons_H cmdDelete 
            Height          =   495
            Left            =   4026
            TabIndex        =   30
            Top             =   90
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            Caption         =   "&Remove From List"
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
            Image           =   "frm_EXPENDITURES.frx":30BC
            ImgSize         =   24
            cBack           =   -2147483633
            mPointer        =   99
            mIcon           =   "frm_EXPENDITURES.frx":35CF
         End
      End
   End
End
Attribute VB_Name = "frm_EXPENDITURES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngExpenditueTypeID As Long

Private Sub cboLoadCategories_Click()
    frm_EXPENDITURE_TYPES.Show 1
End Sub

Private Sub cboTypeOfXpenditure_Click()
    If cboTypeOfXpenditure.ListIndex = -1 Then Exit Sub

    lngExpenditueTypeID = cboTypeOfXpenditure.ItemData(cboTypeOfXpenditure.ListIndex)

    If lvw.ListItems.Count > 0 Then
        If lngExpenditueTypeID <> lvw.ListItems(1).ListSubItems(7).Text Then
            MsgBox "Please save the details before you select another type of Xpenditure", vbInformation, Title
        End If
        Else
            cmdTotal.Caption = ""
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub



Private Sub cmdaddNew_Click()

    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_COMBO(cboTypeOfXpenditure, "Please, Select The Type Of Expenditure") Then GoTo EXITPROCEDURE
    If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtUnitPrice, "Please, Enter The Amount Or Unit Price") Then GoTo EXITPROCEDURE
    
    Set lst = lvw.ListItems.Add(, , DTXpenditure.Value)
            lst.ListSubItems.Add , , cboTypeOfXpenditure.Text
            lst.ListSubItems.Add , , Trim(txtItem.Text)
            lst.ListSubItems.Add , , Val(txtUnitPrice.Text)
            If Trim(txtQty.Text) = "" Or Trim(txtQty.Text) = 0 Then
                lst.ListSubItems.Add , , 1
                Else
                    lst.ListSubItems.Add , , Val(txtQty.Text)
            End If
            lst.ListSubItems.Add , , Val(txtTotal.Text)
            lst.ListSubItems.Add , , Trim(TxtDescription.Text)
    
            lst.ListSubItems.Add , , cboTypeOfXpenditure.ItemData(cboTypeOfXpenditure.ListIndex)
            
            cmdTotal.Caption = Val(cmdTotal.Caption) + Val(txtTotal.Text)
    
    Call sub_CLEAR_FIELD
    
EXITPROCEDURE:
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Call sub_CLEAR_FIELD
    cmdTotal.Caption = 0
    lvw.ListItems.Clear
End Sub

Private Sub cmdDelete_Click()

    If lvw.ListItems.Count = 0 Then GoTo EXITPROCEDURE
    cmdTotal.Caption = Val(cmdTotal.Caption) - Val(lvw.SelectedItem.ListSubItems(5).Text)
    lvw.ListItems.Remove lvw.SelectedItem.Index
    
EXITPROCEDURE:
    Exit Sub
End Sub

Private Sub cmdExpenditers_Click()
    frmAddXpenses.Show 1
End Sub

Private Sub cmdEdit_Click()
    If lvw.ListItems.Count = 0 Then GoTo EXITPROCEDURE
    
    Call sub_CLEAR_FIELD
    
    DTXpenditure.Value = lvw.SelectedItem.Text
    txtItem.Text = lvw.SelectedItem.ListSubItems(2).Text
    txtUnitPrice.Text = lvw.SelectedItem.ListSubItems(3).Text
    txtQty.Text = lvw.SelectedItem.ListSubItems(4).Text
    txtTotal.Text = lvw.SelectedItem.ListSubItems(5).Text
    TxtDescription.Text = lvw.SelectedItem.ListSubItems(6).Text
    
    cmdTotal.Caption = Val(cmdTotal.Caption) - Val(lvw.SelectedItem.ListSubItems(5).Text)
    lvw.ListItems.Remove lvw.SelectedItem.Index
    
    
    
EXITPROCEDURE:
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    
    Dim ctr As Long
       
    If lvw.ListItems.Count = 0 Then
        MsgBox "Please add your details to the list before you save", vbExclamation, Title
        GoTo EXITPROCEDURE
    End If
    
    
    If lvw.ListItems.Count = 0 Then
        If Mdl_FUNCTIONS.fn_CHECK_EMPTY_COMBO(cboTypeOfXpenditure, "Please, Select The Type Of Expenditure") Then GoTo EXITPROCEDURE
        If Mdl_FUNCTIONS.fn_CHECK_EMPTY_TEXT_BOX(txtUnitPrice, "Please, Enter The Amount Or Unit Price") Then GoTo EXITPROCEDURE
    End If
    
    With cls_EXPENDITURES_Obj
        .ExpenditureID = .fn_AUTOGEN
        .ExpenditureTypeID = lngExpenditueTypeID
        .ExpenditureDate = DTXpenditure.Value
        .ExpenditureTime = Now
        
        If lvw.ListItems.Count = 0 Then
            .ExpenditureTotal = Val(txtTotal.Text)
            Else
                .ExpenditureTotal = Val(cmdTotal.Caption)
        End If
        
        .fn_SAVE_EXPENDITURES_RECORDS
    End With
    
    
    For ctr = 1 To lvw.ListItems.Count
        With cls_EXPENDITURES_Obj
'            .ExpenditureID = lngExpenditueTypeID
            .ExpenditureItems = lvw.ListItems(ctr).ListSubItems(2).Text
            .ExpenditurePrice = lvw.ListItems(ctr).ListSubItems(3).Text
            .ExpenditureQty = lvw.ListItems(ctr).ListSubItems(4).Text
            .ExpenditureDetailsTotal = lvw.ListItems(ctr).ListSubItems(5).Text
            .ExpenditureDescription = lvw.ListItems(ctr).ListSubItems(6).Text
            
            .fn_SAVE_EXPENDITURES_DETAILS_RECORDS
        End With
        
    Next
    
    MsgBox "Transaction successfull", vbInformation, Title
    
    Call sub_CLEAR_FIELD
    lvw.ListItems.Clear
    cmdTotal.Caption = ""
    
EXITPROCEDURE:
    Exit Sub
End Sub

Private Sub cmdUpdate_Click()
    
    If lvw.ListItems.Count = 0 Then GoTo EXITPROCEDURE
    Call MdlFunctions.fnEmptyFields(Me)
    
    DTXpenditure.Value = lvw.SelectedItem.Text
    txtItem.Text = lvw.SelectedItem.ListSubItems(2).Text
    txtUnitPrice.Text = lvw.SelectedItem.ListSubItems(3).Text
    txtQty.Text = lvw.SelectedItem.ListSubItems(4).Text
    txtTotal.Text = lvw.SelectedItem.ListSubItems(5).Text
    TxtDescription.Text = lvw.SelectedItem.ListSubItems(6).Text
    
    cmdTotal.Caption = Val(cmdTotal.Caption) - Val(lvw.SelectedItem.ListSubItems(5).Text)
    lvw.ListItems.Remove lvw.SelectedItem.Index
    
EXITPROCEDURE:
    Exit Sub
End Sub

Private Sub Form_Load()
    Call Mdl_FUNCTIONS.sub_FORM_SIZE(Me)
    DTXpenditure.Value = Date
    Move 0, 0
    Call sub_LOAD_EXPENDITURE_TYPES(cboTypeOfXpenditure)
    
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdAddNew, cmdAddNew.hwnd, "Add expenditure to the list.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdEdit, cmdEdit.hwnd, "Edit existing expenditure details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdDelete, cmdDelete.hwnd, "Remove expenditure from the list.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdSave, cmdSave.hwnd, "Save expenditure details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCancel, cmdCancel.hwnd, "Clear expenditure list.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdClose, cmdClose.hwnd, "Close the window.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor

End Sub

Public Sub sub_LOAD_EXPENDITURE_TYPES(Optional cbo As ComboBox)

    Dim rec As New ADODB.Recordset
    Set rec = cls_REFERENCES_Obj.fn_LOAD_EXPENDITURE_TYPE(0)
    
    cbo.Clear
    
    Do While Not rec.EOF
        cbo.AddItem rec!ExpenditureName
        cbo.ItemData(cbo.NewIndex) = rec!ExpenditureTypeID
        rec.MoveNext
    Loop

End Sub

Private Sub txtItem_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_ALPHABET_ONLY(KeyAscii)
End Sub

Private Sub txtItem_Validate(Cancel As Boolean)
    txtItem.Text = StrConv(txtItem.Text, vbProperCase)
End Sub

Private Sub txtQty_Change()
    If Trim(txtQty.Text) = 0 Or Trim(txtQty.Text) = "" Then
        txtTotal.Text = Val(txtUnitPrice.Text)
        Else
            txtTotal.Text = Val(txtUnitPrice.Text) * Val(txtQty.Text)
    End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub txtUnitPrice_Change()
    If Trim(txtQty.Text) = 0 Or Trim(txtQty.Text) = "" Then
        txtTotal.Text = Val(txtUnitPrice.Text)
        Else
            txtTotal.Text = Val(txtUnitPrice.Text) * Val(txtQty.Text)
    End If
End Sub

Private Sub txtUnitPrice_KeyPress(KeyAscii As Integer)
    Call Mdl_FUNCTIONS.sub_NUMERIC_ONLY(KeyAscii)
End Sub

Private Sub sub_CLEAR_FIELD()
    
'    cboTypeOfXpenditure.ListIndex = -1
    DTXpenditure.Value = Date
    txtItem.Text = ""
    txtUnitPrice.Text = ""
    txtQty.Text = ""
    txtTotal.Text = ""
    TxtDescription.Text = ""
    
End Sub
