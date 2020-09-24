VERSION 5.00
Object = "{3B67AE8A-5616-40F4-93CC-FC55261F4C22}#1.0#0"; "OCX Control.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frm_CATEGORIES 
   Caption         =   "CATEGORIES"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   12090
   Begin OCX.b8Container b8Container1 
      Height          =   8715
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   15372
      BorderColor     =   12735512
      BackColor       =   16185592
      Begin OCX.b8Container b8Container5 
         Height          =   8505
         Left            =   90
         TabIndex        =   3
         Top             =   120
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   15002
         BorderColor     =   16185592
         BackColor       =   16185592
         ShadowColor1    =   16185592
         ShadowColor2    =   16185592
         Begin OCX.b8Container ContainerCatDetails 
            Height          =   7635
            Left            =   3330
            TabIndex        =   7
            Top             =   90
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   13467
            BackColor       =   16185592
            Begin VB.Frame fraCatDescription 
               Appearance      =   0  'Flat
               BackColor       =   &H00F6F8F8&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1275
               Left            =   1560
               TabIndex        =   10
               Top             =   210
               Width           =   3615
               Begin VB.TextBox txtCategoryName 
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
                  IMEMode         =   3  'DISABLE
                  Left            =   60
                  MaxLength       =   100
                  TabIndex        =   0
                  Top             =   0
                  Width           =   3465
               End
               Begin VB.TextBox txtCategoryDescription 
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
                  Height          =   645
                  IMEMode         =   3  'DISABLE
                  Left            =   60
                  MaxLength       =   500
                  MultiLine       =   -1  'True
                  TabIndex        =   1
                  Top             =   540
                  Width           =   3465
               End
            End
            Begin VB.Label Label11 
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
               Height          =   345
               Left            =   270
               TabIndex        =   9
               Top             =   720
               Width           =   1365
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Category Name"
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
               Left            =   270
               TabIndex        =   8
               Top             =   180
               Width           =   1365
            End
         End
         Begin OCX.b8Container ContainerList 
            Height          =   7635
            Left            =   60
            TabIndex        =   4
            Top             =   90
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   13467
            BackColor       =   16185592
            Begin VB.ListBox lstCategories 
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
               Height          =   6855
               Left            =   120
               TabIndex        =   5
               Top             =   360
               Width           =   3015
            End
            Begin OCX.b8SideTab b8SideTab2 
               Height          =   285
               Left            =   120
               TabIndex        =   6
               Top             =   90
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   503
               Caption         =   "Categories"
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
         Begin OCX.b8Container b8Container6 
            Height          =   675
            Left            =   60
            TabIndex        =   11
            Top             =   7770
            Width           =   11805
            _ExtentX        =   20823
            _ExtentY        =   1191
            BackColor       =   16185592
            Begin lvButton.lvButtons_H cmdAddNewCategory 
               Height          =   495
               Left            =   210
               TabIndex        =   12
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
               Image           =   "frm_CATEGORIES.frx":0000
               ImgSize         =   24
               cBack           =   -2147483633
               mPointer        =   99
               mIcon           =   "frm_CATEGORIES.frx":0515
            End
            Begin lvButton.lvButtons_H cmdCancelCategory 
               Height          =   495
               Left            =   8130
               TabIndex        =   13
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
               Image           =   "frm_CATEGORIES.frx":082F
               ImgSize         =   24
               cBack           =   -2147483633
               mPointer        =   99
               mIcon           =   "frm_CATEGORIES.frx":0D9E
            End
            Begin lvButton.lvButtons_H cmdEditCategory 
               Height          =   495
               Left            =   2190
               TabIndex        =   14
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
               Image           =   "frm_CATEGORIES.frx":10B8
               ImgSize         =   24
               cBack           =   -2147483633
               mPointer        =   99
               mIcon           =   "frm_CATEGORIES.frx":1435
            End
            Begin lvButton.lvButtons_H cmdCloseCategory 
               Height          =   495
               Left            =   10110
               TabIndex        =   15
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
               Image           =   "frm_CATEGORIES.frx":174F
               ImgSize         =   24
               cBack           =   -2147483633
               mPointer        =   99
               mIcon           =   "frm_CATEGORIES.frx":1CDF
            End
            Begin lvButton.lvButtons_H cmdSaveCategory 
               Height          =   495
               Left            =   6150
               TabIndex        =   16
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
               Image           =   "frm_CATEGORIES.frx":1FF9
               ImgSize         =   24
               cBack           =   -2147483633
               mPointer        =   99
               mIcon           =   "frm_CATEGORIES.frx":2259
            End
            Begin lvButton.lvButtons_H cmdDeleteCategory 
               Height          =   495
               Left            =   4170
               TabIndex        =   17
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
               Image           =   "frm_CATEGORIES.frx":2573
               ImgSize         =   24
               cBack           =   -2147483633
               mPointer        =   99
               mIcon           =   "frm_CATEGORIES.frx":2A86
            End
         End
      End
   End
End
Attribute VB_Name = "frm_CATEGORIES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lngCategoryID As Long
Dim lngProductID As Long

Private Sub cmdAddNewCategory_Click()
    Call Mdl_FUNCTIONS.fn_UNSET_CONTROL_COLOR(Me, "Category")
    
    blnAddCategory = True
    blnEditCategory = False
    
    Call sub_EMPTY_FIELDS("Category")
    Call fn_ENABLE_CONTROLS("Category")
    
    txtCategoryName.SetFocus
End Sub

Private Sub cmdAddNewProduct_Click()
    Call Mdl_FUNCTIONS.fn_UNSET_CONTROL_COLOR(Me, "Product")
    
    blnAddProduct = True
    blnEditProduct = False
    
    Call sub_EMPTY_FIELDS("Product")
    Call fn_ENABLE_CONTROLS("Product")
    
'    txtProductName.SetFocus
End Sub

Private Sub cmdCancelCategory_Click()
    Call sub_EMPTY_FIELDS("Category")
    Call fn_DISABLE_CONTROLS("Category")
    Call sub_LOAD_CATEGORIES(lstCategories)
    Call sub_LOAD_CATEGORIES_DETAILS(lngCategoryID)
    Call fn_SET_CONTROL_COLOR(Me, "Category")
End Sub

Private Sub cmdCancelProduct_Click()
    Call sub_EMPTY_FIELDS("Product")
    Call fn_DISABLE_CONTROLS("Product")
    Call fn_SET_CONTROL_COLOR(Me, "Product")
End Sub

Private Sub cmdCloseCategory_Click()
    Unload Me
End Sub

Private Sub cmdCloseProduct_Click()
    Unload Me
End Sub

Private Sub cmdDeleteCategory_Click()
On Error GoTo errHandler
    If MsgBox("Are you sure you want to Delete Category " & Trim(txtCategoryName.Text) & "?", vbQuestion + vbYesNo, Title) = vbNo Then Exit Sub

    With cls_CATEGORY_Obj
        .fn_CHECK_CATEGORY_IN_PRODUCTS (lngCategoryID)
        If blnCategoryExist = True Then
            MsgBox "Category Can Not Be Deleted.", vbInformation, Title
            Exit Sub
            Else
                Call .fn_DELETE_CATEGORY(lngCategoryID)
                Call Mdl_FUNCTIONS.sub_EMPTY_FIELS(Me)
                Call sub_LOAD_CATEGORIES(lstCategories)
        End If
    End With

EXITPROCEDURE:
    Exit Sub

errHandler:
    MsgBox Err.Description, vbCritical, "Connection"
    Call Mdl_FUNCTIONS.fn_WRITE_ERROR_TO_FILE(Date, Time, Err.Description, Err.Number, "MdlMain", "LoadRegistrySettings")
    GoTo EXITPROCEDURE
End Sub

Private Sub cmdEditCategory_Click()
    
    
    If txtCategoryName.Text = "" Then
        MsgBox "Please select the Category to Edit."
        Exit Sub
    End If
        
    Call fn_UNSET_CONTROL_COLOR(Me, "Category")
    
    blnAddCategory = False
    blnEditCategory = True
    
    Call fn_ENABLE_CONTROLS("Category")
    txtCategoryName.SetFocus
    
End Sub

Private Sub cmdEditProduct_Click()
    If txtProductName.Text = "" Then
        MsgBox "Please select the product to Edit."
        Exit Sub
    End If
        
    Call fn_UNSET_CONTROL_COLOR(Me, "Product")
    
    blnAddProduct = False
    blnEditProduct = True
    
    Call fn_ENABLE_CONTROLS("Product")
'    txtProductName.SetFocus
End Sub

Private Sub cmdLoadPackage_Click()
    blnPackage = True
    frm_PACKAGES.Show 1
End Sub

Private Sub cmdSaveCategory_Click()
On Error GoTo errHandler

    If Mdl_FUNCTIONS.fn_REQUIRE_TEXT_FIELD(txtCategoryName, "Veuillez saisir le nom de la catégory à enrégistrer svp!") Then GoTo EXITPROCEDURE

    With cls_CATEGORY_Obj
        .CategoryID = .fn_AUTOGEN
        .CategoryName = Trim(txtCategoryName.Text)
        .Description = Trim(txtCategoryDescription.Text)
    

    If blnAddCategory = True And blnEditCategory = False Then
        .fn_SAVE_CATEGORY_RECORDS
        MsgBox "Category details saved successfully", vbExclamation, Title
        Else
            .fn_UPDATE_CATEGORY_RECORDS (lngCategoryID)
            MsgBox "Category details saved successfully", vbExclamation, Title
    End If
    
    End With
    
    Call fn_DISABLE_CONTROLS("Category")
    Call sub_LOAD_CATEGORIES(lstCategories)
    Call fn_SET_CONTROL_COLOR(Me, "Category")
    
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
    
    Call fn_SET_CONTROL_COLOR(Me, "All")
    Call fn_DISABLE_CONTROLS("Category")
    Call fn_DISABLE_CONTROLS("Product")
    
    Call sub_LOAD_CATEGORIES(lstCategories)

    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdAddNewCategory, cmdAddNewCategory.hwnd, "Add new category details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdEditCategory, cmdEditCategory.hwnd, "Edit existing category details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdDeleteCategory, cmdDeleteCategory.hwnd, "Delete existing category details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdSaveCategory, cmdSaveCategory.hwnd, "Save category details.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCancelCategory, cmdCancelCategory.hwnd, "Cancel the transaction.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    Mdl_BALLOON_TOOL_TIP.CreateBalloon cmdCloseCategory, cmdCloseCategory.hwnd, "Close the window.", szBalloon, False, Title, etiInfo, tooltipBackColor, tooltipForeColor
    
End Sub

Private Sub sub_LOAD_CATEGORIES(lst As ListBox)
    Dim rec As New ADODB.Recordset
    
    Set rec = cls_CATEGORY_Obj.fN_LOAD_CATEGORIES(0)
    lst.Clear
    If rec.AbsolutePosition <> -1 Then
        Do While Not rec.EOF
            lst.AddItem rec!CategoryName
            lst.ItemData(lst.NewIndex) = rec!CategoryID
            rec.MoveNext
        Loop
    End If
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If blnCategory = True Then
        Call frm_PRODUCTS.sub_LOAD_CATEGORIES(frm_PRODUCTS.cboCategories)
    End If
End Sub

Private Sub lstCategories_Click()
    If lstCategories.ListIndex = -1 Then Exit Sub
    
    lngCategoryID = lstCategories.ItemData(lstCategories.ListIndex)
    Call sub_LOAD_CATEGORIES_DETAILS(lngCategoryID)

End Sub

Private Sub sub_LOAD_CATEGORIES_DETAILS(lngCategoryID As Long)
    Dim rec As New ADODB.Recordset
    
    Set rec = cls_CATEGORY_Obj.fN_LOAD_CATEGORIES(lngCategoryID)
    If rec.AbsolutePosition <> -1 Then
        txtCategoryName.Text = Trim(rec!CategoryName) & ""
        txtCategoryDescription.Text = Trim(rec!Description) & ""
    End If
    
End Sub




Public Sub sub_LOAD_PACKAGES(Optional cbo As ComboBox)

    Dim rec As New ADODB.Recordset
    Set rec = cls_REFERENCES_Obj.fn_LOAD_PACKAGES
    
    cbo.Clear
    
    Do While Not rec.EOF
        cbo.AddItem rec!PackageName
        cbo.ItemData(cbo.NewIndex) = rec!PackageID
        rec.MoveNext
    Loop

End Sub


'**********Function to disable some buttons***********
Private Function fn_DISABLE_CONTROLS(str As String)

    If str = "Category" Then
        cmdSaveCategory.Enabled = False
        cmdCancelCategory.Enabled = False
            
        cmdEditCategory.Enabled = True
        cmdDeleteCategory.Enabled = True
        cmdAddNewCategory.Enabled = True
        
        lstCategories.Enabled = True
        fraCatDescription.Enabled = False
        
    End If
    
End Function

'**********Function to enable some buttons***********
Private Function fn_ENABLE_CONTROLS(str As String)

    If str = "Category" Then
        cmdSaveCategory.Enabled = True
        cmdCancelCategory.Enabled = True
            
        cmdEditCategory.Enabled = False
        cmdDeleteCategory.Enabled = False
        cmdAddNewCategory.Enabled = False
        
        lstCategories.Enabled = False
        fraCatDescription.Enabled = True
        
    ElseIf str = "Product" Then
    
        cmdSaveProduct.Enabled = True
        cmdCancelProduct.Enabled = True
            
        cmdEditProduct.Enabled = False
        cmdDeleteProduct.Enabled = False
        cmdAddNewProduct.Enabled = False
    
        lstProducts.Enabled = False
        fraProductDetails.Enabled = True
    
    End If

End Function

Private Sub sub_EMPTY_FIELDS(str As String)

    If str = "Category" Then
        
        txtCategoryDescription.Text = ""
        txtCategoryName.Text = ""
        
    ElseIf str = "Product" Then
    
        txtProductName.Text = ""
        cboCategories.ListIndex = -1
        txtQtyPerUnit.Text = ""
        txtUnitPrice.Text = ""
        txtUnitInStock.Text = ""
        txtUnitOnOrder.Text = ""
        txtReOrderLevel.Text = ""
'        ChkAlerte.Value = 0
    
    End If

End Sub
