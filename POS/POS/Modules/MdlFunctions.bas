Attribute VB_Name = "MdlFunctions"

'**********Function to highlight the content of a text***********
Public Function fn_HIGHLIGHT_TEXT(ByVal txt As TextBox)

    txt.SelStart = 0
    txt.SelLength = Len(txt.Text)
    txt.SetFocus

End Function

'**********Function to require data in a textbox***********
Public Function fn_REQUIRE_TEXT_FIELD(ByVal txt As TextBox, strMessage As String) As Boolean
    
    fn_REQUIRE_TEXT_FIELD = False
    
    If txt.Text = "" Then
        fn_REQUIRE_TEXT_FIELD = True
        MsgBox strMessage, vbExclamation, Title
        txt.SetFocus
    End If
    
End Function

Public Function fn_REQUIRE_COMBO_FIELD(ByVal cbo As ComboBox, strMessage As String) As Boolean
    fn_REQUIRE_COMBO_FIELD = False
    
    If cbo.ListIndex = -1 Then
        fn_REQUIRE_COMBO_FIELD = True
        MsgBox strMessage, vbExclamation, Title
        cbo.SetFocus
    End If
    
End Function


Public Function fn_REQUIRE_DATE_OF_BIRTH(ByVal dt As DTPicker, strMessage As String) As Boolean
    fn_REQUIRE_DATE_OF_BIRTH = False
    
    If dt.Value = Date Or dt.Value > Now Then
        fn_REQUIRE_DATE_OF_BIRTH = True
        MsgBox strMessage, vbExclamation, Title
        dt.SetFocus
    End If
    
End Function

'**********Function to accept only numeric characters***********
Public Sub sub_NUMERIC_ONLY(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack, vbKeyDelete
        Case Else
            KeyAscii = 0
    End Select

End Sub

'**********Function to accept only alphabet***********
Public Sub sub_ALPHABET_ONLY(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKey0 To vbKey9
            KeyAscii = 0
    End Select

End Sub

'**********Function to center a form on the screen***********
Public Sub sub_CENTER_FORM(ByVal frm As Form)

    frm.Move (Screen.Width - frm.Width) / 2, (Screen.Height - frm.Height) / 3

End Sub

'**********Function to center a form on the screen***********
Public Sub sub_SET_FORM_POSITION(ByVal frm As Form)

    frm.Left = 0
    frm.Top = 0

End Sub

'**********Function to set the tree font***********
Public Sub sub_SET_TREEVIEW_FONT()

    XNode.Bold = True
    XNode.ForeColor = &HC00000

End Sub

'**********Function to unload a form***********
Public Sub sub_UNLOAD_FORMS(ByVal frm As Form)

    If MsgBox("Are you sure you want to close this window ?", vbQuestion + vbYesNo, Title) = vbNo Then
    
        Exit Sub
        
        Else
                
                Unload frm
    
    End If

End Sub

'**********Function to write all the errors that occur to a file***********
Public Function fn_WRITE_ERROR_TO_FILE(dDate As Date, dTime As Date, strErrDescription As String, strErrNumber As String, strModule As String, strProcedure As String) As String

    Set txtStream = file.OpenTextFile(App.Path & "\Errors\ErrorText.txt", ForAppending, True)
    txtStream.WriteLine "***********************************************"
    txtStream.WriteLine "*Date              :" & dDate
    txtStream.WriteLine "*Time              :" & dTime
    txtStream.WriteLine "*Error Description :" & strErrDescription
    txtStream.WriteLine "*Error Number      :" & strErrNumber
    txtStream.WriteLine "*Module            :" & strModule
    txtStream.WriteLine "*Procedure         :" & strProcedure
    txtStream.WriteLine "***********************************************"
    txtStream.Close
    
End Function


'**********Function to set the list index***********
Public Function fn_GET_LIST_INDEX(cbo As ComboBox, lngItemID As Long) As Long

    Dim lngIndex As Long

    fn_GET_LIST_INDEX = -1
    
    For lngIndex = 0 To cbo.ListCount - 1
        If cbo.ItemData(lngIndex) = lngItemID Then
            fn_GET_LIST_INDEX = lngIndex
        End If
    Next lngIndex

End Function

'**********Function to empty fields*******************
Public Sub sub_EMPTY_FIELS(frm As Form)

    For Each ctl In frm
        If TypeOf ctl Is Label And Left(ctl.Name, 3) = "lab" Then ctl.Caption = ""
        If TypeOf ctl Is TextBox Then ctl.Text = ""
        If TypeOf ctl Is ComboBox Then ctl.ListIndex = -1
        If TypeOf ctl Is DTPicker Then ctl.Value = Date
        If TypeOf ctl Is Image Then ctl.Picture = LoadPicture("")
    Next

End Sub


Public Function fn_CHECK_EMPTY_COMBO(cbo As ComboBox, strMessage As String) As Boolean

    fn_CHECK_EMPTY_COMBO = False

    If cbo.ListIndex = -1 Then
        fn_CHECK_EMPTY_COMBO = True
        MsgBox strMessage, vbExclamation, Title
        cbo.SetFocus
    End If


End Function

Public Function fn_CHECK_EMPTY_TEXT_BOX(txt As TextBox, strMessage As String) As Boolean

    fn_CHECK_EMPTY_TEXT_BOX = False

    If txt.Text = "" Then
        fn_CHECK_EMPTY_TEXT_BOX = True
        MsgBox strMessage, vbExclamation, Title
        txt.SetFocus
    End If


End Function


Public Function fn_CHECK_DATE_OF_BIRTH(dtp As DTPicker, strMessage As String) As Boolean

    fn_CHECK_DATE_OF_BIRTH = False

    If dtp.Value > Date Then
        fn_CHECK_DATE_OF_BIRTH = True
        MsgBox strMessage, vbExclamation, Title
        dtp.SetFocus
    End If


End Function


Public Sub sub_CLOSE_ALL_OPENED_FORMS()
    
    Dim I As Integer
    For I = Forms.Count - 1 To 0 Step -1
        If Forms(I).Name <> "frmMain" Then
            Unload Forms(I)
        End If
    Next I
    
    
End Sub
