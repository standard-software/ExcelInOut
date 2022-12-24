'--------------------------------------------------
'Excel MakeApp Framework
'--------------------------------------------------
'ModuleName:    Main Form
'ObjectName:    FormMain
'--------------------------------------------------
'Version:       2015/07/29
'--------------------------------------------------
Option Explicit

'--------------------------------------------------
'���錾
'--------------------------------------------------
'----------------------------------------
'���t���[�����[�N�p
'----------------------------------------
Public Args As String

Private FormProperty As New st_vba_FormProperty

Private AnchorMenuButton As New st_vba_ControlAnchor

'----------------------------------------
'�����[�U�[�p
'----------------------------------------
'------------------------------
'���A���J�[��`
'------------------------------
Private AnchorTabStrip As New st_vba_ControlAnchor
Private AnchorListBox As New st_vba_ControlAnchor
Private AnchorButtonExecute As New st_vba_ControlAnchor
Private AnchorButtonFileAdd As New st_vba_ControlAnchor
Private AnchorButtonFileRemove As New st_vba_ControlAnchor
Private AnchorButtonAllSelect As New st_vba_ControlAnchor

'------------------------------
'���ϐ���`
'------------------------------
Private ListFileItem As Object

Private ArrayListBoxSelectedIndex() As Long





'--------------------------------------------------
'������
'--------------------------------------------------

'----------------------------------------
'���N���E�I��
'----------------------------------------

'------------------------------
'���ϐ��������Ȃ�
'------------------------------
Private Sub UserForm_Initialize()
    Me.StartUpPosition = 2
    Args = ""
    Call IniRead_UserFormInitialize
End Sub

'------------------------------
'��Main����̌Ăяo��
'------------------------------
Public Sub Initialize( _
ByVal TaskBarButton As Boolean, _
ByVal TitleBar As Boolean, _
ByVal SystemMenu As Boolean, _
ByVal FormIcon As Boolean, _
ByVal MinimizeButton As Boolean, _
ByVal MaximizeButton As Boolean, _
ByVal CloseButton As Boolean, _
ByVal ResizeFrame As Boolean, _
ByVal TopMost As Boolean)

    '------------------------------
    '���t���[�����[�N����������
    '------------------------------
    With Nothing
        Call FormProperty.InitializeForm(Me)

        Call FormProperty.InitializeProperty( _
            TaskBarButton:=TaskBarButton, _
            TitleBar:=TitleBar, _
            SystemMenu:=SystemMenu, _
            FormIcon:=FormIcon, _
            MinimizeButton:=MinimizeButton, _
            MaximizeButton:=MaximizeButton, _
            CloseButton:=CloseButton, _
            ResizeFrame:=ResizeFrame, _
            TopMost:=TopMost)

        FormProperty.IconPath = Project_MainIconFilePath
        FormProperty.IconIndex = Project_MainIconIndex

        Me.Caption = Project_FormMainTitle
    End With

    '------------------------------
    '�����[�U�[�p����������
    '------------------------------
    '�ȉ��Ƀ��[�U�[�Ǝ��̏������������L�q���Ă�������
    '------------------------------
    
    Set ListFileItem = CreateObject("System.Collections.ArrayList")

End Sub

Private Sub UserForm_Activate()
    If FormProperty.Initializing Then
        FormProperty.Initializing = False

        Call SetTaskbarButtonAppID(Project_AppID)

        If FormProperty.Handle = 0 Then
            Call FormProperty.InitializeForm(Me)
            FormProperty.GetWindowsProperty
        Else
            FormProperty.SetWindowsProperty
        End If

        '------------------------------
        '�����j���[�{�^�����E��[�ɂ���
        '------------------------------
        Me.ImageMenuButton.Top = 0
        Me.ImageMenuButton.Left = _
            Me.ImageMenuButton.Parent.InsideWidth - _
            Me.ImageMenuButton.Width + 1

        '------------------------------
        '���t���[�����[�N�A���J�[����������
        '------------------------------
        Call AnchorMenuButton.Initialize( _
            Me.ImageMenuButton, _
            HorizonAnchorType.haRight, IIf(FormProperty.ResizeFrame, 0, 0), _
            VerticalAnchorType.vaTop, IIf(FormProperty.ResizeFrame, 0, 0))
            
        'Excel2016�ł́AOffset�l��ResizeFrame�ɂ�����炸0�ɂȂ�
        'Excel2013�ł͉��L�̃R�[�h���L��
        'Call FAnchorMenuButton.Initialize( _
        '   Me.FrameMenuButton, _
        '   HorizonAnchorType.haRight, IIf(FormProperty.ResizeFrame, 8, 0), _
        '   VerticalAnchorType.vaTop, IIf(FormProperty.ResizeFrame, 8, 0))

        '------------------------------
        '�����[�U�[�p�A���J�[����������
        '------------------------------
        '�ȉ��Ƀ��[�U�[�Ǝ��̃A���J�[�������������L�q���Ă�������
        '------------------------------
        Call AnchorTabStrip.Initialize( _
            Me.TabStrip1, _
            HorizonAnchorType.haStretch, IIf(FormProperty.ResizeFrame, 0, 0), _
            VerticalAnchorType.vaStretch, IIf(FormProperty.ResizeFrame, 0, 0))
        Call AnchorListBox.Initialize( _
            Me.ListBox1, _
            HorizonAnchorType.haStretch, IIf(FormProperty.ResizeFrame, 0, 0), _
            VerticalAnchorType.vaStretch, IIf(FormProperty.ResizeFrame, 0, 0))
        Call AnchorButtonExecute.Initialize( _
            Me.ButtonExecute, _
            HorizonAnchorType.haRight, IIf(FormProperty.ResizeFrame, 0, 0), _
            VerticalAnchorType.vaBottom, IIf(FormProperty.ResizeFrame, 0, 0))
        Call AnchorButtonFileAdd.Initialize( _
            Me.ButtonFileAdd, _
            HorizonAnchorType.haLeft, IIf(FormProperty.ResizeFrame, 0, 0), _
            VerticalAnchorType.vaBottom, IIf(FormProperty.ResizeFrame, 0, 0))
        Call AnchorButtonFileRemove.Initialize( _
            Me.ButtonFileRemove, _
            HorizonAnchorType.haLeft, IIf(FormProperty.ResizeFrame, 0, 0), _
            VerticalAnchorType.vaBottom, IIf(FormProperty.ResizeFrame, 0, 0))
        Call AnchorButtonAllSelect.Initialize( _
            Me.ButtonAllSelect, _
            HorizonAnchorType.haLeft, IIf(FormProperty.ResizeFrame, 0, 0), _
            VerticalAnchorType.vaBottom, IIf(FormProperty.ResizeFrame, 0, 0))

        Call IniRead_UserFormActivate

        '���C�A�E�g�A���J�[�𓮍삳����
        Call UserForm_Resize

        Call FormProperty.ForceActiveMouseClick


    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Select Case CloseMode
    Case 0
        Call IniWrite
    Case 1
    End Select
End Sub

'------------------------------
'���I�����ɌĂяo���֐�
'------------------------------
'Me.Hide �� Call Unload(Me) �ł͂Ȃ�
'����FormClose�֐����Ăяo���Ă�������
'Me.Hide �� Call Unload(Me) �ł�
'UserForm_QueryClose�C�x���g���Ăяo���ꂸ
'Ini�t�@�C���ւ̕ۑ����s���܂���B
'------------------------------
Private Sub FormClose()
    Dim Cancel As Integer
    Cancel = False
    Call UserForm_QueryClose(Cancel, 0)
    If Cancel Then Exit Sub
    Call Me.Hide
End Sub

'----------------------------------------
'��Ini�t�@�C��
'----------------------------------------
'Ini�t�@�C���ւ̕ۑ���Ǎ��̏������L�q���Ă�������
'----------------------------------------
Public Sub IniRead_UserFormInitialize()
    '------------------------------
    '�����[�U�[�pIni�t�@�C���Ǎ�����(UserFormInitialize�C�x���g��)
    '------------------------------
    '�ȉ��ɏ���������Ini�t�@�C���Ǎ��������L�q���Ă�������
    '------------------------------


End Sub

Public Sub IniRead_UserFormActivate()
    '------------------------------
    '���t���[�����[�NForm�ʒu���A����
    '------------------------------
    Call Form_IniReadPosition(Me, _
        Project_IniFilePath, "Form", "Rect", False)

    '------------------------------
    '�����[�U�[�pIni�t�@�C���Ǎ�����(UserFormActivate�C�x���g��)
    '------------------------------
    '�ȉ���UserForm�쐬����������Ini�t�@�C���Ǎ��������L�q���Ă�������
    '------------------------------

    Dim FileListCount As String
    FileListCount = _
        IniFile_GetString(Project_IniFilePath, _
            "Data", "FileListCount", "0")
    If IsLong(FileListCount) Then
        Dim I As Long
        For I = 0 To CLng(FileListCount) - 1
            Dim FilePath As String
            FilePath = IniFile_GetString(Project_IniFilePath, _
                "Data", "File" + LongToStrDigitZero(I, 3), "")
            If fso.FileExists(FilePath) Then
                Call ListFileItem.Add(FilePath)
            End If
        Next
        Call ListBox1_SetListFileItem
    End If


    TabStrip1.Value = CLng(IniFile_GetString(Project_IniFilePath, _
        "Status", "TabIndex", "0"))


End Sub

Public Sub IniWrite()
    '------------------------------
    '���t���[�����[�NForm�ʒu�ۑ�����
    '------------------------------
    Call Assert(FormProperty.Handle <> 0)

    If (FormProperty.WindowState = xlNormal) Then
        Call Form_IniWritePosition(Me, _
            Project_IniFilePath, "Form", "Rect")
    End If

    '------------------------------
    '�����[�U�[�pIni�t�@�C����������
    '------------------------------
    '�ȉ��ɏI������Ini�t�@�C�������������L�q���Ă�������
    '------------------------------
    
    Call IniFile_SetString(Project_IniFilePath, _
        "Data", "FileListCount", ListFileItem.Count)
    Dim I As Long
    For I = 0 To ListFileItem.Count - 1
        Call IniFile_SetString(Project_IniFilePath, _
            "Data", "File" + LongToStrDigitZero(I, 3), ListFileItem(I))
    Next
    
    Call IniFile_SetString(Project_IniFilePath, _
        "Status", "TabIndex", TabStrip1.Value)

End Sub

'----------------------------------------
'�����T�C�Y�C�x���g
'----------------------------------------
Private Sub UserForm_Resize()
    If FormProperty.Initializing = False Then
        '------------------------------
        '���t���[�����[�N�A���J�[���C�A�E�g����
        '------------------------------
        Call AnchorMenuButton.Layout

        '------------------------------
        '�����[�U�[�p�A���J�[���C�A�E�g����
        '------------------------------
        '�ȉ��Ƀ��[�U�[�Ǝ��̃A���J�[���C�A�E�g�������L�q���Ă�������
        '------------------------------
        Call AnchorTabStrip.Layout
        Call AnchorListBox.Layout
        Call AnchorButtonExecute.Layout
        Call AnchorButtonFileAdd.Layout
        Call AnchorButtonFileRemove.Layout
        Call AnchorButtonAllSelect.Layout
    End If
    
    Dim ListBoxColumn1Width As Long
    ListBoxColumn1Width = 150
    ListBox1.ColumnWidths = _
        CStr(ListBoxColumn1Width) & ";" & CStr(ListBox1.Width - ListBoxColumn1Width - 15)
    
End Sub

'----------------------------------------
'�����j���[�{�^��
'----------------------------------------
Private Sub ImageMenuButton_Click()
    Dim PopupMenu As CommandBar
    Set PopupMenu = Application.CommandBars.Add(, Position:=msoBarPopup)

    Dim MenuItemCreateAppShortcut As CommandBarControl
    Set MenuItemCreateAppShortcut = _
        PopupMenu.Controls.Add(Type:=msoControlButton)
    MenuItemCreateAppShortcut.Caption = "�A�v���P�[�V�����̃V���[�g�J�b�g���쐬..."
    MenuItemCreateAppShortcut.FaceId = 0
    MenuItemCreateAppShortcut.OnAction = PopupMenu_ActionText("CreateAppShortcut")

    Dim MenuItemVersionInfo As CommandBarControl
    Set MenuItemVersionInfo = _
        PopupMenu.Controls.Add(Type:=msoControlButton)
    MenuItemVersionInfo.Caption = "�o�[�W�������"
    MenuItemVersionInfo.FaceId = 0
    MenuItemVersionInfo.OnAction = PopupMenu_ActionText("VersionInfo")

    Dim MenuItemAppClose As CommandBarControl
    Set MenuItemAppClose = _
        PopupMenu.Controls.Add(Type:=msoControlButton)
    MenuItemAppClose.BeginGroup = True
    MenuItemAppClose.Caption = "�I��"
    MenuItemAppClose.FaceId = 0
    MenuItemAppClose.OnAction = PopupMenu_ActionText("AppClose")

    Dim XOffset As Long: XOffset = 14
    Dim XOffsetResizeOn As Long: XOffsetResizeOn = 8
    Dim XOffsetResizeOff As Long: XOffsetResizeOff = 4
    Dim YOffsetTitleBarOn As Long: YOffsetTitleBarOn = 20
    Dim YOffsetTitleBarOff As Long: YOffsetTitleBarOff = 0
    Dim YOffsetResizeOn As Long: YOffsetResizeOn = 8
    Dim YOffsetResizeOff As Long: YOffsetResizeOff = 4

    XOffset = XOffset * (GetDPI / 96)
    XOffsetResizeOn = XOffsetResizeOn * (GetDPI / 96)
    XOffsetResizeOff = XOffsetResizeOff * (GetDPI / 96)
    YOffsetTitleBarOn = YOffsetTitleBarOn * (GetDPI / 96)
    YOffsetTitleBarOff = YOffsetTitleBarOff * (GetDPI / 96)
    YOffsetResizeOn = YOffsetResizeOn * (GetDPI / 96)
    YOffsetResizeOff = YOffsetResizeOff * (GetDPI / 96)

    Select Case PopupMenu_PopupReturn(PopupMenu, _
        PointToPixel(Me.Left + ImageMenuButton.Left + ImageMenuButton.Width) _
        + IIf(FormProperty.ResizeFrame, XOffsetResizeOn, XOffsetResizeOff) _
        - PopupMenu.Width + XOffset, _
        PointToPixel(Me.Top + ImageMenuButton.Top + ImageMenuButton.Height) _
        + IIf(FormProperty.ResizeFrame, YOffsetResizeOn, YOffsetResizeOff) _
        + IIf(FormProperty.TitleBar, YOffsetTitleBarOn, YOffsetTitleBarOff))
    Case "CreateAppShortcut"
        Call Load(FormCreateAppShortcut)
        Call FormCreateAppShortcut.ShowDialog( _
            Me, FormProperty.TopMost)
        Call Unload(FormCreateAppShortcut)
    Case "VersionInfo"
        Call MsgBox( _
            Project_VersionDialogInstruction + vbNewLine + _
            Project_VersionDialogContent, _
            vbOKOnly, _
            Project_VersionDialogWindowTitle)
    Case "AppClose"
        FormClose
    End Select
End Sub

'--------------------------------------------------
'���v���O�����{��
'--------------------------------------------------
'�ȉ��Ƀv���O�����{�̂̏������L�q���Ă�������
'--------------------------------------------------


'--------------------------------------------------
'ListBox�̕����I��UI�ύX
'--------------------------------------------------

Private Sub ListBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    Call ListBox1_Click

    Erase ArrayListBoxSelectedIndex

    Dim I As Long
    For I = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(I) Then
            Call ArrayAdd(ArrayListBoxSelectedIndex, I)
        End If
    Next
End Sub

Private Sub ListBox1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If ListBox_SelectedCount(ListBox1) = 1 Then
        Dim I As Long
        For I = 0 To ListBox1.ListCount - 1
            If ArrayIndexOf(ArrayListBoxSelectedIndex, I) <> -1 Then
                ListBox1.Selected(I) = Not ListBox1.Selected(I)
            End If
        Next
    End If
End Sub

Public Function ListBox_SelectedCount(ByVal ListBox As Object) As Long
    Dim Result As Long: Result = 0
    Dim I As Long
    For I = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(I) Then
            Result = Result + 1
        End If
    Next
    ListBox_SelectedCount = Result
End Function


'--------------------------------------------------
'�{�^���@�\
'--------------------------------------------------
Private Sub ButtonFileAdd_Click()

    Dim FileList() As String
    FileList = Split(FileDialog_FilePicker(ThisWorkbook.Path + "\", msoFileDialogViewDetails, True, _
        "Excel�u�b�N|*.xls; *.xlsx; *.xlsm"), vbCrLf)
        
    Dim I As Long
    For I = 0 To ArrayCount(FileList) - 1
        Call ListFileItem.Add(FileList(I))
    Next

    Call ListBox1_SetListFileItem
End Sub


Private Sub ButtonFileRemove_Click()
    If ListBox_SelectedCount(ListBox1.Object) = 0 Then
        Call MsgBox(StringCombine(vbCrLf, _
            "�t�@�C�����I������Ă��܂���B", _
            "�I���������ڂ�o�^�������܂��B"))
        Exit Sub
    End If
    
    Dim I As Long
    For I = ListFileItem.Count - 1 To 0 Step -1
        If ListBox1.Selected(I) Then
            Call ListFileItem.Remove(ListFileItem(I))
        End If
    Next
    
    Call ListBox1_SetListFileItem
End Sub

Private Sub ListBox1_SetListFileItem()

    If ListFileItem.Count = 0 Then
        '�N���A����Ƃ��͋�z���ݒ肷��
        ListBox1.Column() = Array()
    Else
        Dim SetListBoxArray() As Variant
        Dim I As Long
        For I = 0 To ListFileItem.Count - 1
            Call Array2dAdd(SetListBoxArray, Array( _
                fso.GetFileName(ListFileItem(I)), _
                fso.GetParentFolderName(ListFileItem(I))))
            '��2������A�񎟌��z��̗��2�ɂ��Ă���
        Next
        ListBox1.Column() = SetListBoxArray
    End If
End Sub

Private Sub ButtonAllSelect_Click()
    Dim SelectFlag As Boolean
    If ListBox_SelectedCount(ListBox1.Object) = ListBox1.ListCount Then
        SelectFlag = False
    Else
        SelectFlag = True
    End If
    
    Dim I As Long
    For I = 0 To ListBox1.ListCount - 1
        ListBox1.Selected(I) = SelectFlag
    Next
End Sub

Private Sub TabStrip1_Change()
    If TabStrip1.Value = 0 Then
        ButtonExecute.Caption = "�\�[�X�R�[�h����"
    ElseIf TabStrip1.Value = 1 Then
        ButtonExecute.Caption = "�\�[�X�R�[�h�o��"
    Else
        Call Assert(False)
    End If
End Sub

Private Sub ButtonExecute_Click()
    Call Assert(OrValue(TabStrip1.Value, 0, 1))
    
    If ListBox_SelectedCount(ListBox1.Object) = 0 Then
        Call MsgBox("���s�Ώۂ�����܂���")
        Exit Sub
    End If

    If TabStrip1.Value = 0 Then
        If MsgBox("�\�[�X�R�[�h����͂��܂��B��낵���ł����H", _
            VbMsgBoxStyle.vbOKCancel) <> VbMsgBoxResult.vbOK Then Exit Sub
    ElseIf TabStrip1.Value = 1 Then
        If MsgBox("�\�[�X�R�[�h���o�͂��܂��B��낵���ł����H", _
            VbMsgBoxStyle.vbOKCancel) <> VbMsgBoxResult.vbOK Then Exit Sub
    End If

    Dim ListSelectFileItem As Object
    Set ListSelectFileItem = CreateObject("System.Collections.ArrayList")
    Call ListSelectFileItem.Clear


    Dim Book As Workbook
    Dim I As Long
    For I = 0 To ListBox1.ListCount - 1
    Do
        If ListBox1.Selected(I) Then
            If Not fso.FileExists(ListFileItem(I)) Then Exit Do
            
            Call ListSelectFileItem.Add(ListFileItem(I))
            
        End If
    Loop While False
    Next
        
    For I = 0 To ListSelectFileItem.Count - 1
    Do
        If TabStrip1.Value = 0 Then
            '����
            Set Book = App_GetOpenedBookOrOpenBook(Application, _
                ListSelectFileItem(I), True, False)
            Call VBACode_Input( _
                FolderPath_VBACode(Book_FullPath(Book), "code"), _
                Book.VBProject)
            Call Book.Save
            If Book_FullPath(ThisWorkbook) <> Book_FullPath(Book) Then
                Call Book_CloseSilence(Book)
            End If
        ElseIf TabStrip1.Value = 1 Then
            '�o��
            Set Book = App_GetOpenedBookOrOpenBook(Application, _
                ListSelectFileItem(I), True, True)
            Call VBACode_Output( _
                FolderPath_VBACode(Book_FullPath(Book), "code"), _
                Book.VBProject, False)
            If Book_FullPath(ThisWorkbook) <> Book_FullPath(Book) Then
                Call Book_CloseSilence(Book)
            End If
        End If

    Loop While False
    Next
    
    'VBModule�𑀍삷��ƁAListBox�̑I�������������Ƃ���
    '�Ӗ��s���ȃo�O���������̂őΏ�����
    '�g���[�X���s���͂��̖��͌����Ȃ�
    Dim SearchIndex As Long
    For I = 0 To ListSelectFileItem.Count - 1
    Do
        SearchIndex = ListFileItem.IndexOf_3(ListSelectFileItem(I))
        ListBox1.Selected(SearchIndex) = True
    Loop While False
    Next
    
    Set ListSelectFileItem = Nothing
    
    If TabStrip1.Value = 0 Then
        MsgBox ("���͊���")
    ElseIf TabStrip1.Value = 1 Then
        MsgBox ("�o�͊���")
    End If
End Sub
