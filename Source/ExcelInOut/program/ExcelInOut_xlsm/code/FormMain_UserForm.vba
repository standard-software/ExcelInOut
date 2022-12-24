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
'■宣言
'--------------------------------------------------
'----------------------------------------
'◆フレームワーク用
'----------------------------------------
Public Args As String

Private FormProperty As New st_vba_FormProperty

Private AnchorMenuButton As New st_vba_ControlAnchor

'----------------------------------------
'◆ユーザー用
'----------------------------------------
'------------------------------
'◇アンカー定義
'------------------------------
Private AnchorTabStrip As New st_vba_ControlAnchor
Private AnchorListBox As New st_vba_ControlAnchor
Private AnchorButtonExecute As New st_vba_ControlAnchor
Private AnchorButtonFileAdd As New st_vba_ControlAnchor
Private AnchorButtonFileRemove As New st_vba_ControlAnchor
Private AnchorButtonAllSelect As New st_vba_ControlAnchor

'------------------------------
'◇変数定義
'------------------------------
Private ListFileItem As Object

Private ArrayListBoxSelectedIndex() As Long





'--------------------------------------------------
'■実装
'--------------------------------------------------

'----------------------------------------
'◆起動・終了
'----------------------------------------

'------------------------------
'◇変数初期化など
'------------------------------
Private Sub UserForm_Initialize()
    Me.StartUpPosition = 2
    Args = ""
    Call IniRead_UserFormInitialize
End Sub

'------------------------------
'◇Mainからの呼び出し
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
    '◇フレームワーク初期化処理
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
    '◇ユーザー用初期化処理
    '------------------------------
    '以下にユーザー独自の初期化処理を記述してください
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
        '◇メニューボタンを右上端にする
        '------------------------------
        Me.ImageMenuButton.Top = 0
        Me.ImageMenuButton.Left = _
            Me.ImageMenuButton.Parent.InsideWidth - _
            Me.ImageMenuButton.Width + 1

        '------------------------------
        '◇フレームワークアンカー初期化処理
        '------------------------------
        Call AnchorMenuButton.Initialize( _
            Me.ImageMenuButton, _
            HorizonAnchorType.haRight, IIf(FormProperty.ResizeFrame, 0, 0), _
            VerticalAnchorType.vaTop, IIf(FormProperty.ResizeFrame, 0, 0))
            
        'Excel2016では、Offset値はResizeFrameにかかわらず0になる
        'Excel2013では下記のコードが有効
        'Call FAnchorMenuButton.Initialize( _
        '   Me.FrameMenuButton, _
        '   HorizonAnchorType.haRight, IIf(FormProperty.ResizeFrame, 8, 0), _
        '   VerticalAnchorType.vaTop, IIf(FormProperty.ResizeFrame, 8, 0))

        '------------------------------
        '◇ユーザー用アンカー初期化処理
        '------------------------------
        '以下にユーザー独自のアンカー初期化処理を記述してください
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

        'レイアウトアンカーを動作させる
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
'◇終了時に呼び出す関数
'------------------------------
'Me.Hide や Call Unload(Me) ではなく
'このFormClose関数を呼び出してください
'Me.Hide や Call Unload(Me) では
'UserForm_QueryCloseイベントが呼び出されず
'Iniファイルへの保存が行われません。
'------------------------------
Private Sub FormClose()
    Dim Cancel As Integer
    Cancel = False
    Call UserForm_QueryClose(Cancel, 0)
    If Cancel Then Exit Sub
    Call Me.Hide
End Sub

'----------------------------------------
'◆Iniファイル
'----------------------------------------
'Iniファイルへの保存や読込の処理を記述してください
'----------------------------------------
Public Sub IniRead_UserFormInitialize()
    '------------------------------
    '◇ユーザー用Iniファイル読込処理(UserFormInitializeイベント時)
    '------------------------------
    '以下に初期化時のIniファイル読込処理を記述してください
    '------------------------------


End Sub

Public Sub IniRead_UserFormActivate()
    '------------------------------
    '◇フレームワークForm位置復帰処理
    '------------------------------
    Call Form_IniReadPosition(Me, _
        Project_IniFilePath, "Form", "Rect", False)

    '------------------------------
    '◇ユーザー用Iniファイル読込処理(UserFormActivateイベント時)
    '------------------------------
    '以下にUserForm作成初期化時のIniファイル読込処理を記述してください
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
    '◇フレームワークForm位置保存処理
    '------------------------------
    Call Assert(FormProperty.Handle <> 0)

    If (FormProperty.WindowState = xlNormal) Then
        Call Form_IniWritePosition(Me, _
            Project_IniFilePath, "Form", "Rect")
    End If

    '------------------------------
    '◇ユーザー用Iniファイル書込処理
    '------------------------------
    '以下に終了時のIniファイル書込処理を記述してください
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
'◆リサイズイベント
'----------------------------------------
Private Sub UserForm_Resize()
    If FormProperty.Initializing = False Then
        '------------------------------
        '◇フレームワークアンカーレイアウト処理
        '------------------------------
        Call AnchorMenuButton.Layout

        '------------------------------
        '◇ユーザー用アンカーレイアウト処理
        '------------------------------
        '以下にユーザー独自のアンカーレイアウト処理を記述してください
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
'◆メニューボタン
'----------------------------------------
Private Sub ImageMenuButton_Click()
    Dim PopupMenu As CommandBar
    Set PopupMenu = Application.CommandBars.Add(, Position:=msoBarPopup)

    Dim MenuItemCreateAppShortcut As CommandBarControl
    Set MenuItemCreateAppShortcut = _
        PopupMenu.Controls.Add(Type:=msoControlButton)
    MenuItemCreateAppShortcut.Caption = "アプリケーションのショートカットを作成..."
    MenuItemCreateAppShortcut.FaceId = 0
    MenuItemCreateAppShortcut.OnAction = PopupMenu_ActionText("CreateAppShortcut")

    Dim MenuItemVersionInfo As CommandBarControl
    Set MenuItemVersionInfo = _
        PopupMenu.Controls.Add(Type:=msoControlButton)
    MenuItemVersionInfo.Caption = "バージョン情報"
    MenuItemVersionInfo.FaceId = 0
    MenuItemVersionInfo.OnAction = PopupMenu_ActionText("VersionInfo")

    Dim MenuItemAppClose As CommandBarControl
    Set MenuItemAppClose = _
        PopupMenu.Controls.Add(Type:=msoControlButton)
    MenuItemAppClose.BeginGroup = True
    MenuItemAppClose.Caption = "終了"
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
'■プログラム本体
'--------------------------------------------------
'以下にプログラム本体の処理を記述してください
'--------------------------------------------------


'--------------------------------------------------
'ListBoxの複数選択UI変更
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
'ボタン機能
'--------------------------------------------------
Private Sub ButtonFileAdd_Click()

    Dim FileList() As String
    FileList = Split(FileDialog_FilePicker(ThisWorkbook.Path + "\", msoFileDialogViewDetails, True, _
        "Excelブック|*.xls; *.xlsx; *.xlsm"), vbCrLf)
        
    Dim I As Long
    For I = 0 To ArrayCount(FileList) - 1
        Call ListFileItem.Add(FileList(I))
    Next

    Call ListBox1_SetListFileItem
End Sub


Private Sub ButtonFileRemove_Click()
    If ListBox_SelectedCount(ListBox1.Object) = 0 Then
        Call MsgBox(StringCombine(vbCrLf, _
            "ファイルが選択されていません。", _
            "選択した項目を登録解除します。"))
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
        'クリアするときは空配列を設定する
        ListBox1.Column() = Array()
    Else
        Dim SetListBoxArray() As Variant
        Dim I As Long
        For I = 0 To ListFileItem.Count - 1
            Call Array2dAdd(SetListBoxArray, Array( _
                fso.GetFileName(ListFileItem(I)), _
                fso.GetParentFolderName(ListFileItem(I))))
            '列が2だから、二次元配列の列を2にしている
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
        ButtonExecute.Caption = "ソースコード入力"
    ElseIf TabStrip1.Value = 1 Then
        ButtonExecute.Caption = "ソースコード出力"
    Else
        Call Assert(False)
    End If
End Sub

Private Sub ButtonExecute_Click()
    Call Assert(OrValue(TabStrip1.Value, 0, 1))
    
    If ListBox_SelectedCount(ListBox1.Object) = 0 Then
        Call MsgBox("実行対象がありません")
        Exit Sub
    End If

    If TabStrip1.Value = 0 Then
        If MsgBox("ソースコードを入力します。よろしいですか？", _
            VbMsgBoxStyle.vbOKCancel) <> VbMsgBoxResult.vbOK Then Exit Sub
    ElseIf TabStrip1.Value = 1 Then
        If MsgBox("ソースコードを出力します。よろしいですか？", _
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
            '入力
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
            '出力
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
    
    'VBModuleを操作すると、ListBoxの選択が解除されるという
    '意味不明なバグがあったので対処する
    'トレース実行時はこの問題は見えない
    Dim SearchIndex As Long
    For I = 0 To ListSelectFileItem.Count - 1
    Do
        SearchIndex = ListFileItem.IndexOf_3(ListSelectFileItem(I))
        ListBox1.Selected(SearchIndex) = True
    Loop While False
    Next
    
    Set ListSelectFileItem = Nothing
    
    If TabStrip1.Value = 0 Then
        MsgBox ("入力完了")
    ElseIf TabStrip1.Value = 1 Then
        MsgBox ("出力完了")
    End If
End Sub
