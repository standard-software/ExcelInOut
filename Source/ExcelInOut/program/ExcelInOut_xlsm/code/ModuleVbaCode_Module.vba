Option Explicit

Public Enum ComponentType
    STANDARD_MODULE = 1
    CLASS_MODULE = 2
    USER_FORM = 3
    OBJECT_MODULE = 100
    OTHER_UNKNOWN = 999
End Enum

'----------------------------------------
'・ VBAソースコードの出力
'----------------------------------------
Public Function VBACode_Output( _
  FolderPath_Output As String, _
  VBProject As VBProject, _
  OutputNoCodeModule As Boolean) As Boolean

    Dim Result As Boolean: Result = False

    Do
        If VBProject.Protection = vbext_pp_locked Then
            Exit Do
        End If
    
        If fso.FolderExists(FolderPath_Output) Then
            Call Folder_DeleteSubItem(FolderPath_Output)
        Else
            Call ForceCreateFolder(FolderPath_Output)
        End If

        'プロジェクトからソースを抽出する
        Dim VBComponent As VBComponent
        For Each VBComponent In VBProject.VBComponents
        Do
            Dim CodeModule As CodeModule
            Set CodeModule = VBComponent.CodeModule

            'ファイルパス取得
            Dim FilePath_Output As String

            Select Case VBComponent.Type
                Case STANDARD_MODULE
                    FilePath_Output = VBComponent.Name & "_Module.vba"
                Case CLASS_MODULE
                    FilePath_Output = VBComponent.Name & "_Class.vba"
                Case USER_FORM
                    FilePath_Output = VBComponent.Name & "_UserForm.vba"
                Case OBJECT_MODULE
                    FilePath_Output = VBComponent.Name & "_Object.vba"
                Case Else
                    FilePath_Output = VBComponent.Name & "_Other.vba"
            End Select


            '生成ファイルのフルパスを取得
            FilePath_Output = PathCombine(FolderPath_Output, FilePath_Output)

            If OutputNoCodeModule = False Then
                'コードがないモジュールは出力しない場合

                If (CodeModule.CountOfLines = 0) Then
                    Exit Do
                ElseIf OrValue(Trim(ExcludeCRLF(CodeModule.Lines(1, CodeModule.CountOfLines))), _
                    "", "Option Explicit") Then
                    '改行削除してトリムして残ったものが
                    '空文字か[Option Explicit]なら空のモジュールとして飛ばす
                    Exit Do
                End If
            End If

            'コードの出力
            Dim SB As New st_vba_StringBuilder
            Call SB.Clear
            Dim I As Long
            For I = 1 To CodeModule.CountOfLines
                SB.Add (CodeModule.Lines(I, 1) + vbCrLf)
            Next
            Call String_SaveToFile(SB.Text, FilePath_Output)

            Set CodeModule = Nothing
        Loop While False
        Next VBComponent
        Result = True

    Loop While False

    VBACode_Output = Result
End Function

Public Sub test_VBACode_Output()
    Call VBACode_Output( _
        FolderPath_VBACode(Book_FullPath(ThisWorkbook), "code"), _
        ThisWorkbook.VBProject, False)
        
    Dim TestFilePath As String
    TestFilePath = AbsolutePath(ThisWorkbook.Path, "..\..\test\test01\Book1.xls")
    Dim Book_Test As Workbook
    Set Book_Test = App_GetOpenedBookOrOpenBook(Application, _
        TestFilePath, False, True)
    Call VBACode_Output( _
        FolderPath_VBACode(Book_FullPath(ThisWorkbook), "code"), _
        Book_Test.VBProject, False)
End Sub

'----------------------------------------
'・ VBAソースコードの入力
'----------------------------------------

Public Function VBACode_Input( _
  FolderPath_Input As String, _
  VBProject As VBProject) As Boolean

    Dim Result As Boolean: Result = False

    Do
        If fso.FolderExists(FolderPath_Input) = False Then
            Exit Do
        End If

        If VBProject.Protection = vbext_pp_locked Then
            Exit Do
        End If

        'ファイルを列挙
        Dim File As Object
        For Each File In fso.GetFolder(FolderPath_Input).Files()

        Do
            Dim FileName As String
            FileName = fso.GetFileName(File)

            Dim ModuleName As String
            ModuleName = ""
            Dim ModuleType As ComponentType

            If IsLastStr(FileName, "_Module.vba") Then
                ModuleType = STANDARD_MODULE
                ModuleName = ExcludeLastStr(FileName, "_Module.vba")
            ElseIf IsLastStr(FileName, "_Class.vba") Then
                ModuleType = CLASS_MODULE
                ModuleName = ExcludeLastStr(FileName, "_Class.vba")
            ElseIf IsLastStr(FileName, "_UserForm.vba") Then
                ModuleType = USER_FORM
                ModuleName = ExcludeLastStr(FileName, "_UserForm.vba")
            ElseIf IsLastStr(FileName, "_Object.vba") Then
                ModuleType = OBJECT_MODULE
                ModuleName = ExcludeLastStr(FileName, "_Object.vba")
            ElseIf IsLastStr(FileName, "_Other.vba") Then
                ModuleType = OTHER_UNKNOWN
                ModuleName = ExcludeLastStr(FileName, "_Other.vba")
            Else
                '該当しないファイルの場合は処理しない
                Exit Do
            End If

            If ModuleExists(VBProject, ModuleName) = False Then
                'モジュールが存在しない場合はソースコード作成
                Dim VBComponent As VBComponent
                Select Case ModuleType
                    Case STANDARD_MODULE
                        Set VBComponent = _
                            VBProject.VBComponents.Add(vbext_ct_StdModule)
                        VBComponent.Name = ModuleName
                    Case CLASS_MODULE
                        Set VBComponent = _
                            VBProject.VBComponents.Add(vbext_ct_ClassModule)
                        VBComponent.Name = ModuleName
                    Case Else
                        '標準モジュールとクラスモジュール以外のモジュール追加は認めない
                        Exit Do
                End Select
            End If

            'ソースコード削除
            'モジュールが元から存在した場合でも追加された場合でも実行。
            '特に追加された場合は Option Explicit が追加されている場合があるので
            '一度削除した方がよい
            Call VBProject.VBComponents(ModuleName).CodeModule.DeleteLines(1, _
                VBProject.VBComponents(ModuleName).CodeModule.CountOfLines)

            Dim CodeModule As CodeModule
            Set CodeModule = VBProject.VBComponents(ModuleName).CodeModule


            Dim FileText() As String
            FileText = Split(String_LoadFromFile(File.Path), vbCrLf)

            'CodeModuleは1オリジン
            Dim I As Long
            For I = 1 To ArrayCount(FileText)

                Call CodeModule.InsertLines(I, FileText(I - 1) + vbCrLf)
                '改行コード無しでInsertLinesすると
                'アンダーバー行が消えるという不具合があるのでこのようにする
            Next

            'ソース中にアンダーバーがある場合空行が増える不具合のために
            '行数がおかしい場合は増加している行を削除するコード
            Dim CodeLineCount As Long: CodeLineCount = ArrayCount(FileText)

            If (1 <= CodeModule.CountOfLines) And (CodeLineCount < CodeModule.CountOfLines) Then
                Dim J As Long
                For J = CodeModule.CountOfLines To CodeLineCount Step -1
                    If OrValue(CodeModule.Lines(J, 1), "", "()") Then
                        Call CodeModule.DeleteLines(J, 1)
                    End If
                Next
                'CodeModuleは1オリジンなので
                'J=CodeLineCountの位置 が 最終行になる。
                'DeleteLinesすることでファイルの最後に空行があっても
                'IDE上には正しく挿入される
            End If

            Set CodeModule = Nothing

            Result = True

        Loop While False
        Next

    Loop While False

    VBACode_Input = Result
End Function

Public Sub test_VBACode_Input()
        
    Dim TestFilePath As String
    TestFilePath = AbsolutePath(ThisWorkbook.Path, "..\..\test\test02\Book1.xls")
    Dim Book_Test As Workbook
    Set Book_Test = App_GetOpenedBookOrOpenBook(Application, _
        TestFilePath, False, True)
    Call VBACode_Output( _
        FolderPath_VBACode(Book_FullPath(ThisWorkbook), "output1"), _
        Book_Test.VBProject, False)

    Call VBACode_Input( _
        FolderPath_VBACode(Book_FullPath(ThisWorkbook), "output1"), _
        Book_Test.VBProject)

    Call VBACode_Output( _
        FolderPath_VBACode(Book_FullPath(ThisWorkbook), "output2"), _
        Book_Test.VBProject, False)
End Sub

'----------------------------------------
'・コンポーネント(モジュール)の存在確認関数
'----------------------------------------
Public Function ModuleExists(VBProject As VBProject, ModuleName As String) As Boolean
    Dim Result As Boolean
    Result = False
    Dim I
    For I = 1 To VBProject.VBComponents.Count
        If LCase(VBProject.VBComponents(I).Name) = LCase(ModuleName) Then
            Result = True
            Exit For
        End If
    Next
    ModuleExists = Result
End Function

'----------------------------------------
'・ 入出力フォルダファイル名
'----------------------------------------
Function FolderPath_VBACode( _
ByVal FilePath As String, _
Optional ByVal SubFolderName As String = "") As String

    Dim Result As String

    Result = ChangeFileExtension(FilePath, _
        IncludeFirstStr( _
            ExcludeFirstStr( _
                GetExtensionIncludePeriod(FilePath), _
                "."), _
            "_"))
    'C:\Temp\Book1.xls に対して C:\Temp\Book1_xls というフォルダを返す
    
    If SubFolderName <> "" Then
        Result = PathCombine(Result, SubFolderName)
    End If

    FolderPath_VBACode = Result
End Function


'----------------------------------------
'・改行コードを取り除く
'----------------------------------------
Public Function ExcludeCRLF(S As String) As String
    Dim Result As String
    Result = Replace(Replace(S, vbCr, ""), vbLf, "")
    ExcludeCRLF = Result
End Function


