Option Explicit

Public Enum ComponentType
    STANDARD_MODULE = 1
    CLASS_MODULE = 2
    USER_FORM = 3
    OBJECT_MODULE = 100
    OTHER_UNKNOWN = 999
End Enum

'----------------------------------------
'�E VBA�\�[�X�R�[�h�̏o��
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

        '�v���W�F�N�g����\�[�X�𒊏o����
        Dim VBComponent As VBComponent
        For Each VBComponent In VBProject.VBComponents
        Do
            Dim CodeModule As CodeModule
            Set CodeModule = VBComponent.CodeModule

            '�t�@�C���p�X�擾
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


            '�����t�@�C���̃t���p�X���擾
            FilePath_Output = PathCombine(FolderPath_Output, FilePath_Output)

            If OutputNoCodeModule = False Then
                '�R�[�h���Ȃ����W���[���͏o�͂��Ȃ��ꍇ

                If (CodeModule.CountOfLines = 0) Then
                    Exit Do
                ElseIf OrValue(Trim(ExcludeCRLF(CodeModule.Lines(1, CodeModule.CountOfLines))), _
                    "", "Option Explicit") Then
                    '���s�폜���ăg�������Ďc�������̂�
                    '�󕶎���[Option Explicit]�Ȃ��̃��W���[���Ƃ��Ĕ�΂�
                    Exit Do
                End If
            End If

            '�R�[�h�̏o��
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
'�E VBA�\�[�X�R�[�h�̓���
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

        '�t�@�C�����
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
                '�Y�����Ȃ��t�@�C���̏ꍇ�͏������Ȃ�
                Exit Do
            End If

            If ModuleExists(VBProject, ModuleName) = False Then
                '���W���[�������݂��Ȃ��ꍇ�̓\�[�X�R�[�h�쐬
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
                        '�W�����W���[���ƃN���X���W���[���ȊO�̃��W���[���ǉ��͔F�߂Ȃ�
                        Exit Do
                End Select
            End If

            '�\�[�X�R�[�h�폜
            '���W���[���������瑶�݂����ꍇ�ł��ǉ����ꂽ�ꍇ�ł����s�B
            '���ɒǉ����ꂽ�ꍇ�� Option Explicit ���ǉ�����Ă���ꍇ������̂�
            '��x�폜���������悢
            Call VBProject.VBComponents(ModuleName).CodeModule.DeleteLines(1, _
                VBProject.VBComponents(ModuleName).CodeModule.CountOfLines)

            Dim CodeModule As CodeModule
            Set CodeModule = VBProject.VBComponents(ModuleName).CodeModule


            Dim FileText() As String
            FileText = Split(String_LoadFromFile(File.Path), vbCrLf)

            'CodeModule��1�I���W��
            Dim I As Long
            For I = 1 To ArrayCount(FileText)

                Call CodeModule.InsertLines(I, FileText(I - 1) + vbCrLf)
                '���s�R�[�h������InsertLines�����
                '�A���_�[�o�[�s��������Ƃ����s�������̂ł��̂悤�ɂ���
            Next

            '�\�[�X���ɃA���_�[�o�[������ꍇ��s��������s��̂��߂�
            '�s�������������ꍇ�͑������Ă���s���폜����R�[�h
            Dim CodeLineCount As Long: CodeLineCount = ArrayCount(FileText)

            If (1 <= CodeModule.CountOfLines) And (CodeLineCount < CodeModule.CountOfLines) Then
                Dim J As Long
                For J = CodeModule.CountOfLines To CodeLineCount Step -1
                    If OrValue(CodeModule.Lines(J, 1), "", "()") Then
                        Call CodeModule.DeleteLines(J, 1)
                    End If
                Next
                'CodeModule��1�I���W���Ȃ̂�
                'J=CodeLineCount�̈ʒu �� �ŏI�s�ɂȂ�B
                'DeleteLines���邱�ƂŃt�@�C���̍Ō�ɋ�s�������Ă�
                'IDE��ɂ͐������}�������
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
'�E�R���|�[�l���g(���W���[��)�̑��݊m�F�֐�
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
'�E ���o�̓t�H���_�t�@�C����
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
    'C:\Temp\Book1.xls �ɑ΂��� C:\Temp\Book1_xls �Ƃ����t�H���_��Ԃ�
    
    If SubFolderName <> "" Then
        Result = PathCombine(Result, SubFolderName)
    End If

    FolderPath_VBACode = Result
End Function


'----------------------------------------
'�E���s�R�[�h����菜��
'----------------------------------------
Public Function ExcludeCRLF(S As String) As String
    Dim Result As String
    Result = Replace(Replace(S, vbCr, ""), vbLf, "")
    ExcludeCRLF = Result
End Function


