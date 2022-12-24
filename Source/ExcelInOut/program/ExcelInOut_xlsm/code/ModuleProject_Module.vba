'--------------------------------------------------
'Excel MakeApp Framework
'--------------------------------------------------
'ModuleName:    Project Module
'ObjectName:    ModuleProject
'--------------------------------------------------
'Version:       2015/07/29
'--------------------------------------------------
Option Explicit

'--------------------------------------------------
'���v���W�F�N�g�ݒ�
'--------------------------------------------------
Public Function Project_Name() As String
    Project_Name = "ExcelInOut"
End Function

Public Function Project_AppID() As String
    Project_AppID = "StandardSoftware.ExcelMakeAppFramework." + Project_Name
End Function

Public Function Project_ScriptFileName() As String
    Project_ScriptFileName = Project_Name + ".vbs"
End Function

Public Function Project_ProgramFolderName() As String
    Project_ProgramFolderName = "program"
End Function

Public Function Project_StartMenuFolderName() As String
    Project_StartMenuFolderName = "Excel MakeApp"
End Function

Public Function Project_ShortcutFileName() As String
    Project_ShortcutFileName = Project_Name
End Function
    
Public Function Project_FormMainTitle() As String
    Project_FormMainTitle = Project_Name + " ver " + Project_VersionNumberText
End Function
    
Public Function Project_FormCreateAppShortcut_Title() As String
    Project_FormCreateAppShortcut_Title = Project_Name
End Function

Public Function Project_MainIconFileName() As String
    Project_MainIconFileName = "FormMainIcon.ico"
End Function
    
Public Function Project_MainIconIndex() As Long
    Project_MainIconIndex = 0
End Function


'--------------------------------------------------
'���o�[�W�������
'--------------------------------------------------
Public Function Project_VersionNumberText() As String
    Project_VersionNumberText = "1.3.0"
End Function

Public Function Project_VersionDialogWindowTitle() As String
    Project_VersionDialogWindowTitle = Project_Name + " �̃o�[�W�������"
End Function
    
Public Function Project_VersionDialogInstruction() As String
    Project_VersionDialogInstruction = "�o�[�W�������"
End Function
    
Public Function Project_VersionDialogContent() As String
    Project_VersionDialogContent = Project_Name + vbCrLf + _
    "   " + Project_VersionNumberText
End Function



