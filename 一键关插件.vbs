'@codepage=65001
Option Explicit

' 定义一个变量来保存开启插件的项目列表
Dim Projects
Projects = ""

' 用户输入
Dim userChoice : userChoice = "f"

' 根据用户的选择执行相应的操作
If userChoice = "t" Then '开
    ProcessFolders GetScriptDirectory(), True
ElseIf userChoice = "f" Then '关
    ProcessFolders GetScriptDirectory(), False
Else
    WScript.Quit
End If

' 获取当前 VBScript 文件所在的目录
Function GetScriptDirectory()
    Dim objFSO, objShell
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objShell = CreateObject("WScript.Shell")
    GetScriptDirectory = objFSO.GetParentFolderName(WScript.ScriptFullName)
End Function

' 主要函数，遍历目录并处理文件
Sub ProcessFolders(ByVal folderPath, ByVal openPlugin)
    On Error Resume Next ' 启用错误处理

    Dim objFSO, objFolder, objSubfolder
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(folderPath)

    If Err.Number <> 0 Then
        WScript.Echo "Error: " & Err.Description
        Exit Sub
    End If

    ' 进入当前目录的所有子目录
    For Each objSubfolder In objFolder.Subfolders
        ProcessSubfolder objSubfolder, openPlugin
    Next

    On Error GoTo 0 ' 关闭错误处理
End Sub

' 处理子目录
Sub ProcessSubfolder(ByVal subfolder, ByVal openPlugin)
    On Error Resume Next ' 启用错误处理

    Dim objFSO, objFolder, objFile
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(subfolder)

    If Err.Number <> 0 Then
        WScript.Echo "Error: " & Err.Description
        Err.Clear
        Exit Sub
    End If
    
	' 遍历子目录下的所有子文件夹
	Dim objSubfolder
    For Each objSubfolder In objFolder.Subfolders
        ' 如果子文件夹名称为 "Content"，进入其中
        If LCase(objSubfolder.Name) = "content" Then
			If openPlugin Then
				' 复制脚本当前目录下的 "\Replace\Plugin\" 中的所有文件到项目的 "Content" 目录
				Dim sourceFolder : sourceFolder = CreateObject("WScript.Shell").CurrentDirectory & "\Replace\Plugin\"
				Dim destFolder : destFolder = objSubfolder.Path
				CopyFiles sourceFolder, destFolder
			Else
				' 删除项目的 Content 目录下的 \Replace\Plugin\ 中的所有文件
                Dim contentFolder : contentFolder = objSubfolder.Path
                DeleteFiles contentFolder
			End If
        End If
    Next
    
    ' 遍历当前目录下的所有文件
    For Each objFile In objFolder.Files
        ' 找到项目文件,开关插件
        If LCase(objFile.Name) = "model.uproject" Then
            Dim projectFilePath, newContent, projectFile, startIndex, endIndex
            ' 读取项目文件内容
            projectFilePath = objFolder.Path & "\" & objFile.Name 
            Dim fileContent : fileContent = objFSO.OpenTextFile(projectFilePath, 1).ReadAll()
            Dim replaceText : replaceText = """Plugins"": []"
			If openPlugin Then
				replaceText = """Plugins"": [{""Name"":""BlueprintFileUtils"",""Enabled"":true},{""Name"":""EditorScriptingUtilities"",""Enabled"":true},{""Name"":""ModelingToolsEditorMode"",""Enabled"":true},{""Name"":""DatasmithImporter"",""Enabled"":true}]"
			End If
			' 开启插件 查找 Plugins 部分并设置
			startIndex = InStr(1, fileContent, """Plugins""", vbTextCompare)
			If startIndex > 0 Then
				endIndex = InStr(startIndex, fileContent, "]", vbTextCompare)
				If endIndex > 0 Then
					newContent = Mid(fileContent, 1, startIndex - 1) & replaceText & Mid(fileContent, endIndex + 1)
					' 保存修改后的内容
					Set projectFile = objFSO.OpenTextFile(projectFilePath, 2) ' 2 表示以写入方式打开文件
					projectFile.Write newContent
				End If
			End If
            ' 将项目名称添加到列表中
            projectFile.Close
            Projects = Projects & vbCrLf & objFolder.Name
        End If
    Next

    On Error GoTo 0 ' 关闭错误处理
End Sub

' 复制文件函数
Sub CopyFiles(ByVal sourceFolder, ByVal destFolder)
    Dim objFSO, objShell, objSourceFolder, objDestFolder
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objShell = CreateObject("WScript.Shell")
    Set objSourceFolder = objFSO.GetFolder(sourceFolder)
    Set objDestFolder = objFSO.GetFolder(destFolder)

    ' 遍历源文件夹中的文件
    Dim objFile
    For Each objFile In objSourceFolder.Files
        Dim sourcePath : sourcePath = objFile.Path
        Dim destPath : destPath = objDestFolder.Path & "\" & objFile.Name
        ' 复制文件
        objFSO.CopyFile sourcePath, destPath
    Next
End Sub
' 删除文件函数
Sub DeleteFiles(ByVal folderPath)
    Dim objFSO, objFolder, objFile
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(folderPath)

    ' 获取插件文件夹路径
    Dim pluginFolderPath : pluginFolderPath = CreateObject("WScript.Shell").CurrentDirectory & "\Replace\Plugin\"

    ' 遍历插件文件夹中的文件
    Dim objPluginFolder, pluginFile
    Set objPluginFolder = objFSO.GetFolder(pluginFolderPath)

    ' 遍历目标文件夹中的文件
    For Each pluginFile In objPluginFolder.Files
        ' 遍历目标文件夹中的文件
        For Each objFile In objFolder.Files
            If LCase(objFile.Name) = LCase(pluginFile.Name) Then
                ' 如果文件名称匹配，删除目标文件夹中的文件
                objFile.Delete
            End If
        Next
    Next
End Sub



' 弹窗显示开启或关闭插件的项目列表
If Projects <> "" Then
    Dim projectListTitle
    If userChoice = "t" Then
        projectListTitle = "Opened Plugins In Projects:"
    Else
        projectListTitle = "Closed Plugins In Projects:"
    End If
    WScript.Echo projectListTitle & vbCrLf & Projects
Else
    Dim noProjectsMessage
    If userChoice = "t" Then
        noProjectsMessage = "No Plugins Opened in Projects."
    Else
        noProjectsMessage = "No Plugins Closed in Projects."
    End If
    WScript.Echo noProjectsMessage
End If
