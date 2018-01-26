Attribute VB_Name = "ExportVisualBasicCode"
' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
Public Sub ExportVBA()
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24

    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim dir_main As String
    Dim extension As String
    Dim fso As New FileSystemObject

    dir_main = "C:\Users\jsikorski\Desktop\Time Card Project - JASON\ALL VBA CODE\" & ThisWorkbook.name & "_VBA_" & Format(Now(), "mm.dd.yy_hh.mm.ss")
    directory = "C:\Users\jsikorski\Desktop\Time Card Project - JASON\hei_time\Time Card Gen BETA.xlsm_VBA"
    count = 0

    If Not fso.FolderExists(directory) Then
        Call fso.CreateFolder(directory)
    End If
    Set fso = Nothing

    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
            Case Form
                extension = ".frm"
            Case Module
                extension = ".bas"
            Case Else
                extension = ".txt"
        End Select


        On Error Resume Next
        Err.Clear

        path = directory & "\" & VBComponent.name & extension
        Call VBComponent.Export(path)

        If Err.Number <> 0 Then
            Call MsgBox("Failed to export " & VBComponent.name & " to " & path, vbCritical)
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.name & ":" & Space(Padding), Padding) & path
        End If

        On Error GoTo 0
    Next

    If Not fso.FolderExists(dir_main) Then
        Call fso.CreateFolder(dir_main)
    End If
    Set fso = Nothing

    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
            Case Form
                extension = ".frm"
            Case Module
                extension = ".bas"
            Case Else
                extension = ".txt"
        End Select


        On Error Resume Next
        Err.Clear

        path = dir_main & "\" & VBComponent.name & extension
        Call VBComponent.Export(path)

        If Err.Number <> 0 Then
            Call MsgBox("Failed to export " & VBComponent.name & " to " & path, vbCritical)
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.name & ":" & Space(Padding), Padding) & path
        End If

        On Error GoTo 0
    Next

    Application.StatusBar = "Successfully exported " & CStr(count) & " VBA files to " & dir_main
    Application.StatusBar = False
End Sub
