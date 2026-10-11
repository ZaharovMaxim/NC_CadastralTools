Imports System.IO
Imports System.Text.RegularExpressions
Imports HostMgd.ApplicationServices
Imports HostMgd.EditorInput
Imports HostMgd.Windows
Imports Teigha.DatabaseServices
Imports Teigha.Geometry
Imports Teigha.GraphicsInterface
Imports Teigha.Runtime
Public Class FuncFiles
    Public Shared Function FuncGetSupportFolderPath(ByRef strPathFolders As String()) As Boolean
        Dim app_com As nanoCAD.InanoCADApplication = HostMgd.ApplicationServices.Application.AcadApplication
        Dim supportFolders As String = app_com.Preferences.Files.SupportPath
        strPathFolders = supportFolders.Split(";")
        If IsArray(strPathFolders) = True Then
            Return True
        Else
            Return False
        End If
        'Dim tempList As HashSet(Of String) = New HashSet(Of String)
        'If arrayStr.Length > 0 Then
        '    For i As Integer = 0 To arrayStr.Length - 1
        '        tempList.Add(arrayStr(i))
        '    Next
        'End If
        'Return tempList
    End Function
End Class
