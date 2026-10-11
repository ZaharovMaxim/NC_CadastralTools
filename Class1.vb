Imports HostMgd.ApplicationServices
Imports HostMgd.EditorInput
Imports HostMgd.Windows
Imports Teigha.DatabaseServices
Imports Teigha.Geometry
Imports Teigha.Runtime
Public Class Class1

    <CommandMethod("HelloWord1")>
    Public Sub TestHelloWorld()
        MsgBox("Hello World!!!")
        Dim app_com As nanoCAD.InanoCADApplication = HostMgd.ApplicationServices.Application.AcadApplication
        Dim p As String = app_com.Preferences.Files.SupportPath
        app_com.Preferences.Files.SetProjectFilePath("", "c:\")
        Dim a = 1
    End Sub



End Class
