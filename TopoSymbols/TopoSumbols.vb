Imports System.IO
Imports System.Windows.Forms.Integration
Imports HostMgd.ApplicationServices
Imports HostMgd.EditorInput
Imports HostMgd.Windows
Imports nanoCAD
Imports Teigha.DatabaseServices
Imports Teigha.Geometry
Imports Teigha.Runtime
Public Class TopoSumbols
    Public nameProgram As String = "Civil3DToolsUtility"
    Public templateXML As String = ""
    Public templateDwg As String = ""
    Public templateFieldData As String = ""

    <CommandMethod("ProjectCadastralSumbols")>
    Public Sub TopoZN()
        '1.Ищем файлы с шаблонами
        Dim arraySupportFiles As String() = {}
        Dim boolFindSupportFiles As Boolean = FuncFiles.FuncGetSupportFolderPath(arraySupportFiles)
        If arraySupportFiles.Length > 0 Then
            For i As Integer = 0 To arraySupportFiles.Length - 1
                Dim tempSampleDir As String = arraySupportFiles.ElementAt(i)
                Dim pos As Integer = tempSampleDir.IndexOf(nameProgram)
                If pos > -1 Then
                    Dim filesSample As String() = Directory.GetFiles(tempSampleDir & "\FileResources\Sample\")
                    If filesSample.Length > 0 Then
                        For j As Integer = 0 To filesSample.Length - 1
                            Dim fileSample As String = filesSample(j)
                            Dim ext As String = Path.GetExtension(fileSample)
                            If ext Like ".xml" Then
                                templateXML = fileSample
                            ElseIf ext Like ".dwg" Then
                                templateDwg = fileSample
                            End If
                        Next j
                    End If
                End If
                If IO.File.Exists(tempSampleDir & "\FileResources\Sample\Semantic\FieldData.xml") = True Then
                    templateFieldData = tempSampleDir & "\FileResources\Sample\Semantic\FieldData.xml"
                End If
            Next
        End If

        Dim arraySampleFiles As String() = {templateXML, templateDwg, templateFieldData}

        Dim PaletteSet As PaletteSet = New HostMgd.Windows.PaletteSet("Земельный кадастр")
        '=========================================================================================
        'форма с условными знаками
        Dim PaletteTopoSumbol As TopoSymbolsWPF = New TopoSymbolsWPF(templateXML)
        PaletteTopoSumbol.templateDWG = templateDwg
        PaletteTopoSumbol.templateDataXML = templateFieldData
        PaletteTopoSumbol.Height = 800
        PaletteTopoSumbol.Width = 400
        'форма с пользовательскими свойствами
        Dim PaletteProperties As UserXRecords = New UserXRecords()
        PaletteProperties.templateDataXML = templateFieldData
        PaletteProperties.templateXML = templateXML
        PaletteProperties.Height = 800
        PaletteProperties.Width = 400
        'форма для выбора объектов
        'Dim PalleteSelectObj As SelectCadastralObject = New SelectCadastralObject
        'PaletteSet.Size = New System.Drawing.Size(PaletteTopoSumbol.Width, PaletteTopoSumbol.Height)
        'PaletteTopoSumbol.userPal = PaletteSet
        'PaletteProperties.userPal = PaletteSet
        'PalleteSelectObj.userPal = PaletteSet
        'PaletteSet.Visible = True

        Dim hostView1 As New ElementHost() With {
            .AutoSize = False,
            .Dock = System.Windows.Forms.DockStyle.Fill,
            .Child = PaletteTopoSumbol ' AnyWPFcontrol -- ваш WPF контрол
        }
        PaletteSet.Add("Создание объектов", hostView1)

        Dim hostView2 As New ElementHost() With {
           .AutoSize = False,
           .Dock = System.Windows.Forms.DockStyle.Fill,
           .Child = PaletteProperties ' AnyWPFcontrol -- ваш WPF контрол
       }
        PaletteSet.Add("Редактирование", hostView2)
        PaletteSet.Visible = True
        ' Dim hostView3 As New ElementHost() With {
        '    .AutoSize = False,
        '    .Dock = System.Windows.Forms.DockStyle.Fill,
        '    .Child = PalleteSelectObj ' AnyWPFcontrol -- ваш WPF контрол
        '}

        ' PaletteSet.Add("Выбор объектов", hostView3)
    End Sub
End Class
