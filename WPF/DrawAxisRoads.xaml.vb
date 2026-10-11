Imports System.Data
Imports System.Windows
Imports System.Windows.Automation.Peers
Imports System.Windows.Controls
Imports System.Windows.Forms
Imports HostMgd.ApplicationServices
Imports HostMgd.EditorInput
Imports Teigha.DatabaseServices
Imports Teigha.Geometry

Public Class AxisParameters
    Public Property NameElement As String
    Public Property boolFixed As Boolean
    Public Property value As Double
End Class



Public Class DrawAxisRoads
    Private Function LoadCollectionData(ByRef arrayParameters As String(,)) As List(Of AxisParameters)
        Dim axisParam As List(Of AxisParameters) = New List(Of AxisParameters)()
        If IsArray(arrayParameters) = True Then
            If arrayParameters.GetUpperBound(0) > 1 Then
                For i As Integer = 0 To arrayParameters.GetUpperBound(1)
                    Dim nameElement As String = arrayParameters(0, i)
                    Dim boolFix As Boolean = False
                    If arrayParameters(1, i) Like "True" Then
                        boolFix = True
                    End If
                    Dim valueParam As Double = Val(arrayParameters(2, i))

                    axisParam.Add(New AxisParameters() With {
                           .NameElement = nameElement,
                           .boolFixed = boolFix,
                           .value = valueParam
                       })
                Next i
            End If
        End If
        Return axisParam
    End Function



    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        Dim items0 As TreeViewItem = New TreeViewItem()
        items0.Header = "Трасса автомобильной дороги"
        items0.Tag = ""
        TreeView1.Items.Add(items0)
        '=====================================================================================
        Dim items1 As TreeViewItem = New TreeViewItem()
        items1.Header = "Прямой участок"
        items1.Tag = ""
        items0.Items.Add(items1)
        ' Добавить код инициализации после вызова InitializeComponent().
        Dim items2 As TreeViewItem = New TreeViewItem()
        items2.Header = "Круговая кривая"
        items2.Tag = ""
        items0.Items.Add(items2)
        ' Добавить код инициализации после вызова InitializeComponent().
        Dim items3 As TreeViewItem = New TreeViewItem()
        items3.Header = "Переходная кривая"
        items3.Tag = ""
        items0.Items.Add(items3)
        '==========================================================================================
        Dim items1_1 As TreeViewItem = New TreeViewItem()
        items1_1.Header = "Прямая"
        items1_1.Tag = ""
        items1.Items.Add(items1_1)

        Dim items1_2 As TreeViewItem = New TreeViewItem()
        items1_2.Header = "Прямая-Прямая"
        items1_2.Tag = ""
        items1.Items.Add(items1_2)

        Dim items1_3 As TreeViewItem = New TreeViewItem()
        items1_3.Header = "Тангенциальный ход"
        items1_3.Tag = ""
        items1.Items.Add(items1_3)

    End Sub
    '================================================================================================================
    'выбор ветки
    Private Sub TreeView1_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles TreeView1.SizeChanged

    End Sub

    Private Sub TreeView1_SelectedItemChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles TreeView1.SelectedItemChanged
        Dim itemSelect As TreeViewItem = TreeView1.SelectedItem

        If itemSelect.Header Like "Прямая" Then
            Dim arrayParam As String(,) = Nothing
            ReDim arrayParam(2, 1)
            arrayParam(0, 0) = "Длина, м"
            arrayParam(1, 0) = "False"
            arrayParam(2, 0) = 0

            arrayParam(0, 1) = "Азимут, дд.mmss"
            arrayParam(1, 1) = "False"
            arrayParam(2, 1) = 0

            DataGrid1.ItemsSource = LoadCollectionData(arrayParam)
            DataGrid1.Tag = "Прямая"

            '    col1.Header = "Тип элемента"
            '    col1.Visibility = True
            '    DataGrid1.Columns.Add(col1)

            '    Dim col2 As DataGridComboBoxColumn = New DataGridComboBoxColumn()
            '    Dim arrayData As String() = {"Фиксированное значение", "Расчетное значение"}
            '    col2.ItemsSource = arrayData
            '    col2.Header = "Действие"
            '    col2.Visibility = True
            '    DataGrid1.Columns.Add(col2)

            '    Dim col3 As DataGridTextColumn = New DataGridTextColumn
            '    col3.Header = "Значение"
            '    col3.Visibility = True
            '    DataGrid1.Columns.Add(col3)

            '    Dim dr As DataGridRow = New DataGridRow()
            '    dr.Header = "Длина 1"
        End If
    End Sub

    Private Sub DataGrid1_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles DataGrid1.SelectionChanged

    End Sub

    Private Sub DataGrid_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)

    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'создать объект
    Private Sub Button1_Click(sender As Object, e As RoutedEventArgs) Handles Button1.Click
        Dim doc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Dim nameObject As String = DataGrid1.Tag
        Select Case nameObject
            Case "Прямая"
                Dim userLenght As Double = 0
                Dim userAzimut As Double = 361
                Dim userItemColl As ItemCollection = DataGrid1.Items
                If userItemColl.Count > 0 Then
                    For i As Integer = 0 To userItemColl.Count - 2
                        Dim axisProp As AxisParameters = userItemColl.Item(i)
                        If axisProp.NameElement Like "Прямая" Then
                            If axisProp.boolFixed = True Then
                                userLenght = axisProp.value
                            End If
                        ElseIf axisProp.NameElement Like "Азимут" Then
                            If axisProp.boolFixed = True Then
                                userAzimut = axisProp.value
                            End If
                        End If
                    Next
                End If
                If userAzimut <> 361 And userLenght <> 0 Then
                    Dim pPtRes As PromptPointResult
                    Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                    pPtOpts.Message = vbLf & "Укажите первую точку прямой: "
                    pPtRes = doc.Editor.GetPoint(pPtOpts)
                    Dim ptStart As Point3d = pPtRes.Value
                    If pPtRes.Status = PromptStatus.Cancel Then Exit Sub
                    pPtOpts.Message = vbLf & "Укажите следующую точку прямого участка: "
                    pPtOpts.UseBasePoint = True
                    pPtOpts.BasePoint = ptStart
                    pPtRes = doc.Editor.GetPoint(pPtOpts)
                    Dim ptEnd As Point3d = pPtRes.Value
                    If pPtRes.Status = PromptStatus.Cancel Then Exit Sub
                    Dim tr As Transaction = db.TransactionManager.StartTransaction()
                    Using tr
                        Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                        Dim btr As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                        Using userLine As Line = New Line(ptStart, ptEnd)
                            btr.AppendEntity(userLine)
                            tr.AddNewlyCreatedDBObject(userLine, True)
                        End Using
                        tr.Commit()
                    End Using
                End If
        End Select
    End Sub
End Class
