Imports System.Collections.ObjectModel
Imports System.Security.Cryptography
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel
Imports System.Windows.Input
Imports CadastralTools
Imports HostMgd.ApplicationServices
Imports HostMgd.EditorInput
Imports nanoCAD
Imports NetTopologySuite.Utilities
Imports Teigha.BoundaryRepresentation
Imports Teigha.DatabaseServices
Imports Teigha.Runtime
Imports Application = HostMgd.ApplicationServices.Application
Imports Document = HostMgd.ApplicationServices.Document
Imports SelectionSet = HostMgd.EditorInput.SelectionSet

Public Class ObjectData
    Private _data As StructureElement
    Private _ps As String
    Private _typeTablePS As StructureElement.tableXRecords
    Private _id As ObjectId
    Public Sub New()
        _data = Nothing
        _ps = ""
        _typeTablePS = StructureElement.tableXRecords.PROJECT_STRUCTURES
        _id = Nothing
    End Sub
    Public Sub New(userData As StructureElement, psTable As String, typeTablePS As StructureElement.tableXRecords, idObject As ObjectId)
        _data = userData
        _ps = psTable
        _typeTablePS = typeTablePS
        _id = idObject
    End Sub
    Public Property UserData As StructureElement
        Get
            Return _data
        End Get
        Set(value As StructureElement)
            _data = value
        End Set
    End Property

    Public Property PSTable As String
        Get
            Return _ps
        End Get
        Set(value As String)
            _ps = value
        End Set
    End Property
    Public Property TypeTablePS As StructureElement.tableXRecords
        Get
            Return _typeTablePS
        End Get
        Set(value As StructureElement.tableXRecords)
            _typeTablePS = value
        End Set
    End Property
    Public Property IDObject As ObjectId
        Get
            Return _id
        End Get
        Set(value As ObjectId)
            _id = value
        End Set
    End Property
End Class
Public Class UserXRecords
    Public userPal As HostMgd.Windows.PaletteSet = Nothing
    Public templateDataXML As String = ""
    Public templateXML As String = ""
    Private _editor As Editor
    Private _editEntity As Entity = Nothing
    Private WPFDataGrid As ObservableCollection(Of CadastralTools.WPFDataGrid) = New ObservableCollection(Of WPFDataGrid)
    Private dictXRef As Dictionary(Of String, String) = New Dictionary(Of String, String)
    Private dictProjectPlaningObject As Dictionary(Of String, List(Of CadastralTools.Topology)) = New Dictionary(Of String, List(Of CadastralTools.Topology))
    Private dictCadastralObject As List(Of Topology) = New List(Of Topology)
    Private listObjectHidlight As List(Of Entity) = New List(Of Entity)
    Private generalNode As TreeViewItem = New TreeViewItem()
    'записать данные в примитив
    Public Function recEntityData() As CadastralTools.StructureElement
        If IsNothing(_editEntity) = False Then
            'читаем все словари из выбранного объекта
            Dim listDictionary As Dictionary(Of String, CadastralTools.StructureElement) = FuncXRecords.readXRecords(_editEntity)
            If IsNothing(listDictionary) = False Then
                If listDictionary.Count > 0 Then
                    For i As Integer = 0 To listDictionary.Count - 1
                        Dim dataStructures As CadastralTools.StructureElement = listDictionary.ElementAt(i).Value
                        Dim classTable As String = listDictionary.ElementAt(i).Key 'читаем имя словаря
                        Dim valueData As StructureElement = listDictionary.ElementAt(i).Value
                        Dim dataObject As ObjectData = DGidPropertiesX.Tag
                        'если в палитре таблица заполнена
                        If IsNothing(DGidPropertiesX.ItemsSource) = False Then
                            If DGidPropertiesX.Items.Count > 0 Then
                                If dataObject.PSTable Like valueData.TablePS Then
                                    Dim dataRec As ObservableCollection(Of CadastralTools.WPFDataGrid) = DGidPropertiesX.ItemsSource
                                    If dataRec.Count > 0 Then
                                        Dim arrayDataRec As String(,) = {}
                                        Dim countArray As Integer = 0
                                        For j As Integer = 0 To dataRec.Count - 1
                                            Dim custData As CadastralTools.WPFDataGrid = dataRec.ElementAt(j)
                                            If IsNothing(custData.nameField) = True Then Continue For
                                            If custData.nameField.Trim.Length = 0 Then Continue For
                                            If IsNothing(custData.Имя_свойства) = True Then Continue For
                                            If custData.Имя_свойства.Trim.Length = 0 Then Continue For
                                            If IsNothing(custData.Значение) = True Then custData.Значение = ""
                                            ReDim Preserve arrayDataRec(2, countArray)
                                            arrayDataRec(0, countArray) = custData.nameField
                                            arrayDataRec(1, countArray) = custData.Значение
                                            arrayDataRec(2, countArray) = custData.Имя_свойства
                                            countArray += 1
                                        Next j
                                        If arrayDataRec.Length > 0 Then
                                            Dim gsonStr As String = CadastralTools.JsonValidator.convertArray2DToJson(arrayDataRec)
                                            dataStructures.KeyParameter = gsonStr
                                            If classTable Like StructureElement.tableXRecords.PROJECT_STRUCTURES.ToString Then
                                                FuncXRecords.writeXRecords(_editEntity, dataStructures)
                                                Return dataStructures
                                            ElseIf classTable Like StructureElement.tableXRecords.CADASTRAL_OBJECT.ToString Then
                                                FuncXRecords.writeXRecords(_editEntity, dataStructures, StructureElement.tableXRecords.CADASTRAL_OBJECT)
                                                Return dataStructures
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next i
                End If
            End If
        End If
        Return Nothing
    End Function
    'заполнить treeView
    Public Function readDictionaryObject(ByRef itemNode As TreeViewItem, ByVal dictObject As Dictionary(Of String, Dictionary(Of String, Topology))) As Boolean
        If IsNothing(itemNode) = True Then Return False
        If IsNothing(dictObject) = False Then
            If dictObject.Count > 0 Then
                If itemNode.Items.Count > 0 Then
                    itemNode.Items.Clear()
                End If
                For i As Integer = 0 To dictObject.Count - 1
                    Dim numberParcel As String = dictObject.ElementAt(i).Key
                    Dim node1 As TreeViewItem = New TreeViewItem()
                    node1.Header = numberParcel
                    itemNode.Items.Add(node1)
                    Dim valueData As Dictionary(Of String, CadastralTools.Topology) = dictObject.ElementAt(i).Value
                    If IsNothing(valueData) = False Then
                        If valueData.Count > 1 Then
                            For j = 0 To valueData.Count - 1
                                Dim userTopoData As Topology = valueData.ElementAt(j).Value
                                Dim structData As StructureElement = userTopoData.StructureElement
                                Dim desk As String = structData.Description
                                Dim keyParam As String = structData.KeyParameter
                                If CadastralTools.JsonValidator.IsValidJson(keyParam) = True Then
                                    'читаем кадастровый номер
                                    desk = CadastralTools.JsonValidator.GetValueFromJson(keyParam, "CadNumber")
                                    'смотрим ксть ли номер у объекта
                                    Dim nameResult As String = CadastralTools.JsonValidator.GetValueFromJson(keyParam, "Number")
                                    If IsNothing(nameResult) = False Then
                                        desk = desk & nameResult
                                        'читаем номер контура
                                        Dim numberCounter = CadastralTools.JsonValidator.GetValueFromJson(keyParam, "NumberCounter")
                                        If Val(numberCounter) > 0 Then
                                            desk = desk & "(" & numberCounter & ")"
                                        End If
                                    ElseIf IsNothing(nameResult) = True Then
                                        nameResult = CadastralTools.JsonValidator.GetValueFromJson(keyParam, "ConditionalNumber")
                                        desk = desk & nameResult
                                    End If
                                End If
                                Dim node2 As TreeViewItem = New TreeViewItem
                                node2.Header = desk
                                node2.Tag = structData
                                node1.Items.Add(node2)
                            Next
                        ElseIf valueData.Count = 1 Then
                            Dim userTopoData As Topology = valueData.First.Value
                            Dim structData As StructureElement = userTopoData.StructureElement
                            Dim desk As String = structData.Description
                            Dim keyParam As String = structData.KeyParameter
                            If CadastralTools.JsonValidator.IsValidJson(keyParam) = True Then
                                'читаем кадастровый номер
                                desk = CadastralTools.JsonValidator.GetValueFromJson(keyParam, "CadNumber")
                                'смотрим ксть ли номер у объекта
                                Dim nameResult As String = CadastralTools.JsonValidator.GetValueFromJson(keyParam, "Number")
                                If IsNothing(nameResult) = True Then
                                    nameResult = CadastralTools.JsonValidator.GetValueFromJson(keyParam, "ConditionalNumber")
                                    desk = desk & nameResult
                                End If
                            End If
                            node1.Tag = structData
                        End If
                    End If
                Next i
            End If
        End If
    End Function
    'функция возвращает customerData
    Public Function returnCustomerData(ByVal userSData As StructureElement, Optional id As ObjectId = Nothing) As Boolean
        WPFDataGrid = New ObservableCollection(Of WPFDataGrid)
        If IsNothing(userSData) = False Then
            Dim keyParam As String = userSData.KeyParameter
            Dim PS As String = userSData.TablePS
            Dim arrayTablePSFromXML As String(,) = {}
            Dim boolReadPS As Boolean = CadastralTools.TemplateXML.FuncReadTablePSToXML(templateXML, PS, arrayTablePSFromXML)
            WPFDataGrid = CadastralTools.WPFDataGrid.createDataForCustomerDataGrid(arrayTablePSFromXML, keyParam)
            If WPFDataGrid.Count > 0 Then
                DGidPropertiesX.SelectedItems.Clear()
                DGidPropertiesX.SelectedIndex = -1
                DGidPropertiesX.CurrentCell = Nothing
                DGidPropertiesX.ItemsSource = WPFDataGrid
                If DGidPropertiesX.Columns.Count > 0 Then
                    DGidPropertiesX.Columns(0).Visibility = Visibility.Collapsed
                End If
                DGidPropertiesX.FrozenColumnCount = 1
                DGidPropertiesX.Items.Refresh()
                Dim newDataObject As ObjectData = New ObjectData(userSData, PS, StructureElement.tableXRecords.PROJECT_STRUCTURES, id)
                DGidPropertiesX.Tag = newDataObject
            End If
        End If
        Return True
    End Function
    Public Sub New()
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        ' Добавить код инициализации после вызова InitializeComponent().
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        _editor = Application.DocumentManager.MdiActiveDocument.Editor
        ' Или на событие запроса выбора объектов
        AddHandler _editor.SelectionAdded, AddressOf OnSelectionAdded
        AddHandler _editor.SelectionRemoved, AddressOf OnSelectionRemoved
        'получаем все загруженные внешние ссылки
        Dim boolFindXRef As Boolean = FuncDWGReference.getXRef(dictXRef)
        If IsNothing(dictXRef) = False Then
            If dictXRef.Count > 0 Then
                CBox_listReference.ItemsSource = dictXRef.Values.ToList()
                CBox_listReference.Text = dictXRef.ElementAt(0).Value
                If CBox_listReference.Text <> "" Then
                    'заносим данные в корневой элемент дерева
                    generalNode = New TreeViewItem
                    generalNode.Header = CBox_listReference.Text
                    Dim dictTag As Dictionary(Of String, String) = New Dictionary(Of String, String)
                    dictTag.Add("PutchDrawing", dictXRef.ElementAt(0).Key)
                    generalNode.Tag = dictTag

                    Dim node0 As New TreeViewItem()
                    node0.Header = "Площадные объекты"
                    Dim dictParentTag1 As Dictionary(Of String, String) = New Dictionary(Of String, String)
                    dictParentTag1.Add("PutchDrawing", dictTag.Item("PutchDrawing"))
                    dictParentTag1.Add("TypeObject", "Polygon")
                    node0.Tag = dictParentTag1
                    generalNode.Items.Add(node0)

                    Dim node1 As New TreeViewItem()
                    node1.Header = "Линейные объекты"
                    Dim dictParentTag2 As Dictionary(Of String, String) = New Dictionary(Of String, String)
                    dictParentTag2.Add("PutchDrawing", dictTag.Item("PutchDrawing"))
                    dictParentTag2.Add("TypeObject", "Line")
                    node1.Tag = dictParentTag2
                    generalNode.Items.Add(node1)

                    Dim node2 As New TreeViewItem()
                    node2.Header = "Точечные объекты"
                    Dim dictParentTag3 As Dictionary(Of String, String) = New Dictionary(Of String, String)
                    dictParentTag3.Add("PutchDrawing", dictTag.Item("PutchDrawing"))
                    dictParentTag3.Add("TypeObject", "Point")
                    node2.Tag = dictParentTag3
                    generalNode.Items.Add(node2)

                    TreeViewProp.Items.Add(generalNode)
                End If
            End If
        End If
    End Sub

    ' Обработчик добавления объектов в выбор
    Private Sub OnSelectionAdded(sender As Object, e As SelectionAddedEventArgs)
        Dim selectionSet As SelectionSet = e.AddedObjects
        Dim objId As ObjectId = Nothing
        Dim boolShowData As Boolean = False
        For Each objId In e.AddedObjects
            If objId.IsNull = False Then
                Continue For
            End If
        Next
        If IsNothing(objId) = False Then
            If objId.IsNull = False Then
                Dim doc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim db As Database = doc.Database
                Using tr As Transaction = db.TransactionManager.StartTransaction()
                    _editEntity = tr.GetObject(objId, OpenMode.ForRead)
                    Dim civilListEntity As Dictionary(Of String, CadastralTools.StructureElement) = FuncXRecords.readXRecords(_editEntity)
                    If IsNothing(civilListEntity) = False Then
                        If civilListEntity.Count > 0 Then
                            For i As Integer = 0 To civilListEntity.Count - 1
                                Dim civilEntity As CadastralTools.StructureElement = civilListEntity.ElementAt(i).Value
                                If IsNothing(civilEntity) = False Then
                                    Dim customerData As ObservableCollection(Of CadastralTools.WPFDataGrid) = New ObservableCollection(Of CadastralTools.WPFDataGrid)
                                    Dim PS As String = civilEntity.TablePS
                                    Dim keyParam As String = civilEntity.KeyParameter
                                    Dim arrayTablePSFromXML As String(,) = {}
                                    Dim boolReadPS As Boolean = CadastralTools.TemplateXML.FuncReadTablePSToXML(templateXML, PS, arrayTablePSFromXML)
                                    WPFDataGrid = CadastralTools.WPFDataGrid.createDataForCustomerDataGrid(arrayTablePSFromXML, keyParam)
                                    If WPFDataGrid.Count > 0 Then
                                        DGidPropertiesX.SelectedItems.Clear()
                                        DGidPropertiesX.SelectedIndex = -1
                                        DGidPropertiesX.CurrentCell = Nothing
                                        DGidPropertiesX.ItemsSource = WPFDataGrid
                                        If DGidPropertiesX.Columns.Count > 0 Then
                                            DGidPropertiesX.Columns(0).Visibility = Visibility.Collapsed
                                        End If
                                        DGidPropertiesX.FrozenColumnCount = 1
                                        DGidPropertiesX.Items.Refresh()
                                        Dim newDataObject As ObjectData = New ObjectData(civilEntity, PS, StructureElement.tableXRecords.PROJECT_STRUCTURES, objId)
                                        DGidPropertiesX.Tag = newDataObject
                                        boolShowData = True
                                    End If
                                End If
                            Next i
                        End If
                    End If
                End Using
            End If
        End If
        If boolShowData = False Then
            Dim arrayTablePSFromXML As String(,) = {}
            WPFDataGrid = CadastralTools.WPFDataGrid.createDataForCustomerDataGrid(arrayTablePSFromXML, "")
            If WPFDataGrid.Count > 0 Then
                DGidPropertiesX.SelectedItems.Clear()
                DGidPropertiesX.SelectedIndex = -1
                DGidPropertiesX.CurrentCell = Nothing
                DGidPropertiesX.ItemsSource = WPFDataGrid
                If DGidPropertiesX.Columns.Count > 0 Then
                    DGidPropertiesX.Columns(0).Visibility = Visibility.Collapsed
                End If
                DGidPropertiesX.FrozenColumnCount = 1
                DGidPropertiesX.Items.Refresh()
                DGidPropertiesX.Tag = Nothing
            End If
        End If
    End Sub

    ' Обработчик удаления объектов из выбора
    Private Sub OnSelectionRemoved(sender As Object, e As SelectionRemovedEventArgs)
        For Each objId As ObjectId In e.RemovedObjects
            _editor.WriteMessage($"Удален объект: {objId.ToString()}")
        Next
    End Sub


    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'смена значений по правой кнопке мыши
    Private Sub DGidPropertiesX_MouseRightButtonDown(sender As Object, e As MouseButtonEventArgs) Handles DGidPropertiesX.MouseRightButtonDown
        Dim dataObject As ObjectData = DGidPropertiesX.Tag
        If IsNothing(dataObject) = True Then
            MsgBox("Не удалось определить тип объекта в структуре проекта! Выберите объект повторно и поменяйте ему свойства еще раз!!!")
            Exit Sub
        End If
        If IsNothing(dataObject.IDObject) = True Then
            MsgBox("Не удалось определить тип объекта в структуре проекта! Выберите объект повторно и поменяйте ему свойства еще раз!!!")
            Exit Sub
        End If
        Dim selectedItem As Object = DGidPropertiesX.SelectedItem
        If IsNothing(selectedItem) = False Then
            Dim newCustData As ObservableCollection(Of WPFDataGrid) = New ObservableCollection(Of WPFDataGrid)
            newCustData = DGidPropertiesX.ItemsSource
            If IsNothing(newCustData) = False Then
                If newCustData.Count > 0 Then
                    If TypeOf selectedItem Is CadastralTools.WPFDataGrid Then
                        Dim userCustomer As CadastralTools.WPFDataGrid = selectedItem
                        Dim nameField As String = userCustomer.nameField
                        If IsNothing(nameField) = False Then
                            If nameField.Trim.Length > 0 Then
                                If IO.File.Exists(templateDataXML) = True Then
                                    Dim arrayValue As String() = {}
                                    Dim boolReadValue As Boolean = CadastralTools.TemplateXML.readFieldDataToXML(templateDataXML, nameField, arrayValue)
                                    If arrayValue.Length > 0 Then
                                        Dim selectedValueForm As CadastralTools.FormSelectObject = New CadastralTools.FormSelectObject
                                        selectedValueForm.SetListBoxValue = arrayValue
                                        selectedValueForm.ShowDialog()
                                        If selectedValueForm.boolShow = True Then
                                            Dim newValue As String = selectedValueForm.listBoxValue
                                            If IsNothing(newValue) = False Then
                                                If newValue.Trim.Length > 0 Then
                                                    userCustomer.Значение = newValue
                                                    'перезаписываем данные в примитиве
                                                    Dim userSData As StructureElement = dataObject.UserData
                                                    Dim idObject As ObjectId = dataObject.IDObject
                                                    If IsNothing(idObject) = False Then
                                                        If idObject.IsNull = False Then
                                                            Dim doc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                                                            Dim db As Database = doc.Database
                                                            Dim tr As Transaction = doc.TransactionManager.StartTransaction()
                                                            Using tr
                                                                _editEntity = tr.GetObject(idObject, OpenMode.ForWrite)
                                                                Dim civilLandEnt As CadastralTools.StructureElement = recEntityData()
                                                                If IsNothing(civilLandEnt) = False Then
                                                                    'selectedItemTreeView.Tag = civilLandEnt
                                                                    'обновляем таблицу
                                                                    DGidPropertiesX.SelectedItems.Clear()
                                                                    DGidPropertiesX.SelectedIndex = -1
                                                                    DGidPropertiesX.CurrentCell = Nothing
                                                                    DGidPropertiesX.ItemsSource = newCustData
                                                                    'DGidProperties.HeadersVisibility = False
                                                                    If DGidPropertiesX.Columns.Count > 0 Then
                                                                        DGidPropertiesX.Columns(0).Visibility = Visibility.Collapsed
                                                                    End If
                                                                    DGidPropertiesX.FrozenColumnCount = 1
                                                                    DGidPropertiesX.Items.Refresh()
                                                                End If
                                                                tr.Commit()
                                                            End Using
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Else
            MsgBox("Сначала надо выбрать ячейку!!!")
        End If
    End Sub
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор чертежа внешней ссылки
    Private Sub CBox_listReference_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles CBox_listReference.SelectionChanged
        Dim indexSelect As Integer = CBox_listReference.SelectedIndex
        Dim nameSelect As String = CBox_listReference.SelectedValue
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        If indexSelect > -1 Then
            If indexSelect < dictXRef.Count Then
                Dim putchRefDrawing As String = dictXRef.ElementAt(indexSelect).Key
                If IO.File.Exists(putchRefDrawing) = True Then
                    If CBox_listReference.Text <> "" Then
                        'заносим данные в корневой элемент дерева
                        generalNode = New TreeViewItem
                        generalNode.Header = CBox_listReference.Text
                        Dim dictTag As Dictionary(Of String, String) = New Dictionary(Of String, String)
                        dictTag.Add("PutchDrawing", dictXRef.ElementAt(0).Key)
                        generalNode.Tag = dictTag

                        Dim node0 As New TreeViewItem()
                        node0.Header = "Площадные объекты"
                        Dim dictParentTag1 As Dictionary(Of String, String) = New Dictionary(Of String, String)
                        dictParentTag1.Add("PutchDrawing", dictTag.Item("PutchDrawing"))
                        dictParentTag1.Add("TypeObject", "Polygon")
                        node0.Tag = dictParentTag1
                        generalNode.Items.Add(node0)

                        Dim node1 As New TreeViewItem()
                        node1.Header = "Линейные объекты"
                        Dim dictParentTag2 As Dictionary(Of String, String) = New Dictionary(Of String, String)
                        dictParentTag2.Add("PutchDrawing", dictTag.Item("PutchDrawing"))
                        dictParentTag2.Add("TypeObject", "Line")
                        node1.Tag = dictParentTag2
                        generalNode.Items.Add(node1)

                        Dim node2 As New TreeViewItem()
                        node2.Header = "Точечные объекты"
                        Dim dictParentTag3 As Dictionary(Of String, String) = New Dictionary(Of String, String)
                        dictParentTag3.Add("PutchDrawing", dictTag.Item("PutchDrawing"))
                        dictParentTag3.Add("TypeObject", "Point")
                        node2.Tag = dictParentTag3
                        generalNode.Items.Add(node2)

                        TreeViewProp.Items.Add(generalNode)
                    End If

                    'Dim node0 As TreeViewItem = New TreeViewItem
                    'If doc.Name Like putchRefDrawing Then
                    '    Dim civilGWGProject As NC_LandProject = New NC_LandProject
                    '    'dictCadastralProject = civilGWGProject.GetObjectToCurrentDWG()
                    '    node0.Tag = doc.Name
                    'Else
                    '    'dictCadastralProject = FuncDWGReference.GetObjectToReferenceDWG(putchRefDrawing)
                    '    node0.Tag = putchRefDrawing
                    'End If
                    'If dictCadastralObject.Count > 0 Then
                    '    If TreeViewProp.Items.Count > 0 Then
                    '        For i As Integer = 0 To TreeViewProp.Items.Count - 1
                    '            Dim tempNode As TreeViewItem = TreeViewProp.Items.Item(i)
                    '            If tempNode.Header Like nameSelect Then
                    '                node0 = tempNode
                    '                Exit For
                    '            End If
                    '        Next i
                    '    End If
                    '    'readDictionaryObject(node0)
                    '    If IsNothing(node0.Header) = True Then
                    '        node0.Header = nameSelect
                    '        TreeViewProp.Items.Add(node0)
                    '    End If
                    'End If
                End If
            End If
        End If
    End Sub
    '========================================================================================================================================
    'выбор узла treeView
    Private Sub TreeViewProp_SelectedItemChanged(sender As Object, e As Windows.RoutedPropertyChangedEventArgs(Of Object)) Handles TreeViewProp.SelectedItemChanged
        Dim selectItem As TreeViewItem = e.NewValue
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        If IsNothing(selectItem) = False Then
            Dim tagItemSelect As Object = selectItem.Tag
            If IsNothing(tagItemSelect) = False Then
                If TypeOf tagItemSelect Is Dictionary(Of String, String) Then
                    Dim tagItem As Dictionary(Of String, String) = selectItem.Tag
                    Dim putchRefDrawing As String = tagItem.Item("PutchDrawing")
                    If tagItem.ContainsKey("DeskObject") = True Then
                        Dim deskObject As String = tagItem.Item("DeskObject")
                        Dim dictParcel As Dictionary(Of String, Dictionary(Of String, Topology)) = FuncDatabase.groupObjectByNumber(deskObject, dictProjectPlaningObject)
                        readDictionaryObject(selectItem, dictParcel)
                    ElseIf tagItem.ContainsKey("TypeObject") = True Then
                        Dim typeObject As String = tagItem.Item("TypeObject")
                        If typeObject Like "Polygon" Then
                            'читаем все полигональные объекты
                            If doc.Name Like putchRefDrawing Then
                                dictProjectPlaningObject = FuncDatabase.getCadastralPolygon()
                            Else
                                dictProjectPlaningObject = FuncDWGReference.GetCadastralPolygonToReferenceDWG(putchRefDrawing)
                            End If
                            'заполняем дерево
                            If dictProjectPlaningObject.Count > 0 Then
                                For i As Integer = 0 To dictProjectPlaningObject.Count - 1
                                    Dim deskObject As String = dictProjectPlaningObject.ElementAt(i).Key
                                    Dim node0 As New TreeViewItem()
                                    node0.Header = deskObject
                                    Dim dictParentTag1 As Dictionary(Of String, String) = New Dictionary(Of String, String)
                                    dictParentTag1.Add("PutchDrawing", putchRefDrawing)
                                    dictParentTag1.Add("TypeObject", "Polygon")
                                    dictParentTag1.Add("DeskObject", deskObject)
                                    node0.Tag = dictParentTag1
                                    selectItem.Items.Add(node0)
                                Next i
                            End If
                        ElseIf typeObject Like "Line" Then
                            'читаем все полигональные объекты
                            If doc.Name Like putchRefDrawing Then
                                dictProjectPlaningObject = FuncDatabase.getCadastralLine()
                            Else
                                dictProjectPlaningObject = FuncDWGReference.GetCadastralPolygonToReferenceDWG(putchRefDrawing)
                            End If
                            'заполняем дерево
                            If dictProjectPlaningObject.Count > 0 Then
                                For i As Integer = 0 To dictProjectPlaningObject.Count - 1
                                    Dim deskObject As String = dictProjectPlaningObject.ElementAt(i).Key
                                    Dim node0 As New TreeViewItem()
                                    node0.Header = deskObject
                                    Dim dictParentTag1 As Dictionary(Of String, String) = New Dictionary(Of String, String)
                                    dictParentTag1.Add("PutchDrawing", putchRefDrawing)
                                    dictParentTag1.Add("TypeObject", "Line")
                                    dictParentTag1.Add("DeskObject", deskObject)
                                    node0.Tag = dictParentTag1
                                    selectItem.Items.Add(node0)
                                Next i
                            End If
                        End If
                    End If
                ElseIf TypeOf tagItemSelect Is StructureElement Then
                    Dim db As Database = doc.Database
                    Dim ed As Editor = doc.Editor
                    Dim userData As ObjectData = DGidPropertiesX.Tag
                    'If IsNothing(userData) = False Then
                    '    Dim idOld As ObjectId = userData.IDObject
                    '    If IsNothing(idOld) = False Then
                    '        If idOld.IsNull = False Then
                    '            Dim tr As Transaction = doc.TransactionManager.StartTransaction()
                    '            Using tr
                    '                Dim ent As Entity = idOld.GetObject(OpenMode.ForWrite)
                    '                If IsDBNull(ent) = False Then
                    '                    ent.Highlight()
                    '                End If
                    '            End Using
                    '        End If
                    '    End If
                    'End If
                    Dim userSData As StructureElement = tagItemSelect
                    If IsNothing(userSData.IdObject) = False Then
                        If userSData.ToString.Length > 0 Then
                            Dim id As ObjectId = FuncDatabase.getObjectID(userSData.IdObject)
                            Using tr As Transaction = db.TransactionManager.StartTransaction()
                                Dim boolCD As Boolean = returnCustomerData(userSData, id)
                                If boolCD = True Then
                                    Dim acEnt As Entity = tr.GetObject(id, OpenMode.ForRead)
                                    acEnt.Highlight()
                                    listObjectHidlight.Add(acEnt)
                                    ed.UpdateScreen()
                                End If
                                tr.Commit()
                            End Using
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'обновить дерево
    Private Sub ButtonRefresh_Click(sender As Object, e As Windows.RoutedEventArgs) Handles ButtonRefresh.Click
        Dim nameDwg As String = CBox_listReference.SelectedValue
        Dim indexDWG As Integer = CBox_listReference.SelectedIndex
        If indexDWG < dictXRef.Count Then
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim putchRefDrawing As String = dictXRef.ElementAt(indexDWG).Key

            If doc.Name Like putchRefDrawing Then
                dictProjectPlaningObject = FuncDatabase.getCadastralPolygon()
            Else
                dictProjectPlaningObject = FuncDWGReference.GetCadastralPolygonToReferenceDWG(putchRefDrawing)
            End If
            Dim node0 As TreeViewItem = New TreeViewItem
            If TreeViewProp.Items.Count > 0 Then
                For i As Integer = 0 To TreeViewProp.Items.Count - 1
                    Dim tempNode As TreeViewItem = TreeViewProp.Items.Item(i)
                    If tempNode.Header Like nameDwg Then
                        node0 = tempNode
                        Exit For
                    End If
                Next i
            End If

            If IsNothing(node0.Header) = True Then
                node0.Header = nameDwg
                TreeViewProp.Items.Add(node0)
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ввод значения
    Private Sub DGidPropertiesX_CellEditEnding(sender As Object, e As DataGridCellEditEndingEventArgs) Handles DGidPropertiesX.CellEditEnding
        If IsNothing(DGidPropertiesX.Tag) = True Then Exit Sub
        If TypeOf DGidPropertiesX.Tag Is ObjectData Then
            Dim userCivilCadastralData As ObjectData = DGidPropertiesX.Tag
            Dim editedColumn As DataGridColumn = e.Column
            Dim element As FrameworkElement = TryCast(e.EditingElement, FrameworkElement)
            ' Определяем действие (Commit - подтверждение, Cancel - отмена)
            If e.EditAction = DataGridEditAction.Commit Then
                ' Данные подтверждены
                Dim newValue As String = ""
                ' Получаем значение в зависимости от типа колонки
                If TypeOf editedColumn Is DataGridTextColumn Then
                    If TypeOf element Is TextBox Then
                        newValue = DirectCast(element, TextBox).Text
                    End If
                ElseIf TypeOf editedColumn Is DataGridCheckBoxColumn Then
                    If TypeOf element Is CheckBox Then
                        newValue = DirectCast(element, CheckBox).IsChecked.ToString()
                    End If
                ElseIf TypeOf editedColumn Is DataGridComboBoxColumn Then
                    If TypeOf element Is ComboBox Then
                        newValue = DirectCast(element, ComboBox).SelectedItem?.ToString()
                    End If
                End If
                Dim selectedItem As WPFDataGrid = DGidPropertiesX.SelectedItem
                selectedItem.Значение = newValue
                Dim nameField As String = selectedItem.nameField
                Dim userSData As StructureElement = userCivilCadastralData.UserData
                Dim keyParam As String = userSData.KeyParameter
                If CadastralTools.JsonValidator.IsValidJson(keyParam) Then
                    keyParam = CadastralTools.JsonValidator.SetJsonFieldValue(keyParam, nameField, newValue)
                End If
                userSData.KeyParameter = keyParam
                Dim doc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                If IsNothing(doc) = False Then
                    Dim db As Database = doc.Database
                    Dim tr As Transaction = doc.TransactionManager.StartTransaction()
                    Using tr
                        Dim id As ObjectId = userCivilCadastralData.IDObject
                        If IsNothing(id) = False Then
                            If id.IsNull = False Then
                                _editEntity = tr.GetObject(id, OpenMode.ForWrite)
                                Dim boolWriteData As Boolean = FuncXRecords.writeXRecords(_editEntity, userSData)
                            End If
                        End If
                        tr.Commit()
                    End Using
                End If
            End If
        ElseIf e.EditAction = DataGridEditAction.Cancel Then
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'изменение размера палитры
    Private Sub UserXRecords_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
        Try
            If IsNothing(userPal) = False Then
                Dim s1 As Double = userPal.Size.Height 'ширина формы (палитры)
                Dim s2 As Double = userPal.Size.Width
                If s1 > 800 Then
                    GridObject.Height = s1 - 20
                    TreeViewProp.Height = s1 - 400
                    'DGidProperties.Height = s1 - 20
                End If
                If s2 > 400 Then
                    GridObject.Width = s2 - 20
                    TreeViewProp.Width = s2 - 20
                    DGidPropertiesX.Width = s2 - 20
                End If
            End If
        Catch ex As System.Exception
        End Try
    End Sub

    Private Sub Unhighlight_Click(sender As Object, e As RoutedEventArgs) Handles Unhighlight.Click
        Dim unHidl As Boolean = FuncSelectedObject.UnhighlightEntity(listObjectHidlight)
    End Sub
End Class
