Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Security
Imports System.Security.Cryptography
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Xml
Imports CadastralTools.StructureElement
Imports HostMgd.ApplicationServices
Imports HostMgd.EditorInput
Imports HostMgd.Windows
Imports NetTopologySuite.Geometries
Imports NetTopologySuite.LinearReferencing
Imports NetTopologySuite.Utilities
Imports Teigha.Colors
Imports Teigha.DatabaseServices
Imports Teigha.Geometry
Imports Teigha.GraphicsInterface
Imports Polyline = Teigha.DatabaseServices.Polyline

Public Class TopoSymbolsWPF
    Public userPal As HostMgd.Windows.PaletteSet = Nothing
    Public userPalette As Palette = Nothing
    Public templateXML As String = ""
    Public templateDWG As String = ""
    Public templateDataXML As String = ""
    Private dataRec As ObservableCollection(Of CadastralTools.WPFDataGrid) = New ObservableCollection(Of CadastralTools.WPFDataGrid)
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'первый запуск палитры
    Public Sub New(ByVal fileXML As String)
        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        templateXML = fileXML
        If IO.File.Exists(templateXML) = True Then
            Dim arrayCategory As String() = {}
            Dim boolFindCategory As Boolean = CadastralTools.TemplateXML.FuncReadCategoryToXML(templateXML, arrayCategory)
            If arrayCategory.Length > 0 Then
                For i As Integer = 0 To arrayCategory.Length - 1
                    Dim nameCategory As String = arrayCategory(i)
                    Dim node0 As TreeViewItem = New TreeViewItem
                    node0.Header = nameCategory
                    node0.Name = "Category"
                    node0.FontSize = 16
                    node0.Tag = {templateXML}
                    TreeViewObject.Items.Add(node0)
                Next i
            End If
        End If
        ' Добавьте эту строку
        AddHandler Me.SizeChanged, AddressOf TopoSymbolsWPF_SizeChanged
        ' Или через лямбду (если используете VB.NET с поддержкой)
        'AddHandler SizeChanged, Sub(s, e) Debug.WriteLine("Size changed!")
        '=====================================================================================================================
        Dim listSelect As List(Of String) = New List(Of String) From {"Замена", "Вставка"}
        CBox_Select.ItemsSource = listSelect
    End Sub
    Private Sub TopoSymbolsWPF_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized

    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'выбор узла дерева
    Private Sub TreeViewObject_SelectedItemChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object)) Handles TreeViewObject.SelectedItemChanged
        Dim TreeViewControl As System.Windows.Controls.TreeView = e.Source
        TreeViewControl.Focus()
        Dim itemSelectObject As Object = TreeViewControl.SelectedValue
        Dim doc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database 'Database
        Dim ed As Editor = doc.Editor 'Editor
        If TypeOf itemSelectObject Is DockPanel Then

        Else
            Dim ItemSelect As TreeViewItem = itemSelectObject 'ссылка на выбранный элемент
            If IsNothing(ItemSelect) = False Then
                Dim typeItem As String = ItemSelect.Name 'тип выбранной ветки
                If typeItem Like "Category" Then
                    'выбрана ветка с категорией
                    If ItemSelect.Items.Count = 0 Then
                        'загружаем таблицы
                        Dim arrayTables As String() = {}
                        Dim nameCategory As String = ItemSelect.Header
                        Dim boolFindTables As Boolean = CadastralTools.TemplateXML.FuncReadTablesToXML(templateXML, nameCategory, arrayTables)
                        If arrayTables.Length > 0 Then
                            For i As Integer = 0 To arrayTables.Length - 1
                                Dim nameTable As String = arrayTables(i)
                                Dim node0 As TreeViewItem = New TreeViewItem
                                node0.Header = nameTable
                                node0.Name = "Table"
                                node0.FontSize = 15
                                node0.Tag = {templateXML, nameCategory}
                                ItemSelect.Items.Add(node0)
                            Next i
                        End If
                        If ItemSelect.Items.Count > 0 Then
                            ItemSelect.ExpandSubtree()
                        End If
                    End If
                ElseIf typeItem Like "Table" Then
                    'выбрана ветка с категорией
                    If ItemSelect.Items.Count = 0 Then
                        'загружаем таблицы
                        Dim arrayTypes As String() = {}
                        Dim tagNode As String() = ItemSelect.Tag
                        Dim nameCategory As String = tagNode(1)
                        Dim nameTable As String = ItemSelect.Header
                        Dim boolFindTypes As Boolean = CadastralTools.TemplateXML.FuncReadTypesObjectToXML(templateXML, nameCategory, nameTable, arrayTypes)
                        If arrayTypes.Length > 0 Then
                            For i As Integer = 0 To arrayTypes.Length - 1
                                Dim nameType As String = arrayTypes(i)
                                Dim node0 As TreeViewItem = New TreeViewItem
                                node0.Header = nameType
                                node0.Name = "Type"
                                node0.FontSize = 14
                                node0.Tag = {templateXML, nameCategory, nameTable}
                                ItemSelect.Items.Add(node0)
                            Next i
                        End If
                        If ItemSelect.Items.Count > 0 Then
                            ItemSelect.ExpandSubtree()
                        End If
                    End If
                ElseIf typeItem Like "Type" Then
                    'выбрана ветка с категорией
                    If ItemSelect.Items.Count = 0 Then
                        'загружаем таблицы
                        Dim arrayObject As String() = {}
                        Dim tagNode As String() = ItemSelect.Tag
                        Dim nameCategory As String = tagNode(1)
                        Dim nameTable As String = tagNode(2)
                        Dim nameType As String = ItemSelect.Header
                        Dim boolFindTypes As Boolean = CadastralTools.TemplateXML.FuncReadObjectToXML(templateXML, nameCategory, nameTable, nameType, arrayObject)
                        If arrayObject.Length > 0 Then
                            For i As Integer = 0 To arrayObject.Length - 1
                                Dim nameObject As String = arrayObject(i)
                                Dim node0 As TreeViewItem = New TreeViewItem
                                node0.Header = nameObject
                                node0.Name = "Object"
                                node0.FontSize = 12
                                node0.Tag = {templateXML, nameCategory, nameTable, nameType}
                                ItemSelect.Items.Add(node0)
                            Next i
                        End If
                        If ItemSelect.Items.Count > 0 Then
                            ItemSelect.ExpandSubtree()
                        End If
                    End If
                ElseIf typeItem Like "Object" Then
                    'выбрана ветка с объектом - чистим список
                    DGidProperties.SelectedItems.Clear()
                    If ItemSelect.Items.Count = 0 Then
                        'загружаем таблицы
                        Dim arrayProperties As String(,) = {}
                        Dim tagNode As String() = ItemSelect.Tag
                        Dim nameCategory As String = tagNode(1)
                        Dim nameTable As String = tagNode(2)
                        Dim nameType As String = tagNode(3)
                        Dim nameObject As String = ItemSelect.Header
                        'читаем свойства объекта
                        Dim boolFindPropertyObject As Boolean = CadastralTools.TemplateXML.FuncReadPropertiesObjectToXML(templateXML, nameCategory, nameTable, nameType, nameObject, arrayProperties)
                        'читаем таблицу PS
                        Dim PS As String = CadastralTools.MathFunction.findValueToArray2d(arrayProperties, "PS", 0, 1)
                        'читаем поля выбранной таблицы
                        Dim arrayFieldDataPs As String(,) = {}
                        Dim boolReadPSTable As Boolean = CadastralTools.TemplateXML.FuncReadTablePSToXML(templateXML, PS, arrayFieldDataPs)
                        If arrayFieldDataPs.Length > 0 Then
                            'заполняем датагрид
                            Dim customerData As ObservableCollection(Of CadastralTools.WPFDataGrid) = CadastralTools.WPFDataGrid.createDataForCustomerDataGrid(arrayFieldDataPs, "")
                            If IsNothing(customerData) = False Then
                                If customerData.Count > 0 Then
                                    Dim dictPropObject As Dictionary(Of String, Object) = New Dictionary(Of String, Object)
                                    dictPropObject.Add("LABEL", nameCategory)
                                    dictPropObject.Add("NAME", nameTable)
                                    dictPropObject.Add("TYPE", nameType)
                                    dictPropObject.Add("KeyParameter", arrayProperties)
                                    If nameType Like "Полилиния" Or nameType Like "Линия" Or nameType Like "Полигон" Then
                                        Dim styleObject As CadastralPolylineStyle = New CadastralPolylineStyle
                                        styleObject.TemplateXML = templateXML
                                        styleObject.DBDrawing = db
                                        Dim boolCreateStyle As Boolean = styleObject.createLineStyle(arrayProperties)
                                        If boolCreateStyle = True Then
                                            dictPropObject.Add("Style", styleObject)
                                        End If
                                    ElseIf nameType Like "Блок" Then
                                        'Dim styleObject As CadastralBlockStyle = New CadastralBlockStyle()
                                        'Dim boolCreateStyle As Boolean = styleObject.createBlockStyle(db, arrayProperties, templateXML)
                                        'If boolCreateStyle = True Then
                                        'dictPropObject.Add("Style", styleObject)
                                        'End If
                                        Dim nameBlk As String = CadastralTools.MathFunction.findValueToArray2d(arrayProperties, "Name", 0, 1)
                                        Dim dataAttrib As List(Of CadastralTools.AttributeData) = FuncBlockReference.GetBlockAttributesWithDesc(nameBlk)
                                        If dataAttrib.Count > 0 Then
                                            Dim newParcel As New CadastralTools.CadastralPoint
                                            customerData = newParcel.createUserDataForCustomerDataGrid(dataAttrib)
                                        End If
                                        dictPropObject.Add("NameBlock", nameBlk)
                                    ElseIf nameType Like "Мультивыноска" Then
                                        Dim styleObject As DWGMLeader = New DWGMLeader()
                                        Dim boolCreateStyle As Boolean = styleObject.createMLeaderStyle(db, arrayProperties, templateXML)
                                        If boolCreateStyle = True Then
                                            dictPropObject.Add("Style", styleObject)
                                        End If
                                        Dim nameStyleMleader As String = CadastralTools.MathFunction.findValueToArray2d(arrayProperties, "Name", 0, 1)
                                        dictPropObject.Add("NameMLeader", nameStyleMleader)
                                        Dim note As String = CadastralTools.MathFunction.findValueToArray2d(arrayProperties, "Note", 0, 1)
                                        Dim htext As String = CadastralTools.MathFunction.findValueToArray2d(arrayProperties, "TextHeight", 0, 1)
                                        customerData = styleObject.createDataForWPFDataGrid(note, htext)
                                    End If
                                    If arrayProperties.Length > 0 Then
                                        If customerData.Count > 0 Then
                                            ButtonDraw.Tag = dictPropObject
                                            DGidProperties.ItemsSource = customerData
                                            DGidProperties.HeadersVisibility = False
                                            DGidProperties.Tag = PS
                                            If DGidProperties.Columns.Count > 0 Then
                                                DGidProperties.Columns(0).Visibility = Visibility.Collapsed
                                            End If
                                            DGidProperties.FrozenColumnCount = 1
                                        End If
                                    End If
                                    If ItemSelect.Items.Count > 0 Then
                                        ItemSelect.ExpandSubtree()
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'создать объект
    Private Sub ButtonDraw_Click(sender As Object, e As RoutedEventArgs) Handles ButtonDraw.Click
        Dim doc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database 'Database
        Dim ed As Editor = doc.Editor 'Editor
        Dim scaleCurrentDwg As Double = FuncPropertiesDrawing.FuncReadCurrentAnnotativeScaleDWG()
        Dim selectUserItem As TreeViewItem = TreeViewObject.SelectedValue
        Dim boolReplase As Boolean = True
        Dim typeSelect As String = CBox_Select.Text
        If typeSelect Like "Замена" Then
            boolReplase = True
        Else
            boolReplase = False
        End If
        If TypeOf ButtonDraw.Tag Is Dictionary(Of String, Object) Then
            Dim dictParamObject As Dictionary(Of String, Object) = ButtonDraw.Tag
            Dim label As String = dictParamObject.Item("LABEL")
            Dim name As String = dictParamObject.Item("NAME")
            Dim typeObject As String = dictParamObject.Item("TYPE")
            Dim PS As String = DGidProperties.Tag
            Dim arrayProperties As String(,) = dictParamObject.Item("KeyParameter")
            Dim nameObject As String = PS
            Dim deskObject As String = name
            Dim note As String = ""
            If IsArray(arrayProperties) = True Then
                Dim tempDeskObject As String = CadastralTools.MathFunction.findValueToArray2d(arrayProperties, "Description", 0, 1)
                If tempDeskObject.Trim.Length > 0 Then
                    deskObject = tempDeskObject
                End If
                Dim tempnameObject As String = CadastralTools.MathFunction.findValueToArray2d(arrayProperties, "Object_code", 0, 1)
                If tempnameObject.Trim.Length > 0 Then
                    nameObject = tempnameObject
                End If
                note = CadastralTools.MathFunction.findValueToArray2d(arrayProperties, "Note", 0, 1)
            End If
            Dim GlobalID As String = ""
            Dim countSelectObject As Integer = 0
            Dim numberObject As Integer = 1
            Dim numberCounter As Integer = 0
            Dim arrayDataRec As String(,) = {}
            Dim countArray As Integer = 0
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'начальные установки для обозначения участка
            If IsNothing(DGidProperties.ItemsSource) = False Then
                If DGidProperties.Items.Count > 0 Then
                    dataRec = DGidProperties.ItemsSource
                    If dataRec.Count > 0 Then
                        For i As Integer = 0 To dataRec.Count - 1
                            Dim custData As CadastralTools.WPFDataGrid = dataRec.ElementAt(i)
                            Dim tagField As String = custData.nameField
                            Dim value As String = custData.Значение
                            If IsNothing(tagField) = False Then
                                'заполняем массив с характеристиками
                                If tagField Like "Number" Or tagField Like "NUMBER" Then
                                    If IsNumeric(value) = True Then
                                        If value < 1 Then value = 1
                                        numberObject = Val(value)
                                    End If
                                ElseIf tagField Like "NumberCounter" Then 'участок многоконтурный (добавляем номер к каждому выбранному контуру)
                                    If IsNumeric(value) = True Then
                                        If Val(value) > 0 Then
                                            numberCounter = Val(value)
                                        End If
                                    End If
                                ElseIf custData.nameField Like "ConditionalNumber" Then
                                    'извлекаем номер контура
                                    Dim tempNumberCounter As Integer = CadastralTools.DataFormatting.GetNumberInParenthesis(custData.Значение)
                                    'извлекаем номер участка
                                    Dim tempNumber As Integer = CadastralTools.DataFormatting.GetNumberBeforeParenthesis(custData.Значение)
                                    'номер участка должен быть больше 0
                                    If tempNumber = 0 Then
                                        numberObject = 1
                                    Else
                                        numberObject = tempNumber
                                    End If
                                    'если номер контура больше 0 то считаем объект многоконтурным
                                    If tempNumberCounter > 0 Then
                                        numberCounter = tempNumberCounter
                                    End If
                                End If
                                ReDim Preserve arrayDataRec(2, countArray)
                                arrayDataRec(0, countArray) = custData.nameField
                                arrayDataRec(1, countArray) = custData.Значение
                                arrayDataRec(2, countArray) = custData.Имя_свойства
                                countArray += 1
                            End If
                        Next i
                    End If
                End If
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'делаем выбор объекта
            Using acLckDoc As DocumentLock = doc.LockDocument()
                Using acTrans As Transaction = db.TransactionManager.StartTransaction()
                    Dim acBlkTbl As BlockTable = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead)
                    Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    '=================================================================================================================================
                    If typeObject Like "Полилиния" Or typeObject Like "Линия" Or typeObject Like "Полигон" Then
                        Dim acSSPrompt As PromptSelectionResult = ed.GetSelection()
                        Dim acSSet1 As SelectionSet
                        If acSSPrompt.Status = PromptStatus.OK Then
                            acSSet1 = acSSPrompt.Value
                            For Each IdSelectObject As ObjectId In acSSet1.GetObjectIds()
                                If IdSelectObject.IsNull() = False Then
                                    Dim entity As Teigha.DatabaseServices.Entity = acTrans.GetObject(IdSelectObject, OpenMode.ForWrite)
                                    If TypeOf entity Is Line Or TypeOf entity Is Teigha.DatabaseServices.Polyline Or TypeOf entity Is Polyline2d Or TypeOf entity Is Spline Or TypeOf entity Is Arc Or TypeOf entity Is Circle Then
                                        Dim selectObject As Entity = entity
                                        If boolReplase = False Then
                                            selectObject = entity.Clone()
                                            acBlkTblRec.AppendEntity(selectObject)
                                            acTrans.AddNewlyCreatedDBObject(selectObject, True)
                                        End If
                                        If dictParamObject.ContainsKey("Style") = True Then
                                            Dim styleLine As CadastralPolylineStyle = dictParamObject.Item("Style")
                                            styleLine.setStyle(selectObject)
                                        End If
                                        For i As Integer = 0 To arrayDataRec.GetUpperBound(1)
                                            If arrayDataRec(0, i) Like "Number" Or arrayDataRec(0, i) Like "NUMBER" Then
                                                arrayDataRec(1, i) = numberObject
                                            ElseIf arrayDataRec(0, i) Like "NumberCounter" Then
                                                arrayDataRec(1, i) = numberCounter
                                            ElseIf arrayDataRec(0, i) Like "ConditionalNumber" Then
                                                If numberCounter > 0 Then
                                                    arrayDataRec(1, i) = numberObject & "(" & numberCounter & ")"
                                                Else
                                                    arrayDataRec(1, i) = numberObject
                                                End If
                                            End If
                                        Next
                                        Dim structData As CadastralTools.StructureElement = New CadastralTools.StructureElement
                                        structData.Label = label
                                        If typeObject Like "Полигон" Then
                                            If TypeOf entity Is Polyline Then
                                                Dim userPolyline As Polyline = entity
                                                If userPolyline.NumberOfVertices < 3 Then
                                                    Continue For
                                                End If
                                                If userPolyline.Closed = False Then
                                                    userPolyline.Closed = True
                                                End If
                                                structData.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polygon
                                                For i As Integer = 0 To arrayDataRec.GetUpperBound(1)
                                                    If arrayDataRec(0, i) Like "Area" Then
                                                        arrayDataRec(1, i) = Math.Round(userPolyline.Area, 2)
                                                    ElseIf arrayDataRec(0, i) Like "Lenght" Then
                                                        arrayDataRec(1, i) = Math.Round(userPolyline.Length, 2)
                                                    End If
                                                Next i
                                            Else
                                                Continue For
                                            End If
                                        Else
                                            structData.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.line
                                        End If
                                        'зелаем запись в объект
                                        Dim gsonStr As String = CadastralTools.JsonValidator.convertArray2DToJson(arrayDataRec)
                                        structData.NameTypeObject = CadastralTools.StructureElement.GetEnumNameByDescription(GetType(nameObject), deskObject)
                                        structData.Description = deskObject
                                        structData.KeyParameter = gsonStr
                                        structData.TablePS = PS
                                        structData.IdElement = Guid.NewGuid.ToString
                                        structData.IdStructure = ""
                                        structData.Note = note
                                        Dim boolWriteData As Boolean = FuncXRecords.writeXRecords(selectObject, structData)
                                        If numberCounter > 0 Then
                                            numberCounter += 1
                                        Else
                                            numberObject += 1
                                        End If
                                    End If
                                End If
                            Next
                            doc.Editor.Regen()
                        End If
                    ElseIf typeObject Like "Блок" Then
                        '===========================================================================================================================
                        Dim nameBlock As String = dictParamObject.Item("NameBlock")
                        If acBlkTbl.Has(nameBlock) = True Then
                            Dim writeCustimerData As ObservableCollection(Of CadastralTools.WPFDataGrid) = DGidProperties.ItemsSource
                            'формируем массив с характеристиками атрибута
                            Dim listAttrib As List(Of CadastralTools.AttributeData) = CadastralTools.CadastralPoint.setAtributeProperties(writeCustimerData, arrayProperties)
                            'вставка нового блока
                            If boolReplase = False Then
                                Do
                                    Dim acSSPrompt As PromptPointResult = ed.GetPoint("^C^C " & vbLf & "Укажите точку вставки блока: ")
                                    If Not (acSSPrompt.Status = PromptStatus.OK) Then
                                        Exit Do
                                    End If
                                    Dim insBlk As BlockReference = CreateDWGObject.FuncInsertDrawingBlk(acSSPrompt.Value, nameBlock)
                                    If dictParamObject.ContainsKey("Style") Then
                                        'Dim styleBlk As CadastralBlockStyle = dictParamObject.Item("Style")
                                        'styleBlk.setStyle(insBlk)
                                    End If

                                    If listAttrib.Count > 0 Then
                                        For i As Integer = 0 To listAttrib.Count - 1
                                            Dim userA As CadastralTools.AttributeData = listAttrib.Item(i)
                                            If userA.Tag Like "NUMBER" Then
                                                userA.Value = numberObject
                                                Exit For
                                            End If
                                        Next i
                                    End If
                                    Dim boolWditeAttrib As Boolean = FuncBlockReference.writeListAttributesBlock(insBlk, listAttrib)
                                    'конвертируем список атрибутов в массив для записи его в PS
                                    Dim arrayAttributeBlk As String(,) = FuncBlockReference.convertListAttributesBlockToArray(listAttrib)
                                    Dim strGSONBlk As String = CadastralTools.JsonValidator.convertArray2DToJson(arrayAttributeBlk)
                                    Dim structData As CadastralTools.StructureElement = New CadastralTools.StructureElement
                                    structData.Label = label
                                    structData.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.block
                                    structData.NameClassObject = CadastralTools.StructureElement.GetEnumNameByDescription(GetType(nameObject), deskObject)
                                    structData.Description = deskObject
                                    structData.KeyParameter = strGSONBlk
                                    structData.TablePS = PS
                                    structData.IdElement = Guid.NewGuid.ToString
                                    structData.IdStructure = ""
                                    structData.Note = note
                                    Dim boolWriteData As Boolean = FuncXRecords.writeXRecords(insBlk, structData)
                                    doc.Editor.Regen()
                                    numberObject += 1
                                Loop
                            Else
                                'замена существующего блока
                                Do
                                    Dim acSSPrompt As PromptEntityResult = ed.GetEntity("^C^C " & vbLf & "Укажите блок который требуется заменить: ")
                                    If Not (acSSPrompt.Status = PromptStatus.OK) Then
                                        Exit Do
                                    End If
                                    If acSSPrompt.Status = PromptStatus.OK Then
                                        Dim insertPoint As Point3d = New Point3d(-1, -1, -1)
                                        Dim idSelectObject As ObjectId = acSSPrompt.ObjectId
                                        Dim selectEnt As Entity = acTrans.GetObject(idSelectObject, OpenMode.ForWrite)
                                        If TypeOf selectEnt Is BlockReference Then
                                            Dim oldBlk As BlockReference = selectEnt
                                            insertPoint = oldBlk.Position
                                            oldBlk.Erase()
                                        End If
                                        If insertPoint.X <> -1 And insertPoint.Y <> -1 Then
                                            Dim insBlk As BlockReference = CreateDWGObject.FuncInsertDrawingBlk(insertPoint, nameBlock)
                                            If dictParamObject.ContainsKey("Style") Then
                                                'Dim styleBlk As CadastralBlockStyle = dictParamObject.Item("Style")
                                                'styleBlk.setStyle(insBlk)
                                            End If

                                            If listAttrib.Count > 0 Then
                                                For i As Integer = 0 To listAttrib.Count - 1
                                                    Dim userA As CadastralTools.AttributeData = listAttrib.Item(i)
                                                    If userA.Tag Like "NUMBER" Then
                                                        userA.Value = numberObject
                                                        Exit For
                                                    End If
                                                Next i
                                            End If
                                            Dim boolWditeAttrib As Boolean = FuncBlockReference.writeListAttributesBlock(insBlk, listAttrib)
                                            'конвертируем список атрибутов в массив для записи его в PS
                                            Dim arrayAttributeBlk As String(,) = FuncBlockReference.convertListAttributesBlockToArray(listAttrib)
                                            Dim strGSONBlk As String = CadastralTools.JsonValidator.convertArray2DToJson(arrayAttributeBlk)
                                            Dim structData As CadastralTools.StructureElement = New CadastralTools.StructureElement
                                            structData.Label = label
                                            structData.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.block
                                            structData.NameTypeObject = CadastralTools.StructureElement.GetEnumNameByDescription(GetType(nameObject), deskObject)
                                            structData.Description = deskObject
                                            structData.KeyParameter = strGSONBlk
                                            structData.TablePS = PS
                                            structData.IdElement = Guid.NewGuid.ToString
                                            structData.IdStructure = ""
                                            structData.Note = note
                                            Dim boolWriteData As Boolean = FuncXRecords.writeXRecords(insBlk, structData)
                                            doc.Editor.Regen()
                                            numberObject += 1
                                        End If
                                    End If
                                Loop
                            End If

                        Else
                            MsgBox("Не удалось скопировать определение блока из файла: " & templateDWG)
                        End If
                    ElseIf typeObject Like "Мультивыноска" Then
                        If IsArray(arrayProperties) = True Then
                            Dim hText As String = CadastralTools.MathFunction.findValueToArray2d(arrayProperties, "TextHeight")
                            If IsArray(arrayDataRec) = True Then
                                Dim tempNoteNumerator As String = ""
                                Dim tempNoteDenominator As String = ""
                                For i As Integer = 0 To arrayDataRec.GetUpperBound(1)
                                    Dim namefield As String = arrayDataRec(2, i)
                                    Dim val As String = arrayDataRec(1, i)
                                    If namefield.Trim Like "denominator" Then
                                        tempNoteDenominator = tempNoteDenominator & val
                                    ElseIf namefield.Trim Like "numerator" Then
                                        tempNoteNumerator = tempNoteNumerator & val
                                    ElseIf namefield.Trim Like "Высота текста" Then
                                        hText = val
                                    End If
                                Next i
                                If tempNoteDenominator.Trim.Length > 0 Then
                                    note = tempNoteNumerator & vbLf & tempNoteDenominator
                                Else
                                    note = tempNoteNumerator
                                End If
                                If Val(hText) > 0 Then
                                    Dim boolWrite As Boolean = CadastralTools.MathFunction.setValueToArray2d(arrayProperties, "TextHeight", hText)
                                End If
                            End If
                            If note.Trim.Length > 0 Then
                                If note.IndexOf("ConditionalNumber") > -1 Or note.IndexOf("Area") > -1 Or note.IndexOf("Number") > -1 Or note.IndexOf("NumberCounter") > -1 Then
                                    Dim acSSPrompt As PromptSelectionResult = ed.GetSelection()
                                    Dim acSSet1 As SelectionSet
                                    If acSSPrompt.Status = PromptStatus.OK Then
                                        acSSet1 = acSSPrompt.Value
                                        For Each IdSelectObject As ObjectId In acSSet1.GetObjectIds()
                                            Dim labelMleader As String = note
                                            If IdSelectObject.IsNull() = False Then
                                                Dim entity As Teigha.DatabaseServices.Entity = acTrans.GetObject(IdSelectObject, OpenMode.ForWrite)
                                                If TypeOf entity Is Line Or TypeOf entity Is Teigha.DatabaseServices.Polyline Or TypeOf entity Is Polyline2d Or TypeOf entity Is Circle Then
                                                    Dim dictStructData As Dictionary(Of String, CadastralTools.StructureElement) = FuncXRecords.readXRecords(entity)
                                                    If IsNothing(dictStructData) = False Then
                                                        If dictStructData.ContainsKey("PROJECT_STRUCTURES") = True Then
                                                            Dim structData As CadastralTools.StructureElement = dictStructData.Item("PROJECT_STRUCTURES")
                                                            If IsNothing(structData) = False Then
                                                                Dim tempKeyParam As String = structData.KeyParameter
                                                                If labelMleader.IndexOf("ConditionalNumber") > -1 Then
                                                                    Dim numberParcel As String = CadastralTools.JsonValidator.GetValueFromJson(tempKeyParam, "ConditionalNumber")
                                                                    labelMleader = labelMleader.Replace("<ConditionalNumber>", numberParcel)
                                                                End If
                                                                If labelMleader.IndexOf("Number") > -1 Then
                                                                    Dim numberParcel As String = CadastralTools.JsonValidator.GetValueFromJson(tempKeyParam, "Number")
                                                                    labelMleader = labelMleader.Replace("<Number>", numberParcel)
                                                                End If
                                                                If labelMleader.IndexOf("NumberCounter") > -1 Then
                                                                    Dim numberParcel As String = CadastralTools.JsonValidator.GetValueFromJson(tempKeyParam, "NumberCounter")
                                                                    labelMleader = labelMleader.Replace("<NumberCounter>", numberParcel)
                                                                End If
                                                                If labelMleader.IndexOf("Area") > -1 Then
                                                                    Dim numberParcel As String = CadastralTools.JsonValidator.GetValueFromJson(tempKeyParam, "Area")
                                                                    labelMleader = labelMleader.Replace("<Area>", numberParcel)
                                                                End If
                                                                If labelMleader.IndexOf("Lenght") > -1 Then
                                                                    Dim numberParcel As String = CadastralTools.JsonValidator.GetValueFromJson(tempKeyParam, "Lenght")
                                                                    labelMleader = labelMleader.Replace("<Lenght>", numberParcel)
                                                                End If
                                                                Dim firstPoint As NetTopologySuite.Geometries.Point = Nothing
                                                                Dim netCivilTopo As CadastralTools.FuncNetTopologySuite = New CadastralTools.FuncNetTopologySuite()
                                                                Dim netNCadTopo As FuncTopologySuite = New FuncTopologySuite()

                                                                If TypeOf entity Is Teigha.DatabaseServices.Polyline Then
                                                                    Dim poly As Teigha.DatabaseServices.Polyline = entity
                                                                    If poly.Closed = True Then
                                                                        Dim netPolygon As NetTopologySuite.Geometries.Polygon = netNCadTopo.CreatePolygon(poly)
                                                                        firstPoint = netCivilTopo.GetInteriorPoint(netPolygon)
                                                                    Else
                                                                        Dim netLine As NetTopologySuite.Geometries.LineString = netNCadTopo.CreateLineString(poly)
                                                                        Dim totalLength As Double = netLine.Length
                                                                        Dim midDistance As Double = totalLength / 2.0
                                                                        Dim lengthIndexedLine = New LengthIndexedLine(netLine)
                                                                        Dim midPointCoordinate As Coordinate = lengthIndexedLine.ExtractPoint(midDistance)
                                                                        firstPoint = New NetTopologySuite.Geometries.Point(midPointCoordinate)
                                                                    End If
                                                                End If
                                                                If IsNothing(firstPoint) = False Then
                                                                    If IsNothing(labelMleader) = False Then
                                                                        Dim newMleader As MLeader = DWGMLeader.createMleader(New Point3d(firstPoint.X, firstPoint.Y, 0), labelMleader, 10, 10)
                                                                        If dictParamObject.ContainsKey("Style") = True Then
                                                                            Dim styleLine As DWGMLeader = dictParamObject.Item("Style")
                                                                            styleLine.setStyle(db, newMleader, entity.Layer)
                                                                        End If
                                                                        Dim structDataMLeder As CadastralTools.StructureElement = New CadastralTools.StructureElement
                                                                        structDataMLeder.Label = label
                                                                        structData.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.leader
                                                                        structDataMLeder.NameTypeObject = CadastralTools.StructureElement.GetEnumNameByDescription(GetType(nameObject), deskObject)
                                                                        structDataMLeder.Description = deskObject
                                                                        structDataMLeder.KeyParameter = note
                                                                        structDataMLeder.TablePS = PS
                                                                        structDataMLeder.IdElement = Guid.NewGuid.ToString
                                                                        structDataMLeder.IdStructure = ""
                                                                        structDataMLeder.Note = note
                                                                        Dim boolRecData As Boolean = FuncXRecords.writeXRecords(newMleader, structDataMLeder)
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Next
                                    End If
                                End If
                            End If
                        End If
                    End If
                    acTrans.Commit()
                End Using
            End Using
            'заново заполняем таблицу данных
            If dataRec.Count > 0 Then
                For i As Integer = 0 To dataRec.Count - 1
                    Dim custData As CadastralTools.WPFDataGrid = dataRec.ElementAt(i)
                    Dim tagField As String = custData.nameField
                    Dim value As String = custData.Значение
                    If IsNothing(tagField) = False Then
                        If tagField Like "Number" Or tagField Like "NUMBER" Then
                            custData.Значение = numberObject
                        ElseIf tagField Like "NumberCounter" Then
                            custData.Значение = numberCounter
                        ElseIf arrayDataRec(0, i) Like "ConditionalNumber" Then
                            If numberCounter > 0 Then
                                custData.Значение = numberObject & "(" & numberCounter & ")"
                            Else
                                custData.Значение = numberObject
                            End If
                        End If
                    End If
                Next
                DGidProperties.SelectedItems.Clear()
                DGidProperties.SelectedIndex = -1
                DGidProperties.CurrentCell = Nothing
                DGidProperties.ItemsSource = dataRec
                If DGidProperties.Columns.Count > 0 Then
                    DGidProperties.Columns(0).Visibility = Visibility.Collapsed
                End If
                DGidProperties.FrozenColumnCount = 1
                DGidProperties.Items.Refresh()
            End If
        End If
    End Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'ввод значения в датагрид из заранее сформированного списка
    Private Sub TopoSymbolsWPF_MouseRightButtonDown(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseRightButtonDown
        Dim selectedItem As Object = DGidProperties.SelectedItem
        If IsNothing(selectedItem) = False Then
            Dim newCustData As ObservableCollection(Of CadastralTools.WPFDataGrid) = New ObservableCollection(Of CadastralTools.WPFDataGrid)
            newCustData = DGidProperties.ItemsSource
            If IsNothing(newCustData) = False Then
                If newCustData.Count > 0 Then
                    If TypeOf selectedItem Is CadastralTools.WPFDataGrid Then
                        Dim userCustomer As CadastralTools.WPFDataGrid = selectedItem
                        Dim nameField As String = userCustomer.nameField
                        If IsNothing(nameField) = False Then
                            If nameField.Trim.Length > 0 Then
                                If File.Exists(templateDataXML) = True Then
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
                                                    DGidProperties.SelectedItems.Clear()
                                                    DGidProperties.SelectedIndex = -1
                                                    DGidProperties.CurrentCell = Nothing
                                                    DGidProperties.ItemsSource = newCustData
                                                    'DGidProperties.HeadersVisibility = False
                                                    If DGidProperties.Columns.Count > 0 Then
                                                        DGidProperties.Columns(0).Visibility = Visibility.Collapsed
                                                    End If
                                                    DGidProperties.FrozenColumnCount = 1
                                                    DGidProperties.Items.Refresh()
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
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'изменить размер палитры
    Private Sub TopoSymbolsWPF_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Me.SizeChanged
        Try
            If IsNothing(userPal) = False Then
                Dim s1 As Double = userPal.Size.Height 'ширина формы (палитры)
                Dim s2 As Double = userPal.Size.Width
                If s1 > 800 Then
                    GridObject.Height = s1 - 10
                    TreeViewObject.Height = s1 - DGidProperties.Height - 20
                    Dim mx As Thickness = New Thickness(0, s1 - DGidProperties.Height + 20, 0, 0)
                    DGidProperties.Margin = mx
                Else
                    GridObject.Height = 790
                    TreeViewObject.Height = 440
                    DGidProperties.Height = 280
                End If
                If s2 > 400 Then
                    GridObject.Width = s2 - 20
                    TreeViewObject.Width = s2 - 20
                    DGidProperties.Width = s2 - 20
                Else
                    GridObject.Width = 390
                    TreeViewObject.Width = 380
                    DGidProperties.Width = 380
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub
End Class
