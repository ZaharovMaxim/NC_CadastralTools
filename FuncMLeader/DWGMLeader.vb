Imports System.Collections.ObjectModel
Imports System.IO
Imports CadastralTools
Imports HostMgd.ApplicationServices
Imports NetTopologySuite.GeometriesGraph
Imports Teigha.DatabaseServices
Imports Teigha.Geometry
Imports Teigha.GraphicsInterface

Public Class DWGMLeader
    Private _layer As String
    Private _typeLine As String
    Private _color As Integer
    Private _scaleTypeLine As Double
    Private _weightLine As Double
    Private _arrowSize As Double
    Private _textStyle As String
    Private _textHeight As Double
    Private _objectIDStyle As ObjectId
    Public Sub New()
        _layer = "0"
        _typeLine = "Continuous"
        _color = 7
        _scaleTypeLine = 1
        _weightLine = 20
        _arrowSize = 0
        _textStyle = ""
        _textHeight = 0
        _objectIDStyle = Nothing
    End Sub

    Public Property Layer As String
        Get
            Return _layer
        End Get
        Set(value As String)
            _layer = value
        End Set
    End Property

    Public Property TypeLine As String
        Get
            Return _typeLine
        End Get
        Set(value As String)
            _typeLine = value
        End Set
    End Property

    Public Property Color As Integer
        Get
            Return _color
        End Get
        Set(value As Integer)
            _color = value
        End Set
    End Property

    Public Property ScaleTypeLine As Double
        Get
            Return _scaleTypeLine
        End Get
        Set(value As Double)
            _scaleTypeLine = value
        End Set
    End Property

    Public Property WeightLine As Double
        Get
            Return _weightLine
        End Get
        Set(value As Double)
            _weightLine = value
        End Set
    End Property

    Public Property ArrowSize As Double
        Get
            Return _arrowSize
        End Get
        Set(value As Double)
            _arrowSize = value
        End Set
    End Property

    Public Property TextHeight As Double
        Get
            Return _textHeight
        End Get
        Set(value As Double)
            _textHeight = value
        End Set
    End Property

    Public Property TextStyle As String
        Get
            Return _textStyle
        End Get
        Set(value As String)
            _textStyle = value
        End Set
    End Property

    Public Property ObjectIDStyle As ObjectId
        Get
            Return _objectIDStyle
        End Get
        Set(value As ObjectId)
            _objectIDStyle = value
        End Set
    End Property

    '===========================================================================================================
    'функция копирует стиль мультивыносок из другого чертежа
    Public Shared Function CopyMLeaderStyleDestDraw(ByVal putchOutDWG As String, ByVal nameMLeaderStyle As String, Optional ByVal boolReplace As Boolean = False) As Boolean
        Dim sourceDb As Database = Nothing
        Try
            ' 1. Открываем исходный DWG
            sourceDb = New Database(False, True)
            sourceDb.ReadDwgFile(putchOutDWG, FileShare.Read, True, String.Empty)
            ' 2. Ищем стиль мультивыноски в исходном чертеже
            Dim sourceStyleId As ObjectId = ObjectId.Null
            Using sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
                Dim mleaderDict As DBDictionary = sourceTr.GetObject(sourceDb.MLeaderStyleDictionaryId, OpenMode.ForRead)
                ' Перебираем записи в словаре (более эффективно, чем Like)
                For Each entry As DBDictionaryEntry In mleaderDict
                    ' Сравниваем точно (без учета регистра)
                    If String.Equals(entry.Key, nameMLeaderStyle, StringComparison.OrdinalIgnoreCase) Then
                        sourceStyleId = entry.Value
                        Exit For
                    End If
                Next
                sourceTr.Commit()
            End Using
            ' 3. Если стиль не найден - выходим
            If sourceStyleId.IsNull Then
                Return False
            End If
            ' 4. Копируем в текущий чертеж
            Dim destDb As Database = HostApplicationServices.WorkingDatabase
            Dim destDictId As ObjectId = destDb.MLeaderStyleDictionaryId
            ' Создаем коллекцию с ID копируемого объекта
            Dim sourceIds As New ObjectIdCollection()
            sourceIds.Add(sourceStyleId)
            ' Настройки копирования
            Dim mapping As New IdMapping()
            Dim cloneOption As DuplicateRecordCloning = If(boolReplace, DuplicateRecordCloning.Replace, DuplicateRecordCloning.Ignore)
            ' Выполняем клонирование
            destDb.WblockCloneObjects(sourceIds, destDictId, mapping, cloneOption, False)
            Return True
        Catch ex As Exception
            Console.WriteLine($"Ошибка при копировании стиля мультивыноски '{nameMLeaderStyle}': {ex.Message}")
            Return False
        Finally
            ' 5. Освобождаем ресурсы
            If sourceDb IsNot Nothing Then
                sourceDb.Dispose()
            End If
        End Try
    End Function

    Public Function createMLeaderStyle(ByRef db As Database, ByRef arrayPropertyStyle As String(,), ByVal templateFile As String) As Boolean
        If IsNothing(db) = True Then Return False
        If IsNothing(arrayPropertyStyle) = True Then Return False
        If arrayPropertyStyle.Length = 0 Then Return False
        If IsNothing(templateFile) = True Then Return False
        If IO.File.Exists(templateFile) = False Then Return False
        Using acTrans As Transaction = db.TransactionManager.StartTransaction()
            Try
                For i As Integer = 0 To arrayPropertyStyle.GetUpperBound(1)
                    If arrayPropertyStyle(0, i) Like "Layer" Then
                        'создаем новый слой
                        Dim layerObject As String = arrayPropertyStyle(1, i)
                        If IsNothing(layerObject) = True Then Continue For
                        If layerObject.Trim.Length = 0 Then Continue For
                        If Left(layerObject, 1) Like "_" Then
                            Layer = layerObject
                            Continue For
                        End If
                        Dim acLyrTbl As LayerTable = acTrans.GetObject(db.LayerTableId, OpenMode.ForRead)
                        Dim acLineTypTbl As LinetypeTable = acTrans.GetObject(db.LinetypeTableId, OpenMode.ForRead)
                        'такого слоя нет
                        If acLyrTbl.Has(layerObject) = False Then
                            'ищем настройки слоя
                            Dim userPropertiesLayer As String(,) = {}
                            Dim boolFindData As Boolean = CadastralTools.TemplateXML.FuncReadLayerPropertiesInfoToXML(templateFile, layerObject, userPropertiesLayer)
                            If IsArray(userPropertiesLayer) = True Then
                                'назначаем тип линии для слоя
                                Dim nameTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesLayer, "LineType", 0, 1)
                                If IsNothing(nameTypeLineLayer) = True Then
                                    nameTypeLineLayer = "Continuous"
                                    Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                ElseIf nameTypeLineLayer.Trim.Length = 0 Then
                                    nameTypeLineLayer = "Continuous"
                                    Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                Else
                                    'загружаем тип линии
                                    If acLineTypTbl.Has(nameTypeLineLayer) = False Then
                                        Dim userPropertiesTypeLine As String(,) = {}
                                        boolFindData = CadastralTools.TemplateXML.FuncReadLineTypeInfoToXML(templateFile, nameTypeLineLayer, userPropertiesTypeLine)
                                        If IsArray(userPropertiesTypeLine) = True Then
                                            Dim deskTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesLayer, "DeskLineType", 0, 1)
                                            Dim patternTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesLayer, "CodeLineType", 0, 1)
                                            If IsNothing(deskTypeLineLayer) = True Then
                                                deskTypeLineLayer = nameTypeLineLayer
                                            ElseIf deskTypeLineLayer.Trim.Length = 0 Then
                                                deskTypeLineLayer = nameTypeLineLayer
                                            ElseIf IsNothing(patternTypeLineLayer) = True Then
                                                Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                            ElseIf patternTypeLineLayer.Trim.Length = 0 Then
                                                Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                            Else
                                                Dim userLayerLineType As LinetypeTableRecord = Nothing 'CadastralLeaderStyle.FuncCreateLineType(nameTypeLineLayer, deskTypeLineLayer, patternTypeLineLayer, templateFile)
                                                If IsNothing(userLayerLineType) = True Then
                                                    Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                                ElseIf userLayerLineType.Id.IsNull = True Then
                                                    Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            Dim createNewLayer As LayerTableRecord = Nothing 'CadastralStyle.FuncCreateLayer(userPropertiesLayer)
                            If IsNothing(createNewLayer) = False Then
                                Layer = createNewLayer.Name
                            End If
                        Else
                            Layer = layerObject
                        End If
                    ElseIf arrayPropertyStyle(0, i) Like "Name" Then
                        Dim nameMLeaderStyle As String = arrayPropertyStyle(1, i)
                        Dim mleaderDict As DBDictionary = acTrans.GetObject(db.MLeaderStyleDictionaryId, OpenMode.ForRead)
                        Dim boolFindMleader As Boolean = False
                        For Each entry As DBDictionaryEntry In mleaderDict
                            ' Сравниваем точно (без учета регистра)
                            If String.Equals(entry.Key, nameMLeaderStyle, StringComparison.OrdinalIgnoreCase) Then
                                boolFindMleader = True
                                ObjectIDStyle = entry.Value
                                Exit For
                            End If
                        Next
                        If boolFindMleader = False Then
                            Dim templateDWG As String = templateFile
                            templateDWG = templateDWG.Replace(".xml", ".dwg")
                            Dim blcCopyStyle As Boolean = CopyMLeaderStyleDestDraw(templateDWG, nameMLeaderStyle)
                            mleaderDict = acTrans.GetObject(db.MLeaderStyleDictionaryId, OpenMode.ForRead)
                            For Each entry As DBDictionaryEntry In mleaderDict
                                ' Сравниваем точно (без учета регистра)
                                If String.Equals(entry.Key, nameMLeaderStyle, StringComparison.OrdinalIgnoreCase) Then
                                    ObjectIDStyle = entry.Value
                                    Dim styleMleader As MLeaderStyle = acTrans.GetObject(ObjectIDStyle, OpenMode.ForWrite)
                                    styleMleader.Annotative = AnnotativeStates.False
                                    Exit For
                                End If
                            Next
                        End If
                    ElseIf arrayPropertyStyle(0, i) Like "LineType" Then
                        Dim nameTypeLine As String = arrayPropertyStyle(1, i)
                        Dim acLineTypTbl As LinetypeTable = acTrans.GetObject(db.LinetypeTableId, OpenMode.ForRead)
                        If acLineTypTbl.Has(nameTypeLine) = False Then
                            Dim userPropertiesTypeLine As String(,) = {}
                            Dim boolFindData As Boolean = CadastralTools.TemplateXML.FuncReadLineTypeInfoToXML(templateFile, nameTypeLine, userPropertiesTypeLine)
                            If IsArray(userPropertiesTypeLine) = True Then
                                Dim deskTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesTypeLine, "DeskLineType", 0, 1)
                                Dim patternTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesTypeLine, "CodeLineType", 0, 1)
                                If IsNothing(deskTypeLineLayer) = True Then
                                    deskTypeLineLayer = nameTypeLine
                                ElseIf deskTypeLineLayer.Trim.Length = 0 Then
                                    deskTypeLineLayer = nameTypeLine
                                ElseIf IsNothing(patternTypeLineLayer) = True Then
                                    Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesTypeLine, "DeskLineType", "Continuous", 0, 1)
                                ElseIf patternTypeLineLayer.Trim.Length = 0 Then
                                    Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesTypeLine, "DeskLineType", "Continuous", 0, 1)
                                Else
                                    Dim userLayerLineType As LinetypeTableRecord = Nothing 'CadastralStyle.FuncCreateLineType(nameTypeLine, deskTypeLineLayer, patternTypeLineLayer, templateFile)
                                    If IsNothing(userLayerLineType) = True Then
                                        nameTypeLine = "Continuous"
                                    ElseIf userLayerLineType.Id.IsNull = True Then
                                        nameTypeLine = "Continuous"
                                    End If
                                End If
                            Else
                                nameTypeLine = "Continuous"
                            End If
                        End If
                        _typeLine = nameTypeLine
                    ElseIf arrayPropertyStyle(0, i) Like "Color" Then
                        _color = Val(arrayPropertyStyle(1, i))
                        If _color < 0 Then _color = 0
                        If _color > 256 Then _color = 256
                    ElseIf arrayPropertyStyle(0, i) Like "LineWidth" Then
                        _weightLine = Val(arrayPropertyStyle(1, i))
                    ElseIf arrayPropertyStyle(0, i) Like "LineScale" Then
                        _scaleTypeLine = Val(arrayPropertyStyle(1, i))
                    ElseIf arrayPropertyStyle(0, i) Like "ArrowSize" Then
                        Dim ArrowSize As Double = Val(arrayPropertyStyle(1, i))
                        _arrowSize = ArrowSize
                    ElseIf arrayPropertyStyle(0, i) Like "TextStyle" Then
                        Dim nameTextStyle As String = arrayPropertyStyle(1, i)
                        _textStyle = nameTextStyle
                    ElseIf arrayPropertyStyle(0, i) Like "TextHeight" Then
                        _textHeight = Val(arrayPropertyStyle(1, i))
                    End If
                Next i
            Catch ex As Teigha.Runtime.Exception
                Return False
            End Try
            acTrans.Commit()
        End Using
        Return True
    End Function

    Function setStyle(ByRef db As Database, ByRef lineObject As Teigha.DatabaseServices.MLeader, Optional layerParent As String = "") As Boolean
        If IsNothing(db) = True Then Return False
        If IsNothing(lineObject) = False Then
            Using acTrans As Transaction = db.TransactionManager.StartTransaction()
                If lineObject.IsWriteEnabled = True Then
                    Try
                        If IsNothing(ObjectIDStyle) = False Then
                            If ObjectIDStyle.IsNull = False Then
                                lineObject.MLeaderStyle = ObjectIDStyle
                            End If
                        End If
                        'назначаем цвет
                        If Color > -1 And Color < 257 Then
                            lineObject.ColorIndex = Color
                        Else
                            lineObject.ColorIndex = 256
                        End If
                        'назначаем слой
                        Dim acLyrTbl As LayerTable = acTrans.GetObject(db.LayerTableId, OpenMode.ForRead)
                        If Left(Layer, 1) Like "_" Then 'блок если слой зависит от родительского объекта
                            If layerParent.Trim.Length > 0 Then
                                Layer = layerParent & Layer
                                Dim idLayerMLeafer As ObjectId = CreateDWGObject.createLayer(Layer)
                            End If
                        End If
                        If acLyrTbl.Has(Layer) = True Then
                            lineObject.Layer = Layer
                        Else
                            lineObject.Layer = "0"
                        End If
                        'тип линии
                        Dim acLineTypTbl As LinetypeTable = acTrans.GetObject(db.LinetypeTableId, OpenMode.ForRead)
                        If acLineTypTbl.Has(TypeLine) = True Then
                            lineObject.Linetype = TypeLine
                        Else
                            lineObject.Linetype = "Continuous"
                        End If
                        'масштаб линии
                        If ScaleTypeLine > 0 Then
                            lineObject.LinetypeScale = ScaleTypeLine
                        End If
                        'толщина линии
                        Try
                            If IsNumeric(WeightLine) = True Then
                                Dim weightValue As Integer = Convert.ToInt32(WeightLine)
                                If [Enum].IsDefined(GetType(LineWeight), weightValue) Then
                                    lineObject.LineWeight = CType(weightValue, LineWeight)
                                Else
                                    lineObject.LineWeight = CType(weightValue, LineWeight)
                                End If
                            ElseIf WeightLine Like "ПоСлою" Then
                                lineObject.LineWeight = LineWeight.ByLayer
                            ElseIf WeightLine Like "ПоБлоку" Then
                                lineObject.LineWeight = LineWeight.ByBlock
                            End If
                        Catch ex As Exception
                            lineObject.LineWeight = LineWeight.LineWeight020 ' 0.20мм
                        End Try
                        'размер стрелки
                        If ArrowSize > 0 Then
                            lineObject.ArrowSize = ArrowSize
                        End If
                        'размер текста
                        If TextHeight > 0 Then
                            Dim acMtext As MText = lineObject.MText
                            If acMtext.Height <> TextHeight Then
                                acMtext.Height = TextHeight
                                lineObject.MText = acMtext
                            End If
                        End If

                        Dim acTextStyleTbl As TextStyleTable = acTrans.GetObject(db.TextStyleTableId, OpenMode.ForRead)
                        For Each id As ObjectId In acTextStyleTbl
                            Dim acTextStyleRec As TextStyleTableRecord = acTrans.GetObject(id, OpenMode.ForRead)
                            Dim NameStyle As String = acTextStyleRec.Name.Trim
                            If TextStyle Like NameStyle Then
                                lineObject.TextStyleId = id
                                Exit For
                            End If
                        Next
                    Catch ex As Teigha.Runtime.Exception
                    End Try
                End If
                acTrans.Commit()
            End Using
        End If
        Return True
    End Function

    '==================================================================================================
    'функция вставляет мультивыноску в чертеж (точка вставки, текст, смещениеX, смещениеY,
    Public Shared Function createMleader(ByVal pointIns As Point3d, ByVal MTextIns As String, ByVal OffsetX As Double, ByVal offsetY As Double) As MLeader
        createMleader = Nothing
        Dim acDoc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Try
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim mlStyles As DBDictionary = acTrans.GetObject(acCurDb.MLeaderStyleDictionaryId, OpenMode.ForRead)
                Dim MLeader1 As MLeader = New MLeader
                MLeader1.SetDatabaseDefaults()
                Dim leaderIndex As Integer = MLeader1.AddLeader()
                Dim leaderLineIndex As Integer = MLeader1.AddLeaderLine(leaderIndex)
                MLeader1.AddFirstVertex(leaderLineIndex, New Point3d(pointIns.X, pointIns.Y, 0))
                MLeader1.AddLastVertex(leaderLineIndex, New Point3d(pointIns.X + OffsetX, pointIns.Y + offsetY, 0))
                MLeader1.ContentType = ContentType.MTextContent
                'добавляем многострочный текст
                Dim acMText As MText = New MText()
                acMText.SetDatabaseDefaults()
                acMText.Contents = MTextIns
                MLeader1.MText = acMText
                acBlkTblRec.AppendEntity(MLeader1)
                acTrans.AddNewlyCreatedDBObject(MLeader1, True)
                createMleader = MLeader1
                acTrans.Commit()
            End Using
        Catch ex As Teigha.Runtime.Exception
        End Try
    End Function
    Public Function createDataForWPFDataGrid(ByVal note As String, ByVal hText As Double) As ObservableCollection(Of WPFDataGrid)
        Dim custdata As ObservableCollection(Of WPFDataGrid) = New ObservableCollection(Of WPFDataGrid)
        If IsNothing(note) = False Then
            Dim strArray As String() = note.Split(";")
            If IsArray(strArray) = True Then
                For i As Integer = 0 To strArray.Length - 1
                    Dim valStr As String = strArray(i)
                    Dim strArray2 As String() = valStr.Split("=")
                    Dim userData As WPFDataGrid = New WPFDataGrid
                    If strArray2.Length = 1 Then
                        userData.nameField = "text"
                        userData.Имя_свойства = "Имя свойства"
                        userData.Значение = strArray2(0)
                    ElseIf strArray2.Length > 1 Then
                        userData.nameField = "text"
                        userData.Имя_свойства = strArray2(0)
                        userData.Значение = strArray2(1)
                    End If
                    custdata.Add(userData)
                Next i
            End If
        End If
        Dim userData1 As WPFDataGrid = New WPFDataGrid
        userData1 = New WPFDataGrid
        userData1.nameField = "TextHeight"
        userData1.Имя_свойства = "Высота текста"
        userData1.Значение = hText
        custdata.Add(userData1)
        Return custdata
    End Function
End Class
