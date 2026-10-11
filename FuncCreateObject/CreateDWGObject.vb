Imports System.Globalization
Imports System.IO
Imports System.Text.RegularExpressions
Imports HostMgd.ApplicationServices
Imports HostMgd.EditorInput
Imports HostMgd.Windows
Imports Teigha
Imports Teigha.DatabaseServices
Imports Teigha.Geometry
Imports Teigha.GraphicsInterface
Imports Teigha.Runtime
Imports Polyline = Teigha.DatabaseServices.Polyline

Public Class CreateDWGObject
    Private _db As Database = Nothing
    Public Sub New(ByRef DBDrawing As Database)
        _db = DBDrawing
    End Sub

    Public Property DBDrawing As Database
        Get
            Return _db
        End Get
        Set(value As Database)
            _db = value
        End Set
    End Property

    'создание полилинии из полигона
    Public Function createPolylineByPolygon(ByVal polygon As NetTopologySuite.Geometries.Polygon, Optional ByRef polyline As Polyline = Nothing) As Polyline
        createPolylineByPolygon = Nothing
        If IsNothing(DBDrawing) = False Then
            If IsNothing(polygon) = False Then
                If polygon.Coordinates.Count > 2 Then
                    Using acTrans As Transaction = DBDrawing.TransactionManager.StartTransaction()
                        Dim acBlkTbl As BlockTable = acTrans.GetObject(DBDrawing.BlockTableId, OpenMode.ForRead)
                        Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                        If IsNothing(polyline) = True Then
                            polyline = New Polyline()
                            polyline.SetDatabaseDefaults()
                        Else
                            If polyline.ObjectId.IsValid AndAlso Not polyline.ObjectId.IsNull Then
                                polyline.Erase()
                                polyline = New Polyline()
                                polyline.SetDatabaseDefaults()
                            End If
                        End If
                        Dim listCoordPolugon As NetTopologySuite.Geometries.Coordinate() = polygon.Coordinates
                        If listCoordPolugon.Length > 2 Then
                            For i As Integer = 0 To listCoordPolugon.Length - 2
                                Dim coord As NetTopologySuite.Geometries.Coordinate = listCoordPolugon(i)
                                polyline.AddVertexAt(i, New Point2d(coord.X, coord.Y), 0, 0, 0)
                            Next i
                            Dim startVertex As NetTopologySuite.Geometries.Coordinate = listCoordPolugon(0)
                            Dim endVertex As NetTopologySuite.Geometries.Coordinate = listCoordPolugon(listCoordPolugon.Length - 1)
                            Dim distance As Double = startVertex.Distance(endVertex)
                            If distance > 0.01 Then
                                polyline.AddVertexAt(polyline.NumberOfVertices - 1, New Point2d(endVertex.X, endVertex.Y), 0, 0, 0)
                            End If
                            acBlkTblRec.AppendEntity(polyline)
                            acTrans.AddNewlyCreatedDBObject(polyline, True)
                            polyline.Closed = True
                        End If
                        acTrans.Commit()
                    End Using
                End If
            End If
            Return polyline
        End If
    End Function
    'создание полилинии из линии
    Public Function createPolylineByLineString(ByVal polygon As NetTopologySuite.Geometries.LineString, Optional ByRef polyline As Polyline = Nothing) As Polyline
        createPolylineByLineString = Nothing
        If IsNothing(DBDrawing) = False Then
            If IsNothing(polygon) = False Then
                If polygon.Coordinates.Count > 2 Then
                    Using acTrans As Transaction = DBDrawing.TransactionManager.StartTransaction()
                        Dim acBlkTbl As BlockTable = acTrans.GetObject(DBDrawing.BlockTableId, OpenMode.ForRead)
                        Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                        If IsNothing(polyline) = True Then
                            polyline = New Polyline()
                            polyline.SetDatabaseDefaults()
                        Else
                            While polyline.NumberOfVertices > 0
                                polyline.RemoveVertexAt(0) ' Удаляем первую вершину, пока не останется ни одной
                            End While
                        End If
                        Dim listCoordPolugon As NetTopologySuite.Geometries.Coordinate() = polygon.Coordinates
                        If listCoordPolugon.Length > 1 Then
                            For i As Integer = 0 To listCoordPolugon.Length - 1
                                Dim coord As NetTopologySuite.Geometries.Coordinate = listCoordPolugon(i)
                                polyline.AddVertexAt(0, New Point2d(coord.X, coord.Y), 0, 0, 0)
                            Next i
                            acBlkTblRec.AppendEntity(polyline)
                            acTrans.AddNewlyCreatedDBObject(polyline, True)
                            polyline.Closed = True
                        End If
                        acTrans.Commit()
                    End Using
                End If
            End If
            Return polyline
        End If
    End Function
    '==================================================================================================
    'функция вставляет мультивыноску в чертеж (точка вставки, текст, смещениеX, смещениеY,) новая версия
    Public Function createLeader(ByVal firstVertex As Point3d, ByVal contextText As String, Optional lastVertex As Point3d = Nothing, Optional ByVal idStyleMLeader As ObjectId = Nothing, Optional ByVal NameLayer As String = "0", Optional ByVal M As Double = 1, Optional ByVal StyleTextID As ObjectId = Nothing, Optional ByVal IndexColorText As Double = -1, Optional ByVal Background As Boolean = True) As MLeader
        createLeader = Nothing
        Dim myMLeaderId As ObjectId = ObjectId.Null
        If IsNothing(DBDrawing) = False Then
            Try
                Using acTrans As Transaction = DBDrawing.TransactionManager.StartTransaction()
                    Dim acBlkTbl As BlockTable = acTrans.GetObject(DBDrawing.BlockTableId, OpenMode.ForRead)
                    Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    Dim mlStyles As DBDictionary = acTrans.GetObject(DBDrawing.MLeaderStyleDictionaryId, OpenMode.ForRead)
                    Dim myleaderStyleId As ObjectId = DBDrawing.MLeaderstyle
                    Dim mlstyle As MLeaderStyle = TryCast(acTrans.GetObject(myleaderStyleId, OpenMode.ForRead), MLeaderStyle)
                    'создаем новый MLeader
                    Dim myMLeader As MLeader = New MLeader()
                    myMLeader.SetDatabaseDefaults()
                    myMLeaderId = myMLeader.ObjectId
                    'назначаем стиль
                    If IsNothing(idStyleMLeader) = False Then
                        If idStyleMLeader.IsNull = False Then
                            If mlStyles.Contains(idStyleMLeader) = True Then
                                myMLeader.MLeaderStyle = idStyleMLeader
                                mlstyle = TryCast(acTrans.GetObject(idStyleMLeader, OpenMode.ForRead), MLeaderStyle)
                            End If
                        End If
                    End If
                    'назначаем положение
                    Dim myMText As MText = TryCast(mlstyle.DefaultMText.Clone(), MText)
                    Dim leaderIndex As Integer = myMLeader.AddLeader()
                    Dim leaderLineIndex As Integer = myMLeader.AddLeaderLine(leaderIndex)
                    If lastVertex.X = 0 And lastVertex.Y = 0 And lastVertex.Z = 0 Then
                        myMLeader.AddFirstVertex(leaderLineIndex, firstVertex)
                        myMLeader.RemoveLeaderLine(leaderLineIndex)
                        myMText.Location = firstVertex
                    Else
                        myMLeader.AddFirstVertex(leaderLineIndex, firstVertex)
                        myMLeader.AddLastVertex(leaderLineIndex, lastVertex)
                    End If
                    'назначаем содержимое
                    If myMText IsNot Nothing Then
                        If IsNothing(StyleTextID) = False Then
                            If StyleTextID.IsNull = False Then
                                Dim acTextStyleTbl As TextStyleTable = acTrans.GetObject(DBDrawing.TextStyleTableId, OpenMode.ForRead)
                                If acTextStyleTbl.Has(StyleTextID) = True Then
                                    Dim acTextStyleRec As TextStyleTableRecord = acTrans.GetObject(StyleTextID, OpenMode.ForRead)
                                    myMText.TextStyleId = StyleTextID
                                End If
                            End If
                        End If
                        myMLeader.ContentType = ContentType.MTextContent
                        myMText.Contents = contextText
                        myMText.TextHeight = myMText.TextHeight * M
                        myMLeader.MText = myMText
                        'характеристики текста
                        If IndexColorText > -1 Then
                            myMText.ColorIndex = IndexColorText
                        End If
                        myMText.BackgroundFill = Background
                    End If
                    Dim acLyrTbl As LayerTable = acTrans.GetObject(DBDrawing.LayerTableId, OpenMode.ForRead)
                    If acLyrTbl.Has(NameLayer) = True Then
                        myMLeader.Layer = NameLayer
                    End If
                    myMLeader.ArrowSize = myMLeader.ArrowSize * M
                    createLeader = myMLeader
                    acBlkTblRec.AppendEntity(myMLeader)
                    acTrans.AddNewlyCreatedDBObject(myMLeader, True)
                    acTrans.Commit()
                End Using
            Catch ex As Teigha.Runtime.Exception
            End Try
        End If
    End Function
    '====================================================================================================
    'обновить полилинию из полигона
    Public Function UpdatePolylineFromPolygon(ByVal polygon As NetTopologySuite.Geometries.Polygon) As Boolean
        If IsNothing(DBDrawing) = True Then Return False
        If IsNothing(polygon) = True Then Return False
        Using Tr As Transaction = DBDrawing.TransactionManager.StartTransaction()
            'обновояем внешний контур
            Dim extendPolygon As NetTopologySuite.Geometries.LineString = polygon.ExteriorRing
            Dim exterialData As CadastralTools.StructureElement = TryCast(polygon.UserData, CadastralTools.StructureElement)
            If IsNothing(exterialData) = False Then
                Dim exteralhandle As String = exterialData.IdObject
                If IsNothing(exteralhandle) = False Then
                    Dim id As ObjectId = FuncDatabase.getObjectID(exteralhandle)
                    If IsNothing(id) = False Then
                        If Not (id.IsNull OrElse id.IsErased) Then
                            ' --- Открываем полилинию для записи ---
                            Dim plineCounter As Polyline = TryCast(Tr.GetObject(id, OpenMode.ForWrite), Polyline)
                            If IsNothing(plineCounter) = False Then
                                ' --- Обновляем вершины ---
                                Dim coords As NetTopologySuite.Geometries.Coordinate() = extendPolygon.Coordinates
                                If Not (coords Is Nothing OrElse coords.Length < 4) Then
                                    ' Реальных вершин — n - 1 (последняя точка = копия первой)
                                    Dim vertexCount As Integer = coords.Length - 1
                                    ' Удаляем все существующие вершины
                                    ' --- 1. Удаляем лишние вершины (если полилиния длиннее) ---
                                    ' Удаляем с конца, пока не останется ровно vertexCount вершин.
                                    While plineCounter.NumberOfVertices > vertexCount
                                        plineCounter.RemoveVertexAt(plineCounter.NumberOfVertices - 1)
                                    End While
                                    For j As Integer = 0 To vertexCount - 1
                                        If j <= plineCounter.NumberOfVertices - 1 Then
                                            plineCounter.SetPointAt(j, New Point2d(coords(j).X, coords(j).Y))
                                        Else
                                            plineCounter.AddVertexAt(j + 1, New Point2d(coords(j).X, coords(j).Y), 0.0, 0.0, 0.0)
                                        End If
                                    Next
                                    ' Замыкаем полилинию
                                    plineCounter.Closed = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            'островки
            For i As Integer = 0 To polygon.NumInteriorRings - 1
                Dim userPolygon As NetTopologySuite.Geometries.LineString = polygon.InteriorRings(i)
                Dim userData As CadastralTools.StructureElement = TryCast(userPolygon.UserData, CadastralTools.StructureElement)
                If userData Is Nothing Then Continue For
                Dim handle As String = userData.IdObject
                If IsNothing(handle) = False Then
                    Dim id As ObjectId = FuncDatabase.getObjectID(handle)
                    If IsNothing(id) = False Then
                        If id.IsNull OrElse id.IsErased Then Continue For
                        ' --- Открываем полилинию для записи ---
                        Dim plineCounter As Polyline = TryCast(Tr.GetObject(id, OpenMode.ForWrite), Polyline)
                        If IsNothing(plineCounter) = True Then Continue For
                        ' --- Обновляем вершины ---
                        Dim coords As NetTopologySuite.Geometries.Coordinate() = userPolygon.Coordinates
                        If coords Is Nothing OrElse coords.Length < 4 Then Continue For
                        ' Реальных вершин — n - 1 (последняя точка = копия первой)
                        Dim vertexCount As Integer = coords.Length - 1
                        ' --- 1. Удаляем лишние вершины (если полилиния длиннее) ---
                        ' Удаляем с конца, пока не останется ровно vertexCount вершин.
                        While plineCounter.NumberOfVertices > vertexCount
                            plineCounter.RemoveVertexAt(plineCounter.NumberOfVertices - 1)
                        End While
                        For j As Integer = 0 To vertexCount - 1
                            If j <= plineCounter.NumberOfVertices - 1 Then
                                plineCounter.SetPointAt(j, New Point2d(coords(j).X, coords(j).Y))
                            Else
                                plineCounter.AddVertexAt(j + 1, New Point2d(coords(j).X, coords(j).Y), 0.0, 0.0, 0.0)
                            End If
                        Next
                        ' Замыкаем полилинию
                        plineCounter.Closed = True
                    End If
                End If
            Next
            Tr.Commit()
        End Using
        Return True
    End Function
    '====================================================================================================
    'обновить полилинию из линии (перенос из SurvCivil2018)
    'Аналог UpdatePolylineFromPolygon: линия не замыкается принудительно — замкнутость
    'берётся из самой геометрии (LineString.IsClosed)
    Public Function UpdatePolylineFromLineString(ByVal line As NetTopologySuite.Geometries.LineString) As Boolean
        If IsNothing(DBDrawing) = True Then Return False
        If IsNothing(line) = True Then Return False
        Using Tr As Transaction = DBDrawing.TransactionManager.StartTransaction()
            Dim userData As CadastralTools.StructureElement = TryCast(line.UserData, CadastralTools.StructureElement)
            If IsNothing(userData) = False Then
                Dim handle As String = userData.IdObject
                If IsNothing(handle) = False Then
                    Dim id As ObjectId = FuncDatabase.getObjectID(handle)
                    If IsNothing(id) = False Then
                        If Not (id.IsNull OrElse id.IsErased) Then
                            ' --- Открываем полилинию для записи ---
                            Dim plineCounter As Polyline = TryCast(Tr.GetObject(id, OpenMode.ForWrite), Polyline)
                            If IsNothing(plineCounter) = False Then
                                ' --- Обновляем вершины ---
                                Dim coords As NetTopologySuite.Geometries.Coordinate() = line.Coordinates
                                UpdatePolylineVertices(plineCounter, coords, line.IsClosed)
                            End If
                        End If
                    End If
                End If
            End If
            Tr.Commit()
        End Using
        Return True
    End Function
    '====================================================================================================
    'переписывает вершины существующей полилинии по массиву координат
    'Если линия замкнута, последняя координата — копия первой, реальных вершин на одну меньше
    Private Shared Function UpdatePolylineVertices(ByVal plineCounter As Polyline, ByVal coords As NetTopologySuite.Geometries.Coordinate(), ByVal boolClosed As Boolean) As Boolean
        If IsNothing(plineCounter) OrElse coords Is Nothing Then Return False

        ' Линии достаточно двух точек (у кольца их минимум 4)
        Dim vertexCount As Integer = coords.Length
        If boolClosed = True Then vertexCount = coords.Length - 1
        If vertexCount < 2 Then Return False

        ' --- 1. Удаляем лишние вершины (если полилиния длиннее) ---
        ' Удаляем с конца, пока не останется ровно vertexCount вершин.
        While plineCounter.NumberOfVertices > vertexCount
            plineCounter.RemoveVertexAt(plineCounter.NumberOfVertices - 1)
        End While
        For j As Integer = 0 To vertexCount - 1
            If j <= plineCounter.NumberOfVertices - 1 Then
                plineCounter.SetPointAt(j, New Point2d(coords(j).X, coords(j).Y))
            Else
                plineCounter.AddVertexAt(j + 1, New Point2d(coords(j).X, coords(j).Y), 0.0, 0.0, 0.0)
            End If
        Next
        ' Замыкание берём из геометрии, а не выставляем всегда True
        plineCounter.Closed = boolClosed
        Return True
    End Function
    '====================================================================================================
    'вставка в Чертеж блока с атрибутами точки
    Public Shared Function CreateBlock(ByVal insertPosition As Point3d,
                                   ByVal styleBlock As CadastralBlockStyle,
                                   ByVal listAttributes As List(Of AttributeCadastralBlock)) As BlockReference

        If styleBlock Is Nothing Then Return Nothing

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        If doc Is Nothing Then Return Nothing
        Dim db As Database = doc.Database

        Using tr As Transaction = doc.TransactionManager.StartTransaction()
            Try
                Dim bt As BlockTable = DirectCast(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                If String.IsNullOrWhiteSpace(styleBlock.NameBlock) OrElse Not bt.Has(styleBlock.NameBlock) Then
                    tr.Abort()
                    Return Nothing
                End If

                Dim blockId As ObjectId = bt(styleBlock.NameBlock)
                Dim br As New BlockReference(insertPosition, blockId)

                br.Layer = styleBlock.NameLayer
                br.ScaleFactors = New Scale3d(styleBlock.ScaleBlock, styleBlock.ScaleBlock, styleBlock.ScaleBlock)
                br.ColorIndex = styleBlock.Color
                br.LinetypeScale = styleBlock.ScaleTypeLine
                br.Linetype = styleBlock.TypeLine
                br.LineWeight = styleBlock.WeightLine

                Dim curSpace As BlockTableRecord = DirectCast(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                curSpace.AppendEntity(br)
                tr.AddNewlyCreatedDBObject(br, True)

                ' --- Атрибуты ---
                Dim btrDef As BlockTableRecord = DirectCast(tr.GetObject(blockId, OpenMode.ForRead), BlockTableRecord)
                If btrDef.HasAttributeDefinitions Then
                    For Each id As ObjectId In btrDef
                        Dim ad As AttributeDefinition = TryCast(tr.GetObject(id, OpenMode.ForRead), AttributeDefinition)
                        If ad Is Nothing Then Continue For

                        ' Ищем пользовательское значение по тегу
                        Dim userAttrib As AttributeCadastralBlock = Nothing
                        If listAttributes IsNot Nothing Then
                            For Each ua In listAttributes
                                If String.Equals(ua.TagAttribute, ad.Tag, StringComparison.OrdinalIgnoreCase) Then
                                    userAttrib = ua
                                    Exit For
                                End If
                            Next
                        End If

                        ' Если для этого атрибута нет данных — не создаём ссылку
                        If userAttrib Is Nothing Then Continue For

                        Dim ar As New AttributeReference()
                        ar.SetAttributeFromBlock(ad, br.BlockTransform)
                        ar.Layer = "0"
                        ar.TextString = If(userAttrib.TextStringAttribute, String.Empty)
                        'ar.Visible = userAttrib.VisibleAttribute

                        ' Локальные координаты внутри блока -> WCS
                        Dim ptWcs As New Point3d(userAttrib.XAttribute, userAttrib.YAttribute, 0)

                        If userAttrib.VisibleAttribute Then
                            ar.Oblique = userAttrib.RotationAttribute
                            ar.Height = userAttrib.HeightAttribute
                            ar.ColorIndex = userAttrib.ColorAttribute

                            If userAttrib.HorizontalAlignment = 1 Then
                                ar.HorizontalMode = TextHorizontalMode.TextLeft
                            Else
                                ar.HorizontalMode = TextHorizontalMode.TextRight
                            End If

                            If userAttrib.VerticalAlignment = 1 Then
                                ar.VerticalMode = TextVerticalMode.TextBottom
                            Else
                                ar.VerticalMode = TextVerticalMode.TextTop
                            End If

                            ar.Position = ptWcs
                            ar.AlignmentPoint = ptWcs
                        Else
                            ar.Position = ptWcs
                        End If

                        br.AttributeCollection.AppendAttribute(ar)
                        tr.AddNewlyCreatedDBObject(ar, True)
                    Next
                End If

                tr.Commit()
                Return br

            Catch ex As System.Exception
                tr.Abort()
                Return Nothing
            End Try
        End Using
    End Function






























    'функция сождает новый слой
    Public Shared Function createLayer(ByVal LayerStr As String, Optional ByVal color1 As Short = 7, Optional ByVal idlinetype As ObjectId = Nothing, Optional ByVal wightline As Integer = 20) As ObjectId
        createLayer = Nothing
        Try
            Dim acDoc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acLyrTbl As LayerTable
                acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)
                Try
                    If acLyrTbl.Has(LayerStr) = False Then
                        Dim acLyrTblRec As LayerTableRecord = New LayerTableRecord()
                        acLyrTblRec.Name = LayerStr 'имя слоя
                        'цвет слоя
                        If IsNothing(color1) = False Then
                            Dim color2 As Teigha.Colors.Color = Teigha.Colors.Color.FromColorIndex(Teigha.Colors.ColorMethod.ByAci, color1)
                            acLyrTblRec.Color = color2
                        End If
                        'тип линии
                        If IsNothing(idlinetype) = False Then
                            acLyrTblRec.LinetypeObjectId = idlinetype
                        End If
                        'вес линии
                        If IsNothing(wightline) = False Then
                            acLyrTblRec.LineWeight = wightline
                        End If
                        acLyrTbl.UpgradeOpen()
                        acLyrTbl.Add(acLyrTblRec)
                        acTrans.AddNewlyCreatedDBObject(acLyrTblRec, True)
                        createLayer = acLyrTblRec.Id
                    Else
                        Return acLyrTbl.Item(LayerStr)
                    End If
                Catch
                    createLayer = Nothing
                End Try
                acTrans.Commit()
            End Using
        Catch
            createLayer = Nothing
        End Try
    End Function



    '===================================================================================
    'вставка в Чертеж блока (в пространство модели) со свойсвами определенными отдельными переменными
    Public Shared Function FuncInsertDrawingBlk(ByVal pointBlkXY As Point3d, ByVal NameBlock As String, Optional NameLayer As String = "0", Optional ColorBlk As Integer = 7, Optional ScaleBlkXY As Double = 1, Optional LineWight As Double = 20, Optional rot As Double = 0, Optional ScaleBlkZ As Double = 0) As BlockReference
        FuncInsertDrawingBlk = Nothing
        Dim doc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim tr As Transaction = doc.TransactionManager.StartTransaction()
        If IsNothing(NameBlock) = True Then Exit Function
        If NameBlock.Trim.Length = 0 Then Exit Function
        Using tr
            Try
                Dim bt As BlockTable = DirectCast(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                If bt.Has(NameBlock) = True Then
                    Dim space As BlockTableRecord = DirectCast(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                    Dim btr As BlockTableRecord = DirectCast(tr.GetObject(bt(NameBlock), OpenMode.ForRead), BlockTableRecord)
                    Dim br As New BlockReference(pointBlkXY, btr.ObjectId)
                    space.AppendEntity(br)
                    tr.AddNewlyCreatedDBObject(br, True)
                    'присваиваем характеристики блоку
                    'слой
                    Dim acLyrTbl As LayerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead)
                    If acLyrTbl.Has(NameLayer) = True Then
                        br.Layer = NameLayer
                    End If
                    'цвет
                    If ColorBlk < 0 Then
                        ColorBlk = 0
                    ElseIf ColorBlk > 256 Then
                        ColorBlk = 256
                    End If
                    br.ColorIndex = ColorBlk
                    'масштаб
                    If ScaleBlkXY > 0 Then
                        br.ScaleFactors = New Scale3d(ScaleBlkXY, ScaleBlkXY, ScaleBlkXY)
                    End If
                    If ScaleBlkZ > 0 Then
                        br.ScaleFactors = New Scale3d(ScaleBlkXY, ScaleBlkXY, ScaleBlkZ)
                    End If
                    'вес линии
                    If LineWight > 0 Then
                        br.LineWeight = LineWight
                    Else
                        If LineWight = -2 Then
                            br.LineWeight = LineWeight.ByBlock
                        Else
                            br.LineWeight = LineWeight.ByLayer
                        End If
                    End If
                    'поворот блока
                    br.Rotation = rot
                    If btr.HasAttributeDefinitions Then
                        Dim myAttColl As DatabaseServices.AttributeCollection = br.AttributeCollection
                        For Each entID As ObjectId In btr
                            Dim myEnt As DatabaseServices.Entity = entID.GetObject(OpenMode.ForWrite)
                            If TypeOf myEnt Is DatabaseServices.AttributeDefinition Then 'если это атрибут
                                'делаем на него ссылку
                                Dim myAttDef As DatabaseServices.AttributeDefinition = myEnt
                                'создаем новое определение атрибута
                                Dim myAttRef As New DatabaseServices.AttributeReference
                                'копируем значения в новый атрибут
                                myAttRef.SetAttributeFromBlock(myAttDef, br.BlockTransform)
                                myAttRef.Position = myAttDef.Position.TransformBy(br.BlockTransform)
                                myAttRef.TextString = myAttDef.TextString
                                myAttRef.Tag = myAttDef.Tag
                                myAttColl.AppendAttribute(myAttRef)
                                'проверяем атрибут на наличие поля
                                If myAttDef.HasFields = True Then
                                    Dim id As ObjectId = myAttDef.GetField() 'смотрим старый атрибут
                                    Dim fd As Field = tr.GetObject(id, OpenMode.ForRead) 'читаем его
                                    Dim CodStrFiled As String = fd.GetFieldCode 'получаем код поля
                                    If CodStrFiled.IndexOf("InsertionPoint \f") Then
                                        Dim h As String = FuncFormatData.FuncFormatFloatNumber(br.Position.Z, 2)
                                        myAttRef.TextString = myAttRef.TextString.Replace("----", h)
                                    End If
                                    'получаем id блока и отделяем от него скобки
                                    '    Dim strID As String = br.Id.ToString() 'получаем id вновь созданного блока
                                    '    strID = Mid(strID, 2, strID.Length - 2) 'убираем из строки первую и последние скобки
                                    '    Dim PosBlk As Point3d = br.Position 'смотрим обновленную позицию блока
                                    '    'строка поиска (ОТОБРАЖАТЬ КООРДИНАТЫ)
                                    '    Dim FindStr1 As String = "InsertionPoint \f "
                                    '    Dim ind1 As Integer = CodStrFiled.IndexOf(FindStr1)
                                    '    Dim indk As Integer = CodStrFiled.LastIndexOf(")")
                                    '    If ind1 <> -1 Then
                                    '        Dim str1 As String = "%<\AcObjProp Object(%<\_ObjId "
                                    '        Dim StartPos As Integer = ind1 + 18
                                    '        Dim SumbCount As Integer = indk - ind1 - 17
                                    '        Dim str2 As String = ">%).InsertionPoint \f " & Chr(34) & "%lu2%pr2%pt4" & Chr(34) & ">%"
                                    '        If SumbCount > 0 Then
                                    '            str2 = ">%).InsertionPoint \f " & Mid(CodStrFiled, ind1 + 18, indk - ind1 - 17) & ">%"
                                    '        End If
                                    '        Dim str As String = str1 + strID + str2
                                    '        Dim fldstr As String = String.Format(str, PosBlk.ToString())
                                    '        Dim field As Field = New Field(fldstr)
                                    '        field.Evaluate()
                                    '        Dim fieldEval As FieldEvaluationStatusResult = field.EvaluationStatus
                                    '        If fieldEval.Status = FieldEvaluationStatus.Success Then
                                    '            myAttRef.SetField(field)
                                    '            tr.AddNewlyCreatedDBObject(field, True)
                                    '        Else
                                    '            Dim Otm As Double = Math.Round(pointBlkXY.Z, 2)
                                    '            myAttDef.TextString = String.Format(CultureInfo.InvariantCulture, "{0:0.00}", Otm)
                                    '        End If
                                    '    End If
                                    '    'строка поиска (ДИНАМИЧЕСКИЙ ПАРАМЕТР ЛИНЕЙНЫЙ РАЗМЕР)
                                    '    Dim FindStr2 As String = ").UpdatedDistance \f"
                                    '    Dim ind2 As Integer = CodStrFiled.LastIndexOf(FindStr2)
                                    '    If ind2 <> -1 Then
                                    '        '1. ищем слово параметр
                                    '        Dim PosFindStrPar1 As String = CodStrFiled.LastIndexOf("Parameter(")
                                    '        Dim PosFlag As String = CodStrFiled.IndexOf("AcExpr")
                                    '        Dim str As String = ""
                                    '        If PosFlag > -1 Then
                                    '            Dim PosFindStrPar2 As String = CodStrFiled.LastIndexOf("\f")
                                    '            Dim str1 As String = "%<\AcObjProp Object(%<\_ObjId "
                                    '            Dim str2 As String = Mid(CodStrFiled, PosFindStrPar1, PosFindStrPar2 - PosFindStrPar1 - 2)
                                    '            str = str1 + strID + ">%)" + str2 + """" & ">%"
                                    '        Else
                                    '            Dim PosFindStrPar2 As String = CodStrFiled.LastIndexOf("\f")
                                    '            Dim str1 As String = "%<\AcObjProp Object(%<\_ObjId "
                                    '            Dim str2 As String = Mid(CodStrFiled, PosFindStrPar1)
                                    '            str = str1 + strID + ">%)" + str2 + ">%"
                                    '        End If
                                    '        Dim fldstr As String = String.Format(str, PosBlk.ToString())
                                    '        Dim field As Field = New Field(fldstr)
                                    '        field.Evaluate()
                                    '        Dim fieldEval As FieldEvaluationStatusResult = field.EvaluationStatus
                                    '        If fieldEval.Status = FieldEvaluationStatus.Success Then
                                    '            myAttRef.SetField(field)
                                    '            tr.AddNewlyCreatedDBObject(field, True)
                                    '        Else
                                    '            Dim Otm As Double = Math.Round(pointBlkXY.Z, 2)
                                    '            myAttDef.TextString = String.Format(CultureInfo.InvariantCulture, "{0:0.00}", Otm)
                                    '        End If
                                    '    End If
                                End If
                            End If
                        Next
                    End If

                    FuncInsertDrawingBlk = br
                    tr.Commit()
                End If
            Catch
                FuncInsertDrawingBlk = Nothing
            End Try
        End Using
    End Function



    '===================================================================================
    'создать полилинию из массива
    Public Shared Function createPolylineFromArrayCoordinates(ByVal arrayCoordinates(,) As String) As Polyline
        If arrayCoordinates Is Nothing OrElse arrayCoordinates.GetLength(0) = 0 Then
            Return Nothing
        End If
        ' Создаем новую полилинию
        Dim polyline As New Polyline()
        Try
            ' Получаем количество точек (первая размерность массива)
            Dim pointCount As Integer = arrayCoordinates.GetLength(0)
            ' Проходим по всем точкам
            For i As Integer = 0 To arrayCoordinates.GetUpperBound(1)
                ' Проверяем, что во второй размерности есть хотя бы 2 элемента (X и Y)
                If arrayCoordinates.GetLength(1) < 2 Then
                    Throw New ArgumentException("Массив должен содержать как минимум X и Y координаты")
                End If
                ' Парсим координаты
                Dim x As Double
                Dim y As Double
                If Not Double.TryParse(arrayCoordinates(1, i), x) Then
                    Throw New ArgumentException($"Не удалось преобразовать X координату в строке {i + 1}: {arrayCoordinates(i, 0)}")
                End If

                If Not Double.TryParse(arrayCoordinates(2, i), y) Then
                    Throw New ArgumentException($"Не удалось преобразовать Y координату в строке {i + 1}: {arrayCoordinates(i, 1)}")
                End If
                ' Создаем точку в 2D (Z = 0)
                Dim point2d As New Point2d(x, y)
                ' Добавляем вершину к полилинии
                polyline.AddVertexAt(i, point2d, 0, 0, 0)
            Next
            ' Закрываем полилинию (если нужно)
            'polyline.Closed = True
        Catch ex As Exception
            ' В случае ошибки освобождаем ресурсы
            Return Nothing
        End Try
        Return polyline
    End Function

    Public Shared Function createPolylineFromArrayCoordinates2(ByVal arrayCoordinates(,) As String) As Polyline
        ' Проверяем, что массив не пустой и содержит достаточно данных
        If arrayCoordinates Is Nothing OrElse arrayCoordinates.GetLength(0) < 2 Then
            Return Nothing
        End If

        ' Проверяем, что во второй размерности есть хотя бы 2 элемента (X и Y)
        If arrayCoordinates.GetLength(1) < 2 Then
            Return Nothing
        End If

        ' Создаем новую полилинию
        Dim polyline As New Polyline()

        Try
            ' Получаем количество точек (количество строк в первой размерности)
            Dim pointCount As Integer = arrayCoordinates.GetLength(0)

            ' Проходим по всем точкам (по строкам массива)
            For i As Integer = 0 To pointCount - 1
                ' Парсим координаты
                Dim x As Double
                Dim y As Double

                ' arrayCoordinates(i, 0) - X координата
                ' arrayCoordinates(i, 1) - Y координата
                If Not Double.TryParse(arrayCoordinates(i, 0), x) Then
                    Throw New ArgumentException($"Не удалось преобразовать X координату в строке {i + 1}: {arrayCoordinates(i, 0)}")
                End If

                If Not Double.TryParse(arrayCoordinates(i, 1), y) Then
                    Throw New ArgumentException($"Не удалось преобразовать Y координату в строке {i + 1}: {arrayCoordinates(i, 1)}")
                End If

                ' Создаем точку в 2D (Z = 0)
                Dim point2d As New Point2d(x, y)

                ' Добавляем вершину к полилинии
                polyline.AddVertexAt(i, point2d, 0, 0, 0)
            Next

            ' Если нужно замкнуть полилинию, раскомментируйте следующую строку
            ' polyline.Closed = True

        Catch ex As Exception
            ' В случае ошибки возвращаем Nothing
            Return Nothing
        End Try

        Return polyline
    End Function
    '==============================================================================================================================
    'нарисовать окружность
    Public Shared Function createCircleFromArrayCoordinates(ByVal arrayCoordinates(,) As String) As Circle
        If arrayCoordinates Is Nothing OrElse arrayCoordinates.GetLength(0) = 0 Then
            Return Nothing
        End If
        ' Создаем новую полилинию
        Dim circle As New Circle
        Try
            ' Проходим по всем точкам
            For i As Integer = 0 To arrayCoordinates.GetUpperBound(1)
                ' Парсим координаты
                Dim x As Double
                Dim y As Double
                Dim r As Double
                If Not Double.TryParse(arrayCoordinates(1, i), x) Then
                    Throw New ArgumentException($"Не удалось преобразовать X координату в строке {i + 1}: {arrayCoordinates(i, 0)}")
                End If

                If Not Double.TryParse(arrayCoordinates(2, i), y) Then
                    Throw New ArgumentException($"Не удалось преобразовать Y координату в строке {i + 1}: {arrayCoordinates(i, 1)}")
                End If

                If Not Double.TryParse(arrayCoordinates(0, i), r) Then
                    r = 0.5
                End If
                ' Создаем точку в 2D (Z = 0)
                Dim point2d As New Point2d(x, y)
                ' Добавляем вершину к полилинии
                circle.Center = New Point3d(point2d.X, point2d.Y, 0)
                circle.Radius = r
            Next
        Catch ex As Exception
            ' В случае ошибки освобождаем ресурсы
            Return Nothing
        End Try
        Return circle
    End Function
    Public Shared Function createCircleFromArrayCoordinates2(ByVal arrayCoordinates(,) As String) As Circle
        ' Проверяем, что массив не пустой и содержит достаточно данных
        If arrayCoordinates Is Nothing OrElse arrayCoordinates.GetLength(0) < 1 Then
            Return Nothing
        End If

        ' Проверяем, что во второй размерности есть хотя бы 2 элемента (X и Y)
        If arrayCoordinates.GetLength(1) < 2 Then
            Return Nothing
        End If

        ' Создаем новую окружность
        Dim circle As New Circle()

        Try
            ' Получаем количество точек (количество строк в первой размерности)
            Dim pointCount As Integer = arrayCoordinates.GetLength(0)

            ' Берем только первую точку для определения центра и радиуса
            ' (обычно окружность задается одной точкой центра и радиусом)
            Dim i As Integer = 0

            ' Парсим координаты
            Dim x As Double
            Dim y As Double
            Dim r As Double

            ' arrayCoordinates(i, 0) - X координата
            ' arrayCoordinates(i, 1) - Y координата
            If Not Double.TryParse(arrayCoordinates(i, 0), x) Then
                Throw New ArgumentException($"Не удалось преобразовать X координату: {arrayCoordinates(i, 0)}")
            End If

            If Not Double.TryParse(arrayCoordinates(i, 1), y) Then
                Throw New ArgumentException($"Не удалось преобразовать Y координату: {arrayCoordinates(i, 1)}")
            End If

            ' Если есть третий элемент - это радиус, иначе используем значение по умолчанию
            If arrayCoordinates.GetLength(1) >= 3 Then
                If Not Double.TryParse(arrayCoordinates(i, 2), r) Then
                    r = 0.5 ' Значение по умолчанию
                End If
            Else
                r = 0.5 ' Значение по умолчанию
            End If

            ' Устанавливаем центр окружности
            circle.Center = New Point3d(x, y, 0)
            circle.Radius = r

        Catch ex As Exception
            ' В случае ошибки возвращаем Nothing
            Return Nothing
        End Try

        Return circle
    End Function
End Class
