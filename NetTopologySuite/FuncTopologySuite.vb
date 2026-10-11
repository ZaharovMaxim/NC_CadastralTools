Imports HostMgd.ApplicationServices
Imports HostMgd.EditorInput
Imports NetTopologySuite.Geometries
Imports Teigha.DatabaseServices
Imports Teigha.Geometry

Public Class FuncTopologySuite
    Private ReadOnly _geometryFactory As GeometryFactory
    Public Sub New()
        ' Создаем фабрику геометрий с точностью 0.001
        _geometryFactory = New GeometryFactory(New PrecisionModel(1000))
    End Sub

    ' Создает полигон из полилинии
    Public Function CreatePolygon(ByVal polyline As Polyline) As Polygon
        If polyline.NumberOfVertices < 3 Then
            Return Nothing
        End If
        Dim points As List(Of (x As Double, y As Double)) = New List(Of (x As Double, y As Double))
        For i As Integer = 0 To polyline.NumberOfVertices - 1
            Dim vert As Point2d = polyline.GetPoint2dAt(i)
            points.Add((vert.X, vert.Y))
        Next i
        If polyline.NumberOfVertices > 2 Then
            If Not (points.Item(0).x = points.Item(points.Count - 1).x And points.Item(0).y = points.Item(points.Count - 1).y) Then
                points.Add((polyline.StartPoint.X, polyline.StartPoint.Y))
            End If
        End If
        Dim coordinates = points.Select(Function(p) New Coordinate(p.x, p.y)).ToArray()
        Dim linearRing = _geometryFactory.CreateLinearRing(coordinates)
        Dim pol As Polygon = New Polygon(Nothing)
        Return _geometryFactory.CreatePolygon(linearRing)
    End Function

    'создание полилинии
    Public Function CreateLineString(ByVal polyline As Polyline) As LineString
        If polyline.NumberOfVertices < 2 Then
            Return Nothing
        End If
        Dim points As List(Of (x As Double, y As Double)) = New List(Of (x As Double, y As Double))
        For i As Integer = 0 To polyline.NumberOfVertices - 1
            Dim vert As Point2d = polyline.GetPoint2dAt(i)
            points.Add((vert.X, vert.Y))
        Next i
        Dim coordinates = points.Select(Function(p) New Coordinate(p.x, p.y)).ToArray()
        Return _geometryFactory.CreateLineString(coordinates)
    End Function

    Public Function CreatePoint(ByVal point As Circle) As Point
        Dim pt As Point = New Point(New Coordinate(point.Center.X, point.Center.Y))
        Return pt
    End Function


    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'работа с точками
    Public Function getPointByListBlockReference(ByRef db As Database, ByVal blockName As String, ByVal tagAttribute As String) As Dictionary(Of String, Point)
        Dim result As New Dictionary(Of String, Point)
        Dim doc As Document = Application.DocumentManager.GetDocument(db)
        If doc Is Nothing Then Return result
        Dim ed As Editor = doc.Editor
        ' Фильтр по типу объекта и имени блока (имя слоя в фильтр не заложить — проверяем вручную)
        Dim filter As New SelectionFilter(New TypedValue() {
            New TypedValue(CInt(DxfCode.Start), "INSERT"),
            New TypedValue(CInt(DxfCode.BlockName), blockName)
        })

        Dim selRes As PromptSelectionResult = ed.SelectAll(filter)
        If selRes.Status <> PromptStatus.OK Then Return result

        Using tr As Transaction = db.TransactionManager.StartTransaction()
            ' Кэш состояний слоёв, чтобы не открывать один и тот же слой многократно
            Dim layerVisible As New Dictionary(Of ObjectId, Boolean)()

            For Each id As ObjectId In selRes.Value.GetObjectIds()
                Dim br As BlockReference = TryCast(tr.GetObject(id, OpenMode.ForRead), BlockReference)
                If br Is Nothing Then Continue For

                Dim layerId As ObjectId = br.LayerId
                Dim isVisible As Boolean

                If Not layerVisible.TryGetValue(layerId, isVisible) Then
                    Dim ltr As LayerTableRecord =
                        TryCast(tr.GetObject(layerId, OpenMode.ForRead), LayerTableRecord)
                    isVisible = (ltr IsNot Nothing) AndAlso (Not ltr.IsOff) AndAlso (Not ltr.IsFrozen)
                    layerVisible(layerId) = isVisible
                End If

                If isVisible Then
                    Dim numberBlk As String = AttributeCadastralBlock.readTextStringByAttribute(db, br, "NUMBER")
                    Dim userPoint As Point = New Point(br.Position.X, br.Position.Y, br.Position.Z)
                    userPoint.UserData = Val(numberBlk)
                    result.Add(br.Handle.ToString, userPoint)
                End If
            Next

            tr.Commit()
        End Using

        Return result
    End Function

End Class
