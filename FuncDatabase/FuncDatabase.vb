Imports System.IO
Imports HostMgd.ApplicationServices
Imports NetTopologySuite
Imports Teigha.DatabaseServices
Imports Teigha.GraphicsInterface
Imports Application = HostMgd.ApplicationServices.Application
Imports Polyline = Teigha.DatabaseServices.Polyline

Public Class FuncDatabase
    ''' <summary>
    ''' Возвращает текущий масштаб аннотаций в формате, совместимом с AutoCAD-версией.
    ''' </summary>
    Public Shared Function getCurrentAnnotativeScaleDWG() As Double
        Dim doc As HostMgd.ApplicationServices.Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        If doc Is Nothing Then Return 1.0
        ' Получаем значение системной переменной CANNOSCALEVALUE
        ' Например, для масштаба 1:1000 вернёт 0.001
        Dim scaleValue As Double = HostMgd.ApplicationServices.Application.GetSystemVariable("CANNOSCALEVALUE")
        ' Возвращаем результат в том же виде, что и в вашей AutoCAD-функции
        getCurrentAnnotativeScaleDWG = scaleValue * 1000
    End Function
    'функция по Hangle объекта возвращает его ObjectID
    Public Shared Function getObjectID(ByVal HandleString As String) As ObjectId
        Dim id As ObjectId = Nothing
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        If IsNothing(HandleString) = False Then
            If HandleString.Trim.Length > 1 Then
                Try
                    'получаем доступ к полилиниии в которую требуется внести изменения
                    Dim ln As Long = Convert.ToInt64(HandleString, 16)
                    Dim hn As Handle = New Handle(ln)
                    id = db.GetObjectId(False, hn, 0)
                Catch ex As System.Exception
                End Try
            End If
        End If
        Return id
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция возвращает все объекты Плагина (полигоны)
    Public Shared Function getCadastralPolygon() As Dictionary(Of String, List(Of CadastralTools.Topology))
        Dim result As New Dictionary(Of String, List(Of CadastralTools.Topology))
        Dim listOstrov As New List(Of NetTopologySuite.Geometries.Polygon)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Try
            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Using tr
                Dim btr As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead, False)
                Dim classTopology As FuncTopologySuite = New FuncTopologySuite
                For Each id As ObjectId In btr
                    If Not id.IsNull AndAlso Not id.IsErased Then
                        Dim acEnt As Entity = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
                        If acEnt IsNot Nothing Then
                            Dim civilListObject As Dictionary(Of String, CadastralTools.StructureElement) = FuncXRecords.readXRecords(acEnt)
                            If IsNothing(civilListObject) = False Then
                                If civilListObject.Count > 0 Then
                                    For i As Integer = 0 To civilListObject.Count - 1
                                        Dim dataStructure As CadastralTools.StructureElement = civilListObject.ElementAt(i).Value
                                        Dim tempClassObject As CadastralTools.StructureElement.classObject = dataStructure.NameClassObject
                                        'объект является площадным
                                        If dataStructure.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polygon Then
                                            'островки отдельно обрабатываем
                                            Dim deskObject As String = dataStructure.Description
                                            If Not (deskObject Like "Внутренний островок") Then
                                                If TypeOf acEnt Is Polyline Then
                                                    Dim poly As Polyline = acEnt
                                                    If poly.Closed = True Then
                                                        Dim polygon As NetTopologySuite.Geometries.Polygon = classTopology.CreatePolygon(poly)
                                                        Dim userPolygon As CadastralTools.Topology = New CadastralTools.Topology()
                                                        userPolygon.GeometryElement = polygon
                                                        userPolygon.HandleElement = acEnt.Handle.ToString
                                                        userPolygon.StructureElement = dataStructure
                                                        If result.ContainsKey(deskObject) = False Then
                                                            Dim newList As New List(Of CadastralTools.Topology)
                                                            newList.Add(userPolygon)
                                                            result.Add(deskObject, newList)
                                                        Else
                                                            Dim newList As List(Of CadastralTools.Topology) = result.Item(deskObject)
                                                            newList.Add(userPolygon)
                                                            result.Item(deskObject) = newList
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If TypeOf acEnt Is Polyline Then
                                                    Dim poly As Polyline = acEnt
                                                    If poly.NumberOfVertices > 2 Then
                                                        Dim polygon As NetTopologySuite.Geometries.Polygon = classTopology.CreatePolygon(poly)
                                                        listOstrov.Add(polygon)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next i
                                End If
                            End If
                        End If
                    End If
                Next
                'проверяем все объекты на наличие внутренних островков
                Dim classNetTopo As CadastralTools.FuncNetTopologySuite = New CadastralTools.FuncNetTopologySuite
                If result.Count > 0 Then
                    'этот блок проверяет чтобы полигоны не входили друг в друга
                    For i As Integer = 0 To result.Count - 1
                        Dim listSourcePolygon As List(Of CadastralTools.Topology) = result.ElementAt(i).Value
                        For j As Integer = 0 To listSourcePolygon.Count - 1
                            Dim sourcePolygon As NetTopologySuite.Geometries.Polygon = listSourcePolygon.Item(j).GeometryElement
                            Dim areaSourcePolygon As Double = sourcePolygon.Area
                            If areaSourcePolygon > 0 Then
                                For k As Integer = 0 To listSourcePolygon.Count - 1
                                    Dim userDestPolygon As CadastralTools.Topology = listSourcePolygon.Item(k)
                                    Dim destPolygon As NetTopologySuite.Geometries.Polygon = userDestPolygon.GeometryElement
                                    Dim areaDestPolygon As Double = destPolygon.Area
                                    If areaDestPolygon > 0 Then
                                        If areaDestPolygon < areaSourcePolygon Then
                                            If classNetTopo.IsPolygonInside(sourcePolygon, destPolygon) Then
                                                sourcePolygon = classNetTopo.AddPolygonAsHole(sourcePolygon, destPolygon)
                                                areaSourcePolygon -= areaDestPolygon
                                                Dim structDataPolygon As CadastralTools.StructureElement = listSourcePolygon.Item(j).StructureElement
                                                Dim areaPoly As Double = CadastralTools.JsonValidator.GetValueFromJson(structDataPolygon.KeyParameter, "Area")
                                                If areaPoly > 0 Then
                                                    CadastralTools.JsonValidator.SetJsonFieldValue(structDataPolygon.KeyParameter, "Area", Math.Round(areaSourcePolygon, 2))
                                                End If
                                                Exit For
                                            End If
                                        End If
                                    End If
                                Next k
                            End If
                        Next j
                    Next i
                    'проверяем островки
                    For i As Integer = 0 To result.Count - 1
                        Dim listSourcePolygon As List(Of CadastralTools.Topology) = result.ElementAt(i).Value
                        For j As Integer = 0 To listSourcePolygon.Count - 1
                            Dim sourcePolygon As NetTopologySuite.Geometries.Polygon = listSourcePolygon.Item(j).GeometryElement
                            Dim areaSourcePolygon As Double = sourcePolygon.Area
                            If areaSourcePolygon > 0 Then
                                For k As Integer = 0 To listOstrov.Count - 1
                                    Dim destPolygon As NetTopologySuite.Geometries.Polygon = listOstrov.Item(k)
                                    Dim areaDestPolygon As Double = destPolygon.Area
                                    If areaDestPolygon Then
                                        If areaDestPolygon < areaSourcePolygon Then
                                            If classNetTopo.IsPolygonInside(sourcePolygon, destPolygon) Then
                                                sourcePolygon = classNetTopo.AddPolygonAsHole(sourcePolygon, destPolygon)
                                                areaSourcePolygon -= areaDestPolygon
                                                Dim structDataPolygon As CadastralTools.StructureElement = listSourcePolygon.Item(j).StructureElement
                                                Dim areaPoly As Double = CadastralTools.JsonValidator.GetValueFromJson(structDataPolygon.KeyParameter, "Area")
                                                If areaPoly > 0 Then
                                                    CadastralTools.JsonValidator.SetJsonFieldValue(structDataPolygon.KeyParameter, "Area", Math.Round(areaSourcePolygon, 2))
                                                End If
                                                Exit For
                                            End If
                                        End If
                                    End If
                                Next k
                            End If
                        Next j
                    Next i
                End If
            End Using
        Catch ex As System.Exception
        End Try
        Return result
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция возвращает все объекты Плагина (линии)
    Public Shared Function getCadastralLine() As Dictionary(Of String, List(Of CadastralTools.Topology))
        Dim result As New Dictionary(Of String, List(Of CadastralTools.Topology))
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Try
            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Using tr
                Dim btr As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead, False)
                Dim classTopology As FuncTopologySuite = New FuncTopologySuite
                For Each id As ObjectId In btr
                    If Not id.IsNull AndAlso Not id.IsErased Then
                        Dim acEnt As Entity = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
                        If acEnt IsNot Nothing Then
                            Dim civilListObject As Dictionary(Of String, CadastralTools.StructureElement) = FuncXRecords.readXRecords(acEnt)
                            If IsNothing(civilListObject) = False Then
                                If civilListObject.Count > 0 Then
                                    For i As Integer = 0 To civilListObject.Count - 1
                                        Dim dataStructure As CadastralTools.StructureElement = civilListObject.ElementAt(i).Value
                                        Dim tempClassObject As CadastralTools.StructureElement.classObject = dataStructure.NameClassObject
                                        'объект является площадным
                                        If dataStructure.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.line Then
                                            'островки отдельно обрабатываем
                                            Dim deskObject As String = dataStructure.Description
                                            If TypeOf acEnt Is Polyline Then
                                                Dim poly As Polyline = acEnt
                                                If poly.NumberOfVertices > 2 Then
                                                    Dim userPolygon As CadastralTools.Topology = New CadastralTools.Topology()
                                                    Dim lineStr As NetTopologySuite.Geometries.LineString = classTopology.CreateLineString(poly)
                                                    userPolygon.GeometryLineElement = lineStr
                                                    userPolygon.StructureElement = dataStructure
                                                    userPolygon.HandleElement = acEnt.Handle.ToString
                                                    If result.Count = 0 Then
                                                        Dim newListLine As List(Of CadastralTools.Topology) = New List(Of CadastralTools.Topology) From {userPolygon}
                                                        result.Add(deskObject, newListLine)
                                                    Else
                                                        If result.ContainsKey(deskObject) = True Then
                                                            Dim newListLine As List(Of CadastralTools.Topology) = result.Item(deskObject)
                                                            newListLine.Add(userPolygon)
                                                            result.Item(deskObject) = newListLine
                                                        Else
                                                            Dim newListLine As List(Of CadastralTools.Topology) = New List(Of CadastralTools.Topology) From {userPolygon}
                                                            result.Add(deskObject, newListLine)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next i
                                End If
                            End If
                        End If
                    End If
                Next
            End Using
        Catch ex As System.Exception
        End Try
        Return result
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция возвращает все точечные объекты Плагина
    Public Shared Function getCadastralPoint() As Dictionary(Of String, List(Of NetTopologySuite.Geometries.Point))
        Dim result As New Dictionary(Of String, List(Of NetTopologySuite.Geometries.Point))
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Try
            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Using tr
                Dim btr As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead, False)
                Dim classTopo As FuncTopologySuite = New FuncTopologySuite
                For Each id As ObjectId In btr
                    If Not id.IsNull AndAlso Not id.IsErased Then
                        Dim acEnt As Entity = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
                        If acEnt IsNot Nothing Then
                            If TypeOf acEnt Is BlockReference Then
                                Dim blk As BlockReference = acEnt
                                If blk.Name Like "CADASTRAL_POINT" Then
                                    Dim arrayData As String(,) = FuncBlockReference.readAttributesBlock(blk)
                                    If arrayData.Length > 0 Then
                                        Dim numberPoint As Integer = CadastralTools.MathFunction.findValueToArray2d(arrayData, "Number", 0, 1)
                                        If numberPoint > 0 Then
                                            Dim pt As NetTopologySuite.Geometries.Point = New NetTopologySuite.Geometries.Point(blk.Position.X, blk.Position.Y)
                                            pt.UserData = arrayData
                                            If result.Count = 0 Then
                                                result.Add(numberPoint, New List(Of NetTopologySuite.Geometries.Point) From {pt})
                                            ElseIf result.ContainsKey(numberPoint) = False Then
                                                result.Add(numberPoint, New List(Of NetTopologySuite.Geometries.Point) From {pt})
                                            Else
                                                Dim listPoint As List(Of NetTopologySuite.Geometries.Point) = result.Item(numberPoint)
                                                listPoint.Add(pt)
                                                result.Item(numberPoint) = listPoint
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            End Using
        Catch ex As System.Exception
        End Try
        If result.Count > 1 Then
            result = result.OrderBy(Function(x) x.Key).ToDictionary(Function(x) x.Key, Function(x) x.Value)
        End If
        Return result
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция возвращает список объектов плагина по его описанию
    Public Shared Function groupObjectByNumber(ByVal deskObject As String, ByVal dictionaryPolygon As Dictionary(Of String, List(Of CadastralTools.Topology))) As Dictionary(Of String, Dictionary(Of String, CadastralTools.Topology))
        Dim result As New Dictionary(Of String, Dictionary(Of String, CadastralTools.Topology))
        If dictionaryPolygon.Count > 0 Then
            If dictionaryPolygon.ContainsKey(deskObject) = True Then
                Dim listCounter As List(Of CadastralTools.Topology) = dictionaryPolygon.Item(deskObject)
                For i As Integer = 0 To listCounter.Count - 1
                    Dim userPolygon As CadastralTools.Topology = listCounter.Item(i)
                    Dim dataPolygon As CadastralTools.StructureElement = userPolygon.StructureElement
                    Dim keyParam As String = dataPolygon.KeyParameter
                    If CadastralTools.JsonValidator.IsValidJson(keyParam) = True Then
                        Dim conditionalPolygon As String = ""
                        'отдельная ветка - если это объект ЕГРН - здесь кадастрвый номер читем а не номер объекта
                        If dataPolygon.NameClassObject = CadastralTools.StructureElement.classObject.LAND_CADASTRAL Then
                            conditionalPolygon = CadastralTools.JsonValidator.GetValueFromJson(keyParam, "Cad_number")
                            If IsNothing(conditionalPolygon) = True Then
                                conditionalPolygon = CadastralTools.JsonValidator.GetValueFromJson(keyParam, "Reg_numb_border")
                            End If
                            If IsNothing(conditionalPolygon) = False Then
                                If result.ContainsKey(conditionalPolygon) = False Then
                                    Dim dictCounter As New Dictionary(Of String, CadastralTools.Topology)
                                    dictCounter.Add(conditionalPolygon, userPolygon)
                                    result.Add(conditionalPolygon, dictCounter)
                                Else
                                    Dim dictCounter As Dictionary(Of String, CadastralTools.Topology) = result.Item(conditionalPolygon)
                                    If dictCounter.ContainsKey(conditionalPolygon) = False Then
                                        dictCounter.Add(conditionalPolygon, userPolygon)
                                        result.Item(conditionalPolygon) = dictCounter
                                    End If
                                End If
                            End If
                        Else
                            conditionalPolygon = CadastralTools.JsonValidator.GetValueFromJson(keyParam, "ConditionalNumber")
                            Dim numberPolygon As Integer = 0
                            Dim numberCounter As Integer = 0
                            If IsNothing(conditionalPolygon) = False Then
                                If conditionalPolygon.Trim.Length > 0 Then
                                    numberPolygon = CadastralTools.DataFormatting.GetNumberBeforeParenthesis(conditionalPolygon)
                                    numberCounter = CadastralTools.DataFormatting.GetNumberInParenthesis(conditionalPolygon)
                                End If
                            End If
                            If numberPolygon > 0 Then
                                If result.ContainsKey(numberPolygon) = False Then
                                    Dim dictCounter As New Dictionary(Of String, CadastralTools.Topology)
                                    dictCounter.Add(numberCounter, userPolygon)
                                    result.Add(numberPolygon, dictCounter)
                                Else
                                    Dim dictCounter As Dictionary(Of String, CadastralTools.Topology) = result.Item(numberPolygon)
                                    If dictCounter.ContainsKey(numberCounter) = False Then
                                        dictCounter.Add(numberCounter, userPolygon)
                                        result.Item(numberPolygon) = dictCounter
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        End If
        If result.Count > 0 Then
            result = result.OrderBy(Function(x) x.Key).ToDictionary(Function(x) x.Key, Function(x) x.Value)
        End If
        Return result
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает объекты с кадастрового квартала
    Public Shared Function getNumberCadastralObject() As List(Of String)
        Dim result As New List(Of String)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Try
            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Using tr
                Dim btr As BlockTableRecord = tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead, False)
                For Each id As ObjectId In btr
                    If Not id.IsNull AndAlso Not id.IsErased Then
                        Dim acEnt As Entity = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
                        If acEnt IsNot Nothing Then
                            Dim civilListObject As Dictionary(Of String, CadastralTools.StructureElement) = FuncXRecords.readXRecords(acEnt)
                            If IsNothing(civilListObject) = False Then
                                If civilListObject.Count > 0 Then
                                    For i As Integer = 0 To civilListObject.Count - 1
                                        Dim dataStructure As CadastralTools.StructureElement = civilListObject.ElementAt(i).Value
                                        Dim keyParam As String = dataStructure.KeyParameter
                                        Dim objectNumber As Object = CadastralTools.JsonValidator.GetValueFromJson(keyParam, "Cad_number")
                                        If TypeOf objectNumber Is String Then
                                            result.Add(objectNumber.ToString)
                                        Else
                                            objectNumber = CadastralTools.JsonValidator.GetValueFromJson(keyParam, "Rec_numb_border")
                                            If TypeOf objectNumber Is String Then
                                                result.Add(objectNumber.ToString)
                                            End If
                                        End If
                                    Next i
                                End If
                            End If
                        End If
                    End If
                Next
            End Using
        Catch ex As System.Exception
        End Try
        Return result
    End Function

End Class
