Imports System.IO
Imports HostMgd.ApplicationServices
Imports Teigha.DatabaseServices

Public Class FuncDWGReference
    'функция получает имена всех внешних ссылок (словарь 1-Путь, 2 - имя)
    Public Shared Function getXRef(ByRef dictionaryRef As Dictionary(Of String, String)) As Boolean
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim putchCurrentDwg As String = doc.Name
        Dim nameDoc As String = Path.GetFileNameWithoutExtension(putchCurrentDwg)
        dictionaryRef = New Dictionary(Of String, String)
        dictionaryRef.Add(putchCurrentDwg, nameDoc)
        Dim directory As String = Path.GetDirectoryName(putchCurrentDwg)
        Using docLock As DocumentLock = doc.LockDocument()
            Using trans As Transaction = db.TransactionManager.StartTransaction()
                Try
                    ' Получаем таблицу блоков
                    Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
                    ' Перебираем все записи в таблице блоков
                    For Each blockId As ObjectId In bt
                        Dim btr As BlockTableRecord = trans.GetObject(blockId, OpenMode.ForRead)
                        If btr.IsFromExternalReference Then
                            Dim putchDwg = btr.PathName
                            If putchDwg.StartsWith("..\") OrElse putchDwg.StartsWith(".\") Then
                                putchDwg = directory & "\" & btr.Name & ".dwg"
                                If File.Exists(putchDwg) = False Then Continue For
                            End If
                            If dictionaryRef.ContainsKey(putchDwg) = False Then
                                dictionaryRef.Add(putchDwg, btr.Name)
                            End If
                        End If
                    Next
                Catch ex As Exception
                    trans.Abort()
                    Return False
                End Try
            End Using
        End Using
        Return True
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'чтение объектов во внешних ссылках
    Public Shared Function GetObjectToReferenceDWG(ByVal putchReferenceDWG As String, ByRef classElements As CadastralTools.StructureElement.classObject) As Dictionary(Of CadastralTools.StructureElement.classObject, List(Of CadastralTools.StructureElement))
        Dim result As New Dictionary(Of CadastralTools.StructureElement.classObject, List(Of CadastralTools.StructureElement))()
        If File.Exists(putchReferenceDWG) = True Then
            Try
                Dim sourceDb As Database = New Database(False, True)
                sourceDb.ReadDwgFile(putchReferenceDWG, FileShare.Read, True, String.Empty)
                Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
                Using sourceTr
                    Dim sourceBtr As BlockTableRecord = sourceTr.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForRead, False)
                    For Each id As ObjectId In sourceBtr
                        If Not id.IsNull AndAlso Not id.IsErased Then
                            Dim acEnt As Entity = TryCast(sourceTr.GetObject(id, OpenMode.ForRead), Entity)
                            If acEnt IsNot Nothing Then
                                Dim civilListObject As Dictionary(Of String, CadastralTools.StructureElement) = FuncXRecords.readXRecords(acEnt)
                                If IsNothing(civilListObject) = False Then
                                    If civilListObject.Count > 0 Then
                                        For i As Integer = 0 To civilListObject.Count - 1
                                            Dim dataStructureObject As CadastralTools.StructureElement = civilListObject.ElementAt(i).Value
                                            Dim classEl As CadastralTools.StructureElement.classObject = dataStructureObject.NameClassObject
                                            If classElements = classEl Then
                                                Dim typeObject As CadastralTools.StructureElement.nameObject = dataStructureObject.NameTypeObject
                                                ' Проверяем, существует ли уже такой ключ в словаре
                                                If result.ContainsKey(typeObject) Then
                                                    ' Добавляем в существующий список
                                                    result(typeObject).Add(dataStructureObject)
                                                Else
                                                    ' Создаем новый список и добавляем элемент
                                                    result(typeObject) = New List(Of CadastralTools.StructureElement) From {dataStructureObject}
                                                End If
                                            End If
                                        Next i
                                    End If
                                End If
                            End If
                        End If
                    Next
                    sourceTr.Commit()
                End Using
                Return result
            Catch ex As System.Exception
            End Try
        End If
    End Function

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'чтение площадных объектов во внешней ссылке
    Public Shared Function GetCadastralPolygonToReferenceDWG(ByVal putchReferenceDWG As String, Optional nameObject As String = "") As Dictionary(Of String, List(Of CadastralTools.Topology))
        Dim result As New Dictionary(Of String, List(Of CadastralTools.Topology))
        Dim listOstrov As New List(Of NetTopologySuite.Geometries.Polygon)
        Dim classTopology As FuncTopologySuite = New FuncTopologySuite
        If File.Exists(putchReferenceDWG) = True Then
            Try
                Dim sourceDb As Database = New Database(False, True)
                sourceDb.ReadDwgFile(putchReferenceDWG, FileShare.Read, True, String.Empty)
                Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
                Using sourceTr
                    Dim sourceBtr As BlockTableRecord = sourceTr.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForRead, False)
                    For Each id As ObjectId In sourceBtr
                        If Not id.IsNull AndAlso Not id.IsErased Then
                            Dim acEnt As Entity = TryCast(sourceTr.GetObject(id, OpenMode.ForRead), Entity)
                            If acEnt IsNot Nothing Then
                                Dim civilListObject As Dictionary(Of String, CadastralTools.StructureElement) = FuncXRecords.readXRecords(acEnt)
                                If IsNothing(civilListObject) = False Then
                                    If civilListObject.Count > 0 Then
                                        For i As Integer = 0 To civilListObject.Count - 1
                                            Dim dataStructure As CadastralTools.StructureElement = civilListObject.ElementAt(i).Value
                                            Dim tempClassObject As CadastralTools.StructureElement.classObject = dataStructure.NameClassObject
                                            'объект является площадным
                                            If dataStructure.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polygon Or dataStructure.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.none Then
                                                'островки отдельно обрабатываем
                                                Dim deskObject As String = dataStructure.Description
                                                If Not (deskObject Like "Внутренний островок") Then
                                                    If nameObject.Trim.Length > 0 Then
                                                        If Not (deskObject Like nameObject) Then
                                                            Continue For
                                                        End If
                                                    End If
                                                    If TypeOf acEnt Is Polyline Then
                                                        Dim poly As Polyline = acEnt
                                                        If poly.Closed = True Or poly.StartPoint = poly.EndPoint Then
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
                End Using
                'проверяем все объекты на наличие внутренних островков
                'проверяем все объекты на наличие внутренних островков
                Dim classNetTopo As CadastralTools.FuncNetTopologySuite = New CadastralTools.FuncNetTopologySuite
                If result.Count > 0 Then
                    'этот блок проверяет чтобы полигоны не входили друг в друга
                    For i As Integer = 0 To result.Count - 1
                        Dim listSourcePolygon As List(Of CadastralTools.Topology) = result.ElementAt(i).Value
                        For j As Integer = 0 To listSourcePolygon.Count - 1
                            Dim sourcePolygon As NetTopologySuite.Geometries.Polygon = listSourcePolygon.Item(j).GeometryElement
                            If sourcePolygon.Area > 0 Then
                                For k As Integer = 0 To listSourcePolygon.Count - 1
                                    Dim userDestPolygon As CadastralTools.Topology = listSourcePolygon.Item(k)
                                    Dim destPolygon As NetTopologySuite.Geometries.Polygon = userDestPolygon.GeometryElement
                                    If destPolygon.Area > 0 Then
                                        If destPolygon.Area < sourcePolygon.Area Then
                                            If classNetTopo.IsPolygonInside(sourcePolygon, destPolygon) Then
                                                sourcePolygon = classNetTopo.AddPolygonAsHole(sourcePolygon, destPolygon)
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
                            If sourcePolygon.Area > 0 Then
                                For k As Integer = 0 To listOstrov.Count - 1
                                    Dim destPolygon As NetTopologySuite.Geometries.Polygon = listOstrov.Item(k)
                                    If destPolygon.Area > 0 Then
                                        If destPolygon.Area < sourcePolygon.Area Then
                                            If classNetTopo.IsPolygonInside(sourcePolygon, destPolygon) Then
                                                sourcePolygon = classNetTopo.AddPolygonAsHole(sourcePolygon, destPolygon)
                                                Exit For
                                            End If
                                        End If
                                    End If
                                Next k
                            End If
                        Next j
                    Next i
                End If
                Return result
            Catch ex As System.Exception
            End Try
        End If
    End Function
End Class
