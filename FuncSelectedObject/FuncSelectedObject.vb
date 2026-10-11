Imports CadastralTools
Imports HostMgd.ApplicationServices
Imports HostMgd.EditorInput
Imports NetTopologySuite
Imports Teigha.DatabaseServices
Imports Teigha.Geometry

Public Class FuncSelectedObject
    'функция снимает выделение Highligh с объектов
    Public Shared Function UnhighlightEntity(ByRef listObject As List(Of Entity)) As Boolean
        If IsNothing(listObject) = False Then
            If listObject.Count > 0 Then
                Dim doc As Document = Application.DocumentManager.MdiActiveDocument
                Dim db As Database = doc.Database
                Dim ed As Editor = doc.Editor
                ' Начинаем транзакцию для безопасного чтения объектов
                Using trans As Transaction = db.TransactionManager.StartTransaction()
                    Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
                    Dim ms As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    For i As Integer = 0 To listObject.Count - 1
                        Dim ent As Entity = listObject.Item(i)
                        If IsDBNull(ent) = False Then
                            If ent.IsWriteEnabled = False Then
                                ent.UpgradeOpen()
                            End If
                            ent.Unhighlight()
                        End If
                    Next
                    trans.Commit()
                End Using
                listObject.Clear()
            End If
        End If
        Return True
    End Function
    'функция выбирает объекты внутри замкнутой полилинии
    Public Shared Function selectAllObject(ByRef Db As Database) As Dictionary(Of StructureElement.nameObject, List(Of Geometries.Geometry))
        Dim result As Dictionary(Of StructureElement.nameObject, List(Of Geometries.Geometry)) = New Dictionary(Of StructureElement.nameObject, List(Of Geometries.Geometry))
        If IsNothing(Db) = True Then Return result
        Dim acadTopoNet As FuncTopologySuite = New FuncTopologySuite
        Using acTrans As Transaction = Db.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(Db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
            For Each id As ObjectId In btr
                Dim selectEntity As Teigha.DatabaseServices.Entity = acTrans.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                Dim dataStructuresObject As StructureElement = Nothing
                Dim boolRecData As Boolean = FuncXRecords.getXRecords(selectEntity, dataStructuresObject)
                If boolRecData = True Then
                    Dim nameObject As StructureElement.nameObject = dataStructuresObject.NameTypeObject
                    Dim typeGeometry As StructureElement.typeGeometryObject = dataStructuresObject.TypeGeometry
                    Dim objectGeometry As Geometries.Geometry = Nothing
                    If typeGeometry = StructureElement.typeGeometryObject.polygon Then
                        If TypeOf selectEntity Is Polyline Then
                            Dim polyline As Polyline = selectEntity
                            If polyline.Area > 0 Then
                                Dim polygon As NetTopologySuite.Geometries.Polygon = acadTopoNet.CreatePolygon(polyline)
                                If IsNothing(polygon) = False Then
                                    polygon.UserData = dataStructuresObject
                                    If result.ContainsKey(nameObject) = False Then
                                        result.Add(nameObject, New List(Of Geometries.Geometry) From {polygon})
                                    Else
                                        Dim listObject As List(Of Geometries.Geometry) = result.Item(nameObject)
                                        listObject.Add(polygon)
                                        result.Item(nameObject) = listObject
                                    End If
                                End If
                            End If
                        End If
                    ElseIf typeGeometry = StructureElement.typeGeometryObject.polyline Then
                        If TypeOf selectEntity Is Polyline Then
                            Dim polyline As Polyline = selectEntity
                            If polyline.Length > 0 Then
                                Dim line As NetTopologySuite.Geometries.LineString = acadTopoNet.CreateLineString(polyline)
                                If IsNothing(line) = False Then
                                    line.UserData = dataStructuresObject
                                    If result.ContainsKey(nameObject) = False Then
                                        result.Add(nameObject, New List(Of Geometries.Geometry) From {line})
                                    Else
                                        Dim listObject As List(Of Geometries.Geometry) = result.Item(nameObject)
                                        listObject.Add(line)
                                        result.Item(nameObject) = listObject
                                    End If
                                End If
                            End If
                        End If
                    Else
                        Dim insertPoint As Point3d = New Point3d()
                        If TypeOf selectEntity Is MLeader Then
                            Dim acLeader As MLeader = selectEntity
                            insertPoint = acLeader.GetFirstVertex(0)
                        ElseIf TypeOf selectEntity Is DBText Then
                            Dim acText As DBText = selectEntity
                            insertPoint = acText.Position
                        ElseIf TypeOf selectEntity Is Circle Then
                            Dim acCircle As Circle = selectEntity
                            insertPoint = acCircle.Center
                        ElseIf TypeOf selectEntity Is BlockReference Then
                            Dim acBlock As BlockReference = selectEntity
                            insertPoint = acBlock.Position
                        End If
                        Dim point As Geometries.Point = New Geometries.Point(insertPoint.X, insertPoint.Y, insertPoint.Z)
                        point.UserData = dataStructuresObject
                        If result.ContainsKey(nameObject) = False Then
                            result.Add(nameObject, New List(Of Geometries.Geometry) From {point})
                        Else
                            Dim listObject As List(Of Geometries.Geometry) = result.Item(nameObject)
                            listObject.Add(point)
                            result.Item(nameObject) = listObject
                        End If
                    End If
                End If
            Next
            acTrans.Commit()
        End Using
        Return result
    End Function

    '==========================================================================================================================================================
    'фильтрация объектов на основе таблицы PS
    Public Shared Function selectCadastralObject(ByRef Db As Database, Optional tablePS As String = "") As Dictionary(Of String, List(Of Geometries.Geometry))
        Dim result As Dictionary(Of String, List(Of Geometries.Geometry)) = New Dictionary(Of String, List(Of Geometries.Geometry))
        If IsNothing(Db) = True Then Return result
        Dim acadTopoNet As FuncTopologySuite = New FuncTopologySuite
        Using acTrans As Transaction = Db.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(Db.BlockTableId, OpenMode.ForRead)
            Dim btr As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForRead)
            For Each id As ObjectId In btr
                If id.IsNull = False Then
                    Dim selectEntity As Teigha.DatabaseServices.Entity = acTrans.GetObject(id, OpenMode.ForRead, False) 'открываем объект на чтение
                    Dim dataStructuresObject As StructureElement = Nothing
                    Dim boolRecData As Boolean = FuncXRecords.getXRecords(selectEntity, dataStructuresObject)
                    If boolRecData = True Then
                        Dim tablePSData As String = dataStructuresObject.TablePS
                        Dim boolFindObject As Boolean = False
                        If tablePS.Trim.Length > 0 Then
                            If tablePS Like tablePSData Then
                                boolFindObject = True
                            End If
                        Else
                            boolFindObject = True
                        End If
                        If tablePSData Like "CADASTRAL_OSTROV" Then
                            boolFindObject = True
                        End If
                        If boolFindObject = True Then
                            If TypeOf selectEntity Is Polyline Then
                                Dim polyline As Polyline = selectEntity
                                If polyline.Area > 0 And dataStructuresObject.TypeGeometry = StructureElement.typeGeometryObject.polygon Then
                                    Dim polygon As NetTopologySuite.Geometries.Polygon = acadTopoNet.CreatePolygon(polyline)
                                    If IsNothing(polygon) = False Then
                                        polygon.UserData = dataStructuresObject
                                        If result.ContainsKey(tablePSData) = False Then
                                            result.Add(tablePSData, New List(Of Geometries.Geometry) From {polygon})
                                        Else
                                            Dim listObject As List(Of Geometries.Geometry) = result.Item(tablePSData)
                                            listObject.Add(polygon)
                                            result.Item(tablePSData) = listObject
                                        End If
                                    End If
                                ElseIf dataStructuresObject.TypeGeometry = StructureElement.typeGeometryObject.polyline Then
                                    If TypeOf selectEntity Is Polyline Then
                                        If polyline.Length > 0 Then
                                            Dim line As NetTopologySuite.Geometries.LineString = acadTopoNet.CreateLineString(polyline)
                                            If IsNothing(line) = False Then
                                                line.UserData = dataStructuresObject
                                                If result.ContainsKey(tablePSData) = False Then
                                                    result.Add(tablePSData, New List(Of Geometries.Geometry) From {line})
                                                Else
                                                    Dim listObject As List(Of Geometries.Geometry) = result.Item(tablePSData)
                                                    listObject.Add(line)
                                                    result.Item(tablePSData) = listObject
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            acTrans.Commit()
        End Using
        Return result
    End Function
End Class
