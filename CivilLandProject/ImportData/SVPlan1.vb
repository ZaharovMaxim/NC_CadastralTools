Imports System.IO
Imports System.Xml
Imports HostMgd.ApplicationServices
Imports HostMgd.EditorInput
Imports HostMgd.Windows
Imports nanoCAD
Imports Teigha.DatabaseServices
Imports Teigha.Geometry
Imports Teigha.Runtime
Imports Document = HostMgd.ApplicationServices.Document
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Windows.Shell
Imports CadastralTools

Public Class SVPlan1
    <CommandMethod("SVCadastralPlan")>
    Public Sub SVCadastralPlan()
        Dim doc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        ' Блокировка документа
        Using locker As DocumentLock = doc.LockDocument()
            ' Начинаем транзакцию в основном чертеже
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Try
                    Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                    Dim ms As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    ' Выбор зоны разработки ДПТ
                    Dim funcTopoSuite As New FuncTopologySuite()
                    Dim boudPline As Polyline = Nothing
                    Dim boundPolygon As NetTopologySuite.Geometries.Polygon = Nothing
                    Dim ptResult As PromptEntityResult = ed.GetEntity("Выберите зону разработки ДПТ:")
                    If ptResult.Status = PromptStatus.OK Then
                        Dim ent As Entity = tr.GetObject(ptResult.ObjectId, OpenMode.ForRead)
                        If TypeOf ent Is Polyline Then
                            boudPline = TryCast(ent, Polyline)
                            If boudPline.Area > 0 Then
                                boundPolygon = funcTopoSuite.CreatePolygon(boudPline)
                            End If
                        Else
                            Dim msRez As MsgBoxResult = MsgBox("Зона не выбрана. Продолжить?", MsgBoxStyle.OkCancel)
                            If msRez = MsgBoxResult.Cancel Then
                                Exit Sub
                            End If
                        End If
                    End If
                    Dim resultPolygonDictionary As Dictionary(Of String, List(Of NetTopologySuite.Geometries.Polygon)) = New Dictionary(Of String, List(Of NetTopologySuite.Geometries.Polygon))
                    Dim resultLineDictionary As Dictionary(Of String, List(Of NetTopologySuite.Geometries.LineString)) = New Dictionary(Of String, List(Of NetTopologySuite.Geometries.LineString))
                    Dim resultPointDictionary As Dictionary(Of String, List(Of NetTopologySuite.Geometries.Point)) = New Dictionary(Of String, List(Of NetTopologySuite.Geometries.Point))

                    Dim putchFolder As String = CadastralTools.FuncFiles.getAbsoluteFolderPathName("Выберите каталог с файлами dwg")
                    If IO.Directory.Exists(putchFolder) Then
                        Dim files As String() = Directory.GetFiles(putchFolder, "*.dwg")
                        If files.Length = 0 Then
                            MsgBox("Не найдено DWG файлов в выбранной папке.")
                            Exit Sub
                        End If
                        For Each fileDwg As String In files
                            ' Пропускаем текущий файл
                            If System.IO.Path.GetFileName(fileDwg).Equals(System.IO.Path.GetFileName(db.Filename), StringComparison.OrdinalIgnoreCase) Then
                                Continue For
                            End If
                            ' Создаем временную базу данных для чтения исходного файла
                            Using sourceDb As New Database(False, True)
                                Try
                                    sourceDb.ReadDwgFile(fileDwg, FileOpenMode.OpenForReadAndAllShare, True, "")
                                    ' Собираем ID объектов для клонирования
                                    Dim idsToClone As New ObjectIdCollection()
                                    ' Используем транзакцию в исходной базе
                                    Using sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
                                        Dim sourceMs As BlockTableRecord = sourceTr.GetObject(sourceDb.CurrentSpaceId, OpenMode.ForRead)
                                        For Each id As ObjectId In sourceMs
                                            If Not id.IsNull AndAlso Not id.IsErased Then
                                                Dim acEnt As Entity = TryCast(sourceTr.GetObject(id, OpenMode.ForRead), Entity)
                                                If acEnt IsNot Nothing Then
                                                    If TypeOf acEnt Is Polyline Then
                                                        Dim cadastralPline As Polyline = TryCast(acEnt, Polyline)
                                                        ' Создаем NTS геометрию из полилинии
                                                        Dim objPolygon As NetTopologySuite.Geometries.Polygon = Nothing
                                                        Dim objLine As NetTopologySuite.Geometries.LineString = Nothing
                                                        Dim objPoint As NetTopologySuite.Geometries.Point = Nothing

                                                        If cadastralPline.Closed OrElse cadastralPline.StartPoint.Equals(cadastralPline.EndPoint) Then
                                                            objPolygon = funcTopoSuite.CreatePolygon(cadastralPline)
                                                        ElseIf cadastralPline.NumberOfVertices > 1 Then
                                                            objLine = funcTopoSuite.CreateLineString(cadastralPline)
                                                        End If

                                                        Dim classTopo As New CadastralTools.FuncNetTopologySuite()
                                                        Dim boolIntersect As Boolean = True
                                                        If objPolygon IsNot Nothing Then
                                                            If IsNothing(boundPolygon) = False Then
                                                                boolIntersect = classTopo.DoPolygonsIntersect(boundPolygon, objPolygon)
                                                            End If
                                                            If boolIntersect Then
                                                                Dim layerObj As String = acEnt.Layer
                                                                If resultPolygonDictionary.ContainsKey(layerObj) = True Then
                                                                    Dim dictObject As List(Of NetTopologySuite.Geometries.Polygon) = resultPolygonDictionary.Item(layerObj)
                                                                    If dictObject.Count > 0 Then
                                                                        Dim boolSr As Boolean = False
                                                                        For k As Integer = 0 To dictObject.Count - 1
                                                                            Dim srPoly As NetTopologySuite.Geometries.Polygon = dictObject.Item(k)
                                                                            If srPoly.EqualsTopologically(objPolygon) Then
                                                                                boolSr = True
                                                                                Exit For
                                                                            End If
                                                                        Next
                                                                        If boolSr = False Then
                                                                            idsToClone.Add(id)
                                                                            dictObject.Add(objPolygon)
                                                                            resultPolygonDictionary.Item(layerObj) = dictObject
                                                                        End If
                                                                    End If
                                                                Else
                                                                    idsToClone.Add(id)
                                                                    Dim listObj As List(Of NetTopologySuite.Geometries.Polygon) = New List(Of NetTopologySuite.Geometries.Polygon)
                                                                    listObj.Add(objPolygon)
                                                                    resultPolygonDictionary.Add(layerObj, listObj)
                                                                End If
                                                            End If

                                                        ElseIf objLine IsNot Nothing Then
                                                            If IsNothing(boundPolygon) = False Then
                                                                boolIntersect = classTopo.DoesLineIntersectPolygon(objLine, boundPolygon)
                                                            End If
                                                            If boolIntersect Then
                                                                Dim layerObj As String = acEnt.Layer
                                                                If resultLineDictionary.ContainsKey(layerObj) = True Then
                                                                    Dim dictObject As List(Of NetTopologySuite.Geometries.LineString) = resultLineDictionary.Item(layerObj)
                                                                    If dictObject.Count > 0 Then
                                                                        Dim boolSr As Boolean = False
                                                                        For k As Integer = 0 To dictObject.Count - 1
                                                                            Dim srLine As NetTopologySuite.Geometries.LineString = dictObject.Item(k)
                                                                            If srLine.EqualsTopologically(srLine) Then
                                                                                boolSr = True
                                                                                Exit For
                                                                            End If
                                                                        Next
                                                                        If boolSr = False Then
                                                                            idsToClone.Add(id)
                                                                            dictObject.Add(objLine)
                                                                            resultLineDictionary.Item(layerObj) = dictObject
                                                                        End If
                                                                    End If
                                                                Else
                                                                    idsToClone.Add(id)
                                                                    Dim listObj As List(Of NetTopologySuite.Geometries.LineString) = New List(Of NetTopologySuite.Geometries.LineString)
                                                                    listObj.Add(objLine)
                                                                    resultLineDictionary.Add(layerObj, listObj)
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Next
                                        sourceTr.Commit()
                                    End Using
                                    ' Если есть объекты для клонирования
                                    If idsToClone.Count > 0 Then
                                        ' Создаем маппинг ID
                                        Dim idMap As New IdMapping()

                                        ' ВАЖНО: Используем WblockCloneObjects для клонирования из sourceDb в текущий db
                                        sourceDb.WblockCloneObjects(
                                                idsToClone,                 ' ObjectId из исходной базы
                                                ms.Id,                      ' ID пространства модели текущего чертежа
                                                idMap,                      ' Маппинг ID
                                                DuplicateRecordCloning.Ignore,
                                                False)
                                    End If

                                Catch ex As Exception
                                    ed.WriteMessage(vbLf & $"Ошибка при обработке файла {fileDwg}: {ex.Message}")
                                End Try
                            End Using
                        Next
                    End If


                    tr.Commit()
                Catch ex As Exception
                    ed.WriteMessage(vbLf & $"Ошибка: {ex.Message}")
                    tr.Abort()
                End Try
            End Using
        End Using
    End Sub
End Class
