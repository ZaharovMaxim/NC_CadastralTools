Imports CadastralTools
Imports CadastralTools.StructureElement
Imports HostMgd.ApplicationServices
Imports HostMgd.EditorInput
Imports HostMgd.Windows.ToolPalette
Imports NetTopologySuite.Geometries
Imports System.IO
Imports System.Windows.Forms
Imports Teigha.DatabaseServices
Imports Teigha.Runtime

Public Class AutoCreateParcels
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'автоматическое создание земельных участков
    <CommandMethod("AutomaticCreateCadastralObject")>
    Public Sub AutomaticCreateParcels()
        Dim acDoc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim boundPolygon As NetTopologySuite.Geometries.Polygon = Nothing
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'класс для работы с топологией
        Dim formParcelCreate As CadastralTools.FormAutomaticCreateParcels = New FormAutomaticCreateParcels
        Dim classCivilUtility As New CadastralTools.FuncNetTopologySuite()
        Dim classTopo As FuncTopologySuite = New FuncTopologySuite
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'запускаем форму
        formParcelCreate.ShowDialog()
        If formParcelCreate.boolShow = True Then
        ElseIf formParcelCreate.boolShow = False Then
            Exit Sub
        End If
        Dim dictionaryProperties As Dictionary(Of String, String) = CadastralTools.WinFormDataGrid.serializeWinFormDataGridToDictionary(formParcelCreate.DataGridView1)
        If dictionaryProperties.Count = 0 Then
            Exit Sub
        End If
        Dim listParcels As Dictionary(Of String, List(Of Topology)) = New Dictionary(Of String, List(Of Topology))
        'Dim styleObject As CadastralLineStyle = New CadastralLineStyle
        'Dim boolCreateStyle As Boolean = styleObject.createLineStyle(db, formParcelCreate.propertiesObject, formParcelCreate.putchTemplateXML)
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'выбираем кадастровые участки
        Using tr As Transaction = ed.Document.TransactionManager.StartTransaction()
            Dim result As String = ""
            Dim promptOpts As New PromptEntityOptions(vbLf & "Выберите внешнюю ссылку: ")
            promptOpts.SetRejectMessage("Выбранный объект не является внешней ссылкой")
            promptOpts.AddAllowedClass(GetType(BlockReference), False)
            Dim promptResult As PromptEntityResult = ed.GetEntity(promptOpts)
            If promptResult.Status = PromptStatus.OK Then
                Dim blkRefId As ObjectId = promptResult.ObjectId
                ' Открываем выбранную вставку блока
                Dim blkRef As BlockReference = TryCast(tr.GetObject(blkRefId, OpenMode.ForRead), BlockReference)
                If blkRef IsNot Nothing Then
                    ' Открываем определение блока (BlockTableRecord) на которое ссылается вставка
                    Dim btr As BlockTableRecord = TryCast(tr.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                    If btr IsNot Nothing Then
                        ' Проверяем, является ли блок внешней ссылкой
                        If btr.IsFromExternalReference Then
                            ' ПОЛНОЕ ИМЯ ФАЙЛА ВНЕШНЕЙ ССЫЛКИ ЗДЕСЬ!
                            result = btr.PathName
                            listParcels = FuncDWGReference.GetCadastralPolygonToReferenceDWG(result, "Кадастровый участок учтенный")
                        Else
                            MsgBox(vbLf & "Выбранный блок не является внешней ссылкой")
                        End If
                    End If
                End If
            End If
        End Using
        Dim startNumberParcels As Integer = formParcelCreate.numberParcel
        Do
            Using tr As Transaction = ed.Document.TransactionManager.StartTransaction()
line1:          Dim ptResult As PromptEntityResult = ed.GetEntity("Выберите внешнюю границу:")
                If ptResult.Status = PromptStatus.OK Then
                    Dim ent As Entity = tr.GetObject(ptResult.ObjectId, OpenMode.ForRead)
                    If TypeOf ent Is Polyline Then
                        Dim boundPline As Polyline = TryCast(ent, Polyline)
                        If boundPline.Area > 0 Then
                            boundPolygon = classTopo.CreatePolygon(boundPline)
                        Else
                            GoTo line1
                        End If
                    Else
                        GoTo line1
                    End If
                    Dim processor As New PolygonIntersectionInnerProcessor()
                    Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                    Dim btr As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    If listParcels.Count > 0 Then
                        For i As Integer = 0 To listParcels.Count - 1
                            Dim listCadastralParcels As List(Of Topology) = listParcels.ElementAt(i).Value
                            If listCadastralParcels.Count > 0 Then
                                For j As Integer = 0 To listCadastralParcels.Count - 1
                                    Dim cadastralParcel As Topology = listCadastralParcels.Item(j)
                                    Dim topoPoly As Polygon = cadastralParcel.GeometryElement
                                    Dim boolIntersectPolygon As Boolean = classCivilUtility.DoPolygonsIntersect(boundPolygon, topoPoly)
                                    If boolIntersectPolygon = True Then
                                        Dim allPolygons As List(Of Polygon) = processor.GetIntersectionPolygonsWithCheck(boundPolygon, topoPoly)
                                        If allPolygons.Count > 0 Then
                                            Dim dataStructure As CadastralTools.StructureElement = cadastralParcel.StructureElement
                                            Dim dictPropertiesParcel As Dictionary(Of String, Object) = CadastralTools.JsonValidator.DeserializeJsonToDictionary(dataStructure.KeyParameter)
                                            Dim isGlobal As String = Guid.NewGuid.ToString
                                            Dim listPolyline As List(Of Polyline) = New List(Of Polyline)
                                            For k As Integer = 0 To allPolygons.Count - 1
                                                Dim workPolygon As Polygon = allPolygons.Item(k)
                                                If workPolygon.Area >= 1 Then
                                                    Dim rezPoly As Polyline = New Polyline
                                                    rezPoly.SetDatabaseDefaults()
                                                    Dim coordPoly As List(Of Coordinate) = workPolygon.Coordinates.ToList
                                                    Dim arrayCoord As String(,) = {}
                                                    Dim countArrayCoord As Integer = 0
                                                    For k1 As Integer = 0 To coordPoly.Count - 1
                                                        Dim coodPoint As Coordinate = coordPoly.Item(k1)
                                                        ReDim Preserve arrayCoord(2, countArrayCoord)
                                                        arrayCoord(0, countArrayCoord) = k1 + 1
                                                        arrayCoord(1, countArrayCoord) = coodPoint.X
                                                        arrayCoord(2, countArrayCoord) = coodPoint.Y
                                                        countArrayCoord += 1
                                                    Next k1
                                                    rezPoly = CreateDWGObject.createPolylineFromArrayCoordinates(arrayCoord)
                                                    If IsNothing(rezPoly) = False Then
                                                        If rezPoly.Length > 0 Then
                                                            btr.AppendEntity(rezPoly)
                                                            tr.AddNewlyCreatedDBObject(rezPoly, True)
                                                            'styleObject.setStyle(db, rezPoly)
                                                            listPolyline.Add(rezPoly)
                                                        End If
                                                    End If
                                                End If
                                            Next k
                                            If listPolyline.Count > 0 Then
                                                Dim finalDict As Boolean = CadastralTools.DataFormatting.ReplaceDictionaryValues(dictionaryProperties, dictPropertiesParcel)
                                                For k As Integer = 0 To listPolyline.Count - 1
                                                    If dictionaryProperties.ContainsKey("ConditionatNumber") = True Then
                                                        If listPolyline.Count > 0 Then
                                                            Dim strNumber As String = startNumberParcels & "(" & k + 1 & ""
                                                            dictionaryProperties.Item("ConditionatNumber") = strNumber
                                                        Else
                                                            dictionaryProperties.Item("ConditionatNumber") = startNumberParcels
                                                        End If
                                                    End If
                                                    Dim strJSON As String = CadastralTools.JsonValidator.convertDictionaryToJson(dictionaryProperties)
                                                    Dim polyRec As Polyline = listPolyline.Item(k)
                                                    Dim structureData As CadastralTools.StructureElement = New StructureElement
                                                    structureData.Label = formParcelCreate.label
                                                    structureData.TablePS = formParcelCreate.tablePS
                                                    structureData.IdElement = Guid.NewGuid.ToString
                                                    structureData.NameClassObject = classObject.TERRITORIAL_PLANNING
                                                    structureData.Description = formParcelCreate.deskObject
                                                    structureData.IdObject = polyRec.Handle.ToString
                                                    If k > 0 Then
                                                        structureData.IdStructure = isGlobal
                                                    End If
                                                    structureData.KeyParameter = strJSON
                                                    structureData.TypeGeometry = typeGeometryObject.polygon
                                                    Dim recData As Boolean = FuncXRecords.writeXRecords(polyRec, structureData)
                                                Next k
                                                startNumberParcels += 1
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        Next
                    End If
                Else
                    tr.Commit()
                    Exit Do
                End If
                tr.Commit()
            End Using
        Loop
    End Sub
End Class
