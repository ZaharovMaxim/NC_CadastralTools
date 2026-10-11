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
Imports CadastralTools.StructureElement
Imports NetTopologySuite.Operation

Public Class ImportKPT
    <CommandMethod("ImportRosreestrKPT")>
    Public Sub ImportKPT()
        Dim doc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database 'Database
        Dim ed As Editor = doc.Editor 'Editor
        Dim categoryZn As String = "Земельный кадастр"
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim locker As DocumentLock = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Using locker
            Try
                'передаем масштаб чертежа в форму
                Dim nameKvartal As String = ""
                Dim FormKPTImport As CadastralTools.FormImportKPT = New CadastralTools.FormImportKPT
                Dim tempPath As String = Path.GetTempPath()
                tempPath = tempPath & "rosreestr"
                FormKPTImport.initDirectory = tempPath
                'запускаем форму
                FormKPTImport.ShowDialog()
                If FormKPTImport.boolShow = False Then
                    Exit Sub
                End If
                Dim templateXML As String = FormKPTImport.PutchTemlate
                If IsNothing(templateXML) = True Then
                    MsgBox("Шаблон оформления не найден!!! Макро прервано.")
                    Exit Sub
                End If
                If File.Exists(templateXML) = False Then
                    MsgBox("Шаблон оформления не найден!!! Макро прервано.")
                    Exit Sub
                End If

                Dim dictCadastralNumber As List(Of String) = New List(Of String)
                If FormKPTImport.boolImportXML = True Then
                    If Directory.Exists(FormKPTImport.initDirectory) = True Then
                        Try
                            tempPath = Path.Combine(FormKPTImport.initDirectory, Guid.NewGuid().ToString())
                            Directory.CreateDirectory(tempPath)
                        Catch ex As Exception
                            MsgBox("Не удалось создать папку для выходных данных.")
                            Exit Sub
                        End Try
                    Else
                        Exit Sub
                    End If
                    If FormKPTImport.boolDubleObject = True Then
                        dictCadastralNumber = FuncDatabase.getNumberCadastralObject()
                    End If
                    Try
                        If Directory.Exists(tempPath) = False Then
                            Directory.CreateDirectory(tempPath)
                        End If
                        If Directory.Exists(tempPath) = True Then
                            CadastralTools.ImportKPT.RosreestrImportKPT(tempPath, nameKvartal)
                        End If
                    Catch ex As System.Exception
                    End Try
                Else
                    tempPath = FormKPTImport.initDirectory
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim files As String() = Directory.GetFiles(tempPath)
                If files.Length = 0 Then
                    MsgBox("Не удалось корректно прочитать XML. Макро прервано.")
                    Exit Sub
                End If
                Dim classNetSuite As FuncTopologySuite = New FuncTopologySuite
                Dim cadastralNetSuit As CadastralTools.FuncNetTopologySuite = New CadastralTools.FuncNetTopologySuite
                'библиотека стилей отображения
                Dim dictonaryStyle As Dictionary(Of String, CadastralPolylineStyle) = New Dictionary(Of String, CadastralPolylineStyle)
                Using tr As Transaction = db.TransactionManager.StartTransaction()
                    Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                    Dim btr As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    '================================================================================================
                    'выбираем внешнюю границу для отсечки лишних объектов
                    Dim classTopo As FuncTopologySuite = New FuncTopologySuite
                    Dim boundPolygon As NetTopologySuite.Geometries.Polygon = Nothing
                    If FormKPTImport.boolSelectBound = True Then
                        Dim ptResult As PromptEntityResult = ed.GetEntity("Выберите внешнюю границу:")
                        If ptResult.Status = PromptStatus.OK Then
                            Dim ent As Entity = tr.GetObject(ptResult.ObjectId, OpenMode.ForRead)
                            If TypeOf ent Is Polyline Then
                                Dim boundPline As Polyline = TryCast(ent, Polyline)
                                If boundPline.Area > 0 Then
                                    boundPolygon = classTopo.CreatePolygon(boundPline)
                                End If
                            End If
                        End If
                    End If
                    Try
                        'читаем земельные участки
                        Dim classTopology As CadastralTools.FuncNetTopologySuite = New CadastralTools.FuncNetTopologySuite
                        For Each file As String In files
                            Dim listDictionaryData As List(Of Dictionary(Of String, Object)) = CadastralTools.JsonValidator.CreateObjectsFromJsonFile(file)
                            If listDictionaryData.Count > 0 Then
                                Dim dataCadastralObject As CadastralTools.StructureElement = New CadastralTools.StructureElement
                                Dim dictLandProperties As Dictionary(Of String, String) = New Dictionary(Of String, String)
                                dataCadastralObject.Label = "Земельный кадастр"
                                Dim idStructure As String = ""
                                Dim tableZn As String = ""
                                Dim PSTable As String = ""
                                Dim typeZn As String = "Полилиния"
                                Dim Label As String = ""
                                Dim idName As Integer = 0
                                Dim countOstrov As Integer = 1
                                Dim listPolygonCoordinates As List(Of String(,)) = New List(Of String(,))
                                For Each cadastralObject As Dictionary(Of String, Object) In listDictionaryData
                                    If IsNothing(cadastralObject) = True Then Continue For
                                    If cadastralObject.Count > 0 Then
                                        If cadastralObject.ContainsKey("PS") = True Then
                                            PSTable = cadastralObject.Item("PS")
                                        End If
                                        If cadastralObject.ContainsKey("LABEL") = True Then
                                            Label = cadastralObject.Item("LABEL")
                                        End If
                                        If cadastralObject.ContainsKey("NameTypeObject") = True Then
                                            idName = Val(cadastralObject.Item("NameTypeObject"))
                                        End If
                                        'читаем свойства объекта
                                        Dim boolReadDict As Boolean = CadastralTools.DataFormatting.ConvertToDictionaryOfString(cadastralObject, dictLandProperties)
                                        If cadastralObject.ContainsKey("ListCoord") Then
                                            Dim dataCoordParcel As Dictionary(Of String, Object) = cadastralObject("ListCoord")

                                            For Each kvp As KeyValuePair(Of String, Object) In dataCoordParcel
                                                Dim dataCoordCounter As List(Of Object) = TryCast(kvp.Value, List(Of Object))
                                                If dataCoordCounter Is Nothing OrElse dataCoordCounter.Count = 0 Then Continue For

                                                For Each coordItem As Object In dataCoordCounter
                                                    Dim listCoordinates As List(Of Object) = TryCast(coordItem, List(Of Object))
                                                    If listCoordinates Is Nothing OrElse listCoordinates.Count < 3 Then Continue For

                                                    ' Извлекаем списки координат (индексы 1 и 2)
                                                    Dim elementCoordX As List(Of Object) = TryCast(listCoordinates(1), List(Of Object))
                                                    Dim elementCoordY As List(Of Object) = TryCast(listCoordinates(2), List(Of Object))

                                                    If elementCoordX Is Nothing OrElse elementCoordY Is Nothing Then Continue For
                                                    If elementCoordX.Count = 0 OrElse elementCoordX.Count <> elementCoordY.Count Then Continue For

                                                    ' Создаём массив координат [Точки x Y]
                                                    Dim pointCount As Integer = elementCoordX.Count
                                                    Dim arrayCoord(pointCount - 1, 1) As String  ' [индекс точки, 0=X, 1=Y]

                                                    For k As Integer = 0 To pointCount - 1
                                                        arrayCoord(k, 0) = elementCoordX(k).ToString()
                                                        arrayCoord(k, 1) = elementCoordY(k).ToString()
                                                    Next k

                                                    If IsNothing(boundPolygon) = False Then
                                                        If boundPolygon.Area > 1 Then
                                                            If Not (PSTable Like "LAND_CADASTRAL_DIVISION") Then
                                                                If arrayCoord.GetUpperBound(1) = 0 Then
                                                                    'это точка
                                                                    Dim point As NetTopologySuite.Geometries.Point = New NetTopologySuite.Geometries.Point(arrayCoord(0, 0), arrayCoord(1, 0))
                                                                    Dim boolIntersect As Boolean = cadastralNetSuit.IsPointInsidePolygon(point, boundPolygon)
                                                                    If boolIntersect = False Then
                                                                        Continue For
                                                                    End If
                                                                Else
                                                                    Dim lineStr As NetTopologySuite.Geometries.LineString = cadastralNetSuit.CreateLineString(arrayCoord)
                                                                    Dim boolIntersect As Boolean = cadastralNetSuit.DoesLineIntersectPolygon(lineStr, boundPolygon)
                                                                    If boolIntersect = False Then
                                                                        Continue For
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                    listPolygonCoordinates.Add(arrayCoord)
                                                Next coordItem
                                            Next kvp
                                        End If
                                    End If
                                Next
                                'преобразуйм в единую строку
                                Dim keyParam As String = CadastralTools.JsonValidator.convertDictionaryToJson(dictLandProperties)
                                '==========================================================================================================
                                'включена опция проверки дубликатов
                                Dim cadNumber As String = ""
                                Dim cadObjectNumber As Object = CadastralTools.JsonValidator.GetValueFromJson(keyParam, "Cad_number")
                                If TypeOf cadObjectNumber Is String Then
                                    cadNumber = cadObjectNumber
                                Else
                                    cadObjectNumber = CadastralTools.JsonValidator.GetValueFromJson(keyParam, "Reg_numb_border")
                                    If TypeOf cadObjectNumber Is String Then
                                        cadNumber = cadObjectNumber
                                    End If
                                End If
                                If dictCadastralNumber.Count > 0 Then
                                    If dictCadastralNumber.Contains(cadNumber) = True Then
                                        Continue For
                                    End If
                                End If
                                '==========================================================================================================
                                'ФОРМИРУЕМ СЕМАНТИКУ
                                dataCadastralObject.NameTypeObject = CadastralTools.StructureElement.GetEnumByValue(GetType(nameObject), idName)
                                dataCadastralObject.Description = CadastralTools.StructureElement.GetEnumDescriptionByValue(GetType(nameObject), idName)
                                dataCadastralObject.NameClassObject = CadastralTools.StructureElement.classObject.LAND_CADASTRAL
                                dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polygon
                                dataCadastralObject.KeyParameter = keyParam
                                dataCadastralObject.IdElement = Guid.NewGuid.ToString
                                dataCadastralObject.TablePS = PSTable
                                Select Case PSTable
                                    Case "LAND_CADASTRAL_DIVISION"
                                        tableZn = "ЕГРН-Кадастровое деление"
                                    Case "LAND_CADASTRAL_PARCEL"
                                        tableZn = "ЕГРН-Кадастровые участки"
                                    Case "LAND_CADASTRAL_PART"
                                        tableZn = "ЕГРН-Кадастровые участки"
                                    Case "LAND_CADASTRAL_BUILDING"
                                        tableZn = "ЕГРН-Здания (строения)"
                                    Case "LAND_CADASTRAL_CONSTRUCTION"
                                        dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.line
                                        tableZn = "ЕГРН-Сооружения"
                                    Case "LAND_CADASTRAL_UNFINISHED_CONSTRUCTION"
                                        dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.line
                                        tableZn = "ЕГРН-Объекты незавершенного строительства"
                                    Case "LAND_CADASTRAL_ZONE"
                                        tableZn = "ЕГРН-Территориальные зоны"
                                    Case "LAND_CADASTRAL_BOUNDARY"
                                        tableZn = "ЕГРН-Границы территорий"
                                        If dataCadastralObject.NameTypeObject = nameObject.WaterBoundary Then
                                            dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.line
                                        End If
                                    Case "LAND_CADASTRAL_REDLINE"
                                        tableZn = "ЕГРН-Границы территорий"
                                        dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.line
                                    Case "LAND_CADASTRAL_COASTLINE"
                                        tableZn = "ЕГРН-Границы территорий"
                                        dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.line
                                End Select
                                '=========================================================================================================================
                                'назначаем стиль
                                Dim newStyle As CadastralPolylineStyle = New CadastralPolylineStyle
                                newStyle.TemplateXML = templateXML
                                newStyle.DBDrawing = db
                                If dictonaryStyle.ContainsKey(Label) = False Then
                                    Dim arrayProperties As String(,) = {}
                                    Dim boolFindPropertyObject As Boolean = CadastralTools.TemplateXML.FuncReadPropertiesObjectToXML(templateXML, categoryZn, tableZn, typeZn, dataCadastralObject.Description, arrayProperties)
                                    If IsArray(arrayProperties) = True Then
                                        newStyle.createLineStyle(arrayProperties)
                                    End If
                                    dictonaryStyle.Add(Label, newStyle)
                                Else
                                    newStyle = dictonaryStyle.Item(Label)
                                End If
                                '=============================================================================================================================
                                'рисуем контур
                                If listPolygonCoordinates.Count > 0 Then
                                    Dim listPolygon As Dictionary(Of Object, NetTopologySuite.Geometries.Polygon) = New Dictionary(Of Object, NetTopologySuite.Geometries.Polygon)
                                    For i As Integer = 0 To listPolygonCoordinates.Count - 1
                                        Dim arrayCoord As String(,) = listPolygonCoordinates.ElementAt(i)
                                        If arrayCoord.Length > 2 Then
                                            Dim acPolyline As Polyline = CreateDWGObject.createPolylineFromArrayCoordinates2(arrayCoord)
                                            If IsNothing(acPolyline) = False Then
                                                btr.AppendEntity(acPolyline)
                                                tr.AddNewlyCreatedDBObject(acPolyline, True)
                                                newStyle.setStyle(acPolyline)
                                                If dataCadastralObject.TypeGeometry = typeGeometryObject.polygon Then
                                                    acPolyline.Closed = True
                                                    Dim sourcePolygon As NetTopologySuite.Geometries.Polygon = classNetSuite.CreatePolygon(acPolyline)
                                                    listPolygon.Add(acPolyline, sourcePolygon)
                                                    If listPolygon.Count > 1 Then
                                                        Dim boolFindOstrov As Boolean = False
                                                        For j As Integer = 0 To listPolygon.Count - 2
                                                            Dim destPolygon As NetTopologySuite.Geometries.Polygon = listPolygon.ElementAt(j).Value
                                                            If destPolygon.Area > sourcePolygon.Area Then
                                                                If cadastralNetSuit.IsPolygonInside(destPolygon, sourcePolygon) = True Then
                                                                    boolFindOstrov = True
                                                                    Exit For
                                                                ElseIf cadastralNetSuit.IsPolygonInside(sourcePolygon, destPolygon) = True Then
                                                                    boolFindOstrov = True
                                                                    Exit For
                                                                End If
                                                            End If
                                                        Next j
                                                        If boolFindOstrov = True Then
                                                            Dim objOstrov As CadastralTools.StructureElement = New CadastralTools.StructureElement
                                                            objOstrov.Label = dataCadastralObject.Label
                                                            objOstrov.NameClassObject = CadastralTools.StructureElement.classObject.LAND_CADASTRAL
                                                            dataCadastralObject.TypeGeometry = CadastralTools.StructureElement.typeGeometryObject.polygon
                                                            objOstrov.NameTypeObject = dataCadastralObject.NameTypeObject
                                                            objOstrov.Description = "Внутренний островок"
                                                            objOstrov.IdStructure = dataCadastralObject.IdElement
                                                            objOstrov.IdElement = Guid.NewGuid.ToString
                                                            objOstrov.TablePS = "CADASTRAL_OSTROV"
                                                            Dim dictOstrov As Dictionary(Of String, String) = New Dictionary(Of String, String)
                                                            dictOstrov.Add("ConditionalNumber", countOstrov)
                                                            dictOstrov.Add("ParentNumber", cadNumber)
                                                            Dim keyParamOstrov As String = CadastralTools.JsonValidator.convertDictionaryToJson(dictOstrov)
                                                            objOstrov.KeyParameter = keyParamOstrov
                                                            Dim boolRecOstrov As Boolean = FuncXRecords.writeXRecords(acPolyline, objOstrov)
                                                            countOstrov += 1
                                                            Continue For
                                                        End If
                                                    End If
                                                End If
                                                Dim boolRec As Boolean = FuncXRecords.writeXRecords(acPolyline, dataCadastralObject)
                                            End If
                                        Else
                                            Dim acCircle As Circle = CreateDWGObject.createCircleFromArrayCoordinates2(arrayCoord)
                                            btr.AppendEntity(acCircle)
                                            tr.AddNewlyCreatedDBObject(acCircle, True)
                                            newStyle.setStyle(acCircle)
                                            dataCadastralObject.TypeGeometry = typeGeometryObject.block
                                            Dim boolRec As Boolean = FuncXRecords.writeXRecords(acCircle, dataCadastralObject)
                                            Continue For
                                        End If
                                    Next i
                                End If
                            End If
                        Next
                    Catch ex As Teigha.Runtime.Exception
                    End Try
                    tr.Commit()
                End Using
                If FormKPTImport.boolImportXML = True Then
                    If nameKvartal.Trim.Length > 0 Then
                        Dim boolKVName As Boolean = CadastralTools.SupportFiles.RenameDirectory(tempPath, nameKvartal)
                    End If
                End If

            Catch ex As Exception
            End Try
        End Using
    End Sub
End Class
