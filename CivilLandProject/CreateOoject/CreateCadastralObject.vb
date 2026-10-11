Imports CadastralTools
Imports HostMgd.ApplicationServices
Imports HostMgd.EditorInput
Imports NetTopologySuite.Geometries
Imports NetTopologySuite.LinearReferencing
Imports Teigha.DatabaseServices
Imports Teigha.Geometry
Imports Teigha.Runtime

Public Class CreateCadastralObject
    'создать проектируемый участок (новая версия)
    <CommandMethod("CreateCadastralObject")>
    Public Sub CreateProjectParcel()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'создаем класс с новой формой
        Dim acDoc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument ' открываем базу данных чертежа
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Try
            Dim strDWGName As String = acDoc.Name
            Dim objBase As Object = Application.GetSystemVariable("DWGTITLED")
            If System.Convert.ToInt16(objBase) > 0 Then
                Dim obj1 As Object = Application.GetSystemVariable("DBMOD")
                If Not (System.Convert.ToInt16(obj1) = 0) Then
                    acDoc.Database.SaveAs(strDWGName, True, DwgVersion.Current, acDoc.Database.SecurityParameters)
                End If
            End If
        Catch ex As Teigha.Runtime.Exception
        End Try
        'Dim GeoParcel As FormCreateParcels = New FormCreateParcels()
        'GeoParcel.ShowDialog()
        '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'вызываем диалоговое окно с первоначальными настройками
        Dim generalDidectory As String = "k:\Civil3DToolsUtility"
        Dim formCreateParcels As CadastralTools.FormCreateCadastralObject = New CadastralTools.FormCreateCadastralObject(generalDidectory)
        formCreateParcels.ShowDialog()
        If formCreateParcels.boolShow = False Then
            Exit Sub
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'читаем префикс участка
        Dim numberParcel As Integer = formCreateParcels.numberParcel
        Dim numberCounter As Integer = formCreateParcels.numberCounter
        Dim fileTemlateXML1 As String = formCreateParcels.fileTemlateXML
        Dim fileTemplateDWG As String = fileTemlateXML1.Replace(".xml", ".dwg")
        Dim dictionaryPropertiesObject As Dictionary(Of String, String) = formCreateParcels.resultDictionary
        Dim arrayPropertyObject As String(,) = formCreateParcels.arrayPropertiesZn
        Dim arrayPropertiesLeader As String(,) = formCreateParcels.arrayPropertiesLeader
        Dim nameObject As CadastralTools.StructureElement.nameObject = formCreateParcels.nameTypeObject
        Dim typeGeometryObject As StructureElement.typeGeometryObject = formCreateParcels.typeGeometry
        Dim deskObject As String = formCreateParcels.nameObject
        Dim deskLeader As String = formCreateParcels.deskLeader
        Dim nameTableObject As String = formCreateParcels.nameTableZn
        Dim globalID As String = ""
        If dictionaryPropertiesObject.ContainsKey("GLOBALID") = True Then
            globalID = dictionaryPropertiesObject.Item("GLOBALID")
        End If
        If dictionaryPropertiesObject.Count > 0 Then
            If arrayPropertyObject.Length > 0 Then
                '=========================================================================================================================
                'назначаем стиль объекту
                Dim styleObject As CadastralPolylineStyle = New CadastralPolylineStyle()
                styleObject.DBDrawing = db
                styleObject.TemplateXML = fileTemlateXML1
                styleObject.TemplateDWG = fileTemplateDWG
                Dim arrayProperties As String(,) = {}
                styleObject.createLineStyle(arrayPropertyObject)
                '===========================================================================================================================
                Dim formatLeader As String = CadastralTools.MathFunction.findValueToArray2d(arrayPropertiesLeader, "NOTE")
                Dim cadastralNetSuit As CadastralTools.FuncNetTopologySuite = New CadastralTools.FuncNetTopologySuite
                Dim localClassTopology As FuncTopologySuite = New FuncTopologySuite()
                Dim classDrawObject As CreateDWGObject = New CreateDWGObject(db)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'выбираем объекты
                Dim tr As Transaction = db.TransactionManager.StartTransaction()
                Using tr
                    If typeGeometryObject = StructureElement.typeGeometryObject.polygon Or typeGeometryObject = StructureElement.typeGeometryObject.polyline Then
                        Dim acSSPrompt As PromptSelectionResult
                        Dim acSSet1 As SelectionSet = Nothing
                        acSSPrompt = ed.GetSelection() 'выбираем контура
                        If acSSPrompt.Status = PromptStatus.OK Then
                            acSSet1 = acSSPrompt.Value
                        Else
                            Exit Sub
                        End If
                        For Each objId1 As ObjectId In acSSet1.GetObjectIds()
                            Dim ent As Teigha.DatabaseServices.Entity = tr.GetObject(objId1, OpenMode.ForWrite)
                            If TypeOf ent Is Polyline Then
                                Dim plineParcels As Polyline = ent
                                'точка вставки мультивыноски
                                Dim innerPoint As NetTopologySuite.Geometries.Point = Nothing
                                'обрабатываем полигон
                                If typeGeometryObject = StructureElement.typeGeometryObject.polygon Then
                                    If plineParcels.NumberOfVertices > 2 Then
                                        Dim polygon As NetTopologySuite.Geometries.Polygon = localClassTopology.CreatePolygon(plineParcels)
                                        If IsNothing(polygon) = False Then
                                            polygon = CadastralTools.FuncNetTopologySuite.RemoveDuplicateVertices(polygon, 0.01)
                                            polygon = CadastralTools.FuncNetTopologySuite.ReversePolygonPreserveOrientation(polygon)
                                            plineParcels = classDrawObject.createPolylineByPolygon(polygon, plineParcels)
                                        End If
                                        'ищем внутреннюю точку
                                        innerPoint = cadastralNetSuit.GetInteriorPoint(polygon)
                                    End If
                                    If dictionaryPropertiesObject.ContainsKey("Area") = True Then
                                        dictionaryPropertiesObject.Item("Area") = Math.Round(plineParcels.Area, 0)
                                    End If
                                    'обрабатываем не замкнутую полилинию
                                ElseIf typeGeometryObject = StructureElement.typeGeometryObject.polyline Then
                                    If plineParcels IsNot Nothing AndAlso plineParcels.NumberOfVertices > 1 Then
                                        Dim cadastralLine As NetTopologySuite.Geometries.LineString = localClassTopology.CreateLineString(plineParcels)
                                        If cadastralLine IsNot Nothing AndAlso Not cadastralLine.IsEmpty AndAlso cadastralLine.Length > 0 Then
                                            ' Создаем полилинию (если нужно)
                                            plineParcels = classDrawObject.createPolylineByLineString(cadastralLine, plineParcels)
                                            ' Ищем внутреннюю точку на ГРАНИЦЕ (используем cadastralLine, а не plineParcels)
                                            Dim locationIndexedLine As New LocationIndexedLine(cadastralLine)
                                            ' Вычисляем половину длины от cadastralLine
                                            Dim totalLength As Double = cadastralLine.Length / 2.0
                                            ' Создаем "карту" для преобразования расстояния в LinearLocation
                                            Dim lengthLocationMap As New LengthLocationMap(cadastralLine)
                                            ' Преобразуем длину в LinearLocation
                                            Dim targetLocation As LinearLocation = lengthLocationMap.GetLocation(totalLength)
                                            ' Извлекаем координату
                                            Dim midCoordinate As Coordinate = locationIndexedLine.ExtractPoint(targetLocation)
                                            innerPoint = New NetTopologySuite.Geometries.Point(midCoordinate)
                                        End If
                                    End If
                                    If dictionaryPropertiesObject.ContainsKey("Lenght") = True Then
                                        dictionaryPropertiesObject.Item("Lenght") = Math.Round(plineParcels.Length, 2)
                                    End If
                                End If
                                'обновляем свойства
                                Dim conditionalParcel As String = numberParcel
                                If numberCounter > 0 Then
                                    conditionalParcel = numberParcel & "(" & numberCounter & ")"
                                    numberCounter += 1
                                Else
                                    numberParcel += 1
                                End If
                                'обновляем условный номер и локальный ID
                                If dictionaryPropertiesObject.ContainsKey("ConditionalNumber") = True Then
                                    dictionaryPropertiesObject.Item("ConditionalNumber") = conditionalParcel
                                End If
                                If dictionaryPropertiesObject.ContainsKey("LOCALID") = True Then
                                    dictionaryPropertiesObject.Item("LOCALID") = Guid.NewGuid.ToString
                                End If
                                'формируем свойства объекта
                                Dim keyParamObject As String = CadastralTools.JsonValidator.convertDictionaryToJson(dictionaryPropertiesObject)
                                Dim dataCadastralObject As CadastralTools.StructureElement = New CadastralTools.StructureElement
                                dataCadastralObject.Label = formCreateParcels.nameCategory
                                dataCadastralObject.NameClassObject = CadastralTools.StructureElement.classObject.TERRITORIAL_PLANNING
                                dataCadastralObject.TypeGeometry = formCreateParcels.typeGeometry
                                dataCadastralObject.NameTypeObject = nameObject
                                dataCadastralObject.TableObject = nameTableObject
                                dataCadastralObject.Description = deskObject
                                dataCadastralObject.KeyParameter = keyParamObject
                                dataCadastralObject.TablePS = formCreateParcels.PS
                                dataCadastralObject.IdStructure = globalID
                                dataCadastralObject.IdElement = Guid.NewGuid.ToString
                                dataCadastralObject.IdObject = plineParcels.Handle.ToString
                                Dim boolWriteDataOstriv As Boolean = FuncXRecords.writeXRecords(plineParcels, dataCadastralObject)
                                styleObject.setStyle(plineParcels)
                                '=========================================================================================================================
                                'рисуем выноску
                                '=========================================================================================================================
                                If arrayPropertiesLeader.Length > 0 Then
                                    'назначаем стиль выноске
                                    Dim styleLeaderObject As CadastralLeaderStyle = New CadastralLeaderStyle()
                                    styleLeaderObject.TemplateXML = fileTemlateXML1
                                    styleLeaderObject.TemplateDWG = fileTemplateDWG
                                    styleLeaderObject.DBDrawing = db
                                    '=======================================================================================================================
                                    Dim tempLayerLeader As String = CadastralTools.MathFunction.findValueToArray2d(arrayPropertiesLeader, "Layer")
                                    If Mid(tempLayerLeader.Trim, 1, 1) Like "_" Then
                                        Dim boolCreateLeaderStyle As Boolean = styleLeaderObject.createStyleLeader(arrayPropertiesLeader, plineParcels.Layer)
                                    Else
                                        Dim boolCreateLeaderStyle As Boolean = styleLeaderObject.createStyleLeader(arrayPropertiesLeader)
                                    End If
                                    '===============================================================
                                    'вторая точка
                                    Dim pointBound As Extents3d = plineParcels.Bounds
                                    Dim maskLabel As String = ""
                                    Dim strNoteLeader As String = CadastralTools.MathFunction.findValueToArray2d(arrayPropertiesLeader, "Note")
                                    If strNoteLeader.Trim.Length > 0 Then
                                        conditionalParcel = CadastralTools.DataFormatting.getTextLeader(strNoteLeader, dictionaryPropertiesObject, maskLabel)
                                        If conditionalParcel.Trim.Length > 0 Then
                                            Dim leaderObject As MLeader = classDrawObject.createLeader(New Point3d(innerPoint.X, innerPoint.Y, 0), conditionalParcel, pointBound.MaxPoint)
                                            If IsNothing(leaderObject) = False Then
                                                styleLeaderObject.setLeaderStyle(leaderObject)
                                                Dim createXdata As StructureElement = CadastralTools.CadastralLeader.getLeaderXData(dataCadastralObject, deskLeader, maskLabel, fileTemlateXML1)
                                                Dim boolWriteDataLeader As Boolean = FuncXRecords.writeXRecords(leaderObject, createXdata)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    ElseIf typeGeometryObject = StructureElement.typeGeometryObject.block Then
                        'загружаем в чертеж блок
                        'CadCursors.GetPoint(userCadView,, "Укажите точку вставки блока: ")
                        'userCadView.SelectionSet.SelectObjectsAtScreen(Function(obj) TypeOf obj Is DwgInsert, "Выборите блоки для преобразования: ")
                    End If
                    tr.Commit()
                End Using
            End If
        End If
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'создание внутренних островков
    <CommandMethod("AddCadastralObjectRing")>
    Public Sub CreateCadastralRings()
        Dim acDoc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Try
            Dim strDWGName As String = acDoc.Name
            Dim objBase As Object = Application.GetSystemVariable("DWGTITLED")
            If System.Convert.ToInt16(objBase) > 0 Then
                Dim obj1 As Object = Application.GetSystemVariable("DBMOD")
                If Not (System.Convert.ToInt16(obj1) = 0) Then
                    acDoc.Database.SaveAs(strDWGName, True, DwgVersion.Current, acDoc.Database.SecurityParameters)
                End If
            End If
        Catch ex As Teigha.Runtime.Exception
        End Try
        '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim generalDir As String = ""
        Dim templateXML As String = ""
        Dim templateDWG As String = ""
        Dim arrayDirSupport As String() = Nothing
        Dim boolFindDirSupport As Boolean = FuncFiles.FuncGetSupportFolderPath(arrayDirSupport)
        If IsArray(arrayDirSupport) = True Then
            For i As Integer = 0 To arrayDirSupport.Length - 1
                Dim nameFolder As String = arrayDirSupport(i)
                Dim pos As Integer = nameFolder.IndexOf("Civil3DToolsUtility")
                If pos > -1 Then
                    generalDir = Mid(nameFolder, 1, pos)
                    generalDir = generalDir & "Civil3DToolsUtility"
                    Exit For
                End If
            Next
        End If
        'шаблон оформления
        Dim dictionaryTemplate As Dictionary(Of String, String) = CadastralTools.TemplateXML.readTemplatesXML(generalDir)
        If dictionaryTemplate.Count > 0 Then
            templateXML = dictionaryTemplate.ElementAt(0).Value
        End If
        If IsNothing(templateXML) = True Or templateXML.Trim.Length = 0 Then
            MsgBox("Файл с шаблонами =Стандарт_ADSK.xml= не найден. Пожалуйста, добавьте его в пути доступа к вспомогательным файлам!!!")
            Exit Sub
        End If


        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim per As PromptEntityResult = ed.GetEntity("Выберите контур к которому будет добавлен островок: ")
            'выбор полигона которому принадлежит этот островок
            Dim selectParentObject As Teigha.DatabaseServices.Entity = tr.GetObject(per.ObjectId, OpenMode.ForWrite) 'ссылка на объект
            If TypeOf selectParentObject Is Polyline Then
                Dim parentPolyline As Polyline = selectParentObject
                Dim dataParent As StructureElement = Nothing
                Dim boolReadData As Boolean = FuncXRecords.getXRecords(parentPolyline, dataParent)
                If IsNothing(dataParent) = True Then
                    MsgBox("Выберите объект с присвоенными кадастрвыми данными")
                    Exit Sub
                End If
                'выбираем объекты внутри выбранного контура
                Dim dictionaryCadastralObject As Dictionary(Of StructureElement.nameObject, List(Of NetTopologySuite.Geometries.Geometry)) = FuncSelectedObject.selectAllObject(db)
                Dim cadastralNetTopoSuite As CadastralTools.FuncNetTopologySuite = New CadastralTools.FuncNetTopologySuite
                Dim acadNetTopoSuite As FuncTopologySuite = New FuncTopologySuite()
                Dim classDrawObject As CreateDWGObject = New CreateDWGObject(db)
                Dim lineStyle As CadastralPolylineStyle = New CadastralPolylineStyle(parentPolyline)
                lineStyle.DBDrawing = db
                If IsNothing(dataParent) = False And IsNothing(parentPolyline) = False Then
                    Dim parentID As String = dataParent.IdElement
                    Dim keyParam As String = dataParent.KeyParameter
                    Dim parentPolygon As NetTopologySuite.Geometries.Polygon = acadNetTopoSuite.CreatePolygon(parentPolyline)
                    Dim arrayPropertiesCadastralOstrov As String(,) = {}
                    Dim boolFindPropertiesOstrov As Boolean = CadastralTools.TemplateXML.FuncReadTablePSToXML(templateXML, "CADASTRAL_OSTROV", arrayPropertiesCadastralOstrov)
                    Dim numberOstrov As Integer = CadastralTools.CadastralOstrov.getLastNumber(parentPolygon, dictionaryCadastralObject)
                    numberOstrov += 1
                    Dim listPolygon As List(Of NetTopologySuite.Geometries.Geometry) = Nothing
                    Dim listOstrov As List(Of NetTopologySuite.Geometries.Geometry) = Nothing
                    If dictionaryCadastralObject.ContainsKey(dataParent.NameTypeObject) = True Then
                        listPolygon = dictionaryCadastralObject.Item(dataParent.NameTypeObject)
                    End If
                    If dictionaryCadastralObject.ContainsKey(StructureElement.nameObject.ProjectOstrov) = True Then
                        listOstrov = dictionaryCadastralObject.Item(StructureElement.nameObject.ProjectOstrov)
                    End If
                    parentPolygon = CadastralTools.CadastralPolygon.getCounter(parentPolygon, listPolygon, listOstrov)
                    Dim areaParcel As Double = parentPolygon.Area
                    If arrayPropertiesCadastralOstrov.Length > 0 Then
                        'список полигонов островков для их упорядочивания
                        Dim listPolygonRing As List(Of NetTopologySuite.Geometries.Polygon) = New List(Of NetTopologySuite.Geometries.Polygon)
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'выбираем островки
                        Dim acSSPrompt As PromptSelectionResult
                        Dim acSSet1 As SelectionSet = Nothing
                        acSSPrompt = ed.GetSelection() 'выбираем контура
                        If acSSPrompt.Status = PromptStatus.OK Then
                            acSSet1 = acSSPrompt.Value
                        Else
                            MsgBox("Прервано пользователем.")
                            Exit Sub
                        End If
                        For Each acSSObj As ObjectId In acSSet1.GetObjectIds()
                            If Not IsDBNull(acSSObj) Then
                                Dim acEnt As Teigha.DatabaseServices.Entity = tr.GetObject(acSSObj, OpenMode.ForWrite)
                                If TypeOf (acEnt) Is Polyline Then
                                    Dim plineOstrov As Polyline = acEnt
                                    If plineOstrov.NumberOfVertices > 2 Then
                                        If plineOstrov.Area > 0 Then
                                            Dim dataObject As StructureElement = Nothing
                                            Dim boolReadDataObject As Boolean = FuncXRecords.getXRecords(plineOstrov, dataObject)
                                            If boolReadDataObject = True Then
                                                plineOstrov = plineOstrov.Clone
                                                acBlkTblRec.AppendEntity(plineOstrov)
                                                tr.AddNewlyCreatedDBObject(plineOstrov, True)
                                            End If
                                        End If
                                        Dim ostrovPolygon As NetTopologySuite.Geometries.Polygon = acadNetTopoSuite.CreatePolygon(plineOstrov)
                                        ostrovPolygon = cadastralNetTopoSuite.RemoveDuplicateVertices(ostrovPolygon, 0.01)
                                        ostrovPolygon = cadastralNetTopoSuite.ReversePolygonPreserveOrientation(ostrovPolygon)
                                        If ostrovPolygon.Coordinates.Count > 2 Then
                                            If ostrovPolygon.Area > 0 Then
                                                ostrovPolygon.UserData = plineOstrov.ObjectId
                                                listPolygonRing.Add(ostrovPolygon)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        If listPolygonRing.Count > 2 Then
                            'упорядочиваем полигоны
                            listPolygonRing = cadastralNetTopoSuite.SortByNearestNeighbor(listPolygonRing)
                        End If
                        For Each userPolygon As NetTopologySuite.Geometries.Polygon In listPolygonRing
                            Dim idPolyline As ObjectId = userPolygon.UserData
                            Dim polylineRing As Polyline = tr.GetObject(idPolyline, OpenMode.ForWrite)
                            If IsNothing(polylineRing) = False Then
                                polylineRing = classDrawObject.createPolylineByPolygon(userPolygon, polylineRing)
                                'формируем свойства объекта
                                Dim dataCadastralObject As CadastralTools.StructureElement = CadastralTools.CadastralOstrov.getProjectOstrovXData(dataParent, numberOstrov, arrayPropertiesCadastralOstrov)
                                dataCadastralObject.IdObject = polylineRing.Handle.ToString
                                Dim boolWriteDataOstriv As Boolean = FuncXRecords.writeXRecords(polylineRing, dataCadastralObject, CadastralTools.StructureElement.tableXRecords.PROJECT_STRUCTURES)
                                lineStyle.setStyle(polylineRing)
                                parentPolygon = cadastralNetTopoSuite.AddPolygonAsHole(parentPolygon, userPolygon)
                                numberOstrov += 1
                            End If
                        Next
                        Dim rezArea As Double = Math.Round(parentPolygon.Area, 0)
                        keyParam = CadastralTools.JsonValidator.SetJsonFieldValue(keyParam, "Area", rezArea)
                        dataParent.KeyParameter = keyParam
                        Dim boolNewData As Boolean = FuncXRecords.writeXRecords(parentPolyline, dataParent, CadastralTools.StructureElement.tableXRecords.PROJECT_STRUCTURES)
                    End If
                End If
            End If
            tr.Commit()
        End Using
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'добавить контур
    <CommandMethod("AddCadastralObjectCounter")>
    Public Sub AddUserKounterObject()
        Dim acDoc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Try
            Dim strDWGName As String = acDoc.Name
            Dim objBase As Object = Application.GetSystemVariable("DWGTITLED")
            If System.Convert.ToInt16(objBase) > 0 Then
                Dim obj1 As Object = Application.GetSystemVariable("DBMOD")
                If Not (System.Convert.ToInt16(obj1) = 0) Then
                    acDoc.Database.SaveAs(strDWGName, True, DwgVersion.Current, acDoc.Database.SecurityParameters)
                End If
            End If
        Catch ex As Teigha.Runtime.Exception
        End Try

        '1.\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim generalDir As String = ""
        Dim templateXML As String = ""
        Dim templateDWG As String = ""
        Dim arrayDirSupport As String() = Nothing
        Dim boolFindDirSupport As Boolean = FuncFiles.FuncGetSupportFolderPath(arrayDirSupport)
        If IsArray(arrayDirSupport) = True Then
            For i As Integer = 0 To arrayDirSupport.Length - 1
                Dim nameFolder As String = arrayDirSupport(i)
                Dim pos As Integer = nameFolder.IndexOf("Civil3DToolsUtility")
                If pos > -1 Then
                    generalDir = Mid(nameFolder, 1, pos)
                    generalDir = generalDir & "Civil3DToolsUtility"
                    Exit For
                End If
            Next
        End If
        'шаблон оформления
        Dim dictionaryTemplate As Dictionary(Of String, String) = CadastralTools.TemplateXML.readTemplatesXML(generalDir)
        If dictionaryTemplate.Count > 0 Then
            templateXML = dictionaryTemplate.ElementAt(0).Value
        End If
        If IsNothing(templateXML) = True Or templateXML.Trim.Length = 0 Then
            MsgBox("Файл с шаблонами =Стандарт_ADSK.xml= не найден. Пожалуйста, добавьте его в пути доступа к вспомогательным файлам!!!")
            Exit Sub
        End If
        'сброс выделения
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
            Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            Dim per As PromptEntityResult = ed.GetEntity("Выберите объект к которому будеь добавлен контур: ")
            'выбор полигона которому принадлежит этот островок
            Dim selectParentObject As Teigha.DatabaseServices.Entity = tr.GetObject(per.ObjectId, OpenMode.ForWrite) 'ссылка на объект
            If TypeOf selectParentObject Is Polyline Then
                Dim parentPolyline As Polyline = selectParentObject
                Dim dataParentParcel As StructureElement = Nothing
                Dim boolReadData As Boolean = FuncXRecords.getXRecords(parentPolyline, dataParentParcel)
                If IsNothing(dataParentParcel) = True Then
                    MsgBox("Выберите объект с присвоенными кадастрвыми данными")
                    Exit Sub
                End If
                'выбираем объекты внутри выбранного контура
                Dim dictionaryCadastralObject As Dictionary(Of StructureElement.nameObject, List(Of NetTopologySuite.Geometries.Geometry)) = FuncSelectedObject.selectAllObject(db)
                Dim cadastralNetTopoSuite As CadastralTools.FuncNetTopologySuite = New CadastralTools.FuncNetTopologySuite
                Dim acadNetTopoSuite As FuncTopologySuite = New FuncTopologySuite()
                Dim classDrawObject As CreateDWGObject = New CreateDWGObject(db)
                Dim lineStyle As CadastralPolylineStyle = New CadastralPolylineStyle(parentPolyline)
                '=================================================================================================================================
                Dim numberParcel As Integer = 1
                Dim counterParcel As Integer = 1
                If IsNothing(dataParentParcel) = False And IsNothing(parentPolyline) = False Then
                    Dim parentID As String = dataParentParcel.IdElement
                    Dim keyParam As String = dataParentParcel.KeyParameter
                    Dim dictionaryDataParent As Dictionary(Of String, Object) = CadastralTools.JsonValidator.DeserializeJsonToDictionary(keyParam)
                    If dictionaryDataParent.Count > 0 Then
                        If dictionaryDataParent.ContainsKey("ConditionalNumber") = True Then
                            Dim conditionalParentParcel As String = dictionaryDataParent.Item("ConditionalNumber")
                            numberParcel = CadastralTools.DataFormatting.GetNumberBeforeParenthesis(conditionalParentParcel)
                            If numberParcel = 0 Then numberParcel = 1
                            counterParcel = CadastralTools.DataFormatting.GetNumberInParenthesis(conditionalParentParcel)
                            If counterParcel = 0 Then
                                counterParcel = 1
                                conditionalParentParcel = numberParcel & "(" & counterParcel & ")"
                                keyParam = CadastralTools.JsonValidator.SetJsonFieldValue(keyParam, "ConditionalNumber", conditionalParentParcel)
                                dataParentParcel.KeyParameter = keyParam
                                If dataParentParcel.IdStructure.Length < 2 Then
                                    dataParentParcel.IdStructure = Guid.NewGuid.ToString
                                End If
                                Dim boolRecData As Boolean = FuncXRecords.writeXRecords(parentPolyline, dataParentParcel, StructureElement.tableXRecords.PROJECT_STRUCTURES)
                            End If
                        End If

                        'список полигонов островков для их упорядочивания
                        Dim listCounter As List(Of NetTopologySuite.Geometries.Geometry) = New List(Of NetTopologySuite.Geometries.Geometry)
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'выбираем контура
                        Dim acSSPrompt As PromptSelectionResult
                        Dim acSSet1 As SelectionSet = Nothing
                        acSSPrompt = ed.GetSelection() 'выбираем контура
                        If acSSPrompt.Status = PromptStatus.OK Then
                            acSSet1 = acSSPrompt.Value
                        Else
                            MsgBox("Прервано пользователем.")
                            Exit Sub
                        End If
                        For Each acSSObj As ObjectId In acSSet1.GetObjectIds()
                            If Not IsDBNull(acSSObj) Then
                                Dim acEnt As Teigha.DatabaseServices.Entity = tr.GetObject(acSSObj, OpenMode.ForWrite)
                                If TypeOf (acEnt) Is Polyline Then
                                    Dim plineCounter As Polyline = acEnt
                                    If dataParentParcel.TypeGeometry = StructureElement.typeGeometryObject.polygon Then
                                        If plineCounter.NumberOfVertices > 2 Then
                                            If plineCounter.Area > 0 Then
                                                Dim counterPolygon As NetTopologySuite.Geometries.Polygon = acadNetTopoSuite.CreatePolygon(plineCounter)
                                                counterPolygon = cadastralNetTopoSuite.RemoveDuplicateVertices(counterPolygon, 0.01)
                                                counterPolygon = cadastralNetTopoSuite.ReversePolygonPreserveOrientation(counterPolygon)
                                                If counterPolygon.Coordinates.Count > 2 Then
                                                    If counterPolygon.Area > 0 Then
                                                        counterPolygon.UserData = plineCounter.ObjectId
                                                        listCounter.Add(counterPolygon)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    ElseIf dataParentParcel.TypeGeometry = StructureElement.typeGeometryObject.line Then
                                        Dim cadastralLine As NetTopologySuite.Geometries.LineString = acadNetTopoSuite.CreateLineString(plineCounter)
                                        If cadastralLine IsNot Nothing AndAlso Not cadastralLine.IsEmpty AndAlso cadastralLine.Length > 0 Then
                                            cadastralLine = cadastralNetTopoSuite.RemoveDuplicateVertices(cadastralLine, 0.01)
                                            If cadastralLine.Coordinates.Count > 2 Then
                                                If cadastralLine.Area > 0 Then
                                                    cadastralLine.UserData = plineCounter.ObjectId
                                                    listCounter.Add(cadastralLine)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        If listCounter.Count > 2 Then
                            'упорядочиваем полигоны
                            listCounter = cadastralNetTopoSuite.SortByNearestNeighbor(listCounter)
                        End If
                        '======================================================================================================================
                        'создаем контура
                        For Each topoObject As NetTopologySuite.Geometries.Geometry In listCounter
                            Dim idPolyline As ObjectId = topoObject.UserData
                            Dim polylineCounter As Polyline = tr.GetObject(idPolyline, OpenMode.ForWrite)
                            If polylineCounter.ObjectId.IsNull = False Then
                                If TypeOf topoObject Is NetTopologySuite.Geometries.Polygon Then
                                    polylineCounter = classDrawObject.createPolylineByPolygon(topoObject, polylineCounter)
                                ElseIf TypeOf topoObject Is LineString Then
                                    polylineCounter = classDrawObject.createPolylineByLineString(topoObject, polylineCounter)
                                End If
                                Dim newDataCounter As StructureElement = New StructureElement
                                newDataCounter = dataParentParcel
                                Dim keyParamCounter As String = newDataCounter.KeyParameter
                                counterParcel += 1
                                keyParamCounter = CadastralTools.JsonValidator.SetJsonFieldValue(keyParam, "ConditionalNumber", numberParcel & "(" & counterParcel & ")")
                                newDataCounter.KeyParameter = keyParamCounter
                                newDataCounter.IdElement = Guid.NewGuid.ToString
                                Dim boolWriteDataOstriv As Boolean = FuncXRecords.writeXRecords(polylineCounter, newDataCounter, StructureElement.tableXRecords.PROJECT_STRUCTURES)
                                lineStyle.setStyle(polylineCounter)
                            End If
                        Next
                    End If
                End If
            End If
            tr.Commit()
        End Using
    End Sub
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'рисование объектов по аналогу
    <CommandMethod("ReplaceCadastralObject")>
    Public Sub ReplaceCadastralObject()
        Dim acDoc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        Dim strDWGName As String = acDoc.Name
        Try
            Dim objBase As Object = Application.GetSystemVariable("DWGTITLED")
            If System.Convert.ToInt16(objBase) > 0 Then
                Dim obj1 As Object = Application.GetSystemVariable("DBMOD")
                If Not (System.Convert.ToInt16(obj1) = 0) Then
                    acDoc.Database.SaveAs(strDWGName, True, DwgVersion.Current, acDoc.Database.SecurityParameters)
                End If
            End If
        Catch ex As Teigha.Runtime.Exception
        End Try
        Dim stylePolyline As CadastralPolylineStyle = Nothing
        Using acTrans As Transaction = db.TransactionManager.StartTransaction()
            Dim acBlkTbl As BlockTable = acTrans.GetObject(db.BlockTableId, OpenMode.ForWrite)
            Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
            'получаем доступ к объекту
            Dim parentObjectSelectResult As PromptEntityResult = ed.GetEntity(vbLf & "Выберите объект в качестве образца (Полилинию): ")
            If parentObjectSelectResult.Status = PromptStatus.Cancel Then
                Exit Sub 'если нажали ESC то выходим из программы
            End If
            'ссылка на объект-аналог
            Dim parentObject As Teigha.DatabaseServices.Entity = acTrans.GetObject(parentObjectSelectResult.ObjectId, OpenMode.ForRead)
            Dim dataObject As StructureElement = Nothing
            Dim boolReadData As Boolean = FuncXRecords.getXRecords(parentObject, dataObject)
            Dim tablePS As String = ""
            If IsNothing(dataObject) = False Then
                tablePS = dataObject.TablePS
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'считываем стиль
            stylePolyline = New CadastralPolylineStyle(parentObject)
            stylePolyline.DBDrawing = db
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'выбираем новый объект
            Dim SelectNewObject As PromptEntityResult = ed.GetEntity(vbLf & "Выберите Полилинию: ")
            If SelectNewObject.Status = PromptStatus.OK Then
                Dim entityObject As Teigha.DatabaseServices.Entity = acTrans.GetObject(SelectNewObject.ObjectId, OpenMode.ForWrite)
                If TypeOf entityObject Is Polyline Then
                    stylePolyline.setStyle(entityObject)
                    Dim boolXRec As Boolean = FuncXRecords.writeXRecords(entityObject, dataObject)
                    If boolXRec = True Then
                        If parentObject.IsWriteEnabled = True Then
                            parentObject.Erase()
                        End If
                    End If
                Else
                    MsgBox("Необходимо выдрать примитив - Полилинию")
                End If
            End If
            acTrans.Commit()
        End Using
    End Sub
End Class
