Imports HostMgd.ApplicationServices
Imports HostMgd.EditorInput
Imports Teigha.DatabaseServices
Imports Teigha.Runtime

Public Class ReportTree
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'автоматическое создание земельных участков
    <CommandMethod("ReportTree")>
    Public Sub ReportTree()
        Dim acDoc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
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
        Dim result As String = ""

        Try
            ' Запрашиваем выбор объекта с фильтром на BlockReference
            Dim promptOpts As New PromptEntityOptions(vbLf & "Выберите внешнюю ссылку: ")
            promptOpts.SetRejectMessage("Выбранный объект не является блоком")
            promptOpts.AddAllowedClass(GetType(BlockReference), False)
            Dim promptResult As PromptEntityResult = ed.GetEntity(promptOpts)
            If promptResult.Status = PromptStatus.OK Then
                Dim blkRefId As ObjectId = promptResult.ObjectId
                Using tr As Transaction = ed.Document.TransactionManager.StartTransaction()
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
                            Else
                                ed.WriteMessage(vbLf & "Выбранный блок не является внешней ссылкой")
                            End If
                        End If
                    End If
                    tr.Commit()
                End Using
            End If
        Catch ex As Exception
            ed.WriteMessage(vbLf & $"Ошибка: {ex.Message}")
        End Try
        If IO.File.Exists(result) = True Then
            Dim zagolovok As String(,) = {}
            ReDim zagolovok(5, 0)
            zagolovok(0, 0) = "№ на плане"
            zagolovok(1, 0) = "Вид объекта строительства"
            zagolovok(2, 0) = "X"
            zagolovok(3, 0) = "Y"
            zagolovok(4, 0) = "Кадастровый номер"
            zagolovok(5, 0) = "Разрешенное использование"

            Dim arrayRezult As String(,) = {}
            Dim countArrayRezult As Integer = 0

            Dim listParcels As List(Of NetTopologySuite.Geometries.Polygon) = New List(Of NetTopologySuite.Geometries.Polygon)
            listParcels = FuncDWGReference.GetCadastralPolygonToReferenceDWG(result, CadastralTools.StructureElement.classObject.LAND_CADASTRAL)

            Dim listDivision As List(Of NetTopologySuite.Geometries.Polygon) = New List(Of NetTopologySuite.Geometries.Polygon)
            listDivision = FuncDWGReference.GetCadastralPolygonToReferenceDWG(result, CadastralTools.StructureElement.classObject.LAND_CADASTRAL)

            Dim listPoint As List(Of NetTopologySuite.Geometries.Point) = FuncBlockReference.GetBlocksByName(db, "СЗН-Дерево отдельно стоящее")
            If listParcels.Count > 0 And listPoint.Count > 0 Then
                Dim classTopo As New CadastralTools.FuncNetTopologySuite()
                For i As Integer = 0 To listPoint.Count - 1
                    Dim pt As NetTopologySuite.Geometries.Point = listPoint.Item(i)
                    If IsNothing(pt) = True Then Continue For
                    Dim number As String = 0
                    Dim arrayAttrib As String(,) = pt.UserData
                    If IsNothing(arrayAttrib) = False Then
                        If IsArray(arrayAttrib) = True Then
                            number = CadastralTools.MathFunction.FuncFindValueToArray2d(arrayAttrib, "НОМЕР", 0, 1)
                        End If
                    End If
                    Dim boolInsPoint As Boolean = False
                    For j As Integer = 0 To listParcels.Count - 1
                        Dim poly As NetTopologySuite.Geometries.Polygon = listParcels.Item(j)
                        If IsNothing(poly) = True Then Continue For
                        boolInsPoint = classTopo.IsPointInsidePolygon(pt, poly)
                        If boolInsPoint = True Then
                            Dim dict As Dictionary(Of String, CadastralTools.StructureElement) = poly.UserData
                            Dim cadNumber As String = ""
                            Dim utilization As String = ""
                            If IsNothing(dict) = False Then
                                Dim dataPoly As CadastralTools.StructureElement = dict.Item("PROJECT_STRUCTURES")
                                If IsNothing(dataPoly) = False Then
                                    Dim userParcels As CadastralTools.LAND_CADASTRAL_PARCEL = Newtonsoft.Json.JsonConvert.DeserializeObject(Of CadastralTools.LAND_CADASTRAL_PARCEL)(dataPoly.KeyParameter)
                                    If IsNothing(userParcels) = False Then
                                        cadNumber = userParcels.CadNumber
                                        utilization = userParcels.Utilization
                                    End If
                                End If
                            End If
                            ReDim Preserve arrayRezult(5, countArrayRezult)
                            arrayRezult(0, countArrayRezult) = number
                            arrayRezult(1, countArrayRezult) = "Дерево отдельно стоящее"
                            arrayRezult(2, countArrayRezult) = pt.Y
                            arrayRezult(3, countArrayRezult) = pt.X
                            arrayRezult(4, countArrayRezult) = cadNumber
                            arrayRezult(5, countArrayRezult) = utilization
                            countArrayRezult += 1
                            Exit For
                        End If
                    Next j
                    'проверяем квартала
                    If boolInsPoint = False Then
                        If IsNothing(listDivision) = False Then
                            If listDivision.Count > 0 Then
                                For j As Integer = 0 To listDivision.Count - 1
                                    Dim poly As NetTopologySuite.Geometries.Polygon = listDivision.Item(j)
                                    If IsNothing(poly) = True Then Continue For
                                    boolInsPoint = classTopo.IsPointInsidePolygon(pt, poly)
                                    If boolInsPoint = True Then
                                        Dim dict As Dictionary(Of String, CadastralTools.StructureElement) = poly.UserData
                                        Dim cadNumber As String = ""
                                        If IsNothing(dict) = False Then
                                            Dim dataPoly As CadastralTools.StructureElement = dict.Item("PROJECT_STRUCTURES")
                                            If IsNothing(dataPoly) = False Then
                                                Dim userDivision As CadastralTools.LAND_CADASTRAL_DIVISION = Newtonsoft.Json.JsonConvert.DeserializeObject(Of CadastralTools.LAND_CADASTRAL_DIVISION)(dataPoly.KeyParameter)
                                                If IsNothing(userDivision) = False Then
                                                    cadNumber = userDivision.CadNumber
                                                End If
                                            End If
                                        End If
                                        ReDim Preserve arrayRezult(5, countArrayRezult)
                                        arrayRezult(0, countArrayRezult) = number
                                        arrayRezult(1, countArrayRezult) = "Дерево отдельно стоящее"
                                        arrayRezult(2, countArrayRezult) = pt.Y
                                        arrayRezult(3, countArrayRezult) = pt.X
                                        arrayRezult(4, countArrayRezult) = cadNumber
                                        arrayRezult(5, countArrayRezult) = ""
                                        countArrayRezult += 1
                                        Exit For
                                    End If
                                Next j
                            End If
                        End If

                        If boolInsPoint = False Then
                            ReDim Preserve arrayRezult(5, countArrayRezult)
                            arrayRezult(0, countArrayRezult) = number
                            arrayRezult(1, countArrayRezult) = "Дерево отдельно стоящее"
                            arrayRezult(2, countArrayRezult) = pt.Y
                            arrayRezult(3, countArrayRezult) = pt.X
                            arrayRezult(4, countArrayRezult) = ""
                            arrayRezult(5, countArrayRezult) = ""
                            countArrayRezult += 1
                        End If
                    End If
                Next i
            End If


            If arrayRezult.GetUpperBound(1) > 0 Then
                Dim classTeport As CadastralTools.ReportExcel = New CadastralTools.ReportExcel
                classTeport.pointForest(zagolovok, arrayRezult)
            End If
        End If
    End Sub
End Class
