Imports HostMgd.ApplicationServices
Imports HostMgd.EditorInput
Imports HostMgd.Windows
Imports Teigha.DatabaseServices
Imports Teigha.Geometry
Imports Teigha.Runtime

Public Class FuncAxisRoads
    'прямой участок - прямой участок (без кривых)
    <CommandMethod("CreateLineLineAlignment")>
    Public Sub CreateLineLineAlignment()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        '===============================================================================================================
        Dim pPtRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
        pPtOpts.Message = vbLf & "Укажите первую точку прямой: "
        pPtRes = doc.Editor.GetPoint(pPtOpts)
        Dim ptStart As Point3d = pPtRes.Value
        If pPtRes.Status = PromptStatus.Cancel Then Exit Sub
        pPtOpts.Message = vbLf & "Укажите следующую точку прямого участка: "
        pPtOpts.UseBasePoint = True
        pPtOpts.BasePoint = ptStart

        Do
            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Using tr
                Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                pPtRes = doc.Editor.GetPoint(pPtOpts)
                Dim ptEnd As Point3d = pPtRes.Value
                If pPtRes.Status = PromptStatus.Cancel Then Exit Do
                Using userLine As Line = New Line(ptStart, ptEnd)
                    btr.AppendEntity(userLine)
                    tr.AddNewlyCreatedDBObject(userLine, True)
                End Using
                ptStart = ptEnd
                pPtOpts.BasePoint = ptEnd
                tr.Commit()
            End Using
        Loop
    End Sub
    'прямой участок - прямой участок (с вписанной круговой кривой)
    <CommandMethod("CreateLineCurveLineAlignment")>
    Public Sub CreateLineCurveLineAlignment()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        '===============================================================================================================

        'начальные элементы прямых вставок
        Dim startPointLine1 As Point3d = New Point3d(-1, -1, -1)
        Dim endPointLine1 As Point3d = New Point3d(-1, -1, -1)
        Dim startPointLine2 As Point3d = New Point3d(-1, -1, -1)
        Dim endPointLine2 As Point3d = New Point3d(-1, -1, -1)
        Dim pointCrossTan As Point2d = New Point2d(-1, -1)

        Dim userLine1 As Line = Nothing
        Dim userLine2 As Line = Nothing

        Dim userForm As FormAddCurveAlignment = New FormAddCurveAlignment
        userForm.ShowDialog()
        If userForm.boolShow = False Then
            Exit Sub
        End If

        Do
            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Using tr
                Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                If userForm.RadioButton5.Checked = True Then
                    '===============================================================================================================
                    Dim pPtRes As PromptPointResult
                    Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                    If startPointLine1 = endPointLine1 Then
                        pPtOpts.Message = vbLf & "Укажите первую точку прямой: "
                        pPtRes = doc.Editor.GetPoint(pPtOpts)
                        If pPtRes.Status = PromptStatus.Cancel Then Exit Sub
                        startPointLine1 = pPtRes.Value

                        pPtOpts.Message = vbLf & "Укажите следующую точку прямого участка: "
                        pPtOpts.UseBasePoint = True
                        pPtOpts.BasePoint = startPointLine1
                        pPtRes = doc.Editor.GetPoint(pPtOpts)
                        If pPtRes.Status = PromptStatus.Cancel Then Exit Do
                        endPointLine1 = pPtRes.Value
                        'рисуем 1 линию
                        userLine1 = New Line(startPointLine1, endPointLine1)
                        btr.AppendEntity(userLine1)
                        tr.AddNewlyCreatedDBObject(userLine1, True)
                    End If

                    startPointLine2 = endPointLine1
                    pointCrossTan = New Point2d(startPointLine1.X, startPointLine1.Y)

                    pPtOpts.Message = vbLf & "Укажите следующую точку прямого участка: "
                    pPtOpts.UseBasePoint = True
                    pPtOpts.BasePoint = startPointLine2
                    pPtRes = doc.Editor.GetPoint(pPtOpts)
                    If pPtRes.Status = PromptStatus.Cancel Then Exit Do
                    endPointLine2 = pPtRes.Value
                    'рисуем 2 линию
                    userLine2 = New Line(startPointLine2, endPointLine2)
                    btr.AppendEntity(userLine2)
                    tr.AddNewlyCreatedDBObject(userLine2, True)
                    pPtOpts.BasePoint = endPointLine2
                Else
                    'выбор прямых участков
                    Dim pEntRes As PromptEntityResult
                    Dim pEntOpts As PromptEntityOptions = New PromptEntityOptions("")
                    pEntOpts.Message = vbLf & "Выберите 1 прямой участок трассы: "
                    pEntRes = doc.Editor.GetEntity(pEntOpts)
                    If pEntRes.Status = PromptStatus.OK Then
                        Dim idEnt As ObjectId = pEntRes.ObjectId
                        Dim userEnt As Entity = tr.GetObject(idEnt, OpenMode.ForRead)
                        If TypeOf userEnt Is Line Then
                            userLine1 = userEnt
                        Else
                            MsgBox("Ошибка выбора. Необходимо выбрать линию!!!")
                            Exit Sub
                        End If
                    End If

                    pEntOpts.Message = vbLf & "Выберите 2 прямой участок трассы: "
                    pEntRes = doc.Editor.GetEntity(pEntOpts)
                    If pEntRes.Status = PromptStatus.OK Then
                        Dim idEnt As ObjectId = pEntRes.ObjectId
                        Dim userEnt As Entity = tr.GetObject(idEnt, OpenMode.ForRead)
                        If TypeOf userEnt Is Line Then
                            userLine2 = userEnt
                        Else
                            MsgBox("Ошибка выбора. Необходимо выбрать линию!!!")
                            Exit Sub
                        End If
                    End If
                    'находим начальные и конечные точки выбранных линий
                    startPointLine1 = userLine1.StartPoint
                    endPointLine1 = userLine1.EndPoint

                    startPointLine2 = userLine2.StartPoint
                    endPointLine2 = userLine2.EndPoint

                    Dim l1 As Double = endPointLine1.GetVectorTo(startPointLine2).Length
                    Dim l2 As Double = endPointLine1.GetVectorTo(endPointLine2).Length
                    Dim l3 As Double = startPointLine1.GetVectorTo(startPointLine2).Length
                    Dim l4 As Double = startPointLine1.GetVectorTo(endPointLine2).Length

                    If l2 < l1 And l2 < l3 And l2 < l4 Then
                        startPointLine2 = endPointLine2
                        endPointLine2 = userLine2.StartPoint
                    ElseIf l3 < l1 And l3 < l2 And l3 < l4 Then
                        startPointLine1 = endPointLine1
                        endPointLine1 = userLine1.StartPoint
                    ElseIf l3 < l1 And l4 < l4 And l4 < l3 Then
                        startPointLine1 = userLine1.EndPoint
                        endPointLine1 = userLine1.StartPoint
                        startPointLine2 = userLine2.EndPoint
                        endPointLine2 = userLine2.StartPoint
                    End If
                    'находим точку пересечения тангенсов
                    If endPointLine1 <> startPointLine2 Then
                        Dim boolFindPointTan As Boolean = MathFunction.FuncIntersectPolyline2(startPointLine1.X, startPointLine1.Y, endPointLine1.X, endPointLine1.Y, startPointLine2.X, startPointLine2.Y, endPointLine2.X, endPointLine2.Y, pointCrossTan)
                    End If
                End If
                If pointCrossTan.X <> -1 And pointCrossTan.Y <> -1 = True Then
                    'рассчитываем длины тангенсов
                    Dim T1 As Double = pointCrossTan.GetVectorTo(startPointLine1.Convert2d(New Plane)).Length
                    Dim T2 As Double = pointCrossTan.GetVectorTo(endPointLine2.Convert2d(New Plane)).Length
                    Dim maxT As Double = 0
                    If T1 < T2 Then
                        maxT = T1
                    Else
                        maxT = T2
                    End If

                    'считаем координаты точек для максимального радиуса
                    'для Т1
                    Dim angleT1 As Double = pointCrossTan.GetVectorTo(startPointLine1.Convert2d(New Plane)).Angle
                    Dim pointXt1 As Double = pointCrossTan.X + T1 * Math.Cos(angleT1)
                    Dim pointYt1 As Double = pointCrossTan.Y + T1 * Math.Sin(angleT1)
                    Dim offsetPointXt1 As Double = pointXt1 + 1 * Math.Cos(angleT1 + Math.PI / 2)
                    Dim offsetPointYt1 As Double = pointYt1 + 1 * Math.Sin(angleT1 + Math.PI / 2)
                    'для Т2
                    Dim angleT2 As Double = pointCrossTan.GetVectorTo(endPointLine2.Convert2d(New Plane)).Angle
                    Dim pointXt2 As Double = pointCrossTan.X + T2 * Math.Cos(angleT2)
                    Dim pointYt2 As Double = pointCrossTan.Y + T2 * Math.Sin(angleT2)
                    Dim offsetPointXt2 As Double = pointXt2 + 1 * Math.Cos(angleT2 + Math.PI / 2)
                    Dim offsetPointYt2 As Double = pointYt2 + 1 * Math.Sin(angleT2 + Math.PI / 2)

                    'вычисляем положение центра круговой кривой
                    Dim pointCentre As Point2d = New Point2d(-1, -1)
                    Dim boolFindMaxCentre As Boolean = MathFunction.FuncIntersectPolyline2(pointXt1, pointYt1, offsetPointXt1, offsetPointYt1, pointXt2, pointYt2, offsetPointXt2, offsetPointYt2, pointCentre)
                    If boolFindMaxCentre = True Then
                        '======================================================================================================================
                        'вставляем дугу
                        If userForm.RadioButton1.Checked = True Then
                            'строим перпендикуляры от конечной точки 1 выбранной линии и начальной точки 2 выбранной линии
                            Dim azLine1 As Double = userLine1.Angle - Math.PI / 2
                            Dim azLine2 As Double = userLine2.Angle - Math.PI / 2
                            Dim tempDist As Double = 1
                            Dim tempXLine1 As Double = endPointLine1.X + tempDist * Math.Cos(azLine1)
                            Dim tempYLine1 As Double = endPointLine1.Y + tempDist * Math.Sin(azLine1)
                            Dim tempXYPointLine1 As Point2d = New Point2d(tempXLine1, tempYLine1)

                            Dim tempXLine2 As Double = startPointLine2.X + tempDist * Math.Cos(azLine2)
                            Dim tempYLine2 As Double = startPointLine2.Y + tempDist * Math.Sin(azLine2)
                            Dim tempXYPointLine2 As Point2d = New Point2d(tempXLine2, tempYLine2)
                            'вычисляем положение центра круговой кривой
                            pointCentre = New Point2d(-1, -1)
                            Dim boolFindCentre As Boolean = MathFunction.FuncIntersectPolyline2(endPointLine1.X, tempXYPointLine1.Y, tempXLine1, tempYLine1, startPointLine2.X, tempXYPointLine2.Y, tempXLine2, tempYLine2, pointCentre)
                            If boolFindCentre = True Then
                                Dim startAngle As Double = pointCentre.GetVectorTo(endPointLine1.Convert2d(New Plane)).Angle
                                Dim EndAngle As Double = pointCentre.GetVectorTo(startPointLine2.Convert2d(New Plane)).Angle
                                If startAngle > EndAngle Then
                                    startAngle = EndAngle
                                    EndAngle = pointCentre.GetVectorTo(endPointLine1.Convert2d(New Plane)).Angle
                                End If
                                Dim r As Double = pointCentre.GetVectorTo(endPointLine1.Convert2d(New Plane)).Length
                                Dim centerAcrPoint As Point3d = New Point3d(pointCentre.X, pointCentre.Y, 0)
                                Using userArc As Arc = New Arc(centerAcrPoint, r, startAngle, EndAngle)
                                    btr.AppendEntity(userArc)
                                    tr.AddNewlyCreatedDBObject(userArc, True)
                                End Using
                            End If
                        End If
                    End If
                End If
                tr.Commit()
            End Using
        Loop
    End Sub
End Class
