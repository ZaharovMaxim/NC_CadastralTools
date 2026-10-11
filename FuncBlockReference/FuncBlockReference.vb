Imports HostMgd.ApplicationServices
Imports Teigha.Editor
Imports Teigha.DatabaseServices
Imports HostMgd.EditorInput

Public Class AttributeData
    Public Property Tag As String        ' Имя (тег) атрибута
    Public Property Description As String ' Описание
    Public Property Value As String       ' Значение
End Class

Public Class FuncBlockReference
    Public Shared Function GetBlocksByName(ByVal db As Database, ByVal blockName As String, Optional ByVal matchCase As Boolean = False) As List(Of NetTopologySuite.Geometries.Point)
        Dim result As New List(Of NetTopologySuite.Geometries.Point)()
        Dim comparison As StringComparison = If(matchCase, StringComparison.Ordinal, StringComparison.OrdinalIgnoreCase)

        Try
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                ' Открываем пространство модели
                Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim ms As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead)
                ' Проходим по всем объектам в пространстве модели
                For Each objId As ObjectId In ms
                    ' Проверяем, является ли объект BlockReference
                    If objId.ObjectClass.Name = "AcDbBlockReference" Then
                        Dim blkRef As BlockReference = TryCast(tr.GetObject(objId, OpenMode.ForRead), BlockReference)
                        If blkRef IsNot Nothing Then
                            ' Получаем определение блока для проверки имени
                            Dim btr As BlockTableRecord = TryCast(tr.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                            If btr IsNot Nothing Then
                                ' Сравниваем имя блока
                                If btr.Name.Equals(blockName, comparison) Then
                                    Dim arrayAttrib As String(,) = {}
                                    'Dim boolFindAttrib As Boolean = FuncBlockReference.FuncReadAllBlockAttr(blkRef, arrayAttrib)
                                    Dim userPoint As NetTopologySuite.Geometries.Point = New NetTopologySuite.Geometries.Point(blkRef.Position.X, blkRef.Position.Y, blkRef.Position.Z)
                                    userPoint.UserData = arrayAttrib
                                    result.Add(userPoint)
                                End If
                            End If
                        End If
                    End If
                Next
                tr.Commit()
            End Using
        Catch ex As Exception
        End Try
        Return result
    End Function
    'функция читает атрибуты блока из блока вставленного в чертеж
    Public Function GetBlockAttributes(blockName As String) As List(Of AttributeData)
        Dim result As New List(Of AttributeData)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                ' Получаем пространство модели
                Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim ms As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead)

                ' Ищем все вставки блока с указанным именем
                For Each objId As ObjectId In ms
                    Dim ent As Entity = tr.GetObject(objId, OpenMode.ForRead)

                    If TypeOf ent Is BlockReference Then
                        Dim blkRef As BlockReference = DirectCast(ent, BlockReference)

                        ' Проверяем имя блока
                        If blkRef.Name = blockName Then
                            ' Получаем коллекцию атрибутов
                            Dim atts As AttributeCollection = blkRef.AttributeCollection

                            For Each attId As ObjectId In atts
                                Dim attRef As AttributeReference = tr.GetObject(attId, OpenMode.ForRead)

                                Dim attrData As New AttributeData()
                                attrData.Tag = attRef.Tag
                                attrData.Description = attRef.TextString ' У атрибутов описание часто хранится в Prompt
                                ' Если нужно получить описание из определения атрибута:
                                ' attrData.Description = GetAttributeDescription(blkRef.BlockTableRecord, attRef.Tag, tr)
                                attrData.Value = attRef.TextString
                                result.Add(attrData)
                            Next

                            ' Если нашли первый блок с нужным именем, можно выйти
                            ' Exit For (если нужен только первый блок)
                        End If
                    End If
                Next

                tr.Commit()
            Catch ex As Exception
                ed.WriteMessage(vbCrLf & $"Ошибка: {ex.Message}")
            End Try
        End Using

        Return result
    End Function
    'функция читает атрибуты блока из определения блока
    Public Shared Function GetBlockAttributesWithDesc(blockName As String) As List(Of CadastralTools.AttributeData)
        Dim result As New List(Of CadastralTools.AttributeData)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Try
                Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                ' Проверяем существование определения блока
                If Not bt.Has(blockName) Then
                    Return result
                End If

                ' Открываем определение блока
                Dim btrId As ObjectId = bt(blockName)
                Dim btr As BlockTableRecord = tr.GetObject(btrId, OpenMode.ForRead)

                ' Сначала собираем описания атрибутов из определения блока
                Dim attributeDefs As New Dictionary(Of String, String)

                For Each objId As ObjectId In btr
                    Dim dbObj As DBObject = tr.GetObject(objId, OpenMode.ForRead)
                    If TypeOf dbObj Is AttributeDefinition Then
                        Dim attDef As AttributeDefinition = DirectCast(dbObj, AttributeDefinition)
                        Dim attrData As New CadastralTools.AttributeData()
                        attrData.Tag = attDef.Tag
                        attrData.Description = attDef.Prompt
                        attrData.Value = attDef.TextString
                        result.Add(attrData)
                    End If
                Next
                tr.Commit()
            Catch ex As Exception
            End Try
        End Using

        Return result
    End Function

    '===================================================================================
    'Функция считывает в массив все атрибуты блока (Массив 0-тег, 1-содержание атрибута, 2-высота атрибута, 3-угол наклона атрибута, 4-текстовый стиль, 5-цвет)
    Public Shared Function readAttributesBlock(ByVal blkPoint As BlockReference) As String(,)
        Dim result As String(,) = {}
        Dim countAttrib As Integer = 0
        Dim acDoc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim bt As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = CType(acTrans.GetObject(bt(blkPoint.Name), OpenMode.ForRead), BlockTableRecord)
                Dim attCol As AttributeCollection = blkPoint.AttributeCollection
                For Each attId As ObjectId In attCol
                    Dim attRef As AttributeReference = acTrans.GetObject(attId, OpenMode.ForRead)
                    ReDim Preserve result(5, countAttrib)
                    result(0, countAttrib) = attRef.Tag 'имя поля
                    result(1, countAttrib) = attRef.TextString 'значение
                    result(2, countAttrib) = attRef.Height 'высота атрибута
                    result(3, countAttrib) = attRef.Rotation 'угол атрибута
                    result(4, countAttrib) = attRef.TextStyleName 'стиль текста
                    result(5, countAttrib) = attRef.ColorIndex 'цвет
                    countAttrib += 1
                Next
            Catch ex As Teigha.Runtime.Exception
            End Try
            Return result
        End Using
    End Function
    '-===================================================================================
    'Функция считывает определенный атрибут блока и возврящает его значение
    Public Shared Function FuncReadValueAttr(ByVal blkPoint As BlockReference, ByVal NameField As String) As String
        FuncReadValueAttr = ""
        Dim acDoc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim attCol As AttributeCollection = blkPoint.AttributeCollection
                If attCol.Count > 0 Then
                    For Each attId As ObjectId In attCol
                        Dim attRef As AttributeReference = acTrans.GetObject(attId, OpenMode.ForRead)
                        If attRef.Tag Like NameField Then
                            FuncReadValueAttr = attRef.TextString
                            Exit Function
                        End If
                    Next
                End If
            Catch
                Exit Function
            End Try
        End Using
    End Function

    '-===================================================================================
    'Функция записывает атрибуты блока (Массив 0-tag, 1-содержание атрибута, 2-высота атрибута, 3-угол наклона атрибута, 4-стиль текста, 5-цвет)
    Public Shared Function writeAttributesBlock(ByVal blk As BlockReference, ByRef ArrayAtribute As String(,)) As Boolean
        writeAttributesBlock = False
        Dim acDoc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        If IsNothing(blk) = True Then Return False
        If IsNothing(ArrayAtribute) = True Then Return False
        If ArrayAtribute.Length = 0 Then Return False
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim attributeBlockCollection As AttributeCollection = blk.AttributeCollection
                If attributeBlockCollection.Count > 0 Then
                    For Each attId As ObjectId In attributeBlockCollection
                        Try
                            Dim attRef As AttributeReference = acTrans.GetObject(attId, OpenMode.ForWrite)
                            Dim atributeTag As String = attRef.Tag
                            For i As Integer = 0 To ArrayAtribute.GetUpperBound(1)
                                Dim nameField As String = ArrayAtribute(0, i)
                                If nameField Like atributeTag Then
                                    Dim valField As String = ArrayAtribute(1, i)
                                    If IsNothing(valField) = False Then
                                        If valField.Trim.Length > 0 Then
                                            attRef.TextString = valField
                                            If attRef.Visible = True Then
                                                If ArrayAtribute.GetUpperBound(0) > 1 Then
                                                    Dim heightAttr As Double = Val(ArrayAtribute(2, i))
                                                    If heightAttr > 0 Then
                                                        attRef.Height = heightAttr
                                                    End If
                                                End If
                                                If ArrayAtribute.GetUpperBound(0) > 2 Then
                                                    Dim angle As Double = Val(ArrayAtribute(3, i))
                                                    attRef.Oblique = angle
                                                End If
                                                If ArrayAtribute.GetUpperBound(0) > 3 Then
                                                    Dim textStyleAttr As String = ArrayAtribute(4, i)
                                                    If textStyleAttr.Trim.Length > 0 Then
                                                        Dim acTextStyleTbl As TextStyleTable = acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead)
                                                        If acTextStyleTbl.Has(textStyleAttr) = True Then
                                                            Dim idTextStyle As ObjectId = acTextStyleTbl.Item(textStyleAttr)
                                                            attRef.TextStyleId = idTextStyle
                                                        End If
                                                    End If
                                                End If
                                                If ArrayAtribute.GetUpperBound(0) > 3 Then
                                                    Dim colorInd As Integer = Val(ArrayAtribute(5, i))
                                                    If colorInd > -1 And colorInd < 257 Then
                                                        attRef.ColorIndex = colorInd
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                    Exit For
                                End If
                            Next i
                        Catch ex As System.Exception
                        End Try
                    Next
                End If
            Catch
                Exit Function
            End Try
            acTrans.Commit()
            Return True
        End Using
    End Function
    '-===================================================================================
    'Функция записывает атрибуты блока (Массив 0-tag, 1-содержание атрибута, 2-высота атрибута, 3-угол наклона атрибута, 4-стиль текста, 5-цвет)
    Public Shared Function writeListAttributesBlock(ByVal blk As BlockReference, ByVal listAttrib As List(Of CadastralTools.AttributeData)) As Boolean
        writeListAttributesBlock = False
        Dim acDoc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        If IsNothing(blk) = True Then Return False
        If IsNothing(listAttrib) = True Then Return False
        If listAttrib.Count = 0 Then Return False
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim attributeBlockCollection As AttributeCollection = blk.AttributeCollection
                If attributeBlockCollection.Count > 0 Then
                    For Each attId As ObjectId In attributeBlockCollection
                        Try
                            Dim attRef As AttributeReference = acTrans.GetObject(attId, OpenMode.ForWrite)
                            Dim atributeTag As String = attRef.Tag
                            For i As Integer = 0 To listAttrib.Count - 1
                                Dim userAttrib As CadastralTools.AttributeData = listAttrib.Item(i)
                                Dim nameField As String = userAttrib.Tag
                                If nameField Like atributeTag Then
                                    Dim valField As String = userAttrib.Value
                                    If IsNothing(valField) = False Then
                                        If valField.Trim.Length > 0 Then
                                            attRef.TextString = valField
                                        End If
                                    End If
                                    If attRef.Visible = True Then
                                        If userAttrib.height > 0 Then
                                            attRef.Height = userAttrib.height
                                        End If
                                        If userAttrib.angle > 0 Then
                                            attRef.Oblique = userAttrib.angle
                                        End If
                                        Dim textStyleAttr As String = userAttrib.textStyle
                                        If IsNothing(textStyleAttr) = False Then
                                            If textStyleAttr.Trim.Length > 0 Then
                                                Dim acTextStyleTbl As TextStyleTable = acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead)
                                                If acTextStyleTbl.Has(textStyleAttr) = True Then
                                                    Dim idTextStyle As ObjectId = acTextStyleTbl.Item(textStyleAttr)
                                                    attRef.TextStyleId = idTextStyle
                                                End If
                                            End If
                                            attRef.ColorIndex = userAttrib.color
                                        End If
                                    End If
                                    Exit For
                                End If
                            Next i
                        Catch ex As System.Exception
                        End Try
                    Next
                End If
            Catch
                Exit Function
            End Try
            acTrans.Commit()
            Return True
        End Using
    End Function

    '-===================================================================================
    'Функция крнвертирует список атрибутов блока в масив (Массив 0-tag, 1-содержание атрибута)
    Public Shared Function convertListAttributesBlockToArray(ByVal listAttrib As List(Of CadastralTools.AttributeData)) As String(,)
        Dim result As String(,) = {}
        Dim countResult As Integer = 0
        If IsNothing(listAttrib) = True Then Return result
        If listAttrib.Count = 0 Then Return result
        For i As Integer = 0 To listAttrib.Count - 1
            Dim attrib As CadastralTools.AttributeData = listAttrib.Item(i)
            Dim tag As String = attrib.Tag
            Dim value As String = attrib.Value
            If IsNothing(value) = True Then value = ""
            If IsNothing(tag) = True Then Continue For
            If tag.Trim.Length = 0 Then Continue For
            'проверяем дубликаты
            Dim boolFindTag As Boolean = False
            If result.Length > 0 Then
                For j As Integer = 0 To result.GetUpperBound(1)
                    Dim tempTag As String = result(0, j)
                    If tempTag Like tag Then
                        boolFindTag = True
                        Exit For
                    End If
                Next j
            End If
            If boolFindTag = False Then
                ReDim Preserve result(1, countResult)
                result(0, countResult) = tag
                result(1, countResult) = value
                countResult += 1
            End If
        Next i
        Return result
    End Function
End Class
