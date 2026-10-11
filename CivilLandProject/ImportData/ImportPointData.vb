Imports HostMgd.ApplicationServices
Imports HostMgd.EditorInput
Imports Teigha.DatabaseServices
Imports Teigha.Runtime

Public Class ImportPointData
    <CommandMethod("ImportCadastralPoint")>
    Public Sub ImportCadPoint()
        Dim doc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        Dim formImportData As CadastralTools.FormImportPoint = New CadastralTools.FormImportPoint
        formImportData.ShowDialog()
        If formImportData.boolShow = False Then
            Exit Sub
        End If
        Dim templateDWG As String = formImportData.putchTemplate
        templateDWG = templateDWG.Replace(".xml", ".dwg")
        If IO.File.Exists(templateDWG) = False Then
            MsgBox("Шаблон не найден")
            Exit Sub
        End If
        Dim typeObject As String = formImportData.typeObject
        Dim deskZn As String = formImportData.nameObject
        Dim arrayPropZn As String(,) = formImportData.arrayProperties
        Dim nameBlk As String = CadastralTools.MathFunction.FuncFindValueToArray2d(arrayPropZn, "Name")
        Dim heightAttr As Double = CadastralTools.MathFunction.FuncFindValueToArray2d(arrayPropZn, "AttributeHeight")
        Dim styleObject As CadastralBlockStyle = New CadastralBlockStyle
        styleObject.createBlockStyle(db, arrayPropZn, formImportData.putchTemplate)
        'Dim boolCopyObject As Boolean = CadastralStyle.FuncCopyBlockDestDraw(templateDWG, nameBlk)
        ' Блокировка документа
        Using locker As DocumentLock = doc.LockDocument()
            ' Начинаем транзакцию в основном чертеже
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim errCount As Integer = 1
                Try
                    Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                    Dim ms As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    If IsArray(formImportData.arrayData) = True Then
                        For i As Integer = 0 To formImportData.arrayData.GetUpperBound(1) - 1
                            Dim x As Double = 0
                            Dim y As Double = 0
                            Dim z As Double = 0
                            Dim arrayAttr As String(,) = {}
                            Dim countArrayAttr As Integer = 0
                            For j As Integer = 0 To formImportData.arrayData.GetUpperBound(0)
                                Dim nameField As String = formImportData.arrayData(j, 0)
                                Dim valueData As String = formImportData.arrayData(j, i + 1)
                                If IsNothing(valueData) = False Then
                                    If nameField Like "X" Then
                                        If IsNumeric(valueData) = True Then
                                            x = formImportData.arrayData(j, i + 1)
                                        Else
                                            Continue For
                                        End If

                                    ElseIf nameField Like "Y" Then
                                        If IsNumeric(valueData) = True Then
                                            y = formImportData.arrayData(j, i + 1)
                                        Else
                                            Continue For
                                        End If

                                    ElseIf nameField Like "Z" Then
                                        If IsNumeric(valueData) = True Then
                                            z = formImportData.arrayData(j, i + 1)
                                        Else
                                            z = 0
                                        End If

                                    Else
                                            ReDim Preserve arrayAttr(2, countArrayAttr)
                                        arrayAttr(0, countArrayAttr) = nameField
                                        arrayAttr(1, countArrayAttr) = formImportData.arrayData(j, i + 1)
                                        arrayAttr(2, countArrayAttr) = heightAttr
                                        countArrayAttr += 1
                                    End If
                                End If
                            Next j
                            Dim insPoint As Teigha.Geometry.Point3d = New Teigha.Geometry.Point3d(x, y, z)
                            Dim blk As BlockReference = FuncCreateObject.FuncInsertDrawingBlk(insPoint, nameBlk)
                            If IsNothing(blk) = False Then
                                If arrayAttr.Length > 0 Then
                                    Dim boolRecAttrib As Boolean = FuncBlockReference.writeAttributesBlock(blk, arrayAttr)
                                End If
                            End If
                            styleObject.setStyle(blk)
                            errCount += 1
                        Next i
                    End If
                Catch ex As Exception
                    MsgBox("Ошибка-" & errCount)
                End Try
                tr.Commit()
            End Using
        End Using
    End Sub
End Class
