Imports System.IO
Imports System.Windows.Controls.Primitives
Imports System.Windows.Documents
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar
Imports HostMgd.ApplicationServices
Imports Teigha.DatabaseServices

Public Class NC_LandProject
    Inherits CadastralTools.StructureElement
    'читает объекты из текущего чертежа
    'Public Function GetObjectToCurrentDWG() As Dictionary(Of String, List(Of CadastralTools.StructureElement))
    '    Dim result As New Dictionary(Of CadastralTools.StructureElement.classObject, List(Of CadastralTools.StructureElement))()
    '    Dim doc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
    '    Dim db As Database = doc.Database

    '    Using tr As Transaction = doc.TransactionManager.StartTransaction()
    '        Try
    '            Dim bt As BlockTable = DirectCast(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
    '            Dim btr As BlockTableRecord = DirectCast(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)

    '            For Each id As ObjectId In btr
    '                If Not id.IsNull AndAlso Not id.IsErased Then
    '                    Dim acEnt As Entity = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)

    '                    If acEnt IsNot Nothing Then
    '                        Dim civilListObject As Dictionary(Of String, CadastralTools.StructureElement) = FuncXRecords.readXRecords(acEnt)
    '                        If IsNothing(civilListObject) = False Then
    '                            If civilListObject.Count > 0 Then
    '                                For i As Integer = 0 To civilListObject.Count - 1
    '                                    Dim civilObject As CadastralTools.StructureElement = civilListObject.ElementAt(i).Value
    '                                    Dim objType As CadastralTools.StructureElement.classObject = civilObject.ClassObject
    '                                    Dim objElement As CadastralTools.StructureElement = civilObject
    '                                    ' Проверяем, существует ли уже такой ключ в словаре
    '                                    If result.ContainsKey(objType) Then
    '                                        ' Добавляем в существующий список
    '                                        result(objType).Add(objElement)
    '                                    Else
    '                                        ' Создаем новый список и добавляем элемент
    '                                        result(objType) = New List(Of CadastralTools.StructureElement) From {objElement}
    '                                    End If
    '                                Next i
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '            Next
    '            tr.Commit()
    '        Catch ex As Exception
    '            ' Логирование ошибки
    '            doc.Editor.WriteMessage(vbCrLf & $"Ошибка при чтении объектов: {ex.Message}")
    '            ' Можно также записать в лог-файл
    '            ' Logger.LogError(ex)
    '        End Try
    '    End Using
    '    Return result
    'End Function
    'читает объекты из внешней ссылки

End Class
