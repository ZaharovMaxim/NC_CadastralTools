Imports System.IO
Imports System.Windows.Media.Animation
Imports System.Xml
Imports Teigha.DatabaseServices
Imports HostMgd.ApplicationServices
Imports HostMgd.EditorInput
Imports HostMgd.Windows
Imports Teigha
Imports Teigha.Geometry
Imports Teigha.GraphicsInterface
Imports Teigha.Runtime
Imports System.Diagnostics.Eventing.Reader
Imports System.Windows
Imports Application = HostMgd.ApplicationServices.Application
Imports NetTopologySuite.Utilities
Imports CadastralTools.StructureElement
Imports CadastralTools
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json

Public Class FuncXRecords
    'читает имена и характеристики полей выбранной таблицы PS из файла XML (Результат - Имя словаря, вложенный словарь имя поля - характеристики)
    Public Shared Function FuncReadDataFieldsByTablePSToXML(ByVal putchFileXml As String, ByVal nameTablePS As String, ByRef dictionaryFields As Dictionary(Of String, String(,)), Optional ByRef deskTablePs As String = "") As Boolean
        FuncReadDataFieldsByTablePSToXML = False
        Dim boolTable As Boolean = False
        If File.Exists(putchFileXml) = True And nameTablePS.Trim.Length > 0 Then
            dictionaryFields = New Dictionary(Of String, String(,))
            'читаем xml и загружаем ветки с категориями
            Dim xDoc As XmlDocument = New XmlDocument()
            If IO.File.Exists(putchFileXml) = True Then
                Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
                While reader.Read()
                    Select Case reader.NodeType
                        Case XmlNodeType.Element
                            Dim NameBlock As String = reader.Name 'читаем ветку
                            Dim countArrayTablePSFields As Integer = 0
                            If NameBlock Like "PropertyTable" Then
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        Dim NameObject As String = reader.Value
                                        If nameN Like "TableName" Then
                                            If Not (NameObject Like nameTablePS) Then
                                                boolTable = False
                                                Exit While
                                            Else
                                                boolTable = True
                                            End If
                                        ElseIf nameN Like "TableDescription" Then
                                            deskTablePs = NameObject
                                        End If
                                    End While
                                End If
                            ElseIf NameBlock Like "Field" Then
                                If boolTable = True Then
                                    Dim arrayRez As String(,) = Nothing
                                    Dim countArrayRez As Integer = 0
                                    If reader.HasAttributes = True Then
                                        Dim countAttr As Integer = reader.AttributeCount - 1
                                        Dim nameField As String = ""
                                        While (reader.MoveToNextAttribute())
                                            Dim nameF As String = reader.Name
                                            Dim valueZn As String = reader.Value
                                            'записываем в массив значения
                                            ReDim Preserve arrayRez(1, countArrayRez)
                                            arrayRez(0, countArrayRez) = nameF
                                            arrayRez(1, countArrayRez) = valueZn
                                            countArrayRez += 1
                                            If nameF Like "Name" Then
                                                nameField = valueZn
                                            End If
                                        End While
                                        If IsArray(arrayRez) = True And nameField.Trim.Length > 0 Then
                                            dictionaryFields.Add(nameField, arrayRez)
                                        End If
                                    End If
                                End If
                            End If
                        Case XmlNodeType.EndElement
                            Dim NameBlock As String = reader.Name 'читаем ветку
                            If NameBlock Like "PropertyTables" Then
                                reader.Close()
                                If dictionaryFields.Count > 0 Then
                                    Return True
                                Else
                                    Return False
                                End If
                            ElseIf NameBlock Like "PropertyTable" Then
                                If dictionaryFields.Count > 0 Then
                                    reader.Close()
                                    Return True
                                End If
                            End If
                    End Select
                End While
                reader.Close()
            End If
        End If
    End Function

    'чтение данных XRecords из примитива
    Public Shared Function readXRecords(ByVal acEnt As Teigha.DatabaseServices.Entity, Optional tr As Transaction = Nothing) As Dictionary(Of String, CadastralTools.StructureElement)
        Dim listDataEntity As Dictionary(Of String, CadastralTools.StructureElement) = New Dictionary(Of String, CadastralTools.StructureElement)
        If IsNothing(acEnt) = True Then Return Nothing
        If IsNothing(tr) = True Then
            Dim doc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            tr = doc.TransactionManager.StartTransaction()
        End If
        If IsNothing(tr) = False Then
            Using tr
                Dim extId As ObjectId = acEnt.ExtensionDictionary
                If extId.IsNull = False Then
                    Dim userDict As DBDictionary = CType(tr.GetObject(extId, OpenMode.ForRead), DBDictionary)
                    If userDict.Count > 0 Then
                        Dim enumValues As Array = [Enum].GetValues(GetType(tableXRecords))
                        For Each value As tableXRecords In enumValues
                            ' Получаем имя элемента
                            Dim name As String = [Enum].GetName(GetType(tableXRecords), value)
                            Dim IdkeyDict As ObjectId = userDict.GetAt(name)
                            Dim dataXRecord As StructureElement = New StructureElement()
                            If IsNothing(IdkeyDict) = False Then
                                If IdkeyDict.IsNull = False Then
                                    Dim rec As Xrecord = IdkeyDict.GetObject(OpenMode.ForRead)
                                    Dim rb As ResultBuffer = rec.Data
                                    Dim tVal As TypedValue() = rb.AsArray
                                    If tVal.Count > 0 Then
                                        Dim countArray As Integer = 0
                                        For j As Integer = 0 To tVal.Count - 1
                                            Dim valStr As String = tVal(j).ToString
                                            valStr = valStr.Replace("(1000,", "")
                                            valStr = valStr.Remove(valStr.Length - 1)
                                            Dim valArrayStr As String() = valStr.ToString.Split("=")
                                            If IsArray(valArrayStr) = True Then
                                                Dim nameField As String = valArrayStr(0)
                                                Dim lenZn As Integer = valArrayStr(0).Trim.Length
                                                Dim valueZn As String = Mid(valStr, lenZn + 2)
                                                Select Case nameField
                                                    Case "Label"
                                                        dataXRecord.Label = valueZn
                                                    Case "ClassObject"
                                                        Try
                                                            dataXRecord.NameClassObject = CType([Enum].Parse(GetType(CadastralTools.StructureElement.classObject), valueZn), CadastralTools.StructureElement.classObject)
                                                        Catch ex As TypeLoadException
                                                            dataXRecord.NameClassObject = classObject.OtherElement
                                                        Catch ex As ArgumentException
                                                            dataXRecord.NameClassObject = classObject.OtherElement
                                                        End Try
                                                    Case "TypeGeometry"
                                                        Try
                                                            dataXRecord.TypeGeometry = CType([Enum].Parse(GetType(CadastralTools.StructureElement.typeGeometryObject), valueZn), CadastralTools.StructureElement.typeGeometryObject)
                                                        Catch ex As TypeLoadException
                                                            dataXRecord.TypeGeometry = typeGeometryObject.none
                                                        Catch ex As ArgumentException
                                                            dataXRecord.TypeGeometry = typeGeometryObject.none
                                                        End Try
                                                        dataXRecord.TypeGeometry = CType([Enum].Parse(GetType(CadastralTools.StructureElement.typeGeometryObject), valueZn), CadastralTools.StructureElement.typeGeometryObject)
                                                    Case "NameTypeObject"
                                                        dataXRecord.NameTypeObject = CType([Enum].Parse(GetType(CadastralTools.StructureElement.nameObject), valueZn), CadastralTools.StructureElement.nameObject)
                                                    Case "Description"
                                                        dataXRecord.Description = valueZn
                                                    Case "tablePS"
                                                        dataXRecord.TablePS = valueZn
                                                    Case "KeyParameter"
                                                        dataXRecord.KeyParameter = valueZn
                                                    Case "IdStructure"
                                                        dataXRecord.IdStructure = valueZn
                                                    Case "IdElement"
                                                        dataXRecord.IdElement = valueZn
                                                    Case "Note"
                                                        dataXRecord.Note = valueZn
                                                    Case "IdObject"
                                                        dataXRecord.IdObject = acEnt.Handle.ToString
                                                End Select
                                            End If
                                        Next j
                                        If listDataEntity.ContainsKey(name) = False Then
                                            listDataEntity.Add(name, dataXRecord)
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        Return listDataEntity
                    End If
                End If
            End Using
        End If
        Return Nothing
    End Function
    'запись данных XRecords в примитив
    Public Shared Function writeXRecords(ByRef entityObject As Entity, ByRef dataRecords As StructureElement, Optional ByVal tableXRecords As CadastralTools.StructureElement.tableXRecords = tableXRecords.PROJECT_STRUCTURES) As Boolean
        writeXRecords = False
        Dim nameTableXRecord As String = tableXRecords.ToString
        If IsNothing(entityObject) = False Then
            If IsNothing(entityObject) = True Then Return False
            If TypeOf entityObject Is Entity Then
                Dim doc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
                Dim db As Database = doc.Database
                Dim jsonString As String = dataRecords.KeyParameter
                Try
                    If CadastralTools.JsonValidator.IsValidJson(jsonString) = True Then
                        Dim jsonObject As JObject = JObject.Parse(jsonString)
                        ' Проверяем, существует ли поле ListCoord
                        If jsonObject("ListCoord") IsNot Nothing Then
                            ' Удаляем поле ListCoord
                            jsonObject.Remove("ListCoord")
                            jsonString = jsonObject.ToString
                        End If
                    End If
                Catch ex As JsonException
                Catch ex As Exception
                End Try
                Try
                    Dim tr As Transaction = doc.TransactionManager.StartTransaction()
                    Using tr
                        Dim entity As Entity = entityObject
                        Dim arrayData As String(,) = {}
                        ReDim arrayData(1, 10)
                        arrayData(0, 0) = "Label"
                        arrayData(1, 0) = dataRecords.Label

                        arrayData(0, 1) = "ClassObject"
                        arrayData(1, 1) = dataRecords.NameClassObject.ToString

                        arrayData(0, 2) = "NameTypeObject"
                        arrayData(1, 2) = dataRecords.NameTypeObject.ToString

                        arrayData(0, 3) = "Description"
                        arrayData(1, 3) = dataRecords.Description

                        arrayData(0, 4) = "KeyParameter"
                        arrayData(1, 4) = jsonString

                        arrayData(0, 5) = "tablePS"
                        arrayData(1, 5) = dataRecords.TablePS

                        arrayData(0, 6) = "IdStructure"
                        arrayData(1, 6) = dataRecords.IdStructure

                        arrayData(0, 7) = "IdElement"
                        arrayData(1, 7) = dataRecords.IdElement

                        arrayData(0, 8) = "Note"
                        arrayData(1, 8) = dataRecords.Note

                        arrayData(0, 9) = "IdObject"
                        arrayData(1, 9) = entity.Handle.ToString

                        arrayData(0, 10) = "TypeGeometry"
                        arrayData(1, 10) = dataRecords.TypeGeometry.ToString

                        Dim extId As ObjectId = entity.ExtensionDictionary
                        If extId.IsNull = True Then
                            entity.UpgradeOpen()
                            entity.CreateExtensionDictionary()
                            extId = entity.ExtensionDictionary
                            Dim userDict As DBDictionary = CType(tr.GetObject(extId, OpenMode.ForRead), DBDictionary)
                            If IsNothing(userDict) = False Then
                                userDict.UpgradeOpen()
                                Dim xRec As Xrecord = New Xrecord()
                                Dim rb As ResultBuffer = New ResultBuffer()
                                For i As Integer = 0 To arrayData.GetUpperBound(1)
                                    Dim key As String = arrayData(0, i)
                                    Dim val As String = arrayData(1, i)
                                    If IsNothing(key) = False Then
                                        If key.Trim.Length > 0 Then
                                            If IsNothing(val) = True Then val = ""
                                            Dim recStr As String = key & "=" & val
                                            rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, recStr))
                                        End If
                                    End If
                                Next i
                                xRec.Data = rb
                                userDict.SetAt(nameTableXRecord, xRec)
                                tr.AddNewlyCreatedDBObject(xRec, True)
                                tr.Commit()
                                Return True
                            End If
                        Else
                            'считываем данные
                            Dim userDict As DBDictionary = CType(tr.GetObject(extId, OpenMode.ForWrite), DBDictionary)
                            If IsNothing(userDict) = False Then
                                If userDict.Contains(nameTableXRecord) = True Then
                                    Dim xRecId As ObjectId = userDict.Item(nameTableXRecord)
                                    If IsNothing(xRecId) = False Then
                                        If xRecId.IsNull = False Then
                                            Dim xRec As Xrecord = tr.GetObject(xRecId, OpenMode.ForWrite)
                                            xRec.Erase()
                                            If IsArray(arrayData) Then
                                                xRec = New Xrecord()
                                                Dim rb As ResultBuffer = New ResultBuffer()
                                                For i As Integer = 0 To arrayData.GetUpperBound(1)
                                                    Dim key As String = arrayData(0, i)
                                                    Dim val As String = arrayData(1, i)
                                                    If IsNothing(key) = False Then
                                                        If key.Trim.Length > 0 Then
                                                            If IsNothing(val) = True Then val = ""
                                                            Dim recStr As String = key & "=" & val
                                                            rb.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, recStr))
                                                        End If
                                                    End If
                                                Next i
                                                xRec.Data = rb
                                                userDict.SetAt(nameTableXRecord, xRec)
                                                tr.AddNewlyCreatedDBObject(xRec, True)
                                                tr.Commit()
                                                Return True
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        Return False
                        'tr.Commit()
                    End Using
                Catch ex As System.Exception
                End Try
            End If
        End If
    End Function


    'чтение данных XRecords из примитива
    Public Shared Function getXRecords(ByVal entity As Teigha.DatabaseServices.Entity, ByRef dataXRecord As StructureElement, Optional tr As Transaction = Nothing, Optional nameTableXRecords As StructureElement.tableXRecords = StructureElement.tableXRecords.PROJECT_STRUCTURES) As Boolean
        getXRecords = False
        If IsNothing(entity) = True Then Return Nothing
        If IsNothing(tr) = True Then
            Dim doc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            tr = doc.TransactionManager.StartTransaction()
        End If
        If IsNothing(tr) = False Then
            Using tr
                Dim extId As ObjectId = entity.ExtensionDictionary
                If extId.IsNull = False Then
                    Dim userDict As DBDictionary = CType(tr.GetObject(extId, OpenMode.ForRead), DBDictionary)
                    If userDict.Count > 0 Then
                        Dim enumValues As Array = [Enum].GetValues(GetType(tableXRecords))
                        For Each value As tableXRecords In enumValues
                            ' Получаем имя элемента
                            If value = tableXRecords.PROJECT_STRUCTURES Then
                                Dim name As String = [Enum].GetName(GetType(tableXRecords), value)
                                Dim IdkeyDict As ObjectId = userDict.GetAt(name)
                                If IsNothing(IdkeyDict) = False Then
                                    If IdkeyDict.IsNull = False Then
                                        Dim rec As Xrecord = IdkeyDict.GetObject(OpenMode.ForRead)
                                        Dim rb As ResultBuffer = rec.Data
                                        Dim tVal As TypedValue() = rb.AsArray
                                        If tVal.Count > 0 Then
                                            Dim countArray As Integer = 0
                                            dataXRecord = New StructureElement
                                            For j As Integer = 0 To tVal.Count - 1
                                                Dim valStr As String = tVal(j).ToString
                                                valStr = valStr.Replace("(1000,", "")
                                                valStr = valStr.Remove(valStr.Length - 1)
                                                Dim valArrayStr As String() = valStr.ToString.Split("=")
                                                If IsArray(valArrayStr) = True Then
                                                    Dim nameField As String = valArrayStr(0)
                                                    Dim lenZn As Integer = valArrayStr(0).Trim.Length
                                                    Dim valueZn As String = Mid(valStr, lenZn + 2)
                                                    Select Case nameField
                                                        Case "Label"
                                                            dataXRecord.Label = valueZn
                                                        Case "ClassObject"
                                                            Try
                                                                dataXRecord.NameClassObject = CType([Enum].Parse(GetType(CadastralTools.StructureElement.classObject), valueZn), CadastralTools.StructureElement.classObject)
                                                            Catch ex As TypeLoadException
                                                                dataXRecord.NameClassObject = classObject.OtherElement
                                                            Catch ex As ArgumentException
                                                                dataXRecord.NameClassObject = classObject.OtherElement
                                                            End Try
                                                        Case "TypeGeometry"
                                                            Try
                                                                dataXRecord.TypeGeometry = CType([Enum].Parse(GetType(CadastralTools.StructureElement.typeGeometryObject), valueZn), CadastralTools.StructureElement.typeGeometryObject)
                                                            Catch ex As TypeLoadException
                                                                dataXRecord.TypeGeometry = typeGeometryObject.none
                                                            Catch ex As ArgumentException
                                                                dataXRecord.TypeGeometry = typeGeometryObject.none
                                                            End Try
                                                            dataXRecord.TypeGeometry = CType([Enum].Parse(GetType(CadastralTools.StructureElement.typeGeometryObject), valueZn), CadastralTools.StructureElement.typeGeometryObject)
                                                        Case "NameTypeObject"
                                                            dataXRecord.NameTypeObject = CType([Enum].Parse(GetType(CadastralTools.StructureElement.nameObject), valueZn), CadastralTools.StructureElement.nameObject)
                                                        Case "TableObject"
                                                            dataXRecord.TableObject = valueZn
                                                        Case "Description"
                                                            dataXRecord.Description = valueZn
                                                        Case "tablePS"
                                                            dataXRecord.TablePS = valueZn
                                                        Case "KeyParameter"
                                                            dataXRecord.KeyParameter = valueZn
                                                        Case "IdStructure"
                                                            dataXRecord.IdStructure = valueZn
                                                        Case "IdElement"
                                                            dataXRecord.IdElement = valueZn
                                                        Case "Note"
                                                            dataXRecord.Note = valueZn
                                                        Case "IdObject"
                                                            dataXRecord.IdObject = entity.Handle.ToString
                                                    End Select
                                                End If
                                            Next j
                                        End If
                                    End If
                                End If
                                Return True
                            End If
                        Next
                    End If
                End If
            End Using
        End If
    End Function

End Class
