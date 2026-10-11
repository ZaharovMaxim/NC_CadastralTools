Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Xml
Imports HostMgd.ApplicationServices
Imports Teigha.DatabaseServices
Imports Teigha.Geometry
Imports Teigha.LayerManager
Imports Teigha.Runtime

Public Class FuncXML
    'функция читает информацию о типе линии из xml и создает тип линии
    Public Shared Function FuncCreateLineTypeToXML(ByVal putchFileXml As String, ByVal nameLineTypeRecords As String) As LinetypeTableRecord
        FuncCreateLineTypeToXML = Nothing
        Dim deskTypeLine As String = ""
        Dim patternTypeLine As String = ""
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim boolFindLineType As Boolean = False
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "LinetypeRecord" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If nameN Like "NameLineType" Then
                                        If nameLineTypeRecords Like reader.Value Then
                                            boolFindLineType = True
                                        Else
                                            boolFindLineType = False
                                            Exit While
                                        End If
                                    ElseIf nameN Like "DeskLineType" Then
                                        deskTypeLine = reader.Value
                                    ElseIf nameN Like "CodeLineType" Then
                                        patternTypeLine = reader.Value
                                    End If
                                End While
                                If boolFindLineType = True Then
                                    Dim arrayProperties As String() = {}
                                    ReDim arrayProperties(2)
                                    arrayProperties(0) = nameLineTypeRecords
                                    arrayProperties(1) = deskTypeLine
                                    arrayProperties(2) = patternTypeLine
                                    Dim recTypeLine As LinetypeTableRecord = Nothing 'CadastralStyle.FuncCreateLineType(nameLineTypeRecords, deskTypeLine, patternTypeLine, putchFileXml)
                                    Return recTypeLine
                                End If
                            End If
                        End If
                    Case XmlNodeType.EndElement
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "LinetypeRecords" Then
                            Exit Function
                        End If
                End Select
            End While
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает информацию о характеристиках слоя из xml и создает слой 
    Public Shared Function FuncCreateLayerToXML(ByVal putchFileXml As String, ByVal nameLayer As String) As ObjectId
        FuncCreateLayerToXML = Nothing
        Dim indColor As Integer = 7
        Dim nameLineType As String = "Continuous"
        Dim idLineTypeRec As ObjectId = Nothing
        Dim widthLine As Double = 20
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim boolFindLayer As Boolean = False
            Dim boolFindLayers As Boolean = False
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Layer" And boolFindLayers = True Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If nameN Like "Name" Then
                                        If nameLayer Like reader.Value Then
                                            boolFindLayer = True
                                        Else
                                            boolFindLayer = False
                                            Exit While
                                        End If
                                    ElseIf nameN Like "Color" Then
                                        indColor = Val(reader.Value)
                                        If indColor < 1 Then
                                            indColor = 7
                                        ElseIf indColor > 255 Then
                                            indColor = 7
                                        End If
                                    ElseIf nameN Like "LineType" Then
                                        nameLineType = reader.Value
                                    ElseIf nameN Like "LineWidth" Then
                                        Dim tempWidthLine As String = reader.Value
                                        If tempWidthLine Like "ПоСлою" Then
                                            widthLine = -1
                                        ElseIf tempWidthLine Like "ПоБлоку" Then
                                            widthLine = -2
                                        ElseIf tempWidthLine Like "ByLineWeightDefault" Then
                                            widthLine = -3
                                        Else
                                            widthLine = Val(tempWidthLine)
                                        End If
                                    End If
                                End While
                                If boolFindLayer = True Then
                                    'idLineTypeRec = AcadFunction.FuncGetLineTypeId(nameLineType)
                                    If IsNothing(idLineTypeRec) = True OrElse idLineTypeRec.IsNull = True Then
                                        Dim recTypeLine As LinetypeTableRecord = FuncCreateLineTypeToXML(putchFileXml, nameLineType)
                                        idLineTypeRec = recTypeLine.Id
                                    End If
                                    Dim arrayParam As String(,) = {}
                                    ReDim arrayParam(1, 3)
                                    arrayParam(0, 0) = "Name"
                                    arrayParam(1, 0) = nameLayer

                                    arrayParam(0, 0) = "Color"
                                    arrayParam(1, 0) = indColor

                                    arrayParam(0, 0) = "LineType"
                                    arrayParam(1, 0) = nameLineType

                                    arrayParam(0, 0) = "LineWidth"
                                    arrayParam(1, 0) = widthLine
                                    Dim Layer As LayerTableRecord = Nothing 'CadastralStyle.FuncCreateLayer(arrayParam)
                                    Return Layer.Id
                                End If
                            End If
                        ElseIf NameBlock Like "Layers" Then
                            boolFindLayers = True
                        End If
                    Case XmlNodeType.EndElement
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "LinetypeRecords" Then
                            Exit Function
                        End If
                End Select
            End While
        End If
    End Function
    ' '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает информацию о характеристиках стиля текста из xml и создает стиль текста
    Public Shared Function FuncCreateTextStyleToXML(ByVal putchFileXml As String, ByVal nameTextStyle As String) As ObjectId
        FuncCreateTextStyleToXML = Nothing
        Dim boolItalic As Boolean = True
        Dim NameFont As String = "Arial"
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim boolFindTextStyle As Boolean = False
            Dim boolFindTextStyles As Boolean = False
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "TextStyle" And boolFindTextStyles = True Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If nameN Like "NameTextStyle" Then
                                        If nameTextStyle Like reader.Value Then
                                            boolFindTextStyle = True
                                        Else
                                            boolFindTextStyle = False
                                            Exit While
                                        End If
                                    ElseIf nameN Like "NameFont" Then
                                        NameFont = reader.Value
                                    ElseIf nameN Like "Italic" Then
                                        If reader.Value Like "True" Then
                                            boolItalic = True
                                        Else
                                            boolItalic = False
                                        End If
                                    End If
                                End While
                                If boolFindTextStyle = True Then
                                    Dim idTS As ObjectId = Nothing
                                    'AcadFunction.FuncCreateTextStyle(nameTextStyle, NameFont, boolItalic)
                                    Return idTS
                                End If
                            End If
                        ElseIf NameBlock Like "TextStyles" Then
                            boolFindTextStyles = True
                        End If
                    Case XmlNodeType.EndElement
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "TextStyles" Then
                            Exit Function
                        End If
                End Select
            End While
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает стоки таблицы PS из xml (Имя категории, Имя таблицы, Тип объекта, Описание объекта, результат - массив с Характеристиками)
    Public Shared Function FuncReadDataTablePSToXML(ByVal putchFileXml As String, ByVal namePSTable As String, ByRef arrayDictionaryPS As Dictionary(Of Integer, String(,))) As String
        FuncReadDataTablePSToXML = False
        Dim countRec As Integer = 1
        arrayDictionaryPS.Clear()
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim boolTablePS As Boolean = False
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "PropertyTable" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If nameN Like "TableName" Then
                                        If namePSTable Like reader.Value Then
                                            boolTablePS = True
                                        Else
                                            boolTablePS = False
                                        End If
                                    ElseIf nameN Like "TableDescription" Then
                                        If boolTablePS = True Then
                                            FuncReadDataTablePSToXML = reader.Value
                                        End If
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "Field" Then
                            If boolTablePS = True Then
                                If reader.HasAttributes = True Then
                                    Dim arrayFieldsPS As String(,) = Nothing
                                    Dim countFieldsPS As Integer = 0
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        Dim val As String = reader.Value
                                        ReDim Preserve arrayFieldsPS(1, countFieldsPS)
                                        arrayFieldsPS(0, countFieldsPS) = nameN
                                        arrayFieldsPS(1, countFieldsPS) = val
                                        countFieldsPS += 1
                                    End While
                                    If IsArray(arrayFieldsPS) = True Then
                                        arrayDictionaryPS.Add(countRec, arrayFieldsPS)
                                        countRec += 1
                                    End If
                                End If
                            End If
                        End If
                    Case XmlNodeType.EndElement
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "PropertyTable" Then
                            If IsArray(arrayDictionaryPS) = True Then
                                Exit Function
                            End If
                        End If
                End Select
            End While
        End If
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'функция читает определенную таблицу из файла xml (Имя категории, Имя таблицы, Тип объекта, результат - массив с Характеристиками)
    Public Shared Function FuncReadDataTableToXml(ByVal putchFileXml As String, ByVal category As String, ByVal table As String, ByVal typeObject As String, ByRef dataObjectProperties As String(,)) As Integer
        FuncReadDataTableToXml = 0
        Dim countDataObjectProperties As Integer = 0
        If IsArray(dataObjectProperties) = True Then
            countDataObjectProperties = dataObjectProperties.GetUpperBound(1)
        End If
        'читаем xml и загружаем ветки с категориями
        Dim xDoc As XmlDocument = New XmlDocument()
        If IO.File.Exists(putchFileXml) = True Then
            Dim boolCategory As Boolean = False
            Dim boolTable As Boolean = False
            Dim boolType As Boolean = False
            Dim reader As XmlTextReader = New XmlTextReader(putchFileXml)
            While reader.Read()
                Select Case reader.NodeType
                    Case XmlNodeType.Element
                        Dim NameBlock As String = reader.Name 'читаем ветку
                        If NameBlock Like "Category" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If category Like reader.Value Then
                                        boolCategory = True
                                    Else
                                        boolCategory = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "Tables" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If table Like reader.Value Then
                                        boolTable = True
                                    Else
                                        boolTable = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "TypeObject" Then
                            If reader.HasAttributes = True Then
                                While (reader.MoveToNextAttribute())
                                    Dim nameN As String = reader.Name
                                    If typeObject Like reader.Value Then
                                        boolType = True
                                    Else
                                        boolType = False
                                    End If
                                End While
                            End If
                        ElseIf NameBlock Like "Rows" Then
                            Dim boolDataFind As Boolean = False
                            If boolTable = True And boolCategory = True And boolType = True Then
                                If reader.HasAttributes = True Then
                                    While (reader.MoveToNextAttribute())
                                        Dim nameN As String = reader.Name
                                        Dim NameObject As String = reader.Value
                                        ReDim Preserve dataObjectProperties(1, countDataObjectProperties)
                                        dataObjectProperties(0, countDataObjectProperties) = nameN
                                        dataObjectProperties(1, countDataObjectProperties) = NameObject
                                        countDataObjectProperties += 1
                                    End While
                                End If
                            End If
                        End If
                    Case XmlNodeType.Text
                        Dim a As String = ""
                    Case XmlNodeType.EndElement
                        Dim a As String = ""
                End Select
            End While
        End If
        FuncReadDataTableToXml = countDataObjectProperties
    End Function
End Class
