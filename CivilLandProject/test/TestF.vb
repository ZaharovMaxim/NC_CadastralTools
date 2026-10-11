Imports Teigha.Runtime

Public Class TestF
    Private Function readArray(ByRef data As String()) As String
        Dim result As String = "Результат: " & vbLf
        If IsArray(data) = True Then
            For i As Integer = 0 To data.Length - 1
                result = result & data(i).ToString & vbLf
            Next i
        End If
        Return result
    End Function
    Private Function readArray2d(ByRef data As String(,), Optional ByVal indexField As Integer = 0, Optional ByVal indexValue As Integer = 1) As String
        Dim result As String = "Результат: " & vbLf
        If IsArray(data) = True Then
            For i As Integer = 0 To data.GetUpperBound(1)
                result = result & "Имя поля: " & data(indexField, i).ToString & " ,Значение: " & data(indexValue, i).ToString & vbLf
            Next i
        End If
        Return result
    End Function

    <CommandMethod("SReadObjectInfoXML")>
    Public Sub SReadObjectInfoXML()
        Dim classCivilUtility As New CadastralTools.FuncNetTopologySuite()
        Dim dictTemplatesXML As Dictionary(Of String, String) = CadastralTools.FuncTemplateXML.readTemplatesXML()
        If dictTemplatesXML.Count > 0 Then
            Dim putchTeplateXML As String = dictTemplatesXML.ElementAt(0).Value
            Dim arrayCategories As String() = {}
            Dim boolReadCategories As Boolean = CadastralTools.FuncTemplateXML.FuncReadCategoryToXML(putchTeplateXML, arrayCategories)
            MsgBox(readArray(arrayCategories))
            Dim nameCategory As String = arrayCategories(0)
            Dim arrayTables As String() = {}
            Dim boolReadTables As Boolean = CadastralTools.FuncTemplateXML.FuncReadTablesToXML(putchTeplateXML, nameCategory, arrayTables)
            MsgBox(readArray(arrayTables))
            Dim nameTables As String = arrayTables(0)
            Dim arrayTypeObject As String() = {}
            Dim boolReadTypeObject As Boolean = CadastralTools.FuncTemplateXML.FuncReadTypesObjectToXML(putchTeplateXML, nameCategory, nameTables, arrayTypeObject)
            MsgBox(readArray(arrayTypeObject))
            Dim nameTypeObject As String = arrayTypeObject(0)
            Dim arrayObject As String() = {}
            Dim boolReadObject As Boolean = CadastralTools.FuncTemplateXML.FuncReadObjectToXML(putchTeplateXML, nameCategory, nameTables, nameTypeObject, arrayObject)
            MsgBox(readArray(arrayObject))
            Dim nameObject As String = arrayObject(0)
            Dim arrayPropertiesObject As String(,) = {}
            Dim boolPropertiesReadObject As Boolean = CadastralTools.FuncTemplateXML.FuncReadPropertiesObjectToXML(putchTeplateXML, nameCategory, nameTables, nameTypeObject, nameObject, arrayPropertiesObject)
            MsgBox(readArray2d(arrayPropertiesObject))
        Else
            MsgBox("Шаблон xml не найден")
        End If
    End Sub

    <CommandMethod("SReadOtherInfoXML")>
    Public Sub SReadOtherInfoXML()
        Dim classCivilUtility As New CadastralTools.FuncNetTopologySuite()
        Dim dictTemplatesXML As Dictionary(Of String, String) = CadastralTools.FuncTemplateXML.readTemplatesXML()
        If dictTemplatesXML.Count > 0 Then
            Dim putchTeplateXML As String = dictTemplatesXML.ElementAt(0).Value
            Dim arrayLayerProperties As String(,) = {}
            Dim nameLayer As String = "ИИ-ТОПО-Водопропускные трубы"
            Dim boolReadProperties As Boolean = CadastralTools.FuncTemplateXML.FuncReadLayerPropertiesInfoToXML(putchTeplateXML, nameLayer, arrayLayerProperties)
            MsgBox(readArray2d(arrayLayerProperties))
            Dim arrayLineTypeProperties As String(,) = {}
            Dim nameLineType As String = "A-Брод"
            boolReadProperties = CadastralTools.FuncTemplateXML.FuncReadLineTypeInfoToXML(putchTeplateXML, nameLineType, arrayLineTypeProperties)
            MsgBox(readArray2d(arrayLineTypeProperties))
            Dim arrayTextStuleProperties As String(,) = {}
            Dim nameTextStyle As String = "Д-431"
            boolReadProperties = CadastralTools.FuncTemplateXML.FuncReadTextStylePropertiesInfoToXML(putchTeplateXML, nameTextStyle, arrayTextStuleProperties)
            MsgBox(readArray2d(arrayTextStuleProperties))
            Dim arrayTablePSProperties As String(,) = {}
            Dim nameTablePS As String = "PROJECT_TREE"
            boolReadProperties = CadastralTools.FuncTemplateXML.FuncReadTablePSToXML(putchTeplateXML, nameTablePS, arrayTablePSProperties)
            MsgBox(readArray2d(arrayTablePSProperties))
        Else
            MsgBox("Шаблон xml не найден")
        End If
    End Sub

End Class
