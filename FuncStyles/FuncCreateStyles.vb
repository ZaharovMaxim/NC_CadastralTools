Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Windows
Imports System.Xml
Imports CadastralTools.StructureElement
Imports HostMgd.ApplicationServices
Imports HostMgd.EditorInput
Imports HostMgd.Windows
Imports Teigha.DatabaseServices
Imports Teigha.Geometry
Imports Teigha.GraphicsInterface
Imports Teigha.Runtime
Public Class StyleCadastralObject
    Private _color As Integer
    Private _layer As String
    Private _typeLine As String
    Private _scaleTypeLine As Double
    Private _weightLine As Double
    Private _db As Database = Nothing
    Private _templateXml As String
    Private _templateDwg As String
    Public Sub New()
        _layer = "0"
        _typeLine = "Continuous"
        _color = 7
        _scaleTypeLine = 1
        _weightLine = 20
    End Sub

    Public Property Color As Integer
        Get
            Return _color
        End Get
        Set(value As Integer)
            _color = value
        End Set
    End Property

    Public Property NameLayer As String
        Get
            Return _layer
        End Get
        Set(value As String)
            _layer = value
        End Set
    End Property

    Public Property TypeLine As String
        Get
            Return _typeLine
        End Get
        Set(value As String)
            _typeLine = value
        End Set
    End Property

    Public Property ScaleTypeLine As Double
        Get
            Return _scaleTypeLine
        End Get
        Set(value As Double)
            _scaleTypeLine = value
        End Set
    End Property

    Public Property WeightLine As Double
        Get
            Return _weightLine
        End Get
        Set(value As Double)
            _weightLine = value
        End Set
    End Property

    Public Property DBDrawing As Database
        Get
            Return _db
        End Get
        Set(value As Database)
            _db = value
        End Set
    End Property

    Public Property TemplateXML As String
        Get
            Return _templateXml
        End Get
        Set(value As String)
            _templateXml = value
        End Set
    End Property

    Public Property TemplateDWG As String
        Get
            Return _templateDwg
        End Get
        Set(value As String)
            _templateDwg = value
        End Set
    End Property
    'функция для вставки нового слоя
    Public Function createLayer(ByVal LayerStr As String, Optional ByVal color1 As Short = Nothing, Optional ByVal idlinetype As ObjectId = Nothing, Optional ByVal wightline As Integer = Nothing) As ObjectId
        createLayer = Nothing
        If IsNothing(DBDrawing) = False Then
            Using acTrans As Transaction = DBDrawing.TransactionManager.StartTransaction()
                Dim acLyrTbl As LayerTable
                acLyrTbl = acTrans.GetObject(DBDrawing.LayerTableId, OpenMode.ForRead)
                Try
                    If acLyrTbl.Has(LayerStr) = False Then
                        Dim acLyrTblRec As LayerTableRecord = New LayerTableRecord()
                        acLyrTblRec.Name = LayerStr 'имя слоя
                        'цвет слоя
                        If IsNothing(color1) = False Then
                            Dim color2 As Teigha.Colors.Color = Teigha.Colors.Color.FromColorIndex(Teigha.Colors.ColorMethod.ByAci, color1)
                            acLyrTblRec.Color = color2
                        End If
                        'тип линии
                        If IsNothing(idlinetype) = False Then
                            acLyrTblRec.LinetypeObjectId = idlinetype
                        End If
                        'вес линии
                        If IsNothing(wightline) = False Then
                            acLyrTblRec.LineWeight = wightline
                        End If
                        acLyrTbl.UpgradeOpen()
                        acLyrTbl.Add(acLyrTblRec)
                        acTrans.AddNewlyCreatedDBObject(acLyrTblRec, True)
                        Return acLyrTblRec.Id
                    Else
                        Return acLyrTbl.Item(LayerStr)
                    End If
                Catch ex As Teigha.Runtime.Exception
                End Try
                acTrans.Commit()
            End Using
        End If
    End Function
    Public Function createLayerByArrayProperties(ByRef arrayProperties As String(,)) As LayerTableRecord
        Dim result As LayerTableRecord = Nothing
        If IsNothing(DBDrawing) = False Then
            If IsArray(arrayProperties) = True Then
                Dim nameLayer As String = "0"
                Dim colorLayer As Integer = 7
                Dim lineTypeLayer As String = "Continuous"
                Dim lineWightLayer As Integer = 20
                For i As Integer = 0 To arrayProperties.GetUpperBound(1)
                    Dim nameFields As String = arrayProperties(0, i)
                    Dim valFields As String = arrayProperties(1, i)
                    If IsNothing(nameFields) = True Then Continue For
                    nameFields = nameFields.Trim
                    If IsNothing(valFields) = True Then Continue For
                    valFields = valFields.Trim
                    If nameFields Like "Name" Then
                        nameLayer = valFields
                    ElseIf nameFields Like "Color" Then
                        colorLayer = Val(valFields)
                        If colorLayer < 0 Then colorLayer = 0
                        If colorLayer > 255 Then colorLayer = 7
                    ElseIf nameFields Like "LineType" Then
                        lineTypeLayer = valFields
                    ElseIf nameFields Like "LineWidth" Then
                        lineWightLayer = Val(valFields)
                    End If
                Next i
                If nameLayer.Trim.Length > 0 Then
                    'Start a transaction
                    Using tr As Transaction = DBDrawing.TransactionManager.StartTransaction()
                        Dim acLyrTbl As LayerTable = tr.GetObject(DBDrawing.LayerTableId, OpenMode.ForRead)
                        Try
                            If acLyrTbl.Has(nameLayer) = False Then
                                result = New LayerTableRecord()
                                result.Name = nameLayer 'имя слоя
                                'цвет слоя
                                Dim color2 As Teigha.Colors.Color = Teigha.Colors.Color.FromColorIndex(Teigha.Colors.ColorMethod.ByAci, colorLayer)
                                result.Color = color2
                                'тип линии
                                Dim acLineTypeTbl As LinetypeTable = tr.GetObject(DBDrawing.LinetypeTableId, OpenMode.ForRead)
                                If acLineTypeTbl.Has(lineTypeLayer) = True Then
                                    result.LinetypeObjectId = acLineTypeTbl.Item(lineTypeLayer)
                                End If
                                'вес линии
                                result.LineWeight = lineWightLayer
                                If acLyrTbl.IsWriteEnabled = False Then
                                    acLyrTbl.UpgradeOpen()
                                End If
                                acLyrTbl.Add(result)
                                tr.AddNewlyCreatedDBObject(result, True)
                            Else
                                Dim idLayer As ObjectId = acLyrTbl.Item(nameLayer)
                                If IsNothing(idLayer) = True Then
                                    If idLayer.IsNull = False Then
                                        result = idLayer.GetObject(OpenMode.ForWrite)
                                    End If
                                End If
                            End If
                        Catch ex As Teigha.Runtime.Exception
                            Return Nothing
                        End Try
                        tr.Commit()
                    End Using
                End If
            End If
        End If
        Return result
    End Function
    'создать текстовый стиль
    Public Function createTextStyle(ByVal NameTextStyle As String, ByVal NameFont As String, ByVal NameItalic As Boolean) As ObjectId
        Dim result As ObjectId = Nothing
        If IsNothing(DBDrawing) Then
            Dim tr As Transaction = DBDrawing.TransactionManager.StartTransaction()
            Try
                Using tr
                    Dim StText As TextStyleTable = tr.GetObject(DBDrawing.TextStyleTableId, OpenMode.ForWrite)
                    If NameTextStyle.Trim.Length > 0 Then
                        If StText.Has(NameTextStyle) = True Then
                            Dim idTextStyle As ObjectId = StText.Item(NameTextStyle)
                            result = idTextStyle
                        Else
                            Dim TextStyleLine As TextStyleTableRecord = New TextStyleTableRecord()
                            TextStyleLine.Name = NameTextStyle
                            Dim acFont As Teigha.GraphicsInterface.FontDescriptor = TextStyleLine.Font
                            Dim acNewFont As Teigha.GraphicsInterface.FontDescriptor
                            If NameFont.Trim.Length = 0 Then
                                For Each fontFamily As System.Windows.Media.FontFamily In System.Windows.Media.Fonts.SystemFontFamilies
                                    If fontFamily.Source Like "Arial" Then
                                        NameFont = fontFamily.Source
                                        Exit For
                                    End If
                                    NameFont = fontFamily.Source
                                Next
                            End If
                            If NameItalic = True Then
                                acNewFont = New Teigha.GraphicsInterface.FontDescriptor(NameFont, acFont.Bold = True, acFont.Italic = True, acFont.CharacterSet, acFont.PitchAndFamily)
                            Else
                                acNewFont = New Teigha.GraphicsInterface.FontDescriptor(NameFont, acFont.Bold = True, acFont.Italic = False, acFont.CharacterSet, acFont.PitchAndFamily)
                            End If
                            TextStyleLine.Font = acNewFont
                            Dim idTextStyle As ObjectId = StText.Add(TextStyleLine)
                            tr.AddNewlyCreatedDBObject(TextStyleLine, True)
                            result = idTextStyle
                        End If
                    Else
                        Dim idTextStyle As ObjectId = GetObjectIdOfTextStyleWithFileName(StText, NameFont)
                        If idTextStyle.IsNull Then
                            Dim tstr As TextStyleTableRecord = New TextStyleTableRecord()
                            tstr.IsShapeFile = True
                            tstr.FileName = NameFont
                            idTextStyle = StText.Add(tstr)
                            tr.AddNewlyCreatedDBObject(tstr, True)
                        End If
                        result = idTextStyle
                    End If
                    tr.Commit()
                End Using
            Catch ex As System.Exception
            End Try
        End If
        Return result
    End Function
    Public Function createTextStyleByArrayProperties(ByRef arrayProperties As String(,)) As ObjectId
        Dim result As ObjectId = Nothing
        If IsNothing(DBDrawing) = False Then
            If IsArray(arrayProperties) = True Then
                Dim nameTextStyle As String = "CADASTRAL_OBJECT"
                Dim nameFont As String = "arial.ttf"
                Dim NameItalic As Boolean = False
                Dim anno As Boolean = False
                Dim TextStyleHeight As Double = 0
                Dim TextStyleRotate As Double = 0
                Dim TextStylePrior As Double = 2.5
                For i As Integer = 0 To arrayProperties.GetUpperBound(1)
                    Dim nameFields As String = arrayProperties(0, i)
                    Dim valFields As String = arrayProperties(1, i)
                    If IsNothing(nameFields) = True Then Continue For
                    nameFields = nameFields.Trim
                    If IsNothing(valFields) = True Then Continue For
                    valFields = valFields.Trim
                    If nameFields Like "NameTextStyle" Then
                        nameTextStyle = valFields
                    ElseIf nameFields Like "NameFont" Then
                        nameFont = valFields
                    ElseIf nameFields Like "Italic" Then
                        NameItalic = CBool(valFields)
                    ElseIf nameFields Like "Annotative" Then
                        anno = CBool(valFields)
                    ElseIf nameFields Like "TextStyleHeight" Then
                        TextStyleHeight = Val(valFields)
                    ElseIf nameFields Like "TextStyleRotate" Then
                        TextStyleRotate = Val(valFields)
                    ElseIf nameFields Like "TextStylePrior" Then
                        TextStylePrior = Val(valFields)
                    End If
                Next i
                If nameTextStyle.Trim.Length > 0 Then
                    'Start a transaction
                    Using tr As Transaction = DBDrawing.TransactionManager.StartTransaction()
                        Dim StText As TextStyleTable = tr.GetObject(DBDrawing.TextStyleTableId, OpenMode.ForWrite)
                        Try
                            If StText.Has(nameTextStyle) = False Then
                                Dim TextStyleLine As TextStyleTableRecord = New TextStyleTableRecord()
                                TextStyleLine.Name = nameTextStyle
                                Dim acFont As Teigha.GraphicsInterface.FontDescriptor = TextStyleLine.Font
                                Dim acNewFont As Teigha.GraphicsInterface.FontDescriptor
                                If nameFont.Trim.Length = 0 Then
                                    For Each fontFamily As System.Windows.Media.FontFamily In System.Windows.Media.Fonts.SystemFontFamilies
                                        If fontFamily.Source Like "Arial" Then
                                            nameFont = fontFamily.Source
                                            Exit For
                                        End If
                                        nameFont = fontFamily.Source
                                    Next
                                End If
                                If NameItalic = True Then
                                    acNewFont = New Teigha.GraphicsInterface.FontDescriptor(nameFont, acFont.Bold = True, acFont.Italic = True, acFont.CharacterSet, acFont.PitchAndFamily)
                                Else
                                    acNewFont = New Teigha.GraphicsInterface.FontDescriptor(nameFont, acFont.Bold = True, acFont.Italic = False, acFont.CharacterSet, acFont.PitchAndFamily)
                                End If
                                TextStyleLine.Font = acNewFont
                                TextStyleLine.TextSize = TextStyleHeight
                                TextStyleLine.ObliquingAngle = TextStyleRotate
                                TextStyleLine.PriorSize = TextStylePrior
                                Dim idTextStyle As ObjectId = StText.Add(TextStyleLine)
                                tr.AddNewlyCreatedDBObject(TextStyleLine, True)
                                result = idTextStyle
                            Else
                                Dim idTextStyle As ObjectId = StText.Item(nameTextStyle)
                                If IsNothing(idTextStyle) = True Then
                                    If idTextStyle.IsNull = False Then
                                        result = idTextStyle
                                    End If
                                End If
                            End If
                        Catch ex As Teigha.Runtime.Exception
                            Return Nothing
                        End Try
                        tr.Commit()
                    End Using
                End If
            End If
        End If
        Return result
    End Function
    'Функция позволяет найти текстовый стиль по строке пути к файлу форм (вспомогательная функция)
    Public Function GetObjectIdOfTextStyleWithFileName(ByVal tst As TextStyleTable, ByVal shxFileName As String) As ObjectId
        Dim result As ObjectId = Nothing
        Try
            Using Tr As Transaction = DBDrawing.TransactionManager.StartTransaction()
                Dim FullShxFileName As String = HostApplicationServices.Current.FindFile(shxFileName, tst.Database, FindFileHint.FontFile).ToLower()
                For Each idtsr As ObjectId In tst
                    Dim tstr As TextStyleTableRecord = TryCast(Tr.GetObject(idtsr, OpenMode.ForRead), TextStyleTableRecord)
                    If Not tstr.IsShapeFile Then Continue For
                    Dim filename As String = tstr.FileName.ToLower()
                    Try
                        filename = HostApplicationServices.Current.FindFile(filename, tst.Database, FindFileHint.FontFile).ToLower()
                    Catch
                    End Try
                    If FullShxFileName = filename Then
                        result = tstr.ObjectId
                    End If
                Next
            End Using
        Catch ex As System.Exception
        End Try
        Return result
    End Function
    'функция создает тип линии по коду
    Public Function createTypeLineByCode(ByVal nameTypeLine As String, ByVal deskTypeLine As String, ByVal patternStr As String) As LinetypeTableRecord
        Dim result As LinetypeTableRecord = Nothing
        Dim patternLine As String() = patternStr.Split(",")
        Dim supportPatchFolder As String() = Nothing
        Dim boolFindSupport As Boolean = FuncFiles.FuncGetSupportFolderPath(supportPatchFolder)
        If IsArray(patternLine) = True Then
            If patternLine.Length > 1 Then
                Dim lenghtPattern As Double = 0
                Dim numDashes As Integer = 0
                For i As Integer = 0 To patternLine.Length - 1
                    Dim valZnStr As String = patternLine(i)
                    If IsNumeric(valZnStr) Then
                        Dim zn As Double = Math.Abs(CDbl(valZnStr))
                        If zn > 0 Then
                            lenghtPattern = lenghtPattern + zn
                            numDashes += 1
                        End If
                    End If
                Next i
                If IsNothing(DBDrawing) = False Then
                    Dim tr As Transaction = DBDrawing.TransactionManager.StartTransaction()
                    Using tr
                        Dim lt As LinetypeTable = CType(tr.GetObject(DBDrawing.LinetypeTableId, OpenMode.ForWrite), LinetypeTable)
                        Dim acTextStyleTbl As TextStyleTable = tr.GetObject(DBDrawing.TextStyleTableId, OpenMode.ForRead)
                        Dim idTextStyle As ObjectId = Nothing
                        If lt.Has(nameTypeLine) = True Then
                            Dim idTypeLine As ObjectId = lt.Item(nameTypeLine)
                            result = idTypeLine.GetObject(OpenMode.ForRead)
                        Else
                            Dim ltr As LinetypeTableRecord = New LinetypeTableRecord()
                            ltr.Name = nameTypeLine
                            ltr.AsciiDescription = deskTypeLine
                            ltr.NumDashes = numDashes
                            ltr.PatternLength = lenghtPattern
                            Dim index As Integer = 0

                            For i As Integer = 1 To patternLine.Length - 1
                                Dim tempstr As String = patternLine(i).Trim
                                If IsNumeric(tempstr) = True Then
                                    ltr.SetDashLengthAt(index, Val(tempstr))
                                    index += 1
                                Else
                                    '==================================================================
                                    'сложный тип линии
                                    If Left(tempstr, 1) Like "[[]" Then
                                        tempstr = Mid(tempstr, 2)
                                        tempstr = tempstr.Replace("""", "")
                                        Dim indexForm As Integer = 0
                                        Dim namefile As String = ""
                                        Dim S As Double = 1
                                        Dim R As Double = 0
                                        Dim X As Double = 0
                                        Dim Y As Double = 0
                                        If index > 0 Then
                                            index -= 1
                                        End If

                                        For j As Integer = i + 1 To patternLine.Length - 1
                                            Dim tempStr2 As String = patternLine(j).Trim
                                            Dim boolEnd As Boolean = False
                                            If Right(tempStr2, 1) Like "[]]" Then
                                                boolEnd = True
                                                tempStr2 = Mid(tempStr2, 1, tempStr2.Length - 2)
                                            End If
                                            If Left(tempStr2, 2) Like "S=" Or Left(tempStr2, 2) Like "s=" Then
                                                S = Val(Mid(tempStr2, 3))
                                            ElseIf Left(tempStr2, 2) Like "R=" Or Left(tempStr2, 2) Like "r=" Then
                                                R = Val(Mid(tempStr2, 3))
                                            ElseIf Left(tempStr2, 2) Like "U=" Or Left(tempStr2, 2) Like "u=" Then
                                                R = Val(Mid(tempStr2, 3))
                                            ElseIf Left(tempStr2, 2) Like "X=" Or Left(tempStr2, 2) Like "x=" Then
                                                X = Val(Mid(tempStr2, 3))
                                            ElseIf Left(tempStr2, 2) Like "Y=" Or Left(tempStr2, 2) Like "y=" Then
                                                Y = Val(Mid(tempStr2, 3))
                                            Else
                                                namefile = tempStr2
                                                If Right(namefile, 4) Like ".shx" Then
                                                    For Each id As ObjectId In acTextStyleTbl
                                                        Dim recTextStyle As TextStyleTableRecord = id.GetObject(OpenMode.ForRead)
                                                        If recTextStyle.FileName Like namefile Then
                                                            idTextStyle = id
                                                            Exit For
                                                        End If
                                                    Next
                                                    If IsNothing(idTextStyle) = True OrElse idTextStyle.IsNull = True Then
                                                        idTextStyle = createTextStyle("", namefile, False)
                                                    End If
                                                    Dim FindFile As String = namefile.Replace("shx", "shp")
                                                    If IsArray(supportPatchFolder) = True Then
                                                        For k As Integer = 0 To supportPatchFolder.Length - 1
                                                            Dim namefiletemp As String = supportPatchFolder(k) & "\" & FindFile
                                                            If File.Exists(namefiletemp) = True Then
                                                                'ищем индекс формы
                                                                Dim inputshp As StreamReader = New StreamReader(namefiletemp, True)
                                                                Dim line2 As String
                                                                Dim regMatch As String = "\*"
                                                                Do Until inputshp.EndOfStream
                                                                    line2 = inputshp.ReadLine().Trim 'считываем строку
                                                                    If line2.Length > 0 Then
                                                                        Dim flagread As String = Left(line2, 1)
                                                                        If Regex.IsMatch(flagread, regMatch) Then 'chr(*)
                                                                            Dim arrayShp As String() = line2.Split(",")
                                                                            If IsArray(arrayShp) = True Then
                                                                                If arrayShp.GetUpperBound(0) > 1 Then
                                                                                    If arrayShp(2).Trim Like tempstr Then
                                                                                        Dim StrNumb As String = arrayShp(0).Trim
                                                                                        If StrNumb.Length > 1 Then
                                                                                            Dim Numbshp As Integer = Val(Mid(StrNumb, 2))
                                                                                            If Numbshp > 0 Then
                                                                                                indexForm = Numbshp
                                                                                                Exit Do
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Loop
                                                                inputshp.Close()
                                                            End If
                                                        Next k
                                                    End If
                                                Else
                                                    'ищем текстовый стиль
                                                    If acTextStyleTbl.Has(namefile) = True Then
                                                        idTextStyle = acTextStyleTbl.Item(namefile)
                                                    Else
                                                        idTextStyle = createTextStyle(namefile, "", False)
                                                    End If
                                                End If
                                            End If
                                            If boolEnd = True Then
                                                i = j
                                                Exit For
                                            End If
                                        Next j
                                        If indexForm > 0 Then
                                            ltr.SetShapeNumberAt(index, indexForm)
                                        Else
                                            ltr.SetTextAt(index, tempstr)
                                        End If
                                        ltr.SetShapeStyleAt(index, idTextStyle)
                                        ltr.SetShapeScaleAt(index, Val(S))
                                        ltr.SetShapeRotationAt(index, Val(R))
                                        ltr.SetShapeOffsetAt(index, New Vector2d(X, Y))
                                        index += 1
                                        idTextStyle = Nothing
                                    End If
                                End If
                            Next i
                            Dim ltId As ObjectId = lt.Add(ltr)
                            tr.AddNewlyCreatedDBObject(ltr, True)
                            result = ltr
                        End If
                        tr.Commit()
                    End Using
                End If
            End If
        End If
        Return result
    End Function
End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'класс выноски
Public Class CadastralLeaderStyle
    Inherits StyleCadastralObject
    Private _idStyle As ObjectId
    Private _textStyle As String
    Private _heightText As Double
    Private _sizeArraow As Double
    Public Sub New()
        MyBase.New()
        _idStyle = Nothing
        _textStyle = ""
        _heightText = 2
        _sizeArraow = 2
    End Sub
    'стиль текста
    Public Property IDLeaderStyle As ObjectId
        Get
            Return _idStyle
        End Get
        Set(value As ObjectId)
            _idStyle = value
        End Set
    End Property

    'стиль текста
    Public Property TextStyleLeader As String
        Get
            Return _textStyle
        End Get
        Set(value As String)
            _textStyle = value
        End Set
    End Property

    'высота текста
    Public Property HeightTextLeader As Double
        Get
            Return _heightText
        End Get
        Set(value As Double)
            _heightText = value
        End Set
    End Property

    'размер стрелки
    Public Property SizeArrow As Double
        Get
            Return _sizeArraow
        End Get
        Set(value As Double)
            _sizeArraow = value
        End Set
    End Property

    '===========================================================================================================================
    'функция создает все необходимые стили для выносок
    Public Function createStyleLeader(ByVal arrayPropertiesLeader As String(,), Optional ByVal parentLayer As String = "") As Boolean
        createStyleLeader = False
        Dim nameTypeLine As String = "Continuous"
        Dim nameObject As String = ""
        Dim deskObject As String = ""
        Dim arrayMaskLine As String() = Nothing
        If IsArray(arrayPropertiesLeader) = False Then Return False
        Using acTrans As Transaction = DBDrawing.TransactionManager.StartTransaction()
            For i As Integer = 0 To arrayPropertiesLeader.GetUpperBound(1)
                Dim nameField As String = arrayPropertiesLeader(0, i)
                Dim value As String = arrayPropertiesLeader(1, i)
                If IsNothing(nameField) = True Then Continue For
                If nameField.Trim.Length = 0 Then Continue For
                If IsNothing(value) = True Then Continue For
                If value.Trim.Length = 0 Then Continue For
                nameField = nameField.Trim
                value = value.Trim
                Try
                    'создаем слой
                    If nameField Like "Layer" Then
                        Dim acLyrTbl As LayerTable = acTrans.GetObject(DBDrawing.LayerTableId, OpenMode.ForRead)
                        NameLayer = parentLayer & value
                        If acLyrTbl.Has(NameLayer) = False Then
                            If parentLayer.Trim.Length > 0 Then
                                Dim acLyrTblRec As LayerTableRecord = Nothing
                                If acLyrTbl.Has(parentLayer) = True Then
                                    Dim oldId As ObjectId = acLyrTbl.Item(parentLayer)
                                    acLyrTblRec = acTrans.GetObject(oldId, OpenMode.ForRead)
                                Else
                                    Dim oldId As ObjectId = acLyrTbl.Item("0")
                                    acLyrTblRec = acTrans.GetObject(oldId, OpenMode.ForRead)
                                    NameLayer = value
                                End If
                                Dim idLayer As ObjectId = createLayer(NameLayer, acLyrTblRec.Color.ColorIndex, acLyrTblRec.LinetypeObjectId, acLyrTblRec.LineWeight)
                            Else
                                Dim arrayPropertiesLayer As String(,) = {}
                                If IO.File.Exists(TemplateXML) = True Then
                                    Dim boolFindLayerXML As Boolean = CadastralTools.TemplateXML.FuncReadLayerPropertiesInfoToXML(TemplateXML, value, arrayPropertiesLayer)
                                End If
                                Dim indexColor As Integer = 7
                                Dim nameLayerTypeLine As String = "Continuous"
                                Dim strDeskTypeLine As String = "Непрерывный"
                                Dim lineTypeCode As String = ""
                                Dim widthLineLayer As Integer = 20
                                Dim idLineType As ObjectId = Nothing
                                If arrayPropertiesLayer.Length > 0 Then
                                    For k As Integer = 0 To arrayPropertiesLayer.GetUpperBound(1)
                                        Dim nameLayerField As String = arrayPropertiesLayer(0, k)
                                        Dim valueLayer As String = arrayPropertiesLayer(1, k)
                                        If IsNothing(nameLayerField) = True Then Continue For
                                        If nameLayerField.Trim.Length = 0 Then Continue For
                                        If IsNothing(valueLayer) = True Then Continue For
                                        If valueLayer.Trim.Length = 0 Then Continue For
                                        nameLayerField = nameField.Trim
                                        valueLayer = valueLayer.Trim
                                        Select Case nameLayerField
                                            Case "Color"
                                                indexColor = Val(valueLayer)
                                                If indexColor <= 0 Then
                                                    indexColor = 7
                                                ElseIf indexColor >= 255 Then
                                                    indexColor = 7
                                                End If
                                            Case "LineType"
                                                Dim nameTypeLineLayer As String = valueLayer
                                                Dim acLineTypTbl As LinetypeTable = acTrans.GetObject(DBDrawing.LinetypeTableId, OpenMode.ForRead)
                                                If acLineTypTbl.Has(nameTypeLineLayer) = False Then
                                                    Dim userPropertiesTypeLine As String(,) = {}
                                                    Dim boolFindData = CadastralTools.TemplateXML.FuncReadLineTypeInfoToXML(TemplateXML, nameTypeLineLayer, userPropertiesTypeLine)
                                                    If IsArray(userPropertiesTypeLine) = True Then
                                                        Dim deskTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesTypeLine, "DeskLineType", 0, 1)
                                                        Dim patternTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesTypeLine, "CodeLineType", 0, 1)
                                                        If IsNothing(deskTypeLineLayer) = True Then
                                                            deskTypeLineLayer = nameTypeLineLayer
                                                        ElseIf deskTypeLineLayer.Trim.Length = 0 Then
                                                            deskTypeLineLayer = nameTypeLineLayer
                                                        End If
                                                        If IsNothing(patternTypeLineLayer) = True Then
                                                            Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesTypeLine, "DeskLineType", "Continuous", 0, 1)
                                                        ElseIf patternTypeLineLayer.Trim.Length = 0 Then
                                                            Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesTypeLine, "DeskLineType", "Continuous", 0, 1)
                                                        Else
                                                            Dim userLayerLineType As LinetypeTableRecord = createTypeLineByCode(nameTypeLineLayer, deskTypeLineLayer, patternTypeLineLayer)
                                                            If IsNothing(userLayerLineType) = True Then
                                                                Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesTypeLine, "DeskLineType", "Continuous", 0, 1)
                                                            ElseIf userLayerLineType.Id.IsNull = True Then
                                                                Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesTypeLine, "DeskLineType", "Continuous", 0, 1)
                                                            End If
                                                            idLineType = userLayerLineType.Id
                                                        End If
                                                    End If
                                                Else
                                                    idLineType = acLineTypTbl.Item(nameTypeLineLayer)
                                                End If
                                            Case "LineWidth"
                                                widthLineLayer = Val(value)
                                        End Select
                                    Next k
                                End If
                                Dim idLayer As ObjectId = createLayer(NameLayer, indexColor, idLineType, widthLineLayer)
                            End If
                        End If
                    ElseIf nameField Like "Name" Then
                        Dim boolCopyMleader As Boolean = False
                        Dim mlstyles As DBDictionary = acTrans.GetObject(DBDrawing.MLeaderStyleDictionaryId, OpenMode.ForRead)
                        Try
                            Dim idMleaderStyle As ObjectId = mlstyles.GetAt(value)
                            boolCopyMleader = False
                        Catch ex As Teigha.Runtime.Exception
                            boolCopyMleader = copyStyle(TemplateDWG, value, False)
                        End Try
                        If boolCopyMleader = True Then
                            Try
                                IDLeaderStyle = mlstyles.GetAt(value)
                            Catch ex As Teigha.Runtime.Exception
                                IDLeaderStyle = DBDrawing.MLeaderstyle
                            End Try
                        End If
                        'дескриптор объекта
                    ElseIf nameField Like "Description" Then
                        deskObject = value
                        'тип линии
                    ElseIf nameField Like "LineType" Then
                        If value Like "ПоСлою" Then
                            nameTypeLine = "ByLayer"
                        ElseIf nameTypeLine Like "ПоБлоку" Then
                            nameTypeLine = "ByBlock"
                        Else
                            Dim acLineTypTbl As LinetypeTable = acTrans.GetObject(DBDrawing.LinetypeTableId, OpenMode.ForRead)
                            If acLineTypTbl.Has(value) = False Then
                                Dim userPropertiesTypeLine As String(,) = {}
                                Dim boolFindData = CadastralTools.TemplateXML.FuncReadLineTypeInfoToXML(TemplateXML, value, userPropertiesTypeLine)
                                If IsArray(userPropertiesTypeLine) = True Then
                                    Dim deskTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesTypeLine, "DeskLineType", 0, 1)
                                    Dim patternTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesTypeLine, "CodeLineType", 0, 1)
                                    If IsNothing(deskTypeLineLayer) = True Then
                                        deskTypeLineLayer = value
                                    ElseIf deskTypeLineLayer.Trim.Length = 0 Then
                                        deskTypeLineLayer = value
                                    End If
                                    If IsNothing(patternTypeLineLayer) = True Then
                                        Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesTypeLine, "DeskLineType", "Continuous", 0, 1)
                                    ElseIf patternTypeLineLayer.Trim.Length = 0 Then
                                        Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesTypeLine, "DeskLineType", "Continuous", 0, 1)
                                    Else
                                        Dim userLayerLineType As LinetypeTableRecord = createTypeLineByCode(value, deskTypeLineLayer, patternTypeLineLayer)
                                        If IsNothing(userLayerLineType) = True Then
                                            Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesTypeLine, "DeskLineType", "Continuous", 0, 1)
                                        ElseIf userLayerLineType.Id.IsNull = True Then
                                            Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesTypeLine, "DeskLineType", "Continuous", 0, 1)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        'цвет
                    ElseIf nameField Like "Color" Then
                        Dim Color As Integer = Val(value)
                        If Color < 0 Then Color = 0
                        If Color > 255 Then Color = 255
                    ElseIf nameField Like "LineWidth" Then
                        If nameField Like "ПоСлою" Then
                            WeightLine = -1
                        ElseIf nameField Like "ПоБлоку" Then
                            WeightLine = -2
                        ElseIf nameField Like "ПоУмолчанию" Then
                            WeightLine = -3
                        Else
                            WeightLine = Val(value)
                        End If
                    ElseIf nameField Like "LineScale" Then
                        ScaleTypeLine = Val(value)
                        If ScaleTypeLine <= 0 Then
                            ScaleTypeLine = 1
                        End If
                    ElseIf nameField Like "TextStyle" Then
                        Dim acTextStyleTbl As TextStyleTable = acTrans.GetObject(DBDrawing.TextStyleTableId, OpenMode.ForRead)
                        If acTextStyleTbl.Has(value) = False Then
                            Dim arrayPropertiesTextStyle As String(,) = {}
                            Dim boolFindTextStyle As Boolean = CadastralTools.TemplateXML.FuncReadTextStylePropertiesInfoToXML(TemplateXML, value, arrayPropertiesTextStyle)
                            Dim idTextStyle As ObjectId = Nothing
                            If arrayPropertiesTextStyle.Length > 0 Then
                                idTextStyle = createTextStyleByArrayProperties(arrayPropertiesTextStyle)
                            End If

                            If IsNothing(idTextStyle) = True Then
                                idTextStyle = DBDrawing.Textstyle
                            End If
                            If idTextStyle.IsNull = True Then
                                idTextStyle = DBDrawing.Textstyle
                            Else
                                TextStyleLeader = value
                            End If
                        Else
                            TextStyleLeader = value
                        End If
                    ElseIf nameField Like "TextHeight" Then
                        HeightTextLeader = Val(value)
                    ElseIf nameField Like "ArrowSize" Then
                        SizeArrow = value
                    End If
                Catch ex As System.Exception
                End Try
            Next i
            acTrans.Commit()
        End Using
        Return True
    End Function
    Public Function setLeaderStyle(ByRef entityLeader As MLeader) As Boolean
        If IsNothing(DBDrawing) = True Then Return False
        If IsNothing(entityLeader) = False Then
            Using acTrans As Transaction = DBDrawing.TransactionManager.StartTransaction()
                If entityLeader.IsWriteEnabled = True Then
                    Try
                        If Color > -1 And Color < 257 Then
                            entityLeader.ColorIndex = Color
                        Else
                            entityLeader.ColorIndex = 256
                        End If
                        Dim acLyrTbl As LayerTable = acTrans.GetObject(DBDrawing.LayerTableId, OpenMode.ForRead)
                        If acLyrTbl.Has(NameLayer) = True Then
                            entityLeader.Layer = NameLayer
                        Else
                            entityLeader.Layer = "0"
                        End If
                        Dim acLineTypTbl As LinetypeTable = acTrans.GetObject(DBDrawing.LinetypeTableId, OpenMode.ForRead)
                        If acLineTypTbl.Has(TypeLine) = True Then
                            entityLeader.Linetype = TypeLine
                        Else
                            entityLeader.Linetype = "Continuous"
                        End If
                        If ScaleTypeLine > 0 Then
                            entityLeader.LinetypeScale = ScaleTypeLine
                        End If
                        Try
                            If IsNumeric(WeightLine) = True Then
                                Dim weightValue As Integer = Convert.ToInt32(WeightLine)
                                If [Enum].IsDefined(GetType(LineWeight), weightValue) Then
                                    entityLeader.LineWeight = CType(weightValue, LineWeight)
                                Else
                                    entityLeader.LineWeight = CType(weightValue, LineWeight)
                                End If
                            ElseIf WeightLine Like "ПоСлою" Then
                                entityLeader.LineWeight = LineWeight.ByLayer
                            ElseIf WeightLine Like "ПоБлоку" Then
                                entityLeader.LineWeight = LineWeight.ByBlock
                            End If
                        Catch ex As Exception
                            entityLeader.LineWeight = LineWeight.LineWeight020 ' 0.20мм
                        End Try
                        Dim acTextStyleTbl As TextStyleTable = acTrans.GetObject(DBDrawing.TextStyleTableId, OpenMode.ForRead)
                        If acTextStyleTbl.Has(TextStyleLeader) = False Then
                            entityLeader.TextStyleId = acTextStyleTbl.Item(TextStyleLeader)
                        Else
                            entityLeader.TextStyleId = DBDrawing.Textstyle
                        End If
                        If HeightTextLeader > 0 Then
                            entityLeader.TextHeight = HeightTextLeader
                        End If
                        If SizeArrow > 0 Then
                            entityLeader.ArrowSize = SizeArrow
                        End If
                    Catch ex As Teigha.Runtime.Exception
                    End Try
                End If
                acTrans.Commit()
            End Using
        End If
        Return True
    End Function
    '===========================================================================================================
    'функция копирует стиль мультивыносок из другого чертежа
    Public Function copyStyle(ByVal templateDwg As String, ByVal NameMLeaderStyle As String, Optional boolReplase As Boolean = False) As Boolean
        copyStyle = False
        Dim sourceDb As Database = New Database(False, True)
        sourceDb.ReadDwgFile(templateDwg, FileShare.Read, True, String.Empty)
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            Dim acMleaderStyleTbl As DBDictionary = sourceTr.GetObject(sourceDb.MLeaderStyleDictionaryId, OpenMode.ForRead)
            Dim MLeaderStyleCollection As ObjectIdCollection = New ObjectIdCollection
            For Each id As DBDictionaryEntry In acMleaderStyleTbl
                If id.Key Like NameMLeaderStyle Then
                    MLeaderStyleCollection.Add(id.Value)
                    Exit For
                End If
            Next
            If MLeaderStyleCollection.Count > 0 Then
                Dim destDictId1 As ObjectId = DBDrawing.MLeaderStyleDictionaryId
                Dim iMap1 As IdMapping = New IdMapping()
                If boolReplase = True Then
                    DBDrawing.WblockCloneObjects(MLeaderStyleCollection, destDictId1, iMap1, DuplicateRecordCloning.Replace, False)
                Else
                    DBDrawing.WblockCloneObjects(MLeaderStyleCollection, destDictId1, iMap1, DuplicateRecordCloning.Ignore, False)
                End If
            Else
                Return False
            End If
            sourceTr.Commit()
        End Using
        copyStyle = True
    End Function
End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'класс полилиний
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Class CadastralPolylineStyle
    Inherits StyleCadastralObject
    Private _lineGWidth As Double

    Public Sub New()
        MyBase.New()
        _lineGWidth = 0
    End Sub
    Public Sub New(parentPolyline As Teigha.DatabaseServices.Polyline)
        Color = parentPolyline.ColorIndex
        NameLayer = parentPolyline.Layer
        TypeLine = parentPolyline.Linetype
        ScaleTypeLine = parentPolyline.LinetypeScale
        WeightLine = parentPolyline.LineWeight
        LineGWidth = parentPolyline.ConstantWidth
    End Sub
    'Получить или установить глобальную ширину линии</summary>
    Public Property LineGWidth As Double
        Get
            Return _lineGWidth
        End Get
        Set(value As Double)
            _lineGWidth = value
        End Set
    End Property

    '===========================================================================================================================
    'фунция создает все необходимые стили для линейных объектов
    Public Function createLineStyle(ByRef arrayPropertyStyle As String(,)) As Boolean
        If IsNothing(DBDrawing) = True Then Return False
        If IsNothing(arrayPropertyStyle) = True Then Return False
        If arrayPropertyStyle.Length = 0 Then Return False
        If IsNothing(TemplateXML) = True Then Return False
        If IO.File.Exists(TemplateXML) = False Then Return False
        Using acTrans As Transaction = DBDrawing.TransactionManager.StartTransaction()
            Try
                If IsArray(arrayPropertyStyle) = True Then
                    If arrayPropertyStyle.GetUpperBound(1) > 0 Then
                        For i As Integer = 0 To arrayPropertyStyle.GetUpperBound(1)
                            If arrayPropertyStyle(0, i) Like "Layer" Then
                                'создаем новый слой
                                Dim layerObject As String = arrayPropertyStyle(1, i)
                                If IsNothing(layerObject) = True Then Continue For
                                If layerObject.Trim.Length = 0 Then Continue For
                                Dim acLyrTbl As LayerTable = acTrans.GetObject(DBDrawing.LayerTableId, OpenMode.ForRead)
                                Dim acLineTypTbl As LinetypeTable = acTrans.GetObject(DBDrawing.LinetypeTableId, OpenMode.ForRead)
                                'такого слоя нет
                                If acLyrTbl.Has(layerObject) = False Then
                                    'ищем настройки слоя
                                    Dim userPropertiesLayer As String(,) = {}
                                    Dim boolFindData As Boolean = CadastralTools.TemplateXML.FuncReadLayerPropertiesInfoToXML(TemplateXML, layerObject, userPropertiesLayer)
                                    If IsArray(userPropertiesLayer) = True Then
                                        'назначаем тип линии для слоя
                                        Dim nameTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesLayer, "LineType", 0, 1)
                                        If IsNothing(nameTypeLineLayer) = True Then
                                            nameTypeLineLayer = "Continuous"
                                            Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                        ElseIf nameTypeLineLayer.Trim.Length = 0 Then
                                            nameTypeLineLayer = "Continuous"
                                            Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                        Else
                                            'загружаем тип линии
                                            If acLineTypTbl.Has(nameTypeLineLayer) = False Then
                                                Dim userPropertiesTypeLine As String(,) = {}
                                                boolFindData = CadastralTools.TemplateXML.FuncReadLineTypeInfoToXML(TemplateXML, nameTypeLineLayer, userPropertiesTypeLine)
                                                If IsArray(userPropertiesTypeLine) = True Then
                                                    Dim deskTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesLayer, "DeskLineType", 0, 1)
                                                    Dim patternTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesLayer, "CodeLineType", 0, 1)
                                                    If IsNothing(deskTypeLineLayer) = True Then
                                                        deskTypeLineLayer = nameTypeLineLayer
                                                    ElseIf deskTypeLineLayer.Trim.Length = 0 Then
                                                        deskTypeLineLayer = nameTypeLineLayer
                                                    ElseIf IsNothing(patternTypeLineLayer) = True Then
                                                        Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                                    ElseIf patternTypeLineLayer.Trim.Length = 0 Then
                                                        Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                                    Else
                                                        Dim userLayerLineType As LinetypeTableRecord = createTypeLineByCode(nameTypeLineLayer, deskTypeLineLayer, patternTypeLineLayer)
                                                        If IsNothing(userLayerLineType) = True Then
                                                            Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                                        ElseIf userLayerLineType.Id.IsNull = True Then
                                                            Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                    Dim createNewLayer As LayerTableRecord = createLayerByArrayProperties(userPropertiesLayer)
                                    If IsNothing(createNewLayer) = False Then
                                        NameLayer = createNewLayer.Name
                                    End If
                                Else
                                    NameLayer = layerObject
                                End If
                                'создать тип линии
                            ElseIf arrayPropertyStyle(0, i) Like "Name" Then
                                Dim nameTypeLine As String = arrayPropertyStyle(1, i)
                                Dim acLineTypTbl As LinetypeTable = acTrans.GetObject(DBDrawing.LinetypeTableId, OpenMode.ForRead)
                                If acLineTypTbl.Has(nameTypeLine) = False Then
                                    Dim userPropertiesTypeLine As String(,) = {}
                                    Dim boolFindData As Boolean = CadastralTools.TemplateXML.FuncReadLineTypeInfoToXML(TemplateXML, nameTypeLine, userPropertiesTypeLine)
                                    If IsArray(userPropertiesTypeLine) = True Then
                                        Dim deskTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesTypeLine, "DeskLineType", 0, 1)
                                        Dim patternTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesTypeLine, "CodeLineType", 0, 1)
                                        If IsNothing(deskTypeLineLayer) = True Then
                                            deskTypeLineLayer = nameTypeLine
                                        ElseIf deskTypeLineLayer.Trim.Length = 0 Then
                                            deskTypeLineLayer = nameTypeLine
                                        ElseIf IsNothing(patternTypeLineLayer) = True Then
                                            Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesTypeLine, "DeskLineType", "Continuous", 0, 1)
                                        ElseIf patternTypeLineLayer.Trim.Length = 0 Then
                                            Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesTypeLine, "DeskLineType", "Continuous", 0, 1)
                                        Else
                                            Dim userLayerLineType As LinetypeTableRecord = createTypeLineByCode(nameTypeLine, deskTypeLineLayer, patternTypeLineLayer)
                                            If IsNothing(userLayerLineType) = True Then
                                                nameTypeLine = "Continuous"
                                            ElseIf userLayerLineType.Id.IsNull = True Then
                                                nameTypeLine = "Continuous"
                                            End If
                                        End If
                                    Else
                                        nameTypeLine = "Continuous"
                                    End If
                                End If
                                TypeLine = nameTypeLine
                            ElseIf arrayPropertyStyle(0, i) Like "Color" Then
                                Color = Val(arrayPropertyStyle(1, i))
                                If Color < 0 Then Color = 0
                                If Color > 256 Then Color = 256
                            ElseIf arrayPropertyStyle(0, i) Like "LineWidth" Then
                                WeightLine = Val(arrayPropertyStyle(1, i))
                            ElseIf arrayPropertyStyle(0, i) Like "LineScale" Then
                                ScaleTypeLine = Val(arrayPropertyStyle(1, i))
                            ElseIf arrayPropertyStyle(0, i) Like "LineGWidth" Then
                                LineGWidth = Val(arrayPropertyStyle(1, i))
                            End If
                        Next i
                    End If
                End If
                acTrans.Commit()
                Return True
            Catch ex As Teigha.Runtime.Exception
                Return False
            End Try
        End Using
    End Function

    Function setStyle(ByRef lineObject As Teigha.DatabaseServices.Entity) As Boolean
        If IsNothing(DBDrawing) = True Then Return False
        If IsNothing(lineObject) = False Then
            Using acTrans As Transaction = DBDrawing.TransactionManager.StartTransaction()
                If lineObject.IsWriteEnabled = True Then
                    Try
                        If Color > -1 And Color < 257 Then
                            lineObject.ColorIndex = Color
                        Else
                            lineObject.ColorIndex = 256
                        End If
                        Dim acLyrTbl As LayerTable = acTrans.GetObject(DBDrawing.LayerTableId, OpenMode.ForRead)
                        If acLyrTbl.Has(NameLayer) = True Then
                            lineObject.Layer = NameLayer
                        Else
                            lineObject.Layer = "0"
                        End If
                        Dim acLineTypTbl As LinetypeTable = acTrans.GetObject(DBDrawing.LinetypeTableId, OpenMode.ForRead)
                        If acLineTypTbl.Has(TypeLine) = True Then
                            lineObject.Linetype = TypeLine
                        Else
                            lineObject.Linetype = "Continuous"
                        End If
                        If ScaleTypeLine > 0 Then
                            lineObject.LinetypeScale = ScaleTypeLine
                        End If
                        If LineGWidth > 0 Then
                            If TypeOf lineObject Is Teigha.DatabaseServices.Polyline Then
                                Dim pLine As Teigha.DatabaseServices.Polyline = lineObject
                                pLine.ConstantWidth = LineGWidth
                            End If
                        End If
                        Try
                            If IsNumeric(WeightLine) = True Then
                                Dim weightValue As Integer = Convert.ToInt32(WeightLine)
                                If [Enum].IsDefined(GetType(LineWeight), weightValue) Then
                                    lineObject.LineWeight = CType(weightValue, LineWeight)
                                Else
                                    lineObject.LineWeight = CType(weightValue, LineWeight)
                                End If
                            ElseIf WeightLine Like "ПоСлою" Then
                                lineObject.LineWeight = LineWeight.ByLayer
                            ElseIf WeightLine Like "ПоБлоку" Then
                                lineObject.LineWeight = LineWeight.ByBlock
                            End If
                        Catch ex As Exception
                            lineObject.LineWeight = LineWeight.LineWeight020 ' 0.20мм
                        End Try
                    Catch ex As Teigha.Runtime.Exception
                    End Try
                End If
                acTrans.Commit()
            End Using
        End If
        Return True
    End Function


End Class
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'класс блоки
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Class CadastralBlockStyle
    Inherits StyleCadastralObject
    Private _name As String
    Private _scale As Double
    Private _rotation As Double
    Private _annotative As AnnotativeStates
    Private _atribyteProperties As List(Of AttributeCadastralBlock)
    Public Sub New()
        MyBase.New()
        _name = ""
        _scale = 1
        _rotation = 0
        _annotative = False
        _atribyteProperties = New List(Of AttributeCadastralBlock)
    End Sub
    Public Sub New(parentBlock As BlockReference)
        NameBlock = parentBlock.Name
        Color = parentBlock.ColorIndex
        NameLayer = parentBlock.Layer
        TypeLine = parentBlock.Linetype
        ScaleTypeLine = parentBlock.LinetypeScale
        WeightLine = parentBlock.LineWeight
        ScaleBlock = parentBlock.ScaleFactors.X
        RotationBlock = parentBlock.Rotation
        AnnotativeBlock = parentBlock.Annotative
    End Sub
    'Получить или установить имя блока
    Public Property NameBlock As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property
    'Получить или установить масштаб
    Public Property ScaleBlock As Double
        Get
            Return _scale
        End Get
        Set(value As Double)
            _scale = value
        End Set
    End Property
    'Получить или установить поворот
    Public Property RotationBlock As Double
        Get
            Return _rotation
        End Get
        Set(value As Double)
            _rotation = value
        End Set
    End Property
    'Получить или установить аннотативность
    Public Property AnnotativeBlock As AnnotativeStates
        Get
            Return _annotative
        End Get
        Set(value As AnnotativeStates)
            _annotative = value
        End Set
    End Property
    '===========================================================================================================================
    'фунция создает все необходимые стили для блоков
    Public Function createBlockStyle(ByRef arrayPropertyStyle As String(,)) As Boolean
        If IsNothing(DBDrawing) = True Then Return False
        If IsNothing(arrayPropertyStyle) = True Then Return False
        If arrayPropertyStyle.Length = 0 Then Return False
        If IsNothing(TemplateXML) = True Then Return False
        If IO.File.Exists(TemplateXML) = False Then Return False
        Using acTrans As Transaction = DBDrawing.TransactionManager.StartTransaction()
            Try
                If IsArray(arrayPropertyStyle) = True Then
                    If arrayPropertyStyle.GetUpperBound(1) > 0 Then
                        For i As Integer = 0 To arrayPropertyStyle.GetUpperBound(1)
                            Dim nameField As String = arrayPropertyStyle(0, i)
                            If IsNothing(nameField) = True Then Continue For
                            If nameField.Trim.Length = 0 Then Continue For
                            Dim value As String = arrayPropertyStyle(1, i)
                            If IsNothing(value) = True Then Continue For
                            If value.Trim.Length = 0 Then Continue For
                            If nameField Like "Layer" Then
                                'создаем новый слой
                                Dim layerObject As String = value
                                Dim acLyrTbl As LayerTable = acTrans.GetObject(DBDrawing.LayerTableId, OpenMode.ForRead)
                                Dim acLineTypTbl As LinetypeTable = acTrans.GetObject(DBDrawing.LinetypeTableId, OpenMode.ForRead)
                                'такого слоя нет
                                If acLyrTbl.Has(layerObject) = False Then
                                    'ищем настройки слоя
                                    Dim userPropertiesLayer As String(,) = {}
                                    Dim boolFindData As Boolean = CadastralTools.TemplateXML.FuncReadLayerPropertiesInfoToXML(TemplateXML, layerObject, userPropertiesLayer)
                                    If IsArray(userPropertiesLayer) = True Then
                                        'назначаем тип линии для слоя
                                        Dim nameTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesLayer, "LineType", 0, 1)
                                        If IsNothing(nameTypeLineLayer) = True Then
                                            nameTypeLineLayer = "Continuous"
                                            Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                        ElseIf nameTypeLineLayer.Trim.Length = 0 Then
                                            nameTypeLineLayer = "Continuous"
                                            Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                        Else
                                            'загружаем тип линии
                                            If acLineTypTbl.Has(nameTypeLineLayer) = False Then
                                                Dim userPropertiesTypeLine As String(,) = {}
                                                boolFindData = CadastralTools.TemplateXML.FuncReadLineTypeInfoToXML(TemplateXML, nameTypeLineLayer, userPropertiesTypeLine)
                                                If IsArray(userPropertiesTypeLine) = True Then
                                                    Dim deskTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesLayer, "DeskLineType", 0, 1)
                                                    Dim patternTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesLayer, "CodeLineType", 0, 1)
                                                    If IsNothing(deskTypeLineLayer) = True Then
                                                        deskTypeLineLayer = nameTypeLineLayer
                                                    ElseIf deskTypeLineLayer.Trim.Length = 0 Then
                                                        deskTypeLineLayer = nameTypeLineLayer
                                                    ElseIf IsNothing(patternTypeLineLayer) = True Then
                                                        Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                                    ElseIf patternTypeLineLayer.Trim.Length = 0 Then
                                                        Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                                    Else
                                                        Dim userLayerLineType As LinetypeTableRecord = createTypeLineByCode(nameTypeLineLayer, deskTypeLineLayer, patternTypeLineLayer)
                                                        If IsNothing(userLayerLineType) = True Then
                                                            Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                                        ElseIf userLayerLineType.Id.IsNull = True Then
                                                            Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesLayer, "DeskLineType", "Continuous", 0, 1)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                    Dim createNewLayer As LayerTableRecord = createLayerByArrayProperties(userPropertiesLayer)
                                    If IsNothing(createNewLayer) = False Then
                                        NameLayer = createNewLayer.Name
                                    End If
                                Else
                                    NameLayer = layerObject
                                End If
                                'создать блок
                            ElseIf nameField Like "Name" Then
                                Dim nameBlk As String = arrayPropertyStyle(1, i)
                                If nameBlk.Trim.Length > 0 Then
                                    Dim boolCopyBlock As Boolean = CopyBlockReferenceByDestDraw(TemplateDWG, nameBlk)
                                    If boolCopyBlock = True Then
                                        NameBlock = nameBlk
                                    End If
                                End If
                            ElseIf nameField Like "LineType" Then
                                If value Like "ПоСлою" Then
                                    TypeLine = "ByLayer"
                                ElseIf value Like "ПоБлоку" Then
                                    TypeLine = "ByBlock"
                                Else
                                    Dim acLineTypTbl As LinetypeTable = acTrans.GetObject(DBDrawing.LinetypeTableId, OpenMode.ForRead)
                                    If acLineTypTbl.Has(value) = False Then
                                        Dim userPropertiesTypeLine As String(,) = {}
                                        Dim boolFindData = CadastralTools.TemplateXML.FuncReadLineTypeInfoToXML(TemplateXML, value, userPropertiesTypeLine)
                                        Dim boolCreateTypeLine As Boolean = False
                                        If IsArray(userPropertiesTypeLine) = True Then
                                            Dim deskTypeLine As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesTypeLine, "DeskLineType", 0, 1)
                                            Dim patternTypeLineLayer As String = CadastralTools.MathFunction.findValueToArray2d(userPropertiesTypeLine, "CodeLineType", 0, 1)
                                            'описание линии (не обязательно)
                                            If IsNothing(deskTypeLine) = True Then
                                                deskTypeLine = value
                                            ElseIf deskTypeLine.Trim.Length = 0 Then
                                                deskTypeLine = value
                                            End If
                                            'сам шаблон линии
                                            If IsNothing(patternTypeLineLayer) = True Then
                                                Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesTypeLine, "DeskLineType", "Continuous", 0, 1)
                                            ElseIf patternTypeLineLayer.Trim.Length = 0 Then
                                                Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesTypeLine, "DeskLineType", "Continuous", 0, 1)
                                            Else
                                                Dim userLayerLineType As LinetypeTableRecord = createTypeLineByCode(value, deskTypeLine, patternTypeLineLayer)
                                                If IsNothing(userLayerLineType) = True Then
                                                    Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesTypeLine, "DeskLineType", "Continuous", 0, 1)
                                                ElseIf userLayerLineType.Id.IsNull = True Then
                                                    Dim boolWriteData As Boolean = CadastralTools.MathFunction.setValueToArray2d(userPropertiesTypeLine, "DeskLineType", "Continuous", 0, 1)
                                                End If
                                            End If
                                        Else
                                            TypeLine = "Continuous"
                                        End If
                                    Else
                                        TypeLine = value
                                    End If
                                End If
                            ElseIf nameField Like "Color" Then
                                Color = Val(value)
                                If Color < 0 Then Color = 0
                                If Color > 256 Then Color = 256
                            ElseIf nameField Like "LineWidth" Then
                                WeightLine = Val(value)
                            ElseIf nameField Like "LineScale" Then
                                ScaleTypeLine = Val(value)
                            End If
                        Next i
                    End If
                End If
                acTrans.Commit()
                Return True
            Catch ex As Teigha.Runtime.Exception
                Return False
            End Try
        End Using
    End Function
    '===========================================================================================================================
    'применить стиль к блоку
    Function setStyle(ByRef userBlock As Teigha.DatabaseServices.BlockReference) As Boolean
        If IsNothing(DBDrawing) = True Then Return False
        If IsNothing(userBlock) = False Then
            Using acTrans As Transaction = DBDrawing.TransactionManager.StartTransaction()
                If userBlock.IsWriteEnabled = True Then
                    Try
                        'цвет
                        If Color > -1 And Color < 257 Then
                            userBlock.ColorIndex = Color
                        Else
                            userBlock.ColorIndex = 256
                        End If
                        'слой
                        Dim acLyrTbl As LayerTable = acTrans.GetObject(DBDrawing.LayerTableId, OpenMode.ForRead)
                        If acLyrTbl.Has(NameLayer) = True Then
                            userBlock.Layer = NameLayer
                        Else
                            userBlock.Layer = "0"
                        End If
                        'тип линии
                        Dim acLineTypTbl As LinetypeTable = acTrans.GetObject(DBDrawing.LinetypeTableId, OpenMode.ForRead)
                        If acLineTypTbl.Has(TypeLine) = True Then
                            userBlock.Linetype = TypeLine
                        Else
                            userBlock.Linetype = "Continuous"
                        End If
                        'масштаб типа линии
                        If ScaleTypeLine > 0 Then
                            userBlock.LinetypeScale = ScaleTypeLine
                        End If
                        'сасштаб блока
                        If ScaleBlock > 0 Then
                            Dim tempScale As Scale3d = New Scale3d(ScaleBlock, ScaleBlock, ScaleBlock)
                            userBlock.ScaleFactors = tempScale
                        End If
                        userBlock.Rotation = RotationBlock
                    Catch ex As Teigha.Runtime.Exception
                    End Try
                End If
                acTrans.Commit()
            End Using
        End If
        Return True
    End Function
    '=================================================================================================================
    'функция копирует определение блока из другого чертежа
    Public Shared Function CopyBlockReferenceByDestDraw(ByVal templateDWGFile As String, ByVal NameBlock As String, Optional ByVal boolReplace As Boolean = False) As Boolean
        Dim result As Boolean = False
        Dim sourceDb As Database = New Database(False, True)
        Try
            sourceDb.ReadDwgFile(templateDWGFile, FileShare.Read, True, String.Empty)
            Dim acDoc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            If acDoc Is Nothing Then Return False
            Dim destDb As Database = acDoc.Database
            Using sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
                Dim acBlockTbl As BlockTable = sourceTr.GetObject(sourceDb.BlockTableId, OpenMode.ForRead)
                If Not acBlockTbl.Has(NameBlock) Then
                    sourceTr.Abort()
                    Return False
                End If
                Dim idBlk As ObjectId = acBlockTbl.Item(NameBlock)
                Dim objColID As New ObjectIdCollection()
                objColID.Add(idBlk)
                Dim iMap As New IdMapping()
                Dim cloning As DuplicateRecordCloning = If(boolReplace, DuplicateRecordCloning.Replace, DuplicateRecordCloning.Ignore)
                destDb.WblockCloneObjects(objColID, destDb.BlockTableId, iMap, cloning, False)
                sourceTr.Commit()
                result = True
            End Using
        Catch ex As Teigha.Runtime.Exception
            ' Здесь можно записать в лог или вывести сообщение
            Dim ed As Editor = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
            ed.WriteMessage(vbLf & "Ошибка копирования блока: " & ex.Message)
            result = False
        Catch ex As System.Exception
            result = False
        Finally
            ' Всегда закрываем внешнюю базу
            If sourceDb IsNot Nothing Then
                sourceDb.CloseInput(False)
                sourceDb.Dispose()
            End If
        End Try
        Return result
    End Function
End Class
Public Class AttributeCadastralBlock
    Private _nameField As String
    Private _valueField As String
    Private _definition As String
    Private _rotation As Double
    Private _height As Double
    Private _X As Double
    Private _Y As Double
    Private _alignmentH As Integer
    Private _alignmentY As Integer
    Private _visible As Boolean
    Private _textStyle As String
    Private _color As Integer
    Public Sub New()
        _nameField = ""
        _valueField = ""
        _definition = ""
        _rotation = 0
        _height = 2
        _X = 0
        _Y = 0
        _alignmentH = 0
        _alignmentY = 0
        _visible = True
        _textStyle = ""
        _color = 7
    End Sub
    'Получить или установить имя поля
    Public Property TagAttribute As String
        Get
            Return _nameField
        End Get
        Set(value As String)
            _nameField = value
        End Set
    End Property

    'Получить или установить значение поля
    Public Property TextStringAttribute As String
        Get
            Return _valueField
        End Get
        Set(value As String)
            _valueField = value
        End Set
    End Property
    'Получить или установить значение поля
    Public Property DefinitionAttribute As String
        Get
            Return _definition
        End Get
        Set(value As String)
            _definition = value
        End Set
    End Property
    'Получить или установить поворот
    Public Property RotationAttribute As Double
        Get
            Return _rotation
        End Get
        Set(value As Double)
            _rotation = value
        End Set
    End Property

    'Получить или установить высоту
    Public Property HeightAttribute As Double
        Get
            Return _height
        End Get
        Set(value As Double)
            _height = value
        End Set
    End Property

    'Получить или установить смещение по X
    Public Property XAttribute As Double
        Get
            Return _X
        End Get
        Set(value As Double)
            _X = value
        End Set
    End Property

    'Получить или установить смещение по Y
    Public Property YAttribute As Double
        Get
            Return _Y
        End Get
        Set(value As Double)
            _Y = value
        End Set
    End Property

    'Получить или установить смещение по X
    Public Property HorizontalAlignment As Integer
        Get
            Return _alignmentH
        End Get
        Set(value As Integer)
            _alignmentH = value
        End Set
    End Property

    'Получить или установить смещение по Y
    Public Property VerticalAlignment As Integer
        Get
            Return _alignmentY
        End Get
        Set(value As Integer)
            _alignmentY = value
        End Set
    End Property

    'Получить или установить видимость
    Public Property VisibleAttribute As Boolean
        Get
            Return _visible
        End Get
        Set(value As Boolean)
            _visible = value
        End Set
    End Property
    'Получить или установить значение поля
    Public Property TextStyleAttribute As String
        Get
            Return _textStyle
        End Get
        Set(value As String)
            _textStyle = value
        End Set
    End Property
    'Получить или установить смещение по Y
    Public Property ColorAttribute As Integer
        Get
            Return _color
        End Get
        Set(value As Integer)
            _color = value
        End Set
    End Property
    '===================================================================================
    'Функция считывает в массив все атрибуты блока (Массив 0-тег, 1-содержание атрибута, 2-описание атрибута, 3-высота атрибута, 4-угол наклона атрибута, 5-текстовый стиль, 6-цвет)
    Public Function setAttributesStyle(ByRef db As Database, ByVal blkPoint As BlockReference) As Boolean
        setAttributesStyle = False
        If IsNothing(db) = True Then
            Return False
        End If
        If IsNothing(blkPoint) = True Then
            Return False
        End If
        Using acTrans As Transaction = db.TransactionManager.StartTransaction()
            Try
                Dim bt As BlockTable = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim btr As BlockTableRecord = CType(acTrans.GetObject(bt(blkPoint.Name), OpenMode.ForRead), BlockTableRecord)
                Dim attCol As AttributeCollection = blkPoint.AttributeCollection
                For Each attId As ObjectId In attCol
                    Dim attRef As AttributeReference = acTrans.GetObject(attId, OpenMode.ForRead)
                    TagAttribute = attRef.Tag 'имя поля
                    TextStringAttribute = attRef.TextString 'значение
                    HeightAttribute = attRef.Height 'высота атрибута
                    RotationAttribute = attRef.Rotation 'угол атрибута
                    TextStringAttribute = attRef.TextStyleName 'стиль текста
                    ColorAttribute = attRef.ColorIndex 'цвет
                    XAttribute = attRef.Position.X
                    YAttribute = attRef.Position.Y
                    VisibleAttribute = attRef.Visible
                    If btr.HasAttributeDefinitions Then
                        For Each ObjId As ObjectId In btr
                            Dim obj As Teigha.DatabaseServices.Entity = acTrans.GetObject(ObjId, OpenMode.ForRead)
                            If TypeOf obj Is AttributeDefinition Then
                                Dim AttDef As AttributeDefinition = obj
                                Dim TAG_string As String = AttDef.Tag
                                Dim PROMPT_String As String = AttDef.Prompt
                                If TAG_string Like attRef.Tag Then
                                    DefinitionAttribute = PROMPT_String 'описание поля
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                Next
            Catch
                Exit Function
            End Try
            Return True
        End Using
    End Function
    '-===================================================================================
    'Функция считывает определенный атрибут блока и возврящает его значение
    Public Shared Function readTextStringByAttribute(ByRef db As Database, ByVal userBlock As BlockReference, ByVal tagAttribute As String) As String
        readTextStringByAttribute = ""
        If IsNothing(db) = False Then
            If IsNothing(userBlock) = False Then
                If IsNothing(tagAttribute) = False Then
                    Using acTrans As Transaction = db.TransactionManager.StartTransaction()
                        Try
                            Dim attCol As AttributeCollection = userBlock.AttributeCollection
                            If attCol.Count > 0 Then
                                For Each attId As ObjectId In attCol
                                    Dim attRef As AttributeReference = acTrans.GetObject(attId, OpenMode.ForRead)
                                    If attRef.Tag Like tagAttribute Then
                                        readTextStringByAttribute = attRef.TextString
                                        Exit Function
                                    End If
                                Next
                            End If
                        Catch
                            Exit Function
                        End Try
                    End Using
                End If
            End If
        End If
    End Function
    '-===================================================================================
    'Функция записывает атрибуты блока 
    Public Shared Function writeAttributesInBlockReference(ByRef db As Database, ByRef userBlock As BlockReference, ByVal listAttributes As List(Of AttributeCadastralBlock)) As Boolean
        writeAttributesInBlockReference = False
        If IsNothing(db) = True Then
            Return False
        End If
        If IsNothing(userBlock) = True Then
            Return False
        End If
        If IsNothing(listAttributes) = True Then
            Return False
        End If
        If listAttributes.Count = 0 Then
            Return False
        End If
        Using acTrans As Transaction = db.TransactionManager.StartTransaction()
            Try
                Dim acTextStyleTbl As TextStyleTable = acTrans.GetObject(db.TextStyleTableId, OpenMode.ForRead)
                Dim attCol As AttributeCollection = userBlock.AttributeCollection
                For Each attId As ObjectId In attCol
                    Try
                        Dim attRef As AttributeReference = acTrans.GetObject(attId, OpenMode.ForWrite)
                        Dim atributeTag As String = attRef.Tag
                        For i As Integer = 0 To listAttributes.Count - 1
                            Dim userAttribute As AttributeCadastralBlock = listAttributes(i)
                            If userAttribute.TagAttribute Like atributeTag Then
                                Dim valField As String = userAttribute.TagAttribute
                                If valField.Trim.Length > 0 Then
                                    'значение атрибута
                                    attRef.TextString = valField
                                    'видимость атрибута
                                    attRef.Visible = userAttribute.VisibleAttribute
                                    If attRef.Visible = True Then
                                        'высота
                                        Dim heightAttr As Double = userAttribute.HeightAttribute
                                        If heightAttr > 0 Then
                                            attRef.Height = heightAttr
                                        End If
                                        'угол
                                        Dim angle As Double = userAttribute.RotationAttribute
                                        attRef.Oblique = angle
                                        'текстовый стиль
                                        Dim textStyleAttr As String = userAttribute.TextStyleAttribute
                                        If textStyleAttr.Trim.Length > 0 Then
                                            If acTextStyleTbl.Has(textStyleAttr) = True Then
                                                Dim idTextStyle As ObjectId = acTextStyleTbl.Item(textStyleAttr)
                                                attRef.TextStyleId = idTextStyle
                                            End If
                                        End If
                                        'цвет
                                        Dim colorInd As Integer = userAttribute.ColorAttribute
                                        If colorInd > -1 And colorInd < 257 Then
                                            attRef.ColorIndex = colorInd
                                        End If
                                        'положение
                                        Dim posAttribute As Point3d = New Point3d(userAttribute.XAttribute, userAttribute.YAttribute, 0)
                                        attRef.Position = posAttribute
                                    End If
                                End If
                            End If
                        Next i
                        acTrans.Commit()
                        writeAttributesInBlockReference = True
                    Catch ex As System.Exception
                    End Try
                Next
            Catch
                Exit Function
            End Try
        End Using
    End Function
    '-===================================================================================
    'Функция возвращает координаты вставки атрибута и параметры точки выравнивания(массив-X,Y,ГорВыр 1-лево 2-право,ВертВыр 1-низ 2-верх)
    Public Function calculatePositionAttributes(ByVal Vert1 As Point2d, ByVal Vert2 As Point2d, ByVal Vert3 As Point2d, ByVal Dist As Double) As Boolean
        calculatePositionAttributes = False
        Try
            'определяем дирекционный угол на первую линию
            Dim Angle1 As Double = Vert2.GetVectorTo(Vert1).Angle
            'определяем дирекционый угол на вторую линию
            Dim Angle2 As Double = Vert2.GetVectorTo(Vert3).Angle
            'определяем внутренний угол
            Dim Angle As Double = Angle2 - Angle1
            If Angle < 0 Then Angle = Angle + Math.PI * 2 'если угол меньше 0 тогда добавим 360
            Dim Angle2D As Double = Angle / 2 'разделим угол пополам
            Dim Napravlenie As Double = Angle1 + Angle2D + Math.PI  'вычисляем дир.напраление
            If Napravlenie > Math.PI * 2 Then Napravlenie = Napravlenie - Math.PI * 2 'если угол больше 360 градусов
            'вычисляем координаты точки вставки атрибута
            Dim Dx As Double = Dist * Math.Cos(Napravlenie) 'вычисляем приращения координат
            Dim Dy As Double = Dist * Math.Sin(Napravlenie)
            XAttribute = Dx
            YAttribute = Dy
            Return True
        Catch ex As System.Exception
            XAttribute = Dist
            YAttribute = 0
        End Try
    End Function
End Class