Imports HostMgd.ApplicationServices
Imports HostMgd.EditorInput
Imports Teigha.DatabaseServices
Imports Teigha.Geometry
Imports Teigha.Runtime

Public Class test
    <CommandMethod("TestCreateXData")>
    Public Sub TestCreateXData()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        '===============================================================================================================
        Dim pPtOpts As PromptEntityOptions = New PromptEntityOptions("Выберите примитив: ")
        Dim pPtRes As PromptEntityResult = doc.Editor.GetEntity(pPtOpts)
        If pPtRes.Status = PromptStatus.Cancel Then Exit Sub
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim ent As Entity = tr.GetObject(pPtRes.ObjectId, OpenMode.ForWrite)
            Dim nameDictionary As String = "PROJECT_PARCEL"
            Dim arrayData As String(,) = {}
            ReDim arrayData(1, 4)
            arrayData(0, 0) = "Reg_numb_border"
            arrayData(1, 0) = "53:11"

            arrayData(0, 1) = "Type_boundary"
            arrayData(1, 1) = "Зона"

            arrayData(0, 2) = "DateRegistration"
            arrayData(1, 2) = "01.01.2025"

            arrayData(0, 3) = "LOCALID"
            arrayData(1, 3) = "-"

            arrayData(0, 4) = "GlobalID"
            arrayData(1, 4) = "-"
            Dim dict As Dictionary(Of Integer, String(,)) = New Dictionary(Of Integer, String(,))
            dict.Add(1, arrayData)
            'Dim boolRec As Boolean = FuncXRecords.FuncWriteXData(ent, "PROJECT_ZONA", arrayData)
            tr.Commit()
        End Using
    End Sub
    <CommandMethod("TestReadXData")>
    Public Sub TestReadXData()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor
        '===============================================================================================================
        Dim pPtOpts As PromptEntityOptions = New PromptEntityOptions("Выберите примитив: ")
        Dim pPtRes As PromptEntityResult = doc.Editor.GetEntity(pPtOpts)
        If pPtRes.Status = PromptStatus.Cancel Then Exit Sub
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim ent As Entity = tr.GetObject(pPtRes.ObjectId, OpenMode.ForWrite)
            Dim arrayData As String(,) = {}
            Dim nameDictionary As String = "PROJECT_PARCEL"
            'Dim boolReadData As Boolean = FuncXRecords.FuncReadXData(ent, nameDictionary, arrayData)
            tr.Commit()
        End Using
    End Sub
End Class
