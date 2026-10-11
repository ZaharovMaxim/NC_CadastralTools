Imports HostMgd.ApplicationServices
Imports Teigha.DatabaseServices

Public Class FuncPropertiesDrawing
    '-=====================================================================================
    '-Функция возвращает текущий масштаб аннотаций
    Public Shared Function FuncReadCurrentAnnotativeScaleDWG() As Double
        Dim doc As Document = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim anno As AnnotationScale = db.Cannoscale
        Return (1 / anno.Scale) * 1000
    End Function
End Class
