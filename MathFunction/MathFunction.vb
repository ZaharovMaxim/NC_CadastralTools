Imports HostMgd.ApplicationServices
Imports HostMgd.EditorInput
Imports HostMgd.Windows
Imports Teigha.DatabaseServices
Imports Teigha.Geometry
Imports Teigha.Runtime


Public Class MathFunction
    '===============================================================================================
    'функция вычисляет точку пересечения двух отрезков
    Public Shared Function FuncIntersectPolyline(ByVal Pt1 As Point2d, ByVal Pt2 As Point2d, ByVal Pt3 As Point2d, ByVal pt4 As Point2d, ByRef InPoint As Point2d) As Boolean
        FuncIntersectPolyline = False
        Dim x1 As Double = Pt1.X
        Dim y1 As Double = Pt1.Y

        Dim x2 As Double = Pt2.X
        Dim y2 As Double = Pt2.Y

        Dim x3 As Double = Pt3.X
        Dim y3 As Double = Pt3.Y

        Dim x4 As Double = pt4.X
        Dim y4 As Double = pt4.Y
        Try
            Dim d As Double = (Pt1.X - Pt2.X) * (pt4.Y - Pt3.Y) - (Pt1.Y - Pt2.Y) * (pt4.X - Pt3.X)
            Dim da As Double = (Pt1.X - Pt3.X) * (pt4.Y - Pt3.Y) - (Pt1.Y - Pt3.Y) * (pt4.X - Pt3.X)
            Dim db As Double = (Pt1.X - Pt2.X) * (Pt1.Y - Pt3.Y) - (Pt1.Y - Pt2.Y) * (Pt1.X - Pt3.X)

            Dim ta As Double = da / d
            Dim tb As Double = db / d

            If (ta >= 0 And ta <= 1 And tb >= 0 And tb <= 1) Then
                Dim dx As Double = Pt1.X + ta * (Pt2.X - Pt1.X)
                Dim dy As Double = Pt1.Y + ta * (Pt2.Y - Pt1.Y)
                InPoint = New Point2d(dx, dy)
                Return True
            End If
        Catch ex As Exception
        End Try
    End Function
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    'точка пересечения бесконечных отрезков
    Public Shared Function FuncIntersectPolyline2(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal x4 As Double, ByVal y4 As Double, ByRef InPoint As Point2d) As Boolean
        Dim n As Double
        If (y2 - y1 <> 0) Then
            Dim q As Double = (x2 - x1) / (y1 - y2)
            Dim sn As Double = (x3 - x4) + (y3 - y4) * q
            If (sn = 0) Then Return False ' c(x) + c(y)*q
            Dim fn As Double = (x3 - x1) + (y3 - y1) * q '   // b(x) + b(y)*q
            n = fn / sn
        Else
            If ((y3 - y4)) = 0 Then Return False '; }  // b(y)
            n = (y3 - y1) / (y3 - y4) ';   // c(y)/b(y)
        End If
        Dim x As Double = x3 + (x4 - x3) * n ';  // x3 + (-b(x))*n
        Dim y As Double = y3 + (y4 - y3) * n ';  // y3 +(-b(y))*n
        InPoint = New Point2d(x, y)
        Return True
    End Function

End Class

