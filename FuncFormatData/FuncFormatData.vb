Imports System.Globalization

Public Class FuncFormatData
    '==================================================================================================================================
    'функция форматирует длины линий под определенный формат
    Public Shared Function FuncFormatFloatNumber(ByVal coordinates As Double, Optional ByVal RoundZn As Integer = 2) As String
        FuncFormatFloatNumber = coordinates
        Try
            If RoundZn = 0 Then
                FuncFormatFloatNumber = String.Format(CultureInfo.InvariantCulture, "{0:0}", Math.Round(coordinates, RoundZn))
            ElseIf RoundZn = 1 Then
                FuncFormatFloatNumber = String.Format(CultureInfo.InvariantCulture, "{0:0.0}", Math.Round(coordinates, RoundZn))
            ElseIf RoundZn = 2 Then
                FuncFormatFloatNumber = String.Format(CultureInfo.InvariantCulture, "{0:0.00}", Math.Round(coordinates, RoundZn))
            ElseIf RoundZn = 3 Then
                FuncFormatFloatNumber = String.Format(CultureInfo.InvariantCulture, "{0:0.000}", Math.Round(coordinates, RoundZn))
            ElseIf RoundZn = 4 Then
                FuncFormatFloatNumber = String.Format(CultureInfo.InvariantCulture, "{0:0.0000}", Math.Round(coordinates, RoundZn))
            Else
                FuncFormatFloatNumber = Math.Round(coordinates, RoundZn)
            End If
        Catch
            Exit Function
        End Try
    End Function

    Public Shared Function SafeConvertToDouble(ByVal value As String) As Double?
        If String.IsNullOrWhiteSpace(value) Then
            Return Nothing
        End If

        ' Проверяем, содержит ли строка дефис
        If value.Contains("-") Then
            ' Логируем проблему
            Console.WriteLine($"Обнаружен дефис в строке: {value}")

            ' Возвращаем Nothing или значение по умолчанию
            Return Nothing
        End If

        Dim result As Double
        If Double.TryParse(value, Globalization.NumberStyles.Any,
                           Globalization.CultureInfo.InvariantCulture, result) Then
            Return result
        Else
            Return Nothing
        End If
    End Function

End Class
