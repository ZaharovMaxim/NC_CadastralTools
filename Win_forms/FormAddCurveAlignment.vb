Public Class FormAddCurveAlignment
    Public boolShow As Boolean = False
    Public Sub New()

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()
        Dim arrayIzmR As String() = {"0", "10", "50", "100", "500", "1000"}
        ' Добавить код инициализации после вызова InitializeComponent().
        ComboBox1.DataSource = arrayIzmR
    End Sub
    Private Sub FormAddCurveAlignment_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        boolShow = False
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        boolShow = True
        Me.Hide()
    End Sub
End Class