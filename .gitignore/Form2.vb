Public Class Form2
    Dim Officer As String = ""

    Public Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub
    Public Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        'Makes the selected officer special. This works with Form1 in assigning the appropriate officer
        'to the correct post. 
        Officer = ComboBox1.Text
        Me.Close()
    End Sub


End Class