Public Class frmERT
    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        'Some wonderful magic allowing text to go from this form to txtERT.
        Dim Task As String = txtERT2.Text
        Form1.txtERT.Items.Add(Task)
        txtERT2.Text = ""
        txtERT2.Focus()
    End Sub

    Private Sub btnERTCan_Click(sender As Object, e As EventArgs) Handles btnERTCan.Click
        'Click here to see the top 10 ways this form is cursed. Number 7 will surprise you!
        Me.Close()
    End Sub
End Class