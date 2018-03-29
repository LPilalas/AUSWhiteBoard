Public Class Form3
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtTask.Focus()
    End Sub

    Private Sub btnFrmAdd_Click(sender As Object, e As EventArgs) Handles btnFrmAdd.Click
        'this is the form that will add tasks to the "Shift Tasks" section of the form. 
        Dim Task As String = txtTask.Text
        Form1.txtTasks.Items.Add(Task)
        txtTask.Text = ""
        txtTask.Focus()

    End Sub

    Private Sub btnFrmCan_Click(sender As Object, e As EventArgs) Handles btnFrmCan.Click
        Me.Close()
    End Sub


End Class