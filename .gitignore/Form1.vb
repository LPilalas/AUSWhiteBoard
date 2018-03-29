Public Class Form1_Main
    'Declares global variables
    Public Position As String = ""
    Public Property ImageLocation As String
    Dim Position1 As String = ""
    Dim Position2 As String = ""
    Dim Rover1 As String = ""
    Dim Rover2 As String = ""
    Dim Rover3 As String = ""
    Dim Rover4 As String = ""
   
   Private Sub Form5_load(sender As Object, e As EventArgs) Handles MyBase.Load
     Form5.Show() 'Brings up form5 (form5 still needs to be physicaly designed)
    End sub

    ' Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    '     'Loading Form 1 automatically loads form2 6 times in order to ascertain the 
    '     'officers working on each position.
       
   
   
    '     Form2.Label1.Text = "Who is Positon 1?"
    '     For i As Integer = 0 To 5
    '         Form2.ShowDialog()
    '         If i = 0 Then
    '             Form2.Label1.Text = "Who is Position 2?"
    '             Position1 = Form2.ComboBox1.Text  
    '             btnPos1.Text = Form2.ComboBox1.Text
    '             pbPos1.ImageLocation = "C:\Resources\" + Form2.ComboBox1.Text + ".jpg"
    '         ElseIf i = 1 Then
    '             Form2.Label1.Text = "Who is Rover 1?"
    '             Position2 = Form2.ComboBox1.Text
    '             btnPos2.Text = Form2.ComboBox1.Text
    '             pbPos2.ImageLocation = "C:\Resources\" + Form2.ComboBox1.Text + ".jpg"
    '         ElseIf i = 2 Then
    '             Form2.Label1.Text = "Who is Rover 2?"
    '             Rover1 = Form2.ComboBox1.Text
    '             btnRov1.Text = Form2.ComboBox1.Text
    '             pbRov1.ImageLocation = "C:\Resources\" + Form2.ComboBox1.Text + ".jpg"
    '         ElseIf i = 3 Then
    '             Form2.Label1.Text = "Who is Rover 3?"
    '             Rover2 = Form2.ComboBox1.Text
    '             btnRov2.Text = Form2.ComboBox1.Text
    '             pbRov2.ImageLocation = "C:\Resources\" + Form2.ComboBox1.Text + ".jpg"
    '         ElseIf i = 4 Then
    '             Form2.Label1.Text = "Who is Rover 4?"
    '             Rover3 = Form2.ComboBox1.Text
    '             btnRov3.Text = Form2.ComboBox1.Text
    '             pbRov3.ImageLocation = "C:\Resources\" + Form2.ComboBox1.Text + ".jpg"
    '         ElseIf i = 5 Then
    '             Rover4 = Form2.ComboBox1.Text
    '             btnRov4.Text = Form2.ComboBox1.Text
    '             pbRov4.ImageLocation = "C:\Resources\" + Form2.ComboBox1.Text + ".jpg"
    '         End If
    '     Next
    ' End Sub


    '¯\_(ツ)_/¯
    'This takes the selected officer from Form2 and assigns the correct post. The data held in "Position" will be 
    'delegated to a specific post. That, or it'll turn off the fan in the bathroom. I dunno...wouldn't risk it tho....
    Public Sub btnPos1_Click_1(sender As Object, e As EventArgs) Handles btnPos1.Click
        Position = Position1
    End Sub

    Public Sub btnPos2_Click(sender As Object, e As EventArgs) Handles btnPos2.Click
        Position = Position2
    End Sub

    Public Sub btnRov1_Click(sender As Object, e As EventArgs) Handles btnRov1.Click
        Position = Rover1
    End Sub

    Public Sub btnRov2_Click(sender As Object, e As EventArgs) Handles btnRov2.Click
        Position = Rover2
    End Sub

    Public Sub btnRov3_Click(sender As Object, e As EventArgs) Handles btnRov3.Click
        Position = Rover3
    End Sub

    Public Sub btnRov4_Click(sender As Object, e As EventArgs) Handles btnRov4.Click
        Position = Rover4
    End Sub

    Public Sub btnSetMG_Click(sender As Object, e As EventArgs) Handles btnSetMG.Click
        'Makes the photobox visible and adds the picture of the selected officer to Main Gate
        pbMainGate.Visible = True
        If Position = Position1 Then
            pbMainGate.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            pbMainGate.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            pbMainGate.Image = pbRov1.Image
        ElseIf Position = Rover1 Then
            pbMainGate.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            pbMainGate.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            pbMainGate.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearMG_Click(sender As Object, e As EventArgs) Handles btnClearMG.Click
        'Clears dispatch to this post. 
        pbMainGate.Image = Nothing
        pbMainGate.Visible = False
    End Sub

    Public Sub btnSetV87_Click(sender As Object, e As EventArgs) Handles btnSetV87.Click
        'Makes the photobox visible and adds the picture of the selected officer to V87
        pbV87.Visible = True
        If Position = Position1 Then
            pbV87.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            pbV87.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            pbV87.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            pbV87.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            pbV87.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            pbV87.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearV87_Click(sender As Object, e As EventArgs) Handles btnClearV87.Click
        'Clears dispatch to this post. 
        pbV87.Image = Nothing
        pbV87.Visible = False
    End Sub

    Public Sub btnSetV88_Click(sender As Object, e As EventArgs)
        'Makes the photobox visible and adds the picture of the selected officer to V88
        pbV88.Visible = True
        If Position = Position1 Then
            pbV88.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            pbV88.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            pbV88.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            pbV88.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            pbV88.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            pbV88.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearV88_Click(sender As Object, e As EventArgs)
        'Clears dispatch to this post. 
        pbV88.Image = Nothing
        pbV88.Visible = False
    End Sub

    Public Sub btnSetV03_Click(sender As Object, e As EventArgs) Handles btnSetV03.Click
        'Makes the photobox visible and adds the picture of the selected officer to V03
        pbV03.Visible = True
        If Position = Position1 Then
            pbV03.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            pbV03.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            pbV03.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            pbV03.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            pbV03.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            pbV03.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearV03_Click(sender As Object, e As EventArgs) Handles btnClearV03.Click
        'Clears dispatch to this post. 
        pbV03.Image = Nothing
        pbV03.Visible = False
    End Sub

    Public Sub btnSetDG_Click(sender As Object, e As EventArgs) Handles btnSetDG.Click
        'Makes the photobox visible and adds the picture of the selected officer to Delivery Gate
        pbDG.Visible = True
        If Position = Position1 Then
            pbDG.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            pbDG.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            pbDG.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            pbDG.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            pbDG.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            pbDG.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearDB_Click(sender As Object, e As EventArgs) Handles btnClearDB.Click
        'Clears dispatch to this post. 
        pbDG.Image = Nothing
        pbDG.Visible = False
    End Sub

    Public Sub btnSetV02_Click(sender As Object, e As EventArgs) Handles btnSetV02.Click
        'Makes the photobox visible and adds the picture of the selected officer to V02
        pbV02.Visible = True
        If Position = Position1 Then
            pbV02.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            pbV02.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            pbV02.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            pbV02.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            pbV02.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            pbV02.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearV02_Click(sender As Object, e As EventArgs) Handles btnClearV02.Click
        'Clears dispatch to this post. 
        pbV02.Image = Nothing
        pbV02.Visible = False
    End Sub

    Public Sub btnSetV04_Click(sender As Object, e As EventArgs) Handles btnSetV04.Click
        'Makes the photobox visible and adds the ....look, if you don't get it by now, there's no helping you....
        pbV04.Visible = True
        If Position = Position1 Then
            pbV04.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            pbV04.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            pbV04.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            pbV04.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            pbV04.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            pbV04.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearV04_Click(sender As Object, e As EventArgs) Handles btnClearV04.Click
        'Clears dispatch to this post. 
        pbV04.Image = Nothing
        pbV04.Visible = False
    End Sub

    Public Sub btnSet09_Click(sender As Object, e As EventArgs) Handles btnSet09.Click
        'We apologise for the fault in the subtitles. Those responsible have been sacked.
        pbV09.Visible = True
        If Position = Position1 Then
            pbV09.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            pbV09.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            pbV09.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            pbV09.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            pbV09.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            pbV09.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearV09_Click(sender As Object, e As EventArgs) Handles btnClearV09.Click
        'Clears dispatch to this post. 
        pbV09.Image = Nothing
        pbV09.Visible = False
    End Sub

    Public Sub btnSetV08_Click(sender As Object, e As EventArgs) Handles btnSetV08.Click
        'We apologise again for the fault in the subtitles. Those responsible for 
        'sacking the people hwo have just been sacked have been sacked.
        pbV08.Visible = True
        If Position = Position1 Then
            pbV08.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            pbV08.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            pbV08.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            pbV08.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            pbV08.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            pbV08.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearV08_Click(sender As Object, e As EventArgs) Handles btnClearV08.Click
        'The directors of the firm hired to continue the credits after the other people had been sacked, wish it to be 
        'known that they have just been sacked. The credits have been completed in an entirely different style at 
        'great expense and at the last minute.
        pbV08.Image = Nothing
        pbV08.Visible = False
    End Sub

    Public Sub btnSetV01_Click(sender As Object, e As EventArgs) Handles btnSetV01.Click
        'I don't want to talk to you no more, you empty headed animal food trough wiper. I fart in your general direction.
        'Your mother was a hamster And your father smelt of elderberries!
        pbV01.Visible = True
        If Position = Position1 Then
            pbV01.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            pbV01.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            pbV01.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            pbV01.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            pbV01.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            pbV01.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearV01_Click(sender As Object, e As EventArgs) Handles btnClearV01.Click

        pbV01.Image = Nothing
        pbV01.Visible = False
    End Sub

    Public Sub btnSetV05_Click(sender As Object, e As EventArgs) Handles btnSetV05.Click
        pbV05.Visible = True
        If Position = Position1 Then
            pbV05.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            pbV05.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            pbV05.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            pbV05.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            pbV05.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            pbV05.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearV05_Click(sender As Object, e As EventArgs) Handles btnClearV05.Click
        pbV05.Image = Nothing
        pbV05.Visible = False
    End Sub

    Public Sub btnSetV10_Click(sender As Object, e As EventArgs) Handles btnSetV10.Click
        pbV10.Visible = True
        If Position = Position1 Then
            pbV10.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            pbV10.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            pbV10.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            pbV10.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            pbV10.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            pbV10.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearV10_Click(sender As Object, e As EventArgs) Handles btnClearV10.Click
        pbV10.Image = Nothing
        pbV10.Visible = False
    End Sub

    Private Sub btnSetExt_Click(sender As Object, e As EventArgs) Handles btnSetExt.Click
        pbEXT.Visible = True
        If Position = Position1 Then
            pbEXT.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            pbEXT.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            pbEXT.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            pbEXT.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            pbEXT.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            pbEXT.Image = pbRov4.Image

        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        PictureBox2.Visible = True
        If Position = Position1 Then
            PictureBox2.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            PictureBox2.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            PictureBox2.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            PictureBox2.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            PictureBox2.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            PictureBox2.Image = pbRov4.Image
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        PictureBox3.Visible = True
        If Position = Position1 Then
            PictureBox3.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            PictureBox3.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            PictureBox3.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            PictureBox3.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            PictureBox3.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            PictureBox3.Image = pbRov4.Image
        End If
    End Sub
    Private Sub btnClearExt_Click(sender As Object, e As EventArgs) Handles btnClearExt.Click
        pbEXT.Image = Nothing
        pbEXT.Visible = False
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles btnSetV88.Click
        pbV88.Visible = True
        If Position = Position1 Then
            pbV88.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            pbV88.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            pbV88.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            pbV88.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            pbV88.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            pbV88.Image = pbRov4.Image
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles btnClearV88.Click
        pbV88.Image = Nothing
        pbV88.Visible = False
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        'Loads Form3, which allows you to enter a new shift task to txtTasks
        Form3.Show()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Removes selected task from da box
        txtTasks.Items.Remove(txtTasks.SelectedItem)
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        PictureBox2.Image = Nothing
        PictureBox2.Visible = False
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        PictureBox3.Image = Nothing
        PictureBox3.Visible = False
    End Sub

    Private Sub btnERT_Click(sender As Object, e As EventArgs) Handles btnERT.Click
        'Shows Form4 to enter new ERT members into txtERT, fo sho
        frmERT.Show()
    End Sub

    Private Sub btnRMERT_Click(sender As Object, e As EventArgs) Handles btnRMERT.Click
        'Removes ERT Peeps from list
        txtERT.Items.Remove(txtERT.SelectedItem)
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        MessageBox.Show("Hands off me pixels!")
    End Sub
End Class
