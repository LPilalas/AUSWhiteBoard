private Class Form5_Main 
 'Declares global variables
    Public Position As String = ""
    Public Property ImageLocation As String
    Dim Position1 As String = ""
    Dim Position2 As String = ""
    Dim Rover1 As String = ""
    Dim Rover2 As String = ""
    Dim Rover3 As String = ""
    Dim Rover4 As String = ""
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object

 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 'To export to excel sheet
 'VVVVVVVVVVVVVVVVVVVVVVVV
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 ' Start a new workbook in Excel
 '  Set oExcel = CreateObject("Excel.Application")
 '  Set oBook = oExcel.Workbooks.Add
 ' Add data to cells of the first worksheet in the new workbook
 '  Set oSheet = oBook.Worksheets(1)
 '  oSheet.Range("A1").Value = "Last Name"
 '  oSheet.Range("B1").Value = "First Name"
 '  oSheet.Range("A1:B1").Font.Bold = True
 '  oSheet.Range("A2").Value = "Doe"
 '  oSheet.Range("B2").Value = "John"
 ' Save the Workbook and Quit Excel
 '  oBook.SaveAs "C:\Book1.xls"
 '  oExcel.Quit
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
 'Export to excel sheet
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

 Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Loading Form 5 automatically loads form2 6 times in order to ascertain the 
        'officers working on each position.
        Form2.Label1.Text = "Who is Positon 1?"
        For i As Integer = 0 To 5
            Form2.ShowDialog()
            If i = 0 Then
                Form2.Label1.Text = "Who is Position 2?"
                Position1 = Form2.ComboBox1.Text  
                btnPos1.Text = Form2.ComboBox1.Text
                Form1.pbPos1.ImageLocation = "C:\Resources\" + Form2.ComboBox1.Text + ".jpg"
                pbPos1.ImageLocation = "C:\Resources\" + Form2.ComboBox1.Text + ".jpg"
            ElseIf i = 1 Then
                Form2.Label1.Text = "Who is Rover 1?"
                Position2 = Form2.ComboBox1.Text
                btnPos2.Text = Form2.ComboBox1.Text
                Form1.pbPos2.ImageLocation = "C:\Resources\" + Form2.ComboBox1.Text + ".jpg"
                pbPos2.ImageLocation = "C:\Resources\" + Form2.ComboBox1.Text + ".jpg"
            ElseIf i = 2 Then
                Form2.Label1.Text = "Who is Rover 2?"
                Rover1 = Form2.ComboBox1.Text
                btnRov1.Text = Form2.ComboBox1.Text
                Form1.pbRov1.ImageLocation = "C:\Resources\" + Form2.ComboBox1.Text + ".jpg"
                pbRov1.ImageLocation = "C:\Resources\" + Form2.ComboBox1.Text + ".jpg"
            ElseIf i = 3 Then
                Form2.Label1.Text = "Who is Rover 3?"
                Rover2 = Form2.ComboBox1.Text
                btnRov2.Text = Form2.ComboBox1.Text
                Form1.pbRov2.ImageLocation = "C:\Resources\" + Form2.ComboBox1.Text + ".jpg"
                pbRov2.ImageLocation = "C:\Resources\" + Form2.ComboBox1.Text + ".jpg"
            ElseIf i = 4 Then
                Form2.Label1.Text = "Who is Rover 4?"
                Rover3 = Form2.ComboBox1.Text
                btnRov3.Text = Form2.ComboBox1.Text
                Form1.pbRov3.ImageLocation = "C:\Resources\" + Form2.ComboBox1.Text + ".jpg"
                pbRov3.ImageLocation = "C:\Resources\" + Form2.ComboBox1.Text + ".jpg"
            ElseIf i = 5 Then
                Rover4 = Form2.ComboBox1.Text
                btnRov4.Text = Form2.ComboBox1.Text
                Form1.pbRov4.ImageLocation = "C:\Resources\" + Form2.ComboBox1.Text + ".jpg"
                pbRov4.ImageLocation = "C:\Resources\" + Form2.ComboBox1.Text + ".jpg"
            End If
        Next
    End Sub

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

 Public Sub btnClearMG_Click(sender As Object, e As EventArgs) Handles btnClearMG.Click
        'Clears dispatch to this post. 
        Form1.pbMainGate.Image = Nothing
        Form1.pbMainGate.Visible = False
    End Sub

    Public Sub btnSetV87_Click(sender As Object, e As EventArgs) Handles btnSetV87.Click
        'Makes the photobox visible and adds the picture of the selected officer to V87
        pbV87.Visible = True
        If Position = Position1 Then
            Form1.pbV87.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            Form1.pbV87.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            Form1.pbV87.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            Form1.pbV87.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            Form1.pbV87.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            Form1.pbV87.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearV87_Click(sender As Object, e As EventArgs) Handles btnClearV87.Click
        'Clears dispatch to this post. 
        Form1.pbV87.Image = Nothing
        Form1.pbV87.Visible = False
    End Sub

    Public Sub btnSetV88_Click(sender As Object, e As EventArgs)
        'Makes the photobox visible and adds the picture of the selected officer to V88
        pbV88.Visible = True
        If Position = Position1 Then
            Form1.pbV88.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            Form1.pbV88.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            Form1.pbV88.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            Form1.pbV88.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            Form1.pbV88.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            Form1.pbV88.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearV88_Click(sender As Object, e As EventArgs)
        'Clears dispatch to this post. 
        Form1.pbV88.Image = Nothing
        Form1.pbV88.Visible = False
    End Sub

    Public Sub btnSetV03_Click(sender As Object, e As EventArgs) Handles btnSetV03.Click
        'Makes the photobox visible and adds the picture of the selected officer to V03
        Form1.pbV03.Visible = True
        If Position = Position1 Then
            Form1.pbV03.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            Form1.pbV03.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            Form1.pbV03.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            Form1.pbV03.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            Form1.pbV03.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            Form1.pbV03.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearV03_Click(sender As Object, e As EventArgs) Handles btnClearV03.Click
        'Clears dispatch to this post. 
        Form1.pbV03.Image = Nothing
        Form1.pbV03.Visible = False
    End Sub

    Public Sub btnSetDG_Click(sender As Object, e As EventArgs) Handles btnSetDG.Click
        'Makes the photobox visible and adds the picture of the selected officer to Delivery Gate
        Form1.pbDG.Visible = True
        If Position = Position1 Then
            Form1.pbDG.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            Form1.pbDG.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            Form1.pbDG.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            Form1.pbDG.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            Form1.pbDG.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            Form1.pbDG.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearDB_Click(sender As Object, e As EventArgs) Handles btnClearDB.Click
        'Clears dispatch to this post. 
        Form1.pbDG.Image = Nothing
        Form1.pbDG.Visible = False
    End Sub

    Public Sub btnSetV02_Click(sender As Object, e As EventArgs) Handles btnSetV02.Click
        'Makes the photobox visible and adds the picture of the selected officer to V02
        Form1.pbV02.Visible = True
        If Position = Position1 Then
            Form1.pbV02.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            Form1.pbV02.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            Form1.pbV02.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            Form1.pbV02.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            Form1.pbV02.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            Form1.pbV02.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearV02_Click(sender As Object, e As EventArgs) Handles btnClearV02.Click
        'Clears dispatch to this post. 
        Form1.pbV02.Image = Nothing
        Form1.pbV02.Visible = False
    End Sub

    Public Sub btnSetV04_Click(sender As Object, e As EventArgs) Handles btnSetV04.Click
        'Makes the photobox visible and adds the ....look, if you don't get it by now, there's no helping you....
        Form1.pbV04.Visible = True
        If Position = Position1 Then
            Form1.pbV04.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            Form1.pbV04.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            Form1.pbV04.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            Form1.pbV04.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            Form1.pbV04.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            Form1.pbV04.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearV04_Click(sender As Object, e As EventArgs) Handles btnClearV04.Click
        'Clears dispatch to this post. 
        Form1.pbV04.Image = Nothing
        Form1.pbV04.Visible = False
    End Sub

    Public Sub btnSet09_Click(sender As Object, e As EventArgs) Handles btnSet09.Click
        'We apologise for the fault in the subtitles. Those responsible have been sacked.
        Form1.pbV09.Visible = True
        If Position = Position1 Then
            Form1.pbV09.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            Form1.pbV09.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            Form1.pbV09.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            Form1.pbV09.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            Form1.pbV09.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            Form1.pbV09.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearV09_Click(sender As Object, e As EventArgs) Handles btnClearV09.Click
        'Clears dispatch to this post. 
        Form1.pbV09.Image = Nothing
        Form1.pbV09.Visible = False
    End Sub

    Public Sub btnSetV08_Click(sender As Object, e As EventArgs) Handles btnSetV08.Click
        'We apologise again for the fault in the subtitles. Those responsible for 
        'sacking the people hwo have just been sacked have been sacked.
        Form1.pbV08.Visible = True
        If Position = Position1 Then
            Form1.pbV08.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            Form1.pbV08.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            Form1.pbV08.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            Form1.pbV08.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            Form1.pbV08.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            Form1.pbV08.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearV08_Click(sender As Object, e As EventArgs) Handles btnClearV08.Click
        'The directors of the firm hired to continue the credits after the other people had been sacked, wish it to be 
        'known that they have just been sacked. The credits have been completed in an entirely different style at 
        'great expense and at the last minute.
        Form1.pbV08.Image = Nothing
        Form1.pbV08.Visible = False
    End Sub

    Public Sub btnSetV01_Click(sender As Object, e As EventArgs) Handles btnSetV01.Click
        'I don't want to talk to you no more, you empty headed animal food trough wiper. I fart in your general direction.
        'Your mother was a hamster And your father smelt of elderberries!
        Form1.pbV01.Visible = True
        If Position = Position1 Then
            Form1.pbV01.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            Form1.pbV01.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            Form1.pbV01.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            Form1.pbV01.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            Form1.pbV01.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            Form1.pbV01.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearV01_Click(sender As Object, e As EventArgs) Handles btnClearV01.Click

        Form1.pbV01.Image = Nothing
        Form1.pbV01.Visible = False
    End Sub

    Public Sub btnSetV05_Click(sender As Object, e As EventArgs) Handles btnSetV05.Click
        Form1.pbV05.Visible = True
        If Position = Position1 Then
            Form1.pbV05.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            Form1.pbV05.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            Form1.pbV05.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            Form1.pbV05.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            Form1.pbV05.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            Form1.pbV05.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearV05_Click(sender As Object, e As EventArgs) Handles btnClearV05.Click
        Form1.pbV05.Image = Nothing
        Form1.pbV05.Visible = False
    End Sub

    Public Sub btnSetV10_Click(sender As Object, e As EventArgs) Handles btnSetV10.Click
        Form1.pbV10.Visible = True
        If Position = Position1 Then
           Form1.pbV10.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            Form1.pbV10.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            Form1.pbV10.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            Form1.pbV10.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            Form1.pbV10.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            Form1.pbV10.Image = pbRov4.Image

        End If
    End Sub

    Public Sub btnClearV10_Click(sender As Object, e As EventArgs) Handles btnClearV10.Click
        Form1.pbV10.Image = Nothing
        Form1.pbV10.Visible = False
    End Sub

    Private Sub btnSetExt_Click(sender As Object, e As EventArgs) Handles btnSetExt.Click
        Form1.pbEXT.Visible = True
        If Position = Position1 Then
            Form1.pbEXT.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            Form1.pbEXT.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            Form1.pbEXT.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            Form1.pbEXT.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            Form1.pbEXT.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            Form1.pbEXT.Image = pbRov4.Image

        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Form1.PictureBox2.Visible = True
        If Position = Position1 Then
            Form1.PictureBox2.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            Form1.PictureBox2.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            Form1.PictureBox2.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            Form1.PictureBox2.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            Form1.PictureBox2.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            Form1.PictureBox2.Image = pbRov4.Image
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Form1.PictureBox3.Visible = True
        If Position = Position1 Then
            Form1.PictureBox3.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            Form1.PictureBox3.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            Form1.PictureBox3.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            Form1.PictureBox3.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            Form1.PictureBox3.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            Form1.PictureBox3.Image = pbRov4.Image
        End If
    End Sub
    
    Private Sub btnClearExt_Click(sender As Object, e As EventArgs) Handles btnClearExt.Click
        pbEXT.Image = Nothing
        pbEXT.Visible = False
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles btnSetV88.Click
        Form1.pbV88.Visible = True
        If Position = Position1 Then
            Form1.pbV88.Image = pbPos1.Image
        ElseIf Position = Position2 Then
            Form1.pbV88.Image = pbPos2.Image
        ElseIf Position = Rover1 Then
            Form1.pbV88.Image = pbRov1.Image
        ElseIf Position = Rover2 Then
            Form1.pbV88.Image = pbRov2.Image
        ElseIf Position = Rover3 Then
            Form1.pbV88.Image = pbRov3.Image
        ElseIf Position = Rover4 Then
            Form1.pbV88.Image = pbRov4.Image
        End If
    End Sub


   Private Sub btnFrmAdd_Click(sender As Object, e As EventArgs) Handles btnFrmAdd.Click
        'this is the function that will add tasks to the "Shift Tasks" section of the form. 
        Dim Task As String = txtTask.Text
        Form1.txtTasks.Items.Add(Task)
        txtTask.Text = ""
        txtTask.Focus()

    End Sub

 End Sub

 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

 Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        'Some wonderful magic allowing text to go from this form to txtERT.
        Dim Task As String = txtERT2.Text
        Form1.txtERT.Items.Add(Task)
        txtERT2.Text = ""
        txtERT2.Focus()
    End Sub


End Class