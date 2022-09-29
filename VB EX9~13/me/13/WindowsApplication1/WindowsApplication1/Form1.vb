Public Class Form1
    Dim dis As Integer = 30
    Dim A As Integer
    Dim min, sec As Integer


    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        sec += 1
        If sec = 60 Then
            min += 1
            sec = 0
        End If

        Label4.Text = min & "分" & sec & "秒"
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label4.Text = min & "分" & sec & "秒"

    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton1.Checked = True Then
            PictureBox1.Left -= dis
            If PictureBox1.Left = -PictureBox1.Width Then
                PictureBox1.Left = Me.Width + PictureBox1.Width
            End If
        End If
        

    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Timer2.Interval = HScrollBar1.Value
        Dim P1, P2, P As Object
        P = PictureBox2.Image
        P1 = PictureBox2.Image
        P2 = PictureBox3.Image

        If RadioButton4.Checked = True And RadioButton1.Checked = True Then
            PictureBox1.Left -= dis
            If PictureBox1.Left <= -PictureBox1.Width Then
                PictureBox1.Left = Me.Width + PictureBox1.Width

            End If
            A += 1
            If A Mod 2 = 0 Then
                PictureBox1.Image = PictureBox3.Image
            Else
                PictureBox1.Image = PictureBox2.Image
            End If
        End If

        If RadioButton3.Checked = True And RadioButton1.Checked = True Then
            PictureBox1.Left += dis
            If PictureBox1.Left >= Me.Width + PictureBox1.Width Then
                PictureBox1.Left = -PictureBox1.Width
            End If
            A += 1
            If A Mod 2 = 0 Then
                PictureBox1.Image = PictureBox3.Image
            Else
                PictureBox1.Image = PictureBox2.Image
            End If
        End If

        If RadioButton6.Checked = True And RadioButton1.Checked = True Then
            PictureBox1.Top -= dis
            If PictureBox1.Top <= -PictureBox1.Height Then
                PictureBox1.Top = Me.Height + PictureBox1.Height

            End If
            A += 1
            If A Mod 2 = 0 Then
                PictureBox1.Image = PictureBox3.Image
            Else
                PictureBox1.Image = PictureBox2.Image
            End If
        End If

        If RadioButton5.Checked = True And RadioButton1.Checked = True Then
            PictureBox1.Top += dis
            If PictureBox1.Top >= Me.Height + PictureBox1.Height Then
                PictureBox1.Top = -PictureBox1.Height

            End If
            A += 1
            If A Mod 2 = 0 Then
                PictureBox1.Image = PictureBox3.Image
            Else
                PictureBox1.Image = PictureBox2.Image
            End If
        End If

        








    End Sub

    Private Sub HScrollBar1_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles HScrollBar1.Scroll
        Timer2.Interval = HScrollBar1.Value
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        Timer1.Enabled = False
        Timer2.Enabled = False
        Timer3.Enabled = False
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        Timer1.Enabled = True
        Timer2.Enabled = True
        Timer3.Enabled = True
    End Sub


End Class
