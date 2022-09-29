Public Class Form1

    Sub ansall()
        Label5.Text = ""
        Dim take As Integer
        Randomize()
        ans = 0
        For i = 0 To 9
            A(i) = 0
        Next


        For i = 0 To 3
            take = Int(Rnd() * 10)
            While A(take) = 1
                take = Int(Rnd() * 10)
            End While
            A(take) = 1
            B(i) = take
        Next

        For i = 0 To 3
            ans = ans & B(i)
        Next


    End Sub

    Dim ans, qwe As Integer
    Dim time As Integer
    Dim A(9), B(3) As Integer

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim guess(3), think As Integer
        Dim numA, numB As Integer
        If (qwe < 1) Then
            Label3.Text &= "     Guess     Result" & vbCrLf
            qwe += 1
        End If
        For i = 0 To 3
            guess(i) = Mid(TextBox1.Text, i + 1, 1)

            think &= guess(i)

        Next

        For i = 0 To 3
            If guess(i) = B(i) Then
                numA += 1
            End If
        Next
        For q = 0 To 3
            For J = 0 To 3
                If guess(q) = B(J) And q <> J Then
                    numB += 1
                End If
            Next
        Next
        time += 1

        If think = ans Then
            Label3.Text &= time & ".  " & think & "      " & numA & "A" & numB & "B" & "，恭喜你答對了!"
            Label5.Text = ans
        Else
            Label3.Text &= time & ".  " & think & "      " & numA & "A" & numB & "B" & vbCrLf
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Label5.Text = ans
        Label5.Visible = True
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call ansall()
        Label3.Text = ""
        TextBox1.Text = ""
        Button1.Enabled = False
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call ansall()
    End Sub
End Class
