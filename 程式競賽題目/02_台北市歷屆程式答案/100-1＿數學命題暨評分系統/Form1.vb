Public Class Form1
    Dim user$, no%, score!, num1%, num2%, right1%, right2%, op%, ans&
    Sub test()
        Dim x%, y%
        no += 1
        Label1.Text = no & "."
        Do
            x = Int(Rnd() * 10000)
            y = Int(Rnd() * 999) + 1
        Loop Until x >= y
        Label2.Text = x : Label4.Text = y

        '+,-:20%,*,/:30% --> +:1,2,-:3,4;*:5~7;/:8~10
        op = Int(Rnd() * 10) + 1
        Select Case op
            Case 1, 2
                Label3.Text = "+"
                num1 += 1
                ans = x + y
            Case 3, 4
                Label3.Text = "-"
                num1 += 1
                ans = x - y
            Case 5 To 7
                Label3.Text = "*"
                num2 += 1
                ans = x * y
            Case 8 To 10
                Label3.Text = "/"
                num2 += 1
                ans = Int(x / y + 0.5)
        End Select
    End Sub
    Sub cal()
        If Val(TextBox1.Text) = ans Then
            Select Case op
                Case 1 To 4
                    right1 += 1
                Case 5 To 10
                    right2 += 1
            End Select
        End If
        Label6.Text = "目前答對" & right1 + right2 & "題"
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Do
            user = InputBox("請輸入姓名：")
        Loop Until user <> ""
        Call test()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call Cal()
        If no < 12 Then
            TextBox1.Text = ""
            Call test()
        Else
            score = right1 * Int((40 / num1) * 100 + 0.5) / 100 + right2 * Int((60 / num2) * 100 + 0.5) / 100
            MsgBox(user & "得分：" & score)
            End
        End If
    End Sub
End Class
