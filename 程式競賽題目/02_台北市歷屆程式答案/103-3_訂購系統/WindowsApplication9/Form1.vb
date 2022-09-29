Public Class Form1

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        If TextBox1.Text <> "" And TextBox2.Text <> "" Then
            ListBox1.Items.Add(TextBox1.Text & "－－$" & TextBox2.Text)
            TextBox1.Text = ""
            TextBox2.Text = ""
        End If
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim nam$
        If ListBox1.SelectedIndex <> -1 Then
            'nam = ListBox1.SelectedItem
            nam = ListBox1.SelectedItem.ToString()
            ListBox2.Items.Add(nam)
            ListBox1.Items.Remove(nam)
        End If

    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Dim nam$
        If ListBox2.SelectedIndex <> -1 Then
            nam = ListBox2.SelectedItem.ToString()
            ListBox1.Items.Add(nam)
            ListBox2.Items.Remove(nam)
        End If
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Dim i%, nam$
        For i = 0 To ListBox1.Items.Count - 1
            ListBox1.SelectedIndex = i
            nam = ListBox1.SelectedItem.ToString()
            ListBox2.Items.Add(nam)
        Next
        ListBox1.Items.Clear()
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Dim i%, nam$
        For i = 0 To ListBox2.Items.Count - 1
            ListBox2.SelectedIndex = i
            nam = ListBox2.SelectedItem.ToString()
            ListBox1.Items.Add(nam)
        Next
        ListBox2.Items.Clear()
    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        Dim i%, nam$, total&
        For i = 0 To ListBox2.Items.Count - 1
            ListBox2.SelectedIndex = i
            nam = ListBox2.SelectedItem.ToString()
            total = total + Val(Strings.Right(nam, Len(nam) - InStr(nam, "$")))
        Next
        MsgBox("總金額：" & total, , "訂購系統")
    End Sub
End Class
