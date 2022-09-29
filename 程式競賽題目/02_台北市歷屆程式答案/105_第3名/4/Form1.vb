Public Class Form1

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim a, b, c As Integer
        Dim l As Label
        a = ComboBox1.SelectedIndex
        'Label73.Text = a
        'l.Name = Label & 24 + b + 1
        'l.Text = "hhh"
        If a <> -1 And b <> -1 And c <> -1 Then
            Select Case a
                Case 0

            End Select
        End If
    End Sub


    Private Sub Label32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label25.Click, Label26.Click, Label27.Click, Label28.Click, Label29.Click, Label30.Click, Label31.Click, Label32.Click, Label33.Click, Label34.Click, Label35.Click, Label36.Click, Label37.Click, Label38.Click, Label39.Click, Label40.Click, Label41.Click, Label42.Click, Label43.Click, Label44.Click, Label45.Click, Label46.Click, Label47.Click, Label48.Click, Label49.Click, Label50.Click, Label51.Click, Label52.Click, Label53.Click, Label54.Click, Label55.Click, Label56.Click, Label57.Click, Label58.Click, Label59.Click, Label60.Click, Label61.Click, Label62.Click, Label63.Click, Label64.Click, Label65.Click, Label66.Click, Label67.Click, Label68.Click, Label69.Click, Label70.Click, Label71.Click, Label72.Click, Label73.Click, Label74.Click, Label75.Click, Label76.Click, Label77.Click, Label78.Click, Label79.Click, Label80.Click, Label81.Click, Label82.Click, Label83.Click, Label84.Click, Label85.Click, Label86.Click, Label87.Click, Label88.Click
        Dim a As Integer
        a = ComboBox1.SelectedIndex
        If a <> -1 Then
            Select Case a
                Case 0
                    sender.text = "O"
                Case 1
                    sender.text = "X"
                Case 2
                    sender.text = ""
            End Select
        End If
    End Sub
    
   
    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        Dim q As New Label
        With q
            .Width = 30
            .Height = 30
            .BackColor = Color.Gray
            .Left = (Label1.Left + Label1.Right) \ 2
            .Top = Label1.Top
            .TabIndex = 0
        End With
        Me.Controls.Add(q)
    End Sub
End Class
