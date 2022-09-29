Public Class Form1

    Private Sub HScrollBar1_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles HScrollBar1.Scroll
        Label4.Text = HScrollBar1.Value
        Label7.BackColor = Color.FromArgb(Val(Label4.Text), Val(Label5.Text), Val(Label6.Text))
    End Sub

    Private Sub HScrollBar2_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles HScrollBar2.Scroll
        Label5.Text = HScrollBar2.Value
        Label7.BackColor = Color.FromArgb(Val(Label4.Text), Val(Label5.Text), Val(Label6.Text))
    End Sub

    Private Sub HScrollBar3_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles HScrollBar3.Scroll
        Label6.Text = HScrollBar3.Value
        Label7.BackColor = Color.FromArgb(Val(Label4.Text), Val(Label5.Text), Val(Label6.Text))
    End Sub
End Class
