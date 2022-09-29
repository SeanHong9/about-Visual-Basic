Public Class Form1
    Dim a%(3, 3), s$
    Sub showa()
        If a(0, 0) = 0 Then Label1.Text = "" Else Label1.Text = a(0, 0)
        If a(0, 1) = 0 Then Label2.Text = "" Else Label2.Text = a(0, 1)
        If a(0, 2) = 0 Then Label3.Text = "" Else Label3.Text = a(0, 2)
        If a(0, 3) = 0 Then Label4.Text = "" Else Label4.Text = a(0, 3)
        If a(1, 0) = 0 Then Label5.Text = "" Else Label5.Text = a(1, 0)
        If a(1, 1) = 0 Then Label6.Text = "" Else Label6.Text = a(1, 1)
        If a(1, 2) = 0 Then Label7.Text = "" Else Label7.Text = a(1, 2)
        If a(1, 3) = 0 Then Label8.Text = "" Else Label8.Text = a(1, 3)
        If a(2, 0) = 0 Then Label9.Text = "" Else Label9.Text = a(2, 0)
        If a(2, 1) = 0 Then Label10.Text = "" Else Label10.Text = a(2, 1)
        If a(2, 2) = 0 Then Label11.Text = "" Else Label11.Text = a(2, 2)
        If a(2, 3) = 0 Then Label12.Text = "" Else Label12.Text = a(2, 3)
        If a(3, 0) = 0 Then Label13.Text = "" Else Label13.Text = a(3, 0)
        If a(3, 1) = 0 Then Label14.Text = "" Else Label14.Text = a(3, 1)
        If a(3, 2) = 0 Then Label15.Text = "" Else Label15.Text = a(3, 2)
        If a(3, 3) = 0 Then Label16.Text = "" Else Label16.Text = a(3, 3)
    End Sub
    Sub adda()
        Dim i%, j%, c%
        Do
            i = Int(Rnd() * 4)
            j = Int(Rnd() * 4)
            If a(i, j) = 0 Then
                c = 1
                a(i, j) = Int(Rnd() * 2 + 1) * 2
                Exit Do
            End If

        Loop Until c <> 0
    End Sub
    Sub upa()
        Dim i%, j%, b(3, 3)
        For j = 0 To 3
            For i = 3 To 1 Step -1
                If a(i, j) <> 0 and a(i - 1, j) = 0 Then
                    a(i - 1, j) = a(i, j)
                    a(i, j) = 0
                End If
            Next
            For i = 3 To 0 Step -1
                b(i, j) = a(i, j)
            Next
            For i = 3 To 1 Step -1
                If b(i, j) <> 0 And b(i - 1, j) = b(i, j) Then
                    a(i - 1, j) = a(i, j) + a(i - 1, j)
                    a(i, j) = 0
                End If
            Next
        Next

    End Sub

    Sub downa()
        Dim i%, j%, b(3, 3)
        For j = 0 To 3
            For i = 0 To 2
                If a(i, j) <> 0 And a(i + 1, j) = 0 Then
                    a(i + 1, j) = a(i, j)
                    a(i, j) = 0
                End If
            Next
            For i = 0 To 3
                b(i, j) = a(i, j)
            Next
            For i = 0 To 2
                If b(i, j) <> 0 And b(i + 1, j) = b(i, j) Then
                    a(i + 1, j) = a(i, j) + a(i + 1, j)
                    a(i, j) = 0
                End If
            Next
        Next

    End Sub
    Sub righta()
        Dim i%, j%, b(3, 3)
        For i = 0 To 3
            For j = 0 To 2
                If a(i, j) <> 0 And a(i, j + 1) = 0 Then
                    a(i, j + 1) = a(i, j)
                    a(i, j) = 0
                End If
            Next
            For j = 0 To 3
                b(i, j) = a(i, j)
            Next
            For j = 0 To 2
                If b(i, j) <> 0 And b(i, j + 1) = b(i, j) Then
                    a(i, j + 1) = a(i, j) + a(i, j + 1)
                    a(i, j) = 0
                End If
            Next
        Next

    End Sub

    Sub lefta()
        Dim i%, j%, b(3, 3)
        For i = 0 To 3
            For j = 3 To 1 Step -1
                If a(i, j) <> 0 And a(i, j - 1) = 0 Then
                    a(i, j - 1) = a(i, j)
                    a(i, j) = 0
                End If
            Next
            For j = 0 To 3
                b(i, j) = a(i, j)
            Next
            For j = 3 To 1 Step -1
                If b(i, j) <> 0 And b(i, j - 1) = b(i, j) Then
                    a(i, j - 1) = a(i, j) + a(i, j - 1)
                    a(i, j) = 0
                End If
            Next
        Next

    End Sub
    Sub finda(ByRef nz%, ByRef num%)
        nz = 0 : num = 0
        For i = 0 To 3
            For j = 0 To 3
                If a(i, j) = 0 Then
                    nz = nz + 1
                ElseIf a(i, j) = 2048 Then
                    num = 1
                End If
            Next
        Next
       
    End Sub
    Private Sub Form1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Dim nz%, num%
        Select Case (e.KeyCode)
            Case Keys.Up '按向上鍵
                Call upa()
            Case Keys.Down
                Call downa()
            Case Keys.Left
                Call lefta()
            Case Keys.Right
                Call righta()
        End Select
        Call finda(nz, num)
        If nz = 0 Then
            MsgBox("Game over")
            End
        ElseIf num = 1 Then
            MsgBox("Find 2048")
            End
        Else
            Call adda()

        End If
        Call showa()
    End Sub
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim i%, j%
       

        Call showa()


    End Sub
End Class
