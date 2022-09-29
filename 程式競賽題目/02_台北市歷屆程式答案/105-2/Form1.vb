Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim n, st, ed, c, i, j As Integer
        Dim s As String
        Dim A(), n1() As String
        ' Dim A%() = {88, 77, 66, 55, 44, 33, 22, 11}
        'n = 8

        Dim p As FileIO.TextFieldParser = My.Computer.FileSystem.OpenTextFieldParser("d:\105-2.txt", " ")
        n1 = p.ReadFields()
        A = p.ReadFields()
        n = Val(n1(0))

        st = 0 : ed = n - 1
        For i = 1 To n
            If i Mod 2 = 1 Then

                For j = st To ed
                    s = s & A(j) & " "
                Next
                ed = ed - 1
            Else
                For j = ed To st Step -1
                    s = s & A(j) & " "
                Next
                st = st + 1
            End If
            s = s & vbCrLf

        Next
        MsgBox(s)


    End Sub
End Class
