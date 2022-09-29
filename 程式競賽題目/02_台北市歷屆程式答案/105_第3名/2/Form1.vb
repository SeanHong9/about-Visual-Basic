Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FileOpen(1, "in2.txt", 1)
        FileOpen(2, "out.txt", 2)
        Dim n, k, d As Integer
        Dim a, c, b As String
        Dim f As Boolean
        f = True
        a = LineInput(1)
        Dim x(Val(a) - 1) As Integer
        a = LineInput(1)
        PrintLine(2, a)
        n = a.Length
        k = 0
        For i = 1 To n
            c = Mid(a, i, 1)
            If Asc(c) <> 32 Then
                b &= c
            ElseIf Asc(c) = 32 Then
                x(k) = Val(b)
                b = ""
                k += 1
            End If
            If i = n Then
                x(k) = Val(b)
                b = ""
            End If
        Next
        k = UBound(x)
        d = 1
        For i = k To 1 Step -1
            ReDim Preserve x(k - d)
            d += 1
            c = ""
            'If f = True And i <> k Then
            '    Array.Reverse(x)
            'End If
            'f = Not f
            'x(i) = -1
            Array.Reverse(x)
            For j = 0 To UBound(x)
                If x(j) <> -1 Then
                    c &= x(j) & " "
                End If
            Next
            PrintLine(2, c)
        Next
        End
    End Sub
End Class
