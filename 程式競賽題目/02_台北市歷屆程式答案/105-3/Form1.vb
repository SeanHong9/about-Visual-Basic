Public Class Form1
    Function f() As Integer
        Dim A$(3, 3), S$(47), T%(47), num%, C%(5), out$, score%, total%
        Dim i%, j%, k%, tm%
        For i = 0 To 47
            S(i) = Chr(65 + i \ 8)
        Next
        Randomize()
        For i = 0 To 3
            For j = 0 To 3
                Do
                    num = Int(Rnd() * 48)
                Loop Until T(num) = 0
                T(num) = 1
                A(i, j) = S(num)
                out = out & S(num)
            Next
            out = out & vbCrLf
        Next
        'MsgBox(out)

        '計算橫列得分
        For i = 0 To 3
            For k = 0 To 5
                C(k) = 0
            Next
            score = 0
            For j = 0 To 3
                tm = Asc(A(i, j)) - 65
                C(tm) = C(tm) + 1
            Next
            For k = 0 To 5
                If C(k) = 4 Then
                    score = score + 10
                ElseIf C(k) = 3 Then
                    score = score + 5
                End If
            Next
            total = total + score
        Next

        '計算直行得分
        For i = 0 To 3
            For k = 0 To 5
                C(k) = 0
            Next
            score = 0
            For j = 0 To 3
                tm = Asc(A(j, i)) - 65
                C(tm) = C(tm) + 1
            Next
            For k = 0 To 5
                If C(k) = 4 Then
                    score = score + 10
                ElseIf C(k) = 3 Then
                    score = score + 5
                End If
            Next
            total = total + score
        Next

        '計算對角線1得分
        For k = 0 To 5
            C(k) = 0
        Next
        score = 0
        For j = 0 To 3
            tm = Asc(A(j, j)) - 65
            C(tm) = C(tm) + 1
        Next
        For k = 0 To 5
            If C(k) = 4 Then
                score = score + 10
            ElseIf C(k) = 3 Then
                score = score + 5
            End If
        Next
        total = total + score

        '計算對角線2得分
        For k = 0 To 5
            C(k) = 0
        Next
        score = 0
        For j = 0 To 3
            tm = Asc(A(j, 3 - j)) - 65
            C(tm) = C(tm) + 1
        Next
        For k = 0 To 5
            If C(k) = 4 Then
                score = score + 10
            ElseIf C(k) = 3 Then
                score = score + 5
            End If
        Next
        total = total + score

        MsgBox(out & vbCrLf & "得分:" & total)
        Return total
    End Function
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim i, all As Integer

        For i = 1 To 10
            all = all + f()
        Next
        MsgBox("平均得分:" & all / 10)
    End Sub
End Class
