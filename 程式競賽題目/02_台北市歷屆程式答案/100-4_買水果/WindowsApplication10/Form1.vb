Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim i%, j%, k%, z%, num%, n%(500, 3)
        For i = 0 To 9
            For j = 0 To 8
                For k = 0 To 7
                    For z = 0 To 6
                        If i + j + k + z = 23 Then
                            n(num, 0) = i : n(num, 1) = j : n(num, 2) = k : n(num, 3) = z
                            num += 1

                        End If
                    Next
                Next
            Next
        Next
        Debug.Print("總共有" & num & "種分法")
        For i = 0 To num - 1
            Debug.Print("老大：" & n(i, 0) & "個  " & "老二：" & n(i, 1) & "個  " & "老三：" & n(i, 2) & "個  " & "老四：" & n(i, 3) & "個")
        Next
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Dim i%, j%, k%, z%, num%, n%(50000, 2)
        For i = 0 To 100
            For j = 0 To 100
                For k = 0 To 100

                    If i + j + k = 100 And 50 * i + 10 * j + 10 * k <= 5000 Then
                        n(num, 0) = i : n(num, 1) = j : n(num, 2) = k
                        num += 1
                    End If
                Next
            Next
        Next

        Debug.Print("總共有" & num & "種分法")
        For i = 0 To num - 1
            Debug.Print("蘋果：" & n(i, 0) & "個  " & "奇異果：" & n(i, 1) & "個  " & "柳丁：" & n(i, 2) & "個")
        Next
    End Sub
End Class
