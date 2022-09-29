Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '建立動態產生的LED燈物件
        Dim i%, j%
        For i = 0 To 8
            For j = 0 To 4
                Dim Q As New Label
                Q.Left = 20 * i + 10
                Q.Top = 20 * j + 15
                Q.Width = 20
                Q.Height = 20
                Q.BackColor = Color.White
                Q.BorderStyle = BorderStyle.Fixed3D
                Q.Text = "○"  '用標點符號螢幕小鍵盤輸入
                Q.Name = "Q" + i.ToString + j.ToString
                Me.Controls.Add(Q)
            Next
        Next
    End Sub

    Private Function Q(ByVal i As Integer, ByVal j As Integer) As Label
        '用名稱當索引取得物件參考(這樣就能用習慣的陣列寫法寫指定位置)
        Return Me.Controls("Q" + i.ToString + j.ToString)
    End Function

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        End
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Timer1.Enabled = True
        m = 1


    End Sub
    Dim i%, m%, c%
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        

    End Sub

    Private Sub Timer0_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer0.Tick
        Label1.Text = "時間：" & TimeString()
        '顯示時間和控制燈號須用同一個Timer，否則換燈和換秒時間無法同步，因為Timer開始時間未必相同
        Select Case m
            Case 1
                If i <= 8 Then Q(i, 0).Text = "●" '用標點符號螢幕小鍵盤輸入
                If i > 0 Then Q(i - 1, 0).Text = "○"
                If i = 9 Then
                    Q(i - 1, 0).Text = "○"
                    i = 0 : m = 0
                Else
                    i = i + 1
                End If
            Case 2
                Select Case i
                    Case 0
                        Q(i, 0).Text = "●" : i = i + 1
                    Case Is <= 4
                        Q(i, 0).Text = "●"
                        Q(i - 1, 0).Text = "○" : i = i + 1
                    Case Is <= 8
                        Q(8 - i, 0).Text = "●"
                        Q(8 - i + 1, 0).Text = "○" : i = i + 1
                    Case 9
                        Q(0, 0).Text = "○"
                        i = 0 : m = 0
                End Select
            Case 3
                Select Case i
                    Case 0
                        Q(0, 0).Text = "●" : i = i + 1
                    Case Is <= 4
                        Q(i * 2, 0).Text = "●"
                        Q((i - 1) * 2, 0).Text = "○" : i = i + 1
                    Case 5
                        Q((i - 5) * 2 + 1, 0).Text = "●"
                        Q(8, 0).Text = "○" : i = i + 1
                    Case Is <= 8
                        Q((i - 5) * 2 + 1, 0).Text = "●"
                        Q((i - 6) * 2 + 1, 0).Text = "○" : i = i + 1
                    Case 9
                        Q(7, 0).Text = "○"
                        i = 0 : m = 0
                End Select
            Case 4
                Select Case i
                    Case 0
                        Q(0, 0).Text = "●" : i = i + 1
                    Case Is <= 4
                        Q(i, i).Text = "●"
                        Q(i - 1, i - 1).Text = "○" : i = i + 1
                    Case 5
                        Q(5, 3).Text = "●"
                        Q(4, 4).Text = "○" : i = i + 1
                    Case Is <= 8
                        Q(i, 8 - i).Text = "●"
                        Q((i - 1), 9 - i).Text = "○" : i = i + 1
                    Case 9
                        Q(8, 0).Text = "○"
                        i = 0 : m = 0
                End Select

            Case 5
                Select Case i
                    Case 0
                        Q(4, 4).Text = "●" : i = i + 1
                    Case Is <= 4
                        Q(4 - i, 4 - i).Text = "●"
                        Q(4 + i, 4 - i).Text = "●"
                        Q(5 - i, 5 - i).Text = "○"
                        Q(4 + (i - 1), 4 - (i - 1)).Text = "○" : i = i + 1

                    Case 5
                        Q(8, 0).Text = "○"
                        Q(0, 0).Text = "○"
                        i = 0 : m = 0
                End Select

            Case 6
                Dim a%, b%
                Select Case i

                    Case Is <= 2
                        For a = 0 To c
                            Q(i, a).Text = "●"
                        Next
                        i = i + 1 : c = c + 1
                    Case Is <= 4
                        If i = 3 Then c = 1
                        For a = 0 To c
                            Q(i, a).Text = "●"
                        Next
                        i = i + 1 : c = c - 1
                    Case Is <= 6
                        If i = 5 Then c = 1
                        For a = 0 To c
                            Q(i, a).Text = "●"
                        Next
                        i = i + 1 : c = c + 1
                    Case Is <= 8
                        If i = 7 Then c = 1
                        For a = 0 To c
                            Q(i, a).Text = "●"
                        Next
                        i = i + 1 : c = c - 1
                    Case 9

                        For a = 0 To 8
                            For b = 0 To 4
                                Q(a, b).Text = "○"
                            Next
                        Next

                        i = 0 : m = 0 : c = 0
                End Select


        End Select
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        m = 2
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        m = 3
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        m = 4
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        m = 5
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        m = 6
    End Sub
End Class
