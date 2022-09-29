Public Class Form1
    Dim hp, vp, lp, n, tp As Integer

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim a(7, 5) As String
        Dim f(7, 5) As Boolean
        Dim x, y As Integer
        If n = 10 Then
            ListBox1.Items.Clear()
            Label17.Text = "十次平均值："
            n = 0
        End If
        n += 1
        For i = 0 To 5
            For j = 0 To 7
                If i = 0 Then
                    a(j, i) = "A"
                ElseIf i = 1 Then
                    a(j, i) = "B"
                ElseIf i = 2 Then
                    a(j, i) = "C"
                ElseIf i = 3 Then
                    a(j, i) = "D"
                ElseIf i = 4 Then
                    a(j, i) = "E"
                ElseIf i = 5 Then
                    a(j, i) = "F"
                End If
            Next
        Next
        For i = 1 To 16
            Do
                x = Int(Rnd() * 7)
                y = Int(Rnd() * 5)
                If f(x, y) <> True Then
                    Exit Do
                End If
            Loop
            Select Case i
                Case 1
                    Label1.Text = a(x, y)
                    Label1.Tag = a(x, y)
                    f(x, y) = True
                Case 2
                    Label2.Text = a(x, y)
                    Label2.Tag = a(x, y)
                    f(x, y) = True
                Case 3
                    Label3.Text = a(x, y)
                    Label3.Tag = a(x, y)
                    f(x, y) = True
                Case 4
                    Label4.Text = a(x, y)
                    Label4.Tag = a(x, y)
                    f(x, y) = True
                Case 5
                    Label5.Text = a(x, y)
                    Label5.Tag = a(x, y)
                    f(x, y) = True
                Case 6
                    Label6.Text = a(x, y)
                    Label6.Tag = a(x, y)
                    f(x, y) = True
                Case 7
                    Label7.Text = a(x, y)
                    Label7.Tag = a(x, y)
                    f(x, y) = True
                Case 8
                    Label8.Text = a(x, y)
                    Label8.Tag = a(x, y)
                    f(x, y) = True
                Case 9
                    Label9.Text = a(x, y)
                    Label9.Tag = a(x, y)
                    f(x, y) = True
                Case 10
                    Label10.Text = a(x, y)
                    Label10.Tag = a(x, y)
                    f(x, y) = True
                Case 11
                    Label11.Text = a(x, y)
                    Label11.Tag = a(x, y)
                    f(x, y) = True
                Case 12
                    Label12.Text = a(x, y)
                    Label12.Tag = a(x, y)
                    f(x, y) = True
                Case 13
                    Label13.Text = a(x, y)
                    Label13.Tag = a(x, y)
                    f(x, y) = True
                Case 14
                    Label14.Text = a(x, y)
                    Label14.Tag = a(x, y)
                    f(x, y) = True
                Case 15
                    Label15.Text = a(x, y)
                    Label15.Tag = a(x, y)
                    f(x, y) = True
                Case 16
                    Label16.Text = a(x, y)
                    Label16.Tag = a(x, y)
                    f(x, y) = True
            End Select
        Next
        Call Chkpointh()
        Call ChkpointV()
        Call ChkpointL()
        ListBox1.Items.Add("局" & n & "橫列" & hp)
        ListBox1.Items.Add("局" & n & "直列" & vp)
        ListBox1.Items.Add("局" & n & "對角線" & lp)
        ListBox1.Items.Add("局" & n & "總得分" & hp + vp + lp)
        tp += hp + vp + lp
        If n = 10 Then
            Label17.Text = "十次平均值：" & tp / 10
            tp = 0
        End If
    End Sub
    Function Chkpointh()
        '--------------------------水平
        hp = 0
        If Label1.Tag = Label2.Tag And Label3.Tag = Label4.Tag And Label1.Tag = Label3.Tag Then
            hp += 10
        ElseIf Label1.Tag = Label2.Tag And Label2.Tag = Label3.Tag Then
            hp += 5
        ElseIf Label1.Tag = Label2.Tag And Label2.Tag = Label4.Tag Then
            hp += 5
        ElseIf Label1.Tag = Label3.Tag And Label3.Tag = Label4.Tag Then
            hp += 5
        ElseIf Label2.Tag = Label3.Tag And Label3.Tag = Label4.Tag Then
            hp += 5
        End If
        If Label5.Tag = Label6.Tag And Label7.Tag = Label8.Tag And Label5.Tag = Label7.Tag Then
            hp += 10
        ElseIf Label5.Tag = Label6.Tag And Label6.Tag = Label7.Tag Then
            hp += 5
        ElseIf Label5.Tag = Label6.Tag And Label6.Tag = Label8.Tag Then
            hp += 5
        ElseIf Label5.Tag = Label7.Tag And Label7.Tag = Label8.Tag Then
            hp += 5
        ElseIf Label6.Tag = Label7.Tag And Label7.Tag = Label8.Tag Then
            hp += 5
        End If
        If Label9.Tag = Label10.Tag And Label11.Tag = Label12.Tag And Label9.Tag = Label11.Tag Then
            hp += 10
        ElseIf Label9.Tag = Label10.Tag And Label10.Tag = Label11.Tag Then
            hp += 5
        ElseIf Label9.Tag = Label10.Tag And Label10.Tag = Label12.Tag Then
            hp += 5
        ElseIf Label9.Tag = Label11.Tag And Label11.Tag = Label12.Tag Then
            hp += 5
        ElseIf Label10.Tag = Label11.Tag And Label11.Tag = Label12.Tag Then
            hp += 5
        End If
        If Label13.Tag = Label14.Tag And Label15.Tag = Label16.Tag And Label13.Tag = Label15.Tag Then
            hp += 10
        ElseIf Label13.Tag = Label14.Tag And Label14.Tag = Label15.Tag Then
            hp += 5
        ElseIf Label13.Tag = Label14.Tag And Label14.Tag = Label16.Tag Then
            hp += 5
        ElseIf Label13.Tag = Label15.Tag And Label15.Tag = Label16.Tag Then
            hp += 5
        ElseIf Label14.Tag = Label15.Tag And Label15.Tag = Label16.Tag Then
            hp += 5
        End If
    End Function
    Function ChkpointV()
        '--------------------------垂直
        vp = 0
        If Label1.Tag = Label5.Tag And Label9.Tag = Label13.Tag And Label1.Tag = Label9.Tag Then
            vp += 10
        ElseIf Label1.Tag = Label5.Tag And Label5.Tag = Label9.Tag Then
            vp += 5
        ElseIf Label1.Tag = Label5.Tag And Label5.Tag = Label13.Tag Then
            vp += 5
        ElseIf Label1.Tag = Label9.Tag And Label9.Tag = Label13.Tag Then
            vp += 5
        ElseIf Label5.Tag = Label9.Tag And Label9.Tag = Label13.Tag Then
            vp += 5
        End If
        If Label2.Tag = Label6.Tag And Label10.Tag = Label14.Tag And Label2.Tag = Label10.Tag Then
            vp += 10
        ElseIf Label2.Tag = Label6.Tag And Label6.Tag = Label10.Tag Then
            vp += 5
        ElseIf Label2.Tag = Label6.Tag And Label6.Tag = Label14.Tag Then
            vp += 5
        ElseIf Label2.Tag = Label10.Tag And Label10.Tag = Label14.Tag Then
            vp += 5
        ElseIf Label6.Tag = Label10.Tag And Label10.Tag = Label14.Tag Then
            vp += 5
        End If
        If Label3.Tag = Label7.Tag And Label11.Tag = Label15.Tag And Label3.Tag = Label11.Tag Then
            vp += 10
        ElseIf Label3.Tag = Label7.Tag And Label7.Tag = Label11.Tag Then
            vp += 5
        ElseIf Label3.Tag = Label7.Tag And Label7.Tag = Label15.Tag Then
            vp += 5
        ElseIf Label3.Tag = Label11.Tag And Label11.Tag = Label15.Tag Then
            vp += 5
        ElseIf Label7.Tag = Label11.Tag And Label11.Tag = Label15.Tag Then
            vp += 5
        End If
        If Label4.Tag = Label8.Tag And Label12.Tag = Label16.Tag And Label4.Tag = Label12.Tag Then
            vp += 10
        ElseIf Label4.Tag = Label8.Tag And Label8.Tag = Label12.Tag Then
            vp += 5
        ElseIf Label4.Tag = Label8.Tag And Label8.Tag = Label16.Tag Then
            vp += 5
        ElseIf Label4.Tag = Label12.Tag And Label12.Tag = Label16.Tag Then
            vp += 5
        ElseIf Label8.Tag = Label12.Tag And Label12.Tag = Label16.Tag Then
            vp += 5
        End If
    End Function
    Function ChkpointL()
        '--------------------------斜角
        lp = 0
        If Label1.Tag = Label6.Tag And Label11.Tag = Label16.Tag And Label1.Tag = Label11.Tag Then
            lp += 10
        ElseIf Label1.Tag = Label6.Tag And Label6.Tag = Label11.Tag Then
            lp += 5
        ElseIf Label1.Tag = Label6.Tag And Label6.Tag = Label16.Tag Then
            lp += 5
        ElseIf Label1.Tag = Label11.Tag And Label11.Tag = Label16.Tag Then
            lp += 5
        ElseIf Label6.Tag = Label11.Tag And Label11.Tag = Label16.Tag Then
            lp += 5
        End If
        If Label4.Tag = Label7.Tag And Label10.Tag = Label13.Tag And Label4.Tag = Label10.Tag Then
            lp += 10
        ElseIf Label4.Tag = Label7.Tag And Label7.Tag = Label10.Tag Then
            lp += 5
        ElseIf Label4.Tag = Label7.Tag And Label7.Tag = Label13.Tag Then
            lp += 5
        ElseIf Label4.Tag = Label10.Tag And Label10.Tag = Label13.Tag Then
            lp += 5
        ElseIf Label7.Tag = Label10.Tag And Label10.Tag = Label13.Tag Then
            lp += 5
        End If
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ListBox1.Items.Clear()
        Label17.Text = "十次平均值："
        n = 0
        tp = 0
        For k = 1 To 10
            Dim a(7, 5) As String
            Dim f(7, 5) As Boolean
            Dim x, y As Integer

            n += 1
            For i = 0 To 5
                For j = 0 To 7
                    If i = 0 Then
                        a(j, i) = "A"
                    ElseIf i = 1 Then
                        a(j, i) = "B"
                    ElseIf i = 2 Then
                        a(j, i) = "C"
                    ElseIf i = 3 Then
                        a(j, i) = "D"
                    ElseIf i = 4 Then
                        a(j, i) = "E"
                    ElseIf i = 5 Then
                        a(j, i) = "F"
                    End If
                Next
            Next
            For i = 1 To 16
                Do
                    x = Int(Rnd() * 7)
                    y = Int(Rnd() * 5)
                    If f(x, y) <> True Then
                        Exit Do
                    End If
                Loop
                Select Case i
                    Case 1
                        Label1.Text = a(x, y)
                        Label1.Tag = a(x, y)
                        f(x, y) = True
                    Case 2
                        Label2.Text = a(x, y)
                        Label2.Tag = a(x, y)
                        f(x, y) = True
                    Case 3
                        Label3.Text = a(x, y)
                        Label3.Tag = a(x, y)
                        f(x, y) = True
                    Case 4
                        Label4.Text = a(x, y)
                        Label4.Tag = a(x, y)
                        f(x, y) = True
                    Case 5
                        Label5.Text = a(x, y)
                        Label5.Tag = a(x, y)
                        f(x, y) = True
                    Case 6
                        Label6.Text = a(x, y)
                        Label6.Tag = a(x, y)
                        f(x, y) = True
                    Case 7
                        Label7.Text = a(x, y)
                        Label7.Tag = a(x, y)
                        f(x, y) = True
                    Case 8
                        Label8.Text = a(x, y)
                        Label8.Tag = a(x, y)
                        f(x, y) = True
                    Case 9
                        Label9.Text = a(x, y)
                        Label9.Tag = a(x, y)
                        f(x, y) = True
                    Case 10
                        Label10.Text = a(x, y)
                        Label10.Tag = a(x, y)
                        f(x, y) = True
                    Case 11
                        Label11.Text = a(x, y)
                        Label11.Tag = a(x, y)
                        f(x, y) = True
                    Case 12
                        Label12.Text = a(x, y)
                        Label12.Tag = a(x, y)
                        f(x, y) = True
                    Case 13
                        Label13.Text = a(x, y)
                        Label13.Tag = a(x, y)
                        f(x, y) = True
                    Case 14
                        Label14.Text = a(x, y)
                        Label14.Tag = a(x, y)
                        f(x, y) = True
                    Case 15
                        Label15.Text = a(x, y)
                        Label15.Tag = a(x, y)
                        f(x, y) = True
                    Case 16
                        Label16.Text = a(x, y)
                        Label16.Tag = a(x, y)
                        f(x, y) = True
                End Select
            Next
            Call Chkpointh()
            Call ChkpointV()
            Call ChkpointL()
            ListBox1.Items.Add("局" & n & "橫列" & hp)
            ListBox1.Items.Add("局" & n & "直列" & vp)
            ListBox1.Items.Add("局" & n & "對角線" & lp)
            ListBox1.Items.Add("局" & n & "總得分" & hp + vp + lp)
            tp += hp + vp + lp
            If n = 10 Then Label17.Text = "十次平均值：" & tp / 10
        Next
    End Sub


End Class
