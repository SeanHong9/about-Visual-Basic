Public Class Form1
    Dim n As Integer
    Dim all As Integer
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim rndnum, count(10) As Integer
        Dim eng(20) As String
        Dim score As Integer
        Dim avg As Single
        Dim this As Integer
        Randomize()


        For i = 1 To 16

            rndnum = Int(Rnd() * 6) + 1
            Select Case rndnum
                Case 1
                    eng(i) = "A"
                Case 2
                    eng(i) = "B"
                Case 3
                    eng(i) = "C"
                Case 4
                    eng(i) = "D"
                Case 5
                    eng(i) = "E"
                Case 6
                    eng(i) = "F"
            End Select
        Next
        Label1.Text = eng(1)
        Label2.Text = eng(2)
        Label3.Text = eng(3)
        Label4.Text = eng(4)
        Label5.Text = eng(5)
        Label6.Text = eng(6)
        Label7.Text = eng(7)
        Label8.Text = eng(8)
        Label9.Text = eng(9)
        Label10.Text = eng(10)
        Label11.Text = eng(11)
        Label12.Text = eng(12)
        Label13.Text = eng(13)
        Label14.Text = eng(14)
        Label15.Text = eng(15)
        Label16.Text = eng(16)
        n += 1                   '次數+1
        '-------------------------------------------------------
        For w = 0 To 3                      '橫排分數

            For qq = 1 To 6
                count(qq) = 0
            Next

            For q = 1 + 4 * w To 4 + 4 * w
                Select Case eng(q)
                    Case "A"
                        count(1) += 1
                    Case "B"
                        count(2) += 1
                    Case "C"
                        count(3) += 1
                    Case "D"
                        count(4) += 1
                    Case "E"
                        count(5) += 1
                    Case "F"
                        count(6) += 1
                End Select
                For qw = 1 To 6
                    If count(qw) = 3 Then
                        score += 5
                        all += 5
                    ElseIf count(qw) = 4 Then
                        score += 10
                        all = +10
                    End If
                Next
            Next
        Next                            '橫排分數
        ListBox1.Items.Add("橫列分數 " & score & "分")
        this += score
        score = 0
        '-------------------------------------------------------
        For w = 0 To 3                      '直排分數

            For qq = 1 To 6
                count(qq) = 0
            Next

            For q = 1 + w To 16 Step 4
                Select Case eng(q)
                    Case "A"
                        count(1) += 1
                    Case "B"
                        count(2) += 1
                    Case "C"
                        count(3) += 1
                    Case "D"
                        count(4) += 1
                    Case "E"
                        count(5) += 1
                    Case "F"
                        count(6) += 1
                End Select

                For qw = 1 To 6
                    If count(qw) = 3 Then
                        score += 5
                        all += 5
                    ElseIf count(qw) = 4 Then
                        score += 10
                        all += 10
                    End If
                Next
            Next
        Next                            '直排分數
        ListBox1.Items.Add("直行分數 " & score & "分")
        this += score
        score = 0
        '---------------------------------------------
        '左上右下

        For qq = 1 To 6
            count(qq) = 0
        Next
        For i = 1 To 16 Step 5
            Select Case eng(i)
                Case "A"
                    count(1) += 1
                Case "B"
                    count(2) += 1
                Case "C"
                    count(3) += 1
                Case "D"
                    count(4) += 1
                Case "E"
                    count(5) += 1
                Case "F"
                    count(6) += 1
            End Select

            For qw = 1 To 6
                If count(qw) = 3 Then
                    score += 5
                    all += 5
                ElseIf count(qw) = 4 Then
                    score += 10
                    all += 10
                End If
            Next
        Next

        '左上右下
        '------------------------------------
        '右上左下

        For qq = 1 To 6
            count(qq) = 0
        Next
        For i = 4 To 13 Step 3     '
            Select Case eng(i)
                Case "A"
                    count(1) += 1
                Case "B"
                    count(2) += 1
                Case "C"
                    count(3) += 1
                Case "D"
                    count(4) += 1
                Case "E"
                    count(5) += 1
                Case "F"
                    count(6) += 1
            End Select

            For qw = 1 To 6
                If count(qw) = 3 Then
                    score += 5
                    all += 5
                ElseIf count(qw) = 4 Then
                    score += 10
                    all += 10
                End If
            Next
        Next

        ListBox1.Items.Add("對角線分數 " & score & "分")
        this += score
        score = 0
        '右上左下
        '-----------------------------
        ListBox1.Items.Add("局" & n & "分數 " & this & "分")
        ListBox1.Items.Add(" ")
        If n = 10 Then
            Label17.BackColor = Color.Green
        End If
        Label17.Text = "第 " & n & " 回合"

        If n Mod 10 = 0 Then
            avg = all / 10
            ListBox1.Items.Add("最近10次平均 " & avg & "分")
            all = 0
            avg = 0
        End If
    End Sub
End Class
