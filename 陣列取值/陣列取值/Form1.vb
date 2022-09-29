Imports System.IO

Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '一次讀取檔案內所有資料
        Dim f As New FileInfo("c:\t\test.txt")
        Dim sr As StreamReader = f.OpenText()   '產生streamReaderr的sr物件
        Dim take(), take1, B(13), C(13) As String
        Dim i, ii, time, A(13), D(13) As Integer
        take1 = sr.ReadToEnd()           '由目前位置開啟並取到資料流的末端
        take = Split(take1, " ")
        time = take(0)
        Dim mum(13) As Integer

        For i = 1 To time                '4*4陣列 A1~D4
            A(i) = take(i + 0 + 4 * i)
            B(i) = take(i + 1 + 4 * i)
            C(i) = take(i + 2 + 4 * i)
            D(i) = take(i + 3 + 4 * i)
            B(i) = Val(Mid(B(i), 1, 2))
            C(i) = Val(Mid(C(i), 1, 2))
        Next

        For ii = 1 To time
            mum(ii) = Val(B(ii)) - Val(C(ii))
            mum(ii) = (mum(ii) ^ 2 ^ 0.5) / D(ii)
        Next

        For 


            sr.Close()
    End Sub
End Class
