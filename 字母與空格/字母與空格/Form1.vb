Imports System.IO      '記得寫此行
Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '一次讀取檔案內一行資料
        Dim f As New FileInfo("c:\t\test.txt")
        Dim sr As StreamReader = f.OpenText()   '產生streamReaderr的sr物件
        Dim q As String
        Dim countspace, counteng As Integer

        Do While sr.Peek() >= 0        'sr.Peek可傳回下一個可讀取的字元，若傳回-1則表示沒有字元可讀取
            q = sr.ReadLine()
        Loop

        For i = 1 To q.Length '0=48 A=65 Z=90 a=97 z=122 space=32
            If Val(Asc(Mid(q, i, 1))) > 64 And Val(Asc(Mid(q, i, 1))) < 91 Or Val(Asc(Mid(q, i, 1))) > 96 And Val(Asc(Mid(q, i, 1))) < 123 Then
                counteng += 1
            Else
                countspace += 1
            End If
        Next
        MsgBox("eng=" & counteng & "個", , )
        MsgBox("num=" & countspace & "個", , )



    End Sub

End Class
