Imports System.IO      '記得寫此行
Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim f As New FileInfo("c:\t\test.txt")   '無法建立c:\test.txt,why?  '開新檔案
        Dim sw As StreamWriter = f.CreateText()   '產生streamwriter的sw物件
        sw.Write("VB 2010" & vbCrLf & "22")            '寫入  '用2行sw.Write不會有2行
        sw.Flush()                     '將存在緩衝區內的資料寫入指定的檔案內
        sw.Close()                  '關閉目前的資料流
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        '一次讀取檔案內所有資料
        Dim f As New FileInfo("c:\t\test.txt")
        Dim sr As StreamReader = f.OpenText()   '產生streamReaderr的sr物件
        MsgBox(sr.ReadToEnd())            '由目前位置開啟並取到資料流的末端
        sr.Close()

    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        '一次讀取檔案內一行資料
        Dim f As New FileInfo("c:\t\test.txt")
        Dim sr As StreamReader = f.OpenText()   '產生streamReaderr的sr物件

        Do While sr.Peek() >= 0        'sr.Peek可傳回下一個可讀取的字元，若傳回-1則表示沒有字元可讀取
            MsgBox(sr.ReadLine())            '由目前資料流開始讀取一行
        Loop

        sr.Close()
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        '一次讀取檔案內一字資料()
        Dim f As New FileInfo("c:\t\test.txt")
        Dim sr As StreamReader = f.OpenText()   '產生streamReaderr的sr物件

        Do While sr.Peek() >= 0        'sr.Peek可傳回下一個可讀取的字元，若傳回-1則表示沒有字元可讀取
            MsgBox(ChrW(sr.Read()))       'sr.Read()會往前移一個字元位置,並取得目前一個字元的字元碼
            ',接著再使用ChrW函式轉成所要顯示的正確字元
        Loop

        sr.Close()
    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        '將新資料附加在資料檔的最後
        Dim f As New FileInfo("c:\t\test.txt")   '開啟舊檔
        Dim sw As StreamWriter = f.AppendText()   '產生streamwriter的sw物件
        sw.Write("33")            '附加在檔案的最後
        sw.Flush()                     '將存在緩衝區內的資料寫入指定的檔案內
        sw.Close()
    End Sub

    Private Sub Button7_Click(sender As System.Object, e As System.EventArgs) Handles Button7.Click
        '資料檔刪除
        Dim f As New FileInfo("c:\t\test.txt")
        If f.Exists Then
            f.Delete()          '刪除檔案
            MsgBox("檔案已成功刪除")
        Else
            MsgBox("檔案找不到")
        End If
        
    End Sub
End Class
