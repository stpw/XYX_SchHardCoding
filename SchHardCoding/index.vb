Imports System.IO
Imports System.Net
Imports System.Net.Dns
Imports System.Net.NetworkInformation
Imports System.Xml
Public Class index
    '輸完後,按Enter鍵,可跳到下一輸入點,Form之KeyPreview需設為True
    Private Sub index_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If (e.KeyValue.ToString = 13) Then
            System.Windows.Forms.SendKeys.Send("{tab}")
        End If
    End Sub
    Private Sub index_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class