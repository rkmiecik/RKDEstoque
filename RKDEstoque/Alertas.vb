Imports System.IO
Imports System.IO.Directory
Public Class Alertas
    Private Sub Alertas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim alturamax As Integer = My.Computer.Screen.Bounds.Height.ToString
        Dim larguramax As Integer = My.Computer.Screen.Bounds.Width.ToString
        Me.Size = New System.Drawing.Size((larguramax - (larguramax / 10)), (alturamax - (alturamax / 10)))
        'centralizar formulario
        Me.Location = New Point((Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2, (alturamax / 20))
        '''''
        TableLayoutPanel1.Font = New Font(TableLayoutPanel1.Font.Name, larguramax / MenuIdx.fonteforms)
        LinkLabel1.Font = New Font(LinkLabel1.Font.Name, larguramax / 180, style:=FontStyle.Bold)
        Dim arquivosemext As String = ""
        Dim arquivosemextebool As String = ""
        Dim boolsign As String = ""
        Dim stringpos As String = ""
        Try
            For Each foundfile As String In My.Computer.FileSystem.GetFiles(MenuIdx.alertasfldr)
                arquivosemextebool = (foundfile.Replace(MenuIdx.alertasfldr, "").Replace("(h).txt", "").Replace("(l).txt", ""))
                arquivosemext = (foundfile.Replace(MenuIdx.alertasfldr, "").Replace(".txt", ""))
                boolsign = (foundfile.Replace(MenuIdx.alertasfldr, "").Replace(".txt", "").Replace(arquivosemextebool, ""))

                If boolsign = "(h)" Then
                    stringpos = "ESTOQUE ALTO"
                Else
                    stringpos = "ESTOQUE BAIXO"
                    End If
                ListBox1.Items.Add(stringpos & " NO ITEM      CÓDIGO       " & arquivosemextebool & "  ")

            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Alertas_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        MenuIdx.Button6.Enabled = True
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("http://www.rkd.com.br")
    End Sub
End Class