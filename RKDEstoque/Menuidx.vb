
Imports System.IO
Public Class MenuIdx
    Public codrapidosalvo As Integer
    Public fonteforms As Integer = 170
    Public diritens As String = (Application.StartupPath & "\cadastros\itens\")
    Public numseq As String = (Application.StartupPath & "\cadastros\sequencia\")
    Public estoque As String = (Application.StartupPath & "\cadastros\estoque\")
    Public alertasfldr As String = (Application.StartupPath & "\cadastros\alertas\")
    Public relatoriositens As String = (Application.StartupPath & "\cadastros\relatorios\")
    Public codigos As String = (Application.StartupPath & "\cadastros\sequencia\codigo.txt")
    Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Integer) As Short
    Dim numeroalertas As Integer
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            If Directory.Exists(diritens) Then
            Else
                Directory.CreateDirectory(diritens)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao criar o diretório: " & diritens, "Erro de diretório", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
        Try
            If Directory.Exists(numseq) Then
            Else
                Directory.CreateDirectory(numseq)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao criar o diretório: " & numseq, "Erro de diretório", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
        Try
            If Directory.Exists(numseq) Then
            Else
                Directory.CreateDirectory(numseq)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao criar o diretório: " & numseq, "Erro de diretório", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
        Try
            If Directory.Exists(estoque) Then
            Else
                Directory.CreateDirectory(estoque)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao criar o diretório: " & estoque, "Erro de diretório", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
        Try
            If Directory.Exists(alertasfldr) Then
            Else
                Directory.CreateDirectory(alertasfldr)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao criar o diretório: " & alertasfldr, "Erro de diretório", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
        Try
            If Directory.Exists(relatoriositens) Then
            Else
                Directory.CreateDirectory(relatoriositens)
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao criar o diretório: " & relatoriositens, "Erro de diretório", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
        Dim alturamax As Integer = My.Computer.Screen.Bounds.Height.ToString
        Dim larguramax As Integer = My.Computer.Screen.Bounds.Width.ToString
        Me.Size = New System.Drawing.Size((larguramax - (larguramax / 1.5)), (alturamax - (alturamax / 2)))
        'centralizar formulario
        Me.Location = New Point((Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2, (alturamax / 4))
        '''''
        TableLayoutPanel1.Font = New Font(TableLayoutPanel1.Font.Name, larguramax / 100, style:=FontStyle.Bold)
        LinkLabel1.Font = New Font(LinkLabel1.Font.Name, larguramax / 180, style:=FontStyle.Bold)
        verificaalertas()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Button1.Enabled = False
        LiberaItem.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Button2.Enabled = False
        IncluirCompra.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Button3.Enabled = False
        CadastrarItem.Show()
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Button4.Enabled = False
        ModificarItem.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Button5.Enabled = False
        Relatorios.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Button6.Enabled = False
        Alertas.Show()
    End Sub
    Public Sub verificaalertas()
        numeroalertas = 0
        Button6.Text = "(0) Alertas"
        Button6.BackColor = Color.White
        Button6.ForeColor = Color.Black
        Try
            For Each foundfile As String In My.Computer.FileSystem.GetFiles(alertasfldr)

                numeroalertas = numeroalertas + 1
                Button6.Text = "(" & numeroalertas & ") Alertas"
                Button6.BackColor = Color.DarkRed
                Button6.ForeColor = Color.White
            Next

        Catch ex As Exception

        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("http://www.rkd.com.br")
    End Sub
End Class
