Imports System.IO
Imports System.IO.Directory
Public Class LiberaItem
    Dim contaitens As Integer
    Dim liberaatualizar As Boolean = False
    Dim nomeprocura As String
    Dim saldo As Integer = 0
    Dim horario As String
    Dim data As String
    Dim datasembarras As String
    Private Sub LiberaItem_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim alturamax As Integer = My.Computer.Screen.Bounds.Height.ToString
        Dim larguramax As Integer = My.Computer.Screen.Bounds.Width.ToString
        Me.Size = New System.Drawing.Size((larguramax - (larguramax / 10)), (alturamax - (alturamax / 10)))
        'centralizar formulario
        Me.Location = New Point((Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2, (alturamax / 20))
        '''''
        TableLayoutPanel1.Font = New Font(TableLayoutPanel1.Font.Name, larguramax / MenuIdx.fonteforms)
        TableLayoutPanel3.Font = New Font(TableLayoutPanel3.Font.Name, larguramax / MenuIdx.fonteforms, style:=FontStyle.Bold)
        LinkLabel1.Font = New Font(LinkLabel1.Font.Name, larguramax / 180, style:=FontStyle.Bold)
        If Not MenuIdx.codrapidosalvo = 0 Then
            txtcodigo.Text = MenuIdx.codrapidosalvo
            buscacodigo()
        End If
    End Sub
    Private Sub IncluirCompra_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        MenuIdx.Button1.Enabled = True
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        limpaetrava()
    End Sub

    Private Sub revisacaracteres()
        txtdescbreve.Text = txtdescbreve.Text.Replace("À", "A").Replace("Á", "A").Replace("Â", "A").Replace("Ã", "A").Replace("Ç", "C").Replace("È", "E").Replace("Ê", "E").Replace("É", "E").Replace("Í", "I").Replace("Ó", "O").Replace("Õ", "O").Replace("Ô", "O").Replace("Ú", "U").Replace("-", "").Replace("*", "")
        txdesccompleta.Text = txdesccompleta.Text.Replace("À", "A").Replace("Á", "A").Replace("Â", "A").Replace("Ã", "A").Replace("Ç", "C").Replace("È", "E").Replace("Ê", "E").Replace("É", "E").Replace("Í", "I").Replace("Ó", "O").Replace("Õ", "O").Replace("Ô", "O").Replace("Ú", "U").Replace("-", "")
        txtfornecedor.Text = txtfornecedor.Text.Replace("À", "A").Replace("Á", "A").Replace("Â", "A").Replace("Ã", "A").Replace("Ç", "C").Replace("È", "E").Replace("Ê", "E").Replace("É", "E").Replace("Í", "I").Replace("Ó", "O").Replace("Õ", "O").Replace("Ô", "O").Replace("Ú", "U").Replace("-", "")
        txtmodelo.Text = txtmodelo.Text.Replace("À", "A").Replace("Á", "A").Replace("Â", "A").Replace("Ã", "A").Replace("Ç", "C").Replace("È", "E").Replace("Ê", "E").Replace("É", "E").Replace("Í", "I").Replace("Ó", "O").Replace("Õ", "O").Replace("Ô", "O").Replace("Ú", "U").Replace("-", "")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        liberaatualizar = False
        ListBox1.Items.Clear()
        Label8.Text = "RESULTADO DA PROCURA"
        contaitens = 0
        If Not txtcodigo.Text = "" Then
            buscacodigo()
        Else
            revisacaracteresbusca()

        End If
    End Sub
    Private Sub revisacaracteresbusca()
        txtdescbreve.Text = txtdescbreve.Text.Replace("À", "A").Replace("Á", "A").Replace("Â", "A").Replace("Ã", "A").Replace("Ç", "C").Replace("È", "E").Replace("Ê", "E").Replace("É", "E").Replace("Í", "I").Replace("Ó", "O").Replace("Õ", "O").Replace("Ô", "O").Replace("Ú", "U").Replace("-", "")
        buscadescricao()
    End Sub
    Private Sub buscacodigo()
        Dim arquivosemext As String = ""
        Try
            For Each foundfile As String In My.Computer.FileSystem.GetFiles(MenuIdx.diritens)
                arquivosemext = (foundfile.Replace(MenuIdx.diritens, "").Replace(".txt", ""))
                If arquivosemext.StartsWith(txtcodigo.Text & "-") Then
                    ListBox1.Items.Add(arquivosemext)
                    contaitens = contaitens + 1
                End If

            Next

        Catch ex As Exception

        End Try
        Label8.Text = "RESULTADO DA PROCURA (" & contaitens & " itens encontrados)"
        If ListBox1.Items.Count > 0 Then
            ListBox1.SelectedIndex = 0
            preenchimento()
        Else
            limpaetrava()
        End If
    End Sub
    Private Sub buscadescricao()
        Dim arquivosemext As String = ""
        Try
            'define o critério para o nome dos arquivos
            Dim criterio As String = IIf(txtdescbreve.Text = String.Empty, "*", txtdescbreve.Text)
            ' obtem os arquivos do diretorio selecioonado conforme o critério
            ' lista apenas os nomes dos arquivos usando uma expressão lambada e o método GetFileName
            Dim arquivos = GetFiles(MenuIdx.diritens, criterio).Select(Function(arq) Path.GetFileName(arq))
            If arquivos.Count > 0 Then
                'percorre e exibe os nomes dos arquivos
                For Each nomeArquivo As String In arquivos
                    ListBox1.Items.Add(nomeArquivo.Replace(".txt", ""))
                    contaitens = contaitens + 1
                Next
            Else
            End If


        Catch ex As Exception
            MessageBox.Show("Erro " + ex.Message)
        End Try
        Label8.Text = "RESULTADO DA PROCURA (" & contaitens & " itens encontrados)"
        If ListBox1.Items.Count > 0 Then
            ListBox1.SelectedIndex = 0
            preenchimento()
        Else
            limpaetrava()
        End If
    End Sub

    Private Sub txtcodigo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtcodigo.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))

        KeyAscii = CShort(numerosspermitidos(KeyAscii))

        If KeyAscii = 0 Then

            e.Handled = True
            MessageBox.Show("Somente Números!", "Somente números", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub
    Private Sub txtdescbreve_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtdescbreve.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))

        KeyAscii = CShort(letrasnumpermitidos(KeyAscii))

        If KeyAscii = 0 Then

            e.Handled = True

        End If

    End Sub

    Private Sub textbox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))

        KeyAscii = CShort(numerosspermitidos(KeyAscii))

        If KeyAscii = 0 Then

            e.Handled = True
            MessageBox.Show("Somente Números!", "Somente números", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        limparesultados()
        preenchimento()
    End Sub
    Private Sub preenchimento()

        Try

            If File.Exists(MenuIdx.diritens & ListBox1.Text & ".txt") Then
                Dim posSeparador0 = ListBox1.Text.IndexOf(" ")
                nomeprocura = ListBox1.Text.Substring(posSeparador0).Trim()
                txtdescbreve.Text = nomeprocura

                txtcodigo.Text = ListBox1.Text.Replace(txtdescbreve.Text, "").Replace("-", "").Trim()
                For Each linha As String In File.ReadAllLines(MenuIdx.diritens & ListBox1.Text & ".txt")
                    If linha.StartsWith("descompleta-") Then
                        Dim posSeparador = linha.IndexOf(" ")
                        txdesccompleta.Text = linha.Substring(posSeparador).Trim()
                    End If
                    If linha.StartsWith("fornecedor-") Then
                        Dim posSeparador = linha.IndexOf(" ")
                        txtfornecedor.Text = linha.Substring(posSeparador).Trim()
                    End If
                    If linha.StartsWith("modelo-") Then
                        Dim posSeparador = linha.IndexOf(" ")
                        txtmodelo.Text = linha.Substring(posSeparador).Trim()
                    End If
                    If linha.StartsWith("estoquemax-") Then
                        Dim posSeparador = linha.IndexOf(" ")
                        txtestoquemax.Text = linha.Substring(posSeparador).Trim()
                    End If
                    If linha.StartsWith("estoquemin-") Then
                        Dim posSeparador = linha.IndexOf(" ")
                        txtestoqueminimo.Text = linha.Substring(posSeparador).Trim()
                    End If


                Next
                txtcodigo.ReadOnly = True
                txtdescbreve.ReadOnly = True
                TextBox2.ReadOnly = False
                TextBox3.ReadOnly = False
                txtcodigo.BackColor = Color.Black
                txtcodigo.ForeColor = Color.White
                txtdescbreve.BackColor = Color.Black
                txtdescbreve.ForeColor = Color.White
                TextBox2.BackColor = Color.White
                TextBox2.ForeColor = Color.Black
                TextBox3.BackColor = Color.White
                TextBox3.ForeColor = Color.Black
            End If
            TextBox1.Text = saldo
            If File.Exists(MenuIdx.estoque & txtcodigo.Text & ".txt") Then

                For Each linha As String In File.ReadAllLines(MenuIdx.estoque & txtcodigo.Text & ".txt")
                    If linha.StartsWith("saldo-") Then
                        Dim posSeparador = linha.IndexOf(" ")
                        saldo = Convert.ToInt64(linha.Substring(posSeparador).Trim())
                        TextBox1.Text = saldo
                    End If

                Next
            End If
            liberaatualizar = True
        Catch ex As Exception

        End Try
        MenuIdx.codrapidosalvo = 0
    End Sub


    Private Sub limpaetrava()
        txtcodigo.ReadOnly = False
        txtdescbreve.ReadOnly = False
        TextBox2.ReadOnly = True
        TextBox3.ReadOnly = True
        txtcodigo.BackColor = Color.White
        txtcodigo.ForeColor = Color.Black
        txtdescbreve.BackColor = Color.White
        txtdescbreve.ForeColor = Color.Black
        TextBox2.BackColor = Color.Black
        TextBox2.ForeColor = Color.White
        TextBox3.BackColor = Color.Black
        TextBox3.ForeColor = Color.White
        txtcodigo.Clear()
        txtdescbreve.Clear()
        txdesccompleta.Clear()
        txtfornecedor.Clear()
        txtmodelo.Clear()
        txtestoquemax.Clear()
        txtestoqueminimo.Clear()
        ListBox1.Items.Clear()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
    End Sub
    Private Sub limparesultados()
        txdesccompleta.Clear()
        txtfornecedor.Clear()
        txtmodelo.Clear()
        txtestoquemax.Clear()
        txtestoqueminimo.Clear()
        saldo = 0
        TextBox2.Clear()
        TextBox3.Clear()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            If TextBox2.Text = "" Or TextBox2.Text = "0" Then
                MessageBox.Show("É necessário inserir um valor maior do que zero!", "Valor inválido", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf Convert.ToInt64(TextBox2.Text) > Convert.ToInt64(TextBox1.Text) Then
                MessageBox.Show("O estoque não possui saldo suficiente para substrair!", "Saldo insuficiente", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

            Else
                saldo = (saldo - Convert.ToInt64(TextBox2.Text))
                TextBox1.Text = saldo
                Try
                    Dim lines As New List(Of String)
                    Using sw As New StreamWriter(MenuIdx.estoque & "\" & txtcodigo.Text & ".txt")
                        For Each line As String In lines
                            sw.WriteLine(line)
                            sw.Close()
                        Next
                    End Using
                    Dim fluxoTexto As IO.StreamWriter
                    fluxoTexto = New IO.StreamWriter(MenuIdx.estoque & "\" & txtcodigo.Text & ".txt", True)
                    fluxoTexto.WriteLine("saldo- " & TextBox1.Text)


                    fluxoTexto.Close()
                    fluxoTexto.Dispose()
                    MessageBox.Show("Itens substraídos com sucesso!", "Pronto", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    adicionarelatorio()
                Catch ex As Exception
                    MessageBox.Show("Erro ao Salvar saldo de estoque!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                End Try
                TextBox2.Clear()
                TextBox3.Clear()
            End If

            limpaalertas()

            If Convert.ToInt64(TextBox1.Text) > Convert.ToInt64(txtestoquemax.Text) Then
                MessageBox.Show("O saldo atual em estoque é maior do que o estoque máximo!", "Estoque máximo ultrapassado", MessageBoxButtons.OK, MessageBoxIcon.Information)
                saldoalto()
            End If
            If Convert.ToInt64(TextBox1.Text) < Convert.ToInt64(txtestoqueminimo.Text) Then
                MessageBox.Show("O saldo atual está abaixo do estoque mínimo!", "Estoque mínimo abaixo do desejado", MessageBoxButtons.OK, MessageBoxIcon.Information)
                saldobaixo()
            End If
        Catch ex As Exception

        End Try
        MenuIdx.verificaalertas()
    End Sub
    Private Sub adicionarelatorio()
        Dim arquivorel As String = (MenuIdx.relatoriositens & txtcodigo.Text & "- " & datasembarras & ".txt")
        Try
            If File.Exists(arquivorel) Then
                Using sw As StreamWriter = File.AppendText(arquivorel)

                    sw.WriteLine("out- " & TextBox2.Text & "          " & TextBox3.Text)
                    sw.Close()


                End Using
            Else

                Dim lines2 As New List(Of String)
                Using sw2 As New StreamWriter(arquivorel)

                    sw2.WriteLine("out- " & TextBox2.Text & "          " & TextBox3.Text)
                    sw2.Close()
                End Using


            End If
        Catch ex As Exception

        End Try


    End Sub
    Private Sub limpaalertas()
        If File.Exists(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "(h).txt") Then
            File.Delete(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "(h).txt")
        End If
        If File.Exists(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "(l).txt") Then
            File.Delete(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "(l).txt")
        End If
    End Sub
    Private Sub saldoalto()
        Try
            If File.Exists(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "(l).txt") Then
                File.Delete(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "(l).txt")
            End If
            Dim lines As New List(Of String)
            Using sw As New StreamWriter(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "(h).txt")
                For Each line As String In lines
                    sw.WriteLine(line)
                    sw.Close()
                Next
            End Using
            Dim fluxoTexto As IO.StreamWriter
            fluxoTexto = New IO.StreamWriter(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "(h).txt", True)

            fluxoTexto.Close()
            fluxoTexto.Dispose()


        Catch ex As Exception
        End Try
    End Sub
    Private Sub saldobaixo()
        Try
            If File.Exists(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "(h).txt") Then
                File.Delete(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "(h).txt")
            End If
            Dim lines As New List(Of String)
            Using sw As New StreamWriter(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "(l).txt")
                For Each line As String In lines
                    sw.WriteLine(line)
                    sw.Close()
                Next
            End Using
            Dim fluxoTexto As IO.StreamWriter
            fluxoTexto = New IO.StreamWriter(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "(l).txt", True)

            fluxoTexto.Close()
            fluxoTexto.Dispose()


        Catch ex As Exception
        End Try
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try
            Dim hoje As DateTime = DateTime.Now
            data = hoje.ToString("yyyy/MM/dd")
            datasembarras = (Data.Replace("/", ""))
            horario = Microsoft.VisualBasic.Left(TimeString, 8)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("http://www.rkd.com.br")
    End Sub
End Class