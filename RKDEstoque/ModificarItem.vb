Imports System.IO
Imports System.IO.Directory
Public Class ModificarItem
    Dim contaitens As Integer
    Dim liberaatualizar As Boolean = False
    Dim nomeantigo As String
    Dim saldo As Integer = 0
    Private Sub ModificarItem_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
    Private Sub Modificaritem_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        MenuIdx.Button4.Enabled = True
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        limpaetrava()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            If txtdescbreve.Text.Length < 3 Then
                MessageBox.Show("A descrição Breve deve ter ao menos 3 caracteres!", "Descrição Breve muito curta!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf txdesccompleta.Text.Length < 3 Then
                MessageBox.Show("A descrição Completa deve ter ao menos 3 caracteres!", "Descrição Completa muito curta!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf txtfornecedor.Text.Length < 2 Then
                MessageBox.Show("O campo do nome da Marca/Fornecedor deve ter ao menos 2 caracteres!", "Marca ou Fornecedor muito curto!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf txtmodelo.Text.Length < 2 Then
                MessageBox.Show("O campo Modelo do Item deve ter ao menos 2 caracteres!", "Modelo muito curto!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf txtestoquemax.Text = "" Then
                MessageBox.Show("O campo Estoque máximo está vazio!", "Estoque máximo vazio!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf txtestoqueminimo.Text = "" Then
                MessageBox.Show("O campo Estoque mínimo está vazio!", "Estoque mínimo vazio!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf Convert.ToInt64(txtestoqueminimo.Text) > Convert.ToInt64(txtestoquemax.Text) Then
                MessageBox.Show("O estoque máximo deve ser maior ou igual ao estoque mínimo!", "Estoque máximo menor do que estoque mínimo!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf txtdescbreve.Text.Contains("  ") Then
                MessageBox.Show("Não se pode ter dois espaços seguidos na descrição breve!", "Descrição Breve com dois espaços seguidos!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf txdesccompleta.Text.Contains("  ") Then
                MessageBox.Show("Não se pode ter dois espaços seguidos na descrição completa!", "Descrição completa com dois espaços seguidos!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf txtfornecedor.Text.Contains("  ") Then
                MessageBox.Show("Não se pode ter dois espaços seguidos no nome da Marca/Fornecedor!", "Fornecedor com dois espaços seguidos!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf txtmodelo.Text.Contains("  ") Then
                MessageBox.Show("Não se pode ter dois espaços seguidos no modelo!", "Modelo com dois espaços seguidos!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf txtdescbreve.Text.StartsWith(" ") Then
                MessageBox.Show("Não se pode iniciar com espaço na descrição breve!", "Descrição Breve iniciando com espaço!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf txdesccompleta.Text.StartsWith(" ") Then
                MessageBox.Show("Não se pode ter iniciar com espaço na descrição completa!", "Descrição completa iniciando com espaço!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf txtfornecedor.Text.StartsWith(" ") Then
                MessageBox.Show("Não se pode ter iniciar com espaço no nome da Marca/Fornecedor!", "Fornecedor iniciando com espaço!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf txtmodelo.Text.StartsWith(" ") Then
                MessageBox.Show("Não se pode iniciar com espaço o modelo!", "Modelo iniciando com espaço!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf txtdescbreve.Text.EndsWith(" ") Then
                MessageBox.Show("Não se pode terminar com espaço na descrição breve!", "Descrição Breve terminando com espaço!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf txdesccompleta.Text.EndsWith(" ") Then
                MessageBox.Show("Não se pode ter terminar com espaço na descrição completa!", "Descrição completa terminando com espaço!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf txtfornecedor.Text.EndsWith(" ") Then
                MessageBox.Show("Não se pode ter terminar com espaço no nome da Marca/Fornecedor!", "Fornecedor terminando com espaço!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf txtmodelo.Text.EndsWith(" ") Then
                MessageBox.Show("Não se pode terminar com espaço o modelo!", "Modelo terminando com espaço!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf txtestoquemax.Text.length > 9 Then
                MessageBox.Show("Valor inserido em estoque máximo muito alto!", "Valor inválido", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf txtestoqueminimo.Text.length > 9 Then
                MessageBox.Show("Valor inserido em estoque mínimo muito alto!", "Valor inválido", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                revisacaracteres()
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub revisacaracteres()
        txtdescbreve.Text = txtdescbreve.Text.Replace("À", "A").Replace("Á", "A").Replace("Â", "A").Replace("Ã", "A").Replace("Ç", "C").Replace("È", "E").Replace("Ê", "E").Replace("É", "E").Replace("Í", "I").Replace("Ó", "O").Replace("Õ", "O").Replace("Ô", "O").Replace("Ú", "U").Replace("-", "").Replace("*", "")
        txdesccompleta.Text = txdesccompleta.Text.Replace("À", "A").Replace("Á", "A").Replace("Â", "A").Replace("Ã", "A").Replace("Ç", "C").Replace("È", "E").Replace("Ê", "E").Replace("É", "E").Replace("Í", "I").Replace("Ó", "O").Replace("Õ", "O").Replace("Ô", "O").Replace("Ú", "U").Replace("-", "")
        txtfornecedor.Text = txtfornecedor.Text.Replace("À", "A").Replace("Á", "A").Replace("Â", "A").Replace("Ã", "A").Replace("Ç", "C").Replace("È", "E").Replace("Ê", "E").Replace("É", "E").Replace("Í", "I").Replace("Ó", "O").Replace("Õ", "O").Replace("Ô", "O").Replace("Ú", "U").Replace("-", "")
        txtmodelo.Text = txtmodelo.Text.Replace("À", "A").Replace("Á", "A").Replace("Â", "A").Replace("Ã", "A").Replace("Ç", "C").Replace("È", "E").Replace("Ê", "E").Replace("É", "E").Replace("Í", "I").Replace("Ó", "O").Replace("Õ", "O").Replace("Ô", "O").Replace("Ú", "U").Replace("-", "")
        Dim result As Integer = MessageBox.Show("O item foi criado. Deseja ir diretamente para a tela a tela Adicionar Itens?", "Item Criado!", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            atualiza()
            Me.Close()

        ElseIf result = DialogResult.Yes Then
            atualiza()
            MenuIdx.codrapidosalvo = txtcodigo.Text
            IncluirCompra.Show()
            MenuIdx.Button2.Enabled = False
            Me.Close()
        End If
    End Sub
    Private Sub atualiza()
        Try
            Dim lines As New List(Of String)
            Using sw As New StreamWriter(MenuIdx.diritens & txtcodigo.Text & "- " & txtdescbreve.Text & ".txt")
                For Each line As String In lines
                    sw.WriteLine(line)
                    sw.Close()
                Next
            End Using
            Dim fluxoTexto As IO.StreamWriter
            fluxoTexto = New IO.StreamWriter(MenuIdx.diritens & txtcodigo.Text & "- " & txtdescbreve.Text & ".txt", True)
            fluxoTexto.WriteLine("descompleta- " & txdesccompleta.Text)
            fluxoTexto.WriteLine("fornecedor- " & txtfornecedor.Text)
            fluxoTexto.WriteLine("modelo- " & txtmodelo.Text)
            fluxoTexto.WriteLine("estoquemax- " & txtestoquemax.Text)
            fluxoTexto.WriteLine("estoquemin- " & txtestoqueminimo.Text)

            fluxoTexto.Close()
            fluxoTexto.Dispose()
            If Not nomeantigo = txtdescbreve.Text Then
                File.Delete(MenuIdx.diritens & txtcodigo.Text & "- " & nomeantigo & ".txt")
            End If
            limpaalertas()
            If saldo > Convert.ToInt64(txtestoquemax.Text) Then
                saldoalto()
            End If
            If saldo < Convert.ToInt64(txtestoqueminimo.Text) Then
                saldobaixo()
            End If
        Catch ex As Exception
            MessageBox.Show("Erro ao Salvar código!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        End Try
        MenuIdx.verificaalertas()
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
    Private Sub txdesccompleta_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txdesccompleta.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))

        KeyAscii = CShort(letrasnumpermitidos(KeyAscii))

        If KeyAscii = 0 Then

            e.Handled = True

        End If

    End Sub

    Private Sub txtestoquemax_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtestoquemax.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))

        KeyAscii = CShort(numerosspermitidos(KeyAscii))

        If KeyAscii = 0 Then

            e.Handled = True
            MessageBox.Show("Somente Números!", "Somente números", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub
    Private Sub txtestoqueminimo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtestoqueminimo.KeyPress
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
                nomeantigo = ListBox1.Text.Substring(posSeparador0).Trim()
                txtdescbreve.Text = nomeantigo

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
                txtcodigo.BackColor = Color.Black
                txtcodigo.ForeColor = Color.White
            End If
            If File.Exists(MenuIdx.estoque & txtcodigo.Text & ".txt") Then
                For Each linha As String In File.ReadAllLines(MenuIdx.estoque & txtcodigo.Text & ".txt")
                    If linha.StartsWith("saldo-") Then
                        Dim posSeparador = linha.IndexOf(" ")
                        saldo = Convert.ToInt64(linha.Substring(posSeparador).Trim())
                    End If
                Next
            End If
            readonlyfalse()
            liberaatualizar = True
        Catch ex As Exception

        End Try

    End Sub
    Private Sub readonlytrue()
        txdesccompleta.ReadOnly = True
        txdesccompleta.BackColor = Color.Black
        txdesccompleta.ForeColor = Color.White
        txtfornecedor.ReadOnly = True
        txtfornecedor.BackColor = Color.Black
        txtfornecedor.ForeColor = Color.White
        txtmodelo.ReadOnly = True
        txtmodelo.BackColor = Color.Black
        txtmodelo.ForeColor = Color.White
        txtestoquemax.ReadOnly = True
        txtestoquemax.BackColor = Color.Black
        txtestoquemax.ForeColor = Color.White
        txtestoqueminimo.ReadOnly = True
        txtestoqueminimo.BackColor = Color.Black
        txtestoqueminimo.ForeColor = Color.White
    End Sub
    Private Sub readonlyfalse()
        txdesccompleta.ReadOnly = False
        txdesccompleta.BackColor = Color.White
        txdesccompleta.ForeColor = Color.Black
        txtfornecedor.ReadOnly = False
        txtfornecedor.BackColor = Color.White
        txtfornecedor.ForeColor = Color.Black
        txtmodelo.ReadOnly = False
        txtmodelo.BackColor = Color.White
        txtmodelo.ForeColor = Color.Black
        txtestoquemax.ReadOnly = False
        txtestoquemax.BackColor = Color.White
        txtestoquemax.ForeColor = Color.Black
        txtestoqueminimo.ReadOnly = False
        txtestoqueminimo.BackColor = Color.White
        txtestoqueminimo.ForeColor = Color.Black
    End Sub
    Private Sub limpaetrava()
        txtcodigo.ReadOnly = False
        txtcodigo.BackColor = Color.White
        txtcodigo.ForeColor = Color.Black
        txtcodigo.Clear()
        txtdescbreve.Clear()
        txdesccompleta.Clear()
        txtfornecedor.Clear()
        txtmodelo.Clear()
        txtestoquemax.Clear()
        txtestoqueminimo.Clear()
        ListBox1.Items.Clear()
        readonlytrue()
    End Sub
    Private Sub limparesultados()
        txdesccompleta.Clear()
        txtfornecedor.Clear()
        txtmodelo.Clear()
        txtestoquemax.Clear()
        txtestoqueminimo.Clear()
    End Sub
    Private Sub limpaalertas()
        If File.Exists(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "h.txt") Then
            File.Delete(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "h.txt")
        End If
        If File.Exists(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "l.txt") Then
            File.Delete(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "l.txt")
        End If
    End Sub
    Private Sub saldoalto()
        Try
            If File.Exists(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "l.txt") Then
                File.Delete(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "l.txt")
            End If
            Dim lines As New List(Of String)
            Using sw As New StreamWriter(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "h.txt")
                For Each line As String In lines
                    sw.WriteLine(line)
                    sw.Close()
                Next
            End Using
            Dim fluxoTexto As IO.StreamWriter
            fluxoTexto = New IO.StreamWriter(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "h.txt", True)

            fluxoTexto.Close()
            fluxoTexto.Dispose()


        Catch ex As Exception
        End Try
    End Sub
    Private Sub saldobaixo()
        Try
            If File.Exists(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "h.txt") Then
                File.Delete(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "h.txt")
            End If
            Dim lines As New List(Of String)
            Using sw As New StreamWriter(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "l.txt")
                For Each line As String In lines
                    sw.WriteLine(line)
                    sw.Close()
                Next
            End Using
            Dim fluxoTexto As IO.StreamWriter
            fluxoTexto = New IO.StreamWriter(MenuIdx.alertasfldr & "\" & txtcodigo.Text & "l.txt", True)

            fluxoTexto.Close()
            fluxoTexto.Dispose()


        Catch ex As Exception
        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("http://www.rkd.com.br")
    End Sub
End Class