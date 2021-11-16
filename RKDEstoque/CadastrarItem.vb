Imports System.IO

Public Class CadastrarItem
    Dim ultimateclaapertada As Integer
    Dim codigoatual As Integer = 0
    Private Sub CadastrarItem_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim alturamax As Integer = My.Computer.Screen.Bounds.Height.ToString
        Dim larguramax As Integer = My.Computer.Screen.Bounds.Width.ToString
        Me.Size = New System.Drawing.Size((larguramax - (larguramax / 10)), (alturamax - (alturamax / 10)))
        'centralizar formulario
        Me.Location = New Point((Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2, (alturamax / 20))
        '''''
        TableLayoutPanel1.Font = New Font(TableLayoutPanel1.Font.Name, larguramax / MenuIdx.fonteforms)
        TableLayoutPanel3.Font = New Font(TableLayoutPanel3.Font.Name, larguramax / MenuIdx.fonteforms, style:=FontStyle.Bold)
        LinkLabel1.Font = New Font(LinkLabel1.Font.Name, larguramax / 180, style:=FontStyle.Bold)
        leiturasiniciais()
    End Sub
    Private Sub Cadastraritem_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        MenuIdx.Button3.Enabled = True
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        txtdescbreve.Clear()
        txdesccompleta.Clear()
        txtfornecedor.Clear()
        txtmodelo.Clear()
        txtestoquemax.Clear()
        txtestoqueminimo.Clear()
    End Sub
    Private Sub revisacaracteres()
        txtdescbreve.Text = txtdescbreve.Text.Replace("À", "A").Replace("Á", "A").Replace("Â", "A").Replace("Ã", "A").Replace("Ç", "C").Replace("È", "E").Replace("Ê", "E").Replace("É", "E").Replace("Í", "I").Replace("Ó", "O").Replace("Õ", "O").Replace("Ô", "O").Replace("Ú", "U").Replace("-", "").Replace("*", "")
        txdesccompleta.Text = txdesccompleta.Text.Replace("À", "A").Replace("Á", "A").Replace("Â", "A").Replace("Ã", "A").Replace("Ç", "C").Replace("È", "E").Replace("Ê", "E").Replace("É", "E").Replace("Í", "I").Replace("Ó", "O").Replace("Õ", "O").Replace("Ô", "O").Replace("Ú", "U").Replace("-", "")
        txtfornecedor.Text = txtfornecedor.Text.Replace("À", "A").Replace("Á", "A").Replace("Â", "A").Replace("Ã", "A").Replace("Ç", "C").Replace("È", "E").Replace("Ê", "E").Replace("É", "E").Replace("Í", "I").Replace("Ó", "O").Replace("Õ", "O").Replace("Ô", "O").Replace("Ú", "U").Replace("-", "")
        txtmodelo.Text = txtmodelo.Text.Replace("À", "A").Replace("Á", "A").Replace("Â", "A").Replace("Ã", "A").Replace("Ç", "C").Replace("È", "E").Replace("Ê", "E").Replace("É", "E").Replace("Í", "I").Replace("Ó", "O").Replace("Õ", "O").Replace("Ô", "O").Replace("Ú", "U").Replace("-", "")

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            If txtcodigo.Text.Length < 1 Then
                MessageBox.Show("O código não pode ficar em branco!", "Código em branco!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf txtdescbreve.Text.Length < 3 Then
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
            ElseIf txtestoquemax.Text.Length > 9 Then
                MessageBox.Show("Valor inserido em estoque máximo muito alto!", "Valor inválido", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            ElseIf txtestoqueminimo.Text.Length > 9 Then
                MessageBox.Show("Valor inserido em estoque mínimo muito alto!", "Valor inválido", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                revisacaracteres()
                salva()
            End If

        Catch ex As Exception

        End Try


    End Sub
    Private Sub leiturasiniciais()

        Try

            If File.Exists(MenuIdx.codigos) Then
                For Each linha As String In File.ReadAllLines(MenuIdx.codigos)
                    If linha.StartsWith("codigo-") Then
                        Dim posSeparador = linha.IndexOf(" ")

                        codigoatual = Convert.ToInt64(linha.Substring(posSeparador).Trim())

                    End If


                Next

            End If

        Catch ex As Exception

        End Try
        txtcodigo.Text = codigoatual + 1
        If txtcodigo.Text = "" Then
            MessageBox.Show("Não foi gerado um código sequencial para o item!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Me.Close()
        End If
    End Sub
    Private Sub salva()
        Dim codusado As Boolean = False
        For Each foundfile As String In My.Computer.FileSystem.GetFiles(MenuIdx.diritens)
            If foundfile.Replace(MenuIdx.diritens, "").StartsWith(txtcodigo.Text & "- ") Then
                codusado = True
            End If
        Next
        If codusado = False Then
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

                If Convert.ToInt64(txtestoqueminimo.Text) > 0 Then
                    saldobaixo()
                End If
            Catch ex As Exception
                MessageBox.Show("Erro ao Salvar código!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            End Try


            ''''
            salvacodigo()
            ''''
            MenuIdx.verificaalertas()
            Dim result As Integer = MessageBox.Show("O item foi criado. Deseja ir diretamente para a tela Adicionar Itens?", "Item Criado!", MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                Me.Close()

            ElseIf result = DialogResult.Yes Then
                MenuIdx.codrapidosalvo = txtcodigo.Text
                IncluirCompra.Show()
                MenuIdx.Button2.Enabled = False
                Me.Close()
            End If

        Else
            Dim result As Integer = MessageBox.Show("O código solicitado já existe, deseja utilizar o próximo sequencial disponível?", "Código já utilizado!", MessageBoxButtons.YesNo)
            If result = DialogResult.No Then

            ElseIf result = DialogResult.Yes Then
                Dim proxcod As Integer = (Convert.ToInt64(txtcodigo.Text))
                Do While (codusado = True)
                    proxcod = (proxcod + 1)
                    codusado = False
                    For Each foundfile As String In My.Computer.FileSystem.GetFiles(MenuIdx.diritens)
                        If foundfile.Replace(MenuIdx.diritens, "").StartsWith(proxcod & "- ") Then
                            codusado = True
                        End If
                    Next
                Loop
                txtcodigo.Text = proxcod
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

                    If Convert.ToInt64(txtestoqueminimo.Text) > 0 Then
                        saldobaixo()
                    End If
                Catch ex As Exception
                    MessageBox.Show("Erro ao Salvar código!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                End Try

                ''''
                salvacodigo()
                ''''
                MenuIdx.verificaalertas()
                Dim result2 As Integer = MessageBox.Show("O item foi criado. Deseja ir diretamente para a tela Adicionar Itens?", "Item Criado!", MessageBoxButtons.YesNo)
                If result2 = DialogResult.No Then
                    Me.Close()

                ElseIf result2 = DialogResult.Yes Then
                    MenuIdx.codrapidosalvo = txtcodigo.Text
                    IncluirCompra.Show()
                    MenuIdx.Button2.Enabled = False
                    Me.Close()
                End If
            End If

        End If

    End Sub
    Private Sub salvacodigo()
        Try
            Dim lines As New List(Of String)
            Using sw As New StreamWriter(MenuIdx.codigos)
                For Each line As String In lines
                    sw.WriteLine(line)
                    sw.Close()
                Next
            End Using
            Dim fluxoTexto As IO.StreamWriter
            fluxoTexto = New IO.StreamWriter(MenuIdx.codigos, True)
            ''''
            If Not Convert.ToInt64(txtcodigo.Text) > 999999990 Then
                fluxoTexto.WriteLine("codigo- " & (Convert.ToInt64(txtcodigo.Text)))
            End If
            ''''


            fluxoTexto.Close()
            fluxoTexto.Dispose()


        Catch ex As Exception
            MessageBox.Show("Erro ao Salvar código!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Stop)
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
    Private Sub txtcodigo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtcodigo.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))

        KeyAscii = CShort(numerosspermitidos(KeyAscii))

        If KeyAscii = 0 Then

            e.Handled = True
            MessageBox.Show("Somente Números!", "Somente números", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("http://www.rkd.com.br")
    End Sub


End Class