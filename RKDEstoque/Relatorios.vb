Imports System.IO
Imports System.IO.Directory
Public Class Relatorios
    Dim datacap As String
    Dim datainicialinv As String
    Dim datainicialinvsembarras As Integer
    Dim datafinalinv As String
    Dim datafinalinvsembarras As Integer
    Dim codigopesquisado As String
    Dim arquivosemext As String
    Dim dataarquivo As Integer
    Private Sub Relatorios_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim alturamax As Integer = My.Computer.Screen.Bounds.Height.ToString
        Dim larguramax As Integer = My.Computer.Screen.Bounds.Width.ToString
        Me.Size = New System.Drawing.Size((larguramax - (larguramax / 10)), (alturamax - (alturamax / 10)))
        'centralizar formulario
        Me.Location = New Point((Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2, (alturamax / 20))
        '''''
        TableLayoutPanel1.Font = New Font(TableLayoutPanel1.Font.Name, larguramax / MenuIdx.fonteforms)
        TableLayoutPanel3.Font = New Font(TableLayoutPanel3.Font.Name, larguramax / MenuIdx.fonteforms, style:=FontStyle.Bold)
        LinkLabel1.Font = New Font(LinkLabel1.Font.Name, larguramax / 180, style:=FontStyle.Bold)
    End Sub
    Private Sub Relatorios_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        MenuIdx.Button5.Enabled = True
    End Sub
    Private Sub textbox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        Dim KeyAscii As Short = CShort(Asc(e.KeyChar))

        KeyAscii = CShort(numerosspermitidos(KeyAscii))

        If KeyAscii = 0 Then

            e.Handled = True
            MessageBox.Show("Somente Números!", "Somente números", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        RichTextBox1.Clear()
        datainicialinvsembarras = 0
        datafinalinvsembarras = 99999999
        Try
            If Not MaskedTextBox1.Text = "  /  /" Then
                Dim datainicial As DateTime = MaskedTextBox1.Text
                datainicialinv = datainicial.ToString("yyyy/MM/dd")
                datainicialinvsembarras = Convert.ToInt64(datainicialinv.Replace("/", ""))
            End If
            If Not MaskedTextBox2.Text = "  /  /" Then
                Dim datafinal As DateTime = MaskedTextBox2.Text
                datafinalinv = datafinal.ToString("yyyy/MM/dd")
                datafinalinvsembarras = Convert.ToInt64(datafinalinv.Replace("/", ""))
            End If
            If Not TextBox1.Text = "" Then
                codigopesquisado = TextBox1.Text & "-"
            Else
                codigopesquisado = ""
            End If

            If Not datainicialinvsembarras > datafinalinvsembarras Then
                For Each foundfile As String In My.Computer.FileSystem.GetFiles(MenuIdx.relatoriositens)
                    arquivosemext = (foundfile.Replace(MenuIdx.relatoriositens, "").Replace(".txt", ""))
                    dataarquivo = Convert.ToInt64(Microsoft.VisualBasic.Right(arquivosemext, 8))
                    If arquivosemext.StartsWith(codigopesquisado) And datainicialinvsembarras <= dataarquivo And dataarquivo <= datafinalinvsembarras Then
                        Dim codigoinserir As String = Microsoft.VisualBasic.Left(arquivosemext, 2)
                        Dim ano As String = Microsoft.VisualBasic.Left(dataarquivo, 4)
                        Dim dia As String = Microsoft.VisualBasic.Right(dataarquivo, 2)
                        Dim mesano As String = Microsoft.VisualBasic.Left(dataarquivo, 6)
                        Dim mes As String = Microsoft.VisualBasic.Right(mesano, 2)
                        For Each linha As String In File.ReadAllLines(foundfile)

                            If ComboBox1.Text = "ADICIONADOS" Then
                                If linha.StartsWith("in-") Then
                                    RichTextBox1.AppendText(codigoinserir & "          " & dia & "/" & mes & "/" & ano & "          " & linha.Replace("in-", "ADICIONOU ").Replace("out-", "SUBTRAIU ") & Environment.NewLine)
                                End If
                                Dim posSeparador = linha.IndexOf(" ")
                            ElseIf ComboBox1.Text = "SUBTRAÍDOS" Then
                                If linha.StartsWith("out-") Then
                                    RichTextBox1.AppendText(codigoinserir & "          " & dia & "/" & mes & "/" & ano & "          " & linha.Replace("in-", "ADICIONOU ").Replace("out-", "SUBTRAIU ") & Environment.NewLine)
                                End If
                                Dim posSeparador = linha.IndexOf(" ")
                            Else
                                RichTextBox1.AppendText(codigoinserir & "          " & dia & "/" & mes & "/" & ano & "          " & linha.Replace("in-", "ADICIONOU ").Replace("out-", "SUBTRAIU ") & Environment.NewLine)
                            End If
                        Next

                    End If

                Next
            Else
                MessageBox.Show("A data inicial não pode ser menor do que a data final!", "Datas erradas", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If




        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ComboBox1.SelectedIndex = 0
        TextBox1.Clear()
        MaskedTextBox1.Clear()
        MaskedTextBox2.Clear()
        RichTextBox1.Clear()
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("http://www.rkd.com.br")
    End Sub
End Class