Module teclaspermitidasvb
    Function letraspermitidas(ByVal KeyAscii As Integer) As Integer
        'Transformar letras minusculas em Maiúsculas
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        ' Intercepta um código ASCII recebido e admite somente letras
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(KeyAscii)) = 0 Then
            letraspermitidas = 0
        Else
            letraspermitidas = KeyAscii
        End If
        ' teclas adicionais permitidas
        If KeyAscii = 8 Then letraspermitidas = KeyAscii ' Backspace
        If KeyAscii = 127 Then letraspermitidas = KeyAscii ' Delete
        If KeyAscii = 32 Then letraspermitidas = KeyAscii ' ~Espaço
    End Function
    Function numerosspermitidos(ByVal KeyAscii As Integer) As Integer
        'Transformar letras minusculas em Maiúsculas
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        ' Intercepta um código ASCII recebido e admite somente letras
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then
            numerosspermitidos = 0
        Else
            numerosspermitidos = KeyAscii
        End If
        ' teclas adicionais permitidas
        If KeyAscii = 8 Then numerosspermitidos = KeyAscii ' Backspace
        If KeyAscii = 127 Then numerosspermitidos = KeyAscii ' Delete
    End Function
    Function letrasnumpermitidos(ByVal KeyAscii As Integer) As Integer
        'Transformar letras minusculas em Maiúsculas
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        ' Intercepta um código ASCII recebido e admite somente letras
        If InStr("0123456789AÀÁÂÃBCÇDEÈÊÉFGHIÍJKLMNOÓÕÔPQRSTUÚVWXYZ*()", Chr(KeyAscii)) = 0 Then
            letrasnumpermitidos = 0
        Else
            letrasnumpermitidos = KeyAscii
        End If
        ' teclas adicionais permitidas
        If KeyAscii = 8 Then letrasnumpermitidos = KeyAscii ' Backspace
        If KeyAscii = 127 Then letrasnumpermitidos = KeyAscii ' Delete
        If KeyAscii = 32 Then letrasnumpermitidos = KeyAscii ' ~Espaço
    End Function
    Function letrasnumpontospermitidos(ByVal KeyAscii As Integer) As Integer
        'Transformar letras minusculas em Maiúsculas
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        ' Intercepta um código ASCII recebido e admite somente letras
        If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ,/:()", Chr(KeyAscii)) = 0 Then
            letrasnumpontospermitidos = 0
        Else
            letrasnumpontospermitidos = KeyAscii
        End If
        ' teclas adicionais permitidas
        If KeyAscii = 8 Then letrasnumpontospermitidos = KeyAscii ' Backspace
        If KeyAscii = 127 Then letrasnumpontospermitidos = KeyAscii ' Delete
        If KeyAscii = 32 Then letrasnumpontospermitidos = KeyAscii ' ~Espaço
    End Function
End Module
