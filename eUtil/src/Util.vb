Imports System.ComponentModel
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Security.Cryptography

Public Class Util
    Function CompletaZero(ByVal txt_texto As String, ByVal cod_quantidade As Integer) As String
        While Len(txt_texto) < cod_quantidade
            txt_texto = "0" & txt_texto
        End While
        CompletaZero = txt_texto
    End Function

    Function remove_acento(ByVal texto As String) As String
        If texto <> "" Then
            Do While InStr(texto, " ")
                texto = Replace(texto, " ", "_")
            Loop
            texto = Replace(texto, "á", "a")
            texto = Replace(texto, "â", "a")
            texto = Replace(texto, "ã", "a")
            texto = Replace(texto, "à", "a")
            texto = Replace(texto, "ä", "a")
            texto = Replace(texto, "Â", "A")
            texto = Replace(texto, "À", "A")
            texto = Replace(texto, "Á", "A")
            texto = Replace(texto, "Ä", "A")
            texto = Replace(texto, "Ã", "A")
            texto = Replace(texto, "Ê", "E")
            texto = Replace(texto, "È", "E")
            texto = Replace(texto, "É", "E")
            texto = Replace(texto, "Ë", "E")
            texto = Replace(texto, "é", "e")
            texto = Replace(texto, "ê", "e")
            texto = Replace(texto, "è", "e")
            texto = Replace(texto, "ë", "e")
            texto = Replace(texto, "í", "i")
            texto = Replace(texto, "ì", "i")
            texto = Replace(texto, "ï", "i")
            texto = Replace(texto, "î", "i")
            texto = Replace(texto, "Í", "I")
            texto = Replace(texto, "Ì", "I")
            texto = Replace(texto, "Ï", "I")
            texto = Replace(texto, "Î", "I")
            texto = Replace(texto, "ó", "o")
            texto = Replace(texto, "ò", "o")
            texto = Replace(texto, "ô", "o")
            texto = Replace(texto, "õ", "o")
            texto = Replace(texto, "ö", "o")
            texto = Replace(texto, "Ó", "O")
            texto = Replace(texto, "Ô", "O")
            texto = Replace(texto, "Ò", "O")
            texto = Replace(texto, "Õ", "O")
            texto = Replace(texto, "Ö", "O")
            texto = Replace(texto, "ú", "u")
            texto = Replace(texto, "ù", "u")
            texto = Replace(texto, "û", "u")
            texto = Replace(texto, "ü", "u")
            texto = Replace(texto, "Ú", "U")
            texto = Replace(texto, "Ù", "U")
            texto = Replace(texto, "Ü", "U")
            texto = Replace(texto, "Û", "U")
            texto = Replace(texto, "ç", "c")
            texto = Replace(texto, "Ç", "C")
            texto = Replace(texto, " ", "_")
            texto = Replace(texto, "!", "_")
            texto = Replace(texto, "@", "_")
            texto = Replace(texto, "#", "_")
            texto = Replace(texto, "$", "_")
            texto = Replace(texto, "%", "_")
            texto = Replace(texto, "¨", "_")
            texto = Replace(texto, "&", "_")
            texto = Replace(texto, "*", "_")
            texto = Replace(texto, "(", "_")
            texto = Replace(texto, ")", "_")
            texto = Replace(texto, "-", "_")
            texto = Replace(texto, "+", "_")
            texto = Replace(texto, "=", "_")
            texto = Replace(texto, "§", "_")
            texto = Replace(texto, "'", "_")
            texto = Replace(texto, "´", "_")
            texto = Replace(texto, "`", "_")
            texto = Replace(texto, "{", "_")
            texto = Replace(texto, "}", "_")
            texto = Replace(texto, "[", "_")
            texto = Replace(texto, "]", "_")
            texto = Replace(texto, "ª", "_")
            texto = Replace(texto, "º", "_")
            texto = Replace(texto, "°", "_")
            texto = Replace(texto, "|", "_")
            texto = Replace(texto, ",", "_")
            texto = Replace(texto, ":", "_")
            texto = Replace(texto, ";", "_")
            texto = Replace(texto, "^", "_")
            texto = Replace(texto, "~", "_")
            texto = Replace(texto, ",", "_")
            texto = Replace(texto, Chr(166), "_")
            texto = Replace(texto, Chr(167), "_")
            texto = Replace(texto, Chr(248), "_")
            texto = Replace(texto, Chr(176), "_")
            texto = Replace(texto, Chr(186), "_")
        End If
        remove_acento = texto
    End Function

    Function FormatarCpfCnpj(ByVal strCpfCnpj As String) As String
        If (strCpfCnpj.Length <= 11) Then
            Dim mtpCpf As MaskedTextProvider = New MaskedTextProvider("000\.000\.000-00")
            mtpCpf.Set(ZerosEsquerda(strCpfCnpj, 11))
            Return mtpCpf.ToString()
        Else
            Dim mtpCnpj As MaskedTextProvider = New MaskedTextProvider("00\.000\.000/0000-00")
            mtpCnpj.Set(ZerosEsquerda(strCpfCnpj, 11))
            Return mtpCnpj.ToString()
        End If
    End Function

    Function ZerosEsquerda(ByVal strString As String, ByVal intTamanho As Integer) As String
        Dim strResult As String = ""
        Dim intCont As Integer = 1
        While intCont <= (intTamanho - strString.Length)
            strResult += "0"
        End While
        Return strResult + strString
    End Function

    Function primeira_maiscula(ByVal valor As String) As String
        Dim a_valor() As String = Split(LCase(valor), " ")
        For i As Integer = LBound(a_valor) To UBound(a_valor)
            If Len(a_valor(i)) > 2 Then
                a_valor(i) = UCase(Left(a_valor(i), 1)) & Mid(a_valor(i), 2, Len(a_valor(i)))
            End If
        Next
        primeira_maiscula = Join(a_valor, " ")
    End Function

    Function limpar_comparacao(ByVal texto As String) As Object
        If texto <> "" Then
            texto = Trim(texto)
            If IsNumeric(texto) Then
                texto = CLng(texto)
            End If
        End If
        limpar_comparacao = texto
    End Function

    Function monta_matriz(ByVal codigo As String) As String()
        codigo = Mid(codigo, 2, Len(codigo))
        codigo = Mid(codigo, 1, Len(codigo) - 1)
        Dim a_codigo() As String = Split(codigo, "][")
        monta_matriz = a_codigo
    End Function

    Function retornaHashNomeArquivo(ByVal strArquivo As String) As String
        Dim retorno As String = ""

        Dim padraoRegex As String = "{([A-Z0-9-]*)}"
        Dim nomeArquivo As New RegularExpressions.Regex(padraoRegex, RegexOptions.IgnorePatternWhitespace)

        Try
            retorno = nomeArquivo.Match(strArquivo).Captures.Item(0).Value
        Catch ex As Exception
            Dim a_retorno As String() = strArquivo.Split("\")

            retorno = a_retorno(UBound(a_retorno))
        End Try

        Return retorno
    End Function

    Function retorna_cor(ByVal cont As Integer) As String
        Dim txt_cor As String = "#ffffff"
        If cont Mod 2 = 1 Then
            txt_cor = "#c3c3c3"
        End If
        Return txt_cor
    End Function

    Function MD5(ByVal texto As String) As String
        Dim provider As New MD5CryptoServiceProvider
        Dim byHash() As Byte
        Dim hash As String = String.Empty

        byHash = provider.ComputeHash(System.Text.Encoding.UTF8.GetBytes(texto))
        provider.Clear()

        hash = BitConverter.ToString(byHash).Replace("-", String.Empty)

        Return hash
    End Function

    Public Function ConverterListParaDataTable(Of T)(items As List(Of T)) As DataTable
        Dim dataTable As New DataTable(GetType(T).Name)
        'Pega todas as propriedades
        Dim Propriedades As System.Reflection.PropertyInfo() = GetType(T).GetProperties(System.Reflection.BindingFlags.[Public] Or System.Reflection.BindingFlags.Instance)
        For Each _propriedade As System.Reflection.PropertyInfo In Propriedades
            'Define os nomes das colunas como os nomes das propriedades
            dataTable.Columns.Add(_propriedade.Name)
        Next
        For Each item As T In items
            Dim valores = New Object(Propriedades.Length - 1) {}
            For i As Integer = 0 To Propriedades.Length - 1
                'inclui os valores das propriedades nas linhas do datatable
                valores(i) = Propriedades(i).GetValue(item, Nothing)
            Next
            dataTable.Rows.Add(valores)
        Next
        Return dataTable
    End Function

    Public Shared Function RetornarPrimeiroNome(ByVal valor As String) As String
        If Not IsNothing(valor) Then
            If valor.IndexOf(" ") > 0 Then
                Return valor.Substring(0, valor.IndexOf(" "))
            Else
                Return valor
            End If
        Else
            Return ""
        End If
    End Function

    Public Shared Function IsCPFValido(ByVal CPF As String) As Boolean
        'Declaração das Variáveis 
        Dim strCPFOriginal As String = CPF.Replace(".", "").Replace("-", "")
        Dim strCPF As String = Mid(strCPFOriginal, 1, 9)
        Dim strCPFTemp As String
        Dim intSoma As Integer
        Dim intResto As Integer
        Dim strDigito As String
        Dim intMultiplicador As Integer = 10
        Const constIntMultiplicador As Integer = 11
        Dim i As Integer

        For i = 0 To strCPF.ToString.Length - 1
            intSoma += CInt(strCPF.ToString.Chars(i).ToString) * intMultiplicador
            intMultiplicador -= 1
        Next

        If (intSoma Mod constIntMultiplicador) < 2 Then
            intResto = 0
        Else
            intResto = constIntMultiplicador - (intSoma Mod constIntMultiplicador)
        End If

        strDigito = intResto
        intSoma = 0

        strCPFTemp = strCPF & strDigito
        intMultiplicador = 11
        For i = 0 To strCPFTemp.Length - 1
            intSoma += CInt(strCPFTemp.Chars(i).ToString) * intMultiplicador
            intMultiplicador -= 1
        Next

        If (intSoma Mod constIntMultiplicador) < 2 Then
            intResto = 0
        Else
            intResto = constIntMultiplicador - (intSoma Mod constIntMultiplicador)
        End If
        strDigito &= intResto

        If strDigito = Mid(strCPFOriginal, 10, strCPFOriginal.Length) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Function MontaDecimal(ByVal valor As String, Optional ByVal direcao As Boolean = True) As String
        Dim retorno As String = valor

        If direcao Then
            If Not String.IsNullOrWhiteSpace(valor) Then
                retorno = valor.Replace(".", ",")
            End If
        Else
            If Not String.IsNullOrWhiteSpace(valor) Then
                retorno = valor.Replace(",", ".")
            End If
        End If

        Return retorno.ToString.Trim
    End Function

    Public Shared Function IsAlphaNum(ByVal strInputText As String) As Boolean
        Dim IsAlpha As Boolean = False
        If System.Text.RegularExpressions.Regex.IsMatch(strInputText, "^(?=.*[A-Za-z])(?=.*\d)[A-Za-z\d]{10,}$") Then
            IsAlpha = True
        Else
            IsAlpha = False
        End If
        Return IsAlpha
    End Function

    Public Shared Function MontaMatriz(ByVal codigo As String) As String()
        codigo = Mid(codigo, 2, Len(codigo))
        codigo = Mid(codigo, 1, Len(codigo) - 1)
        Dim a_codigo() As String = Split(codigo, "][")

        Return a_codigo
    End Function

    Public Shared Function AddDiasUteis(ByVal data As Date, ByVal dias As Integer) As Date
        Dim newDate As Date = data

        While dias > 0
            newDate = newDate.AddDays(1)
            If newDate.DayOfWeek = DayOfWeek.Saturday Or newDate.DayOfWeek = DayOfWeek.Sunday Then Continue While
            dias -= 1
        End While

        Return newDate
    End Function

    Public Shared Function RemoveDiasUteis(ByVal data As Date, ByVal dias As Integer) As Date
        Dim newDate As Date = data

        While dias > 0
            newDate = newDate.AddDays(-1)
            If newDate.DayOfWeek = DayOfWeek.Saturday Or newDate.DayOfWeek = DayOfWeek.Sunday Then Continue While
            dias -= 1
        End While

        Return newDate
    End Function

    Public Shared Function NomeDoMes(ByVal mes As Integer) As String
        Select Case mes
            Case 1
                Return "janeiro"
            Case 2
                Return "fevereiro"
            Case 3
                Return "março"
            Case 4
                Return "abril"
            Case 5
                Return "maio"
            Case 6
                Return "junho"
            Case 7
                Return "julho"
            Case 8
                Return "agosto"
            Case 9
                Return "setembro"
            Case 10
                Return "outubro"
            Case 11
                Return "novembro"
            Case 12
                Return "dezembro"
            Case Else
                Return ""
        End Select

    End Function

    Public Shared Function RetornaTipoArquivo(ByVal txt_arquivo As String) As String
        Dim txt_tipo As String = "application/octet-stream"
        txt_arquivo = LCase(StrReverse(txt_arquivo))
        txt_arquivo = Left(txt_arquivo, InStr(txt_arquivo, ".") - 1)
        txt_arquivo = StrReverse(txt_arquivo)

        Select Case txt_arquivo
            Case "pdf"
                txt_tipo = "application/pdf"
            Case "gif"
                txt_tipo = "image/gif"
            Case "jpg", "jpeg"
                txt_tipo = "image/jpeg"
            Case "png"
                txt_tipo = "image/png"
            Case "bmp"
                txt_tipo = "image/bmp"
            Case "docx", "doc"
                txt_tipo = "application/ms-word"
            Case "xlsx", "xls"
                txt_tipo = "application/vnd.xls"
            Case "ppt", "pptx"
                txt_tipo = "application/vnd.ms-powerpoint"
            Case "txt"
                txt_tipo = "text/plain"
            Case "mp3"
                txt_tipo = "audio/mpeg"
            Case Else
                txt_tipo = "application/octet-stream"

        End Select

        Return txt_tipo

    End Function

    Public Shared Function ConverterParaRbg(ByVal HexColor As String) As System.Drawing.Color
        Dim Red, Green, Blue As String

        HexColor = Replace(HexColor, "#", "")
        Red = Val("&H" & Mid(HexColor, 1, 2))
        Green = Val("&H" & Mid(HexColor, 3, 2))
        Blue = Val("&H" & Mid(HexColor, 5, 2))

        Return System.Drawing.Color.FromArgb(Red, Green, Blue)

    End Function

    Public Shared Function ValidaDataUtil(data As Date) As Integer
        Dim retorno = 0

        If Not String.IsNullOrEmpty(data) Then
            If Weekday(data) = 1 Or Weekday(data) = 7 Then
                retorno = -1
            End If

            If data < Date.Now.Date Then
                retorno = -2
            End If

        End If

        Return retorno

    End Function

    Public Shared Function ValidaCNPJ(ByVal CNPJ As String) As Boolean
        CNPJ = limpar_cnpj_cpf(CNPJ)

        Dim i As Integer
        Dim dadosArray() As String = {"111.111.111-11", "222.222.222-22", "333.333.333-33", "444.444.444-44", "555.555.555-55", "666.666.666-66", "777.777.777-77", "888.888.888-88", "999.999.999-99"}

        CNPJ = CNPJ.Trim

        For i = 0 To dadosArray.Length - 1
            If CNPJ.Length <> 14 OrElse dadosArray(i).Equals(CNPJ) Then
                Return False
            End If

        Next

        Dim Numero(13) As Integer
        Dim soma As Integer
        Dim resultado1 As Integer
        Dim resultado2 As Integer

        For i = 0 To Numero.Length - 1
            Numero(i) = CInt(CNPJ.Substring(i, 1))
        Next

        soma = Numero(0) * 5 + Numero(1) * 4 + Numero(2) * 3 + Numero(3) * 2 + Numero(4) * 9 + Numero(5) * 8 + Numero(6) * 7 + Numero(7) * 6 + Numero(8) * 5 + Numero(9) * 4 + Numero(10) * 3 + Numero(11) * 2
        soma = soma - (11 * (Int(soma / 11)))

        If soma = 0 OrElse soma = 1 Then
            resultado1 = 0
        Else
            resultado1 = 11 - soma
        End If

        If resultado1 = Numero(12) Then
            soma = Numero(0) * 6 + Numero(1) * 5 + Numero(2) * 4 + Numero(3) * 3 + Numero(4) * 2 + Numero(5) * 9 + Numero(6) * 8 + Numero(7) * 7 + Numero(8) * 6 + Numero(9) * 5 + Numero(10) * 4 + Numero(11) * 3 + Numero(12) * 2
            soma = soma - (11 * (Int(soma / 11)))

            If soma = 0 OrElse soma = 1 Then
                resultado2 = 0
            Else
                resultado2 = 11 - soma
            End If

            If resultado2.Equals(Numero(13)) Then
                Return True
            Else
                Return False
            End If

        Else
            Return False

        End If

    End Function

    Public Shared Function limpar_cnpj_cpf(ByVal txt_cnpj_cpf As String) As String
        Return Replace(Replace(Replace(Trim(txt_cnpj_cpf), "-", ""), "/", ""), ".", "")

    End Function

    Public Shared Function EnviarEmail(ByVal txt_remetente As String, ByVal txt_destinatario As String, ByVal txt_assunto As String, ByVal txt_mensagem As String, ByVal txt_servidor_smtp_usuario As String,
                                       ByVal txt_servidor_smtp As String, ByVal cod_porta As Integer, ByVal cod_sll As Boolean, ByVal txt_servidor_smtp_senha As String) As Boolean

        Dim retorno As String = String.Empty

        'Cria objeto com dados do e-mail.
        Dim objEmail As New System.Net.Mail.MailMessage()

        txt_servidor_smtp_usuario = txt_servidor_smtp_usuario.Trim
        txt_servidor_smtp = txt_servidor_smtp.Trim
        cod_porta = cod_porta
        cod_sll = cod_sll
        txt_servidor_smtp_senha = txt_servidor_smtp_senha

        objEmail.From = New System.Net.Mail.MailAddress(txt_servidor_smtp_usuario, txt_remetente.Trim, System.Text.Encoding.UTF8)

        'Define os destinatários do e-mail.
        objEmail.To.Add(txt_destinatario.Trim)

        'Define a prioridade do e-mail.
        objEmail.Priority = System.Net.Mail.MailPriority.Normal

        'Define o formato do e-mail HTML (caso não queira HTML alocar valor false)
        objEmail.IsBodyHtml = True

        'Define o título do e-mail.
        objEmail.Subject = txt_assunto.Trim

        'Define o corpo do e-mail.        
        objEmail.Body = txt_mensagem.Trim

        'Para evitar problemas com caracteres "estranhos", configuramos o Charset para "ISO-8859-1"
        objEmail.SubjectEncoding = System.Text.Encoding.GetEncoding("ISO-8859-1")
        objEmail.BodyEncoding = System.Text.Encoding.GetEncoding("ISO-8859-1")

        'Cria objeto com os dados do SMTP
        Dim objSmtp As New System.Net.Mail.SmtpClient

        'Alocamos o endereço do host para enviar os e-mails:
        objSmtp.Host = txt_servidor_smtp
        objSmtp.UseDefaultCredentials = False
        objSmtp.Credentials = New System.Net.NetworkCredential(txt_servidor_smtp_usuario, txt_servidor_smtp_senha)
        objSmtp.Port = cod_porta
        objSmtp.EnableSsl = cod_sll

        Try
            objSmtp.Send(objEmail)
            retorno = "SUCESSO"
        Catch ex As Exception
            retorno = ex.Source & "<br />" & ex.Message & "<br />" & ex.StackTrace
        Finally
            objEmail.Dispose()
        End Try

        Return retorno
    End Function
End Class
