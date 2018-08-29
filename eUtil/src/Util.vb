Imports System.ComponentModel
Imports System.Text
Imports System.Text.RegularExpressions

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
End Class
