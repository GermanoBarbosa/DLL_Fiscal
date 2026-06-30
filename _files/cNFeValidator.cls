VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNFeValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Errors As New Collection
Public Warnings As New Collection

Public Function Validate(Doc As hNFe4_Doc, Ambiente As Long) As Boolean
    Set Errors = Nothing
    Set Warnings = Nothing
    Set Errors = New Collection
    Set Warnings = New Collection

    Validate_IDE Doc.m_ide, Ambiente
    Validate_Emit Doc.m_emit, Doc.m_ide.m_mod
    Validate_Dest Doc.m_dest
    Validate_Itens Doc
    Validate_Total Doc.m_total
    Validate_Pagamentos Doc
    Validate_Transp Doc

    Validate = (Errors.Count = 0)
End Function

Public Function GetAllErrors() As String
    Dim sb As String
    Dim i As Long
    For i = 1 To Errors.Count
        If Len(sb) > 0 Then sb = sb & "; "
        sb = sb & Errors(i)
    Next
    GetAllErrors = sb
End Function

Public Function GetAllWarnings() As String
    Dim sb As String
    Dim i As Long
    For i = 1 To Warnings.Count
        If Len(sb) > 0 Then sb = sb & "; "
        sb = sb & Warnings(i)
    Next
    GetAllWarnings = sb
End Function

Private Sub AddError(ByVal Campo As String, ByVal Descricao As String)
    Errors.Add Campo & ": " & Descricao
End Sub

Private Sub AddWarning(ByVal Campo As String, ByVal Descricao As String)
    Warnings.Add Campo & ": " & Descricao
End Sub

Private Sub Validate_IDE(ide As hNFe4_Doc_ide, Ambiente As Long)
    If Len(Trim(ide.m_cUF)) = 0 Then
        AddError "ide.cUF", "cUF obrigatorio"
    ElseIf Len(ide.m_cUF) <> 2 Or Not IsNumeric(ide.m_cUF) Then
        AddError "ide.cUF", "cUF deve ter 2 digitos numericos"
    End If

    If Len(Trim(ide.m_cNF)) = 0 Then
        AddError "ide.cNF", "cNF obrigatorio"
    ElseIf Len(ide.m_cNF) <> 8 Or Not IsNumeric(ide.m_cNF) Then
        AddError "ide.cNF", "cNF deve ter 8 digitos numericos"
    End If

    If ide.m_mod <> "55" And ide.m_mod <> "65" Then
        AddError "ide.mod", "mod deve ser 55 (NF-e) ou 65 (NFC-e)"
    End If

    If Len(Trim(ide.m_serie)) = 0 Then
        AddError "ide.serie", "serie obrigatoria"
    ElseIf Not IsNumeric(ide.m_serie) Then
        AddError "ide.serie", "serie deve ser numerica"
    ElseIf Val(ide.m_serie) < 1 Or Val(ide.m_serie) > 999 Then
        AddError "ide.serie", "serie deve estar entre 1 e 999"
    End If

    If Len(Trim(ide.m_nNF)) = 0 Then
        AddError "ide.nNF", "nNF obrigatorio"
    ElseIf Not IsNumeric(ide.m_nNF) Then
        AddError "ide.nNF", "nNF deve ser numerico"
    ElseIf Val(ide.m_nNF) < 1 Or Val(ide.m_nNF) > 999999999 Then
        AddError "ide.nNF", "nNF deve estar entre 1 e 999999999"
    End If

    If ide.m_dhEmi < #1/1/2020# Then
        AddError "ide.dhEmi", "Data de emissao invalida ou nao informada"
    End If

    If Len(Trim(ide.m_cMunFG)) = 0 Then
        AddError "ide.cMunFG", "cMunFG obrigatorio"
    ElseIf Len(ide.m_cMunFG) <> 7 Or Not IsNumeric(ide.m_cMunFG) Then
        AddError "ide.cMunFG", "cMunFG deve ter 7 digitos numericos"
    End If

    If Len(Trim(ide.m_tpEmis)) = 0 Then
        AddError "ide.tpEmis", "tpEmis obrigatorio"
    ElseIf ide.m_tpEmis < "1" Or ide.m_tpEmis > "9" Then
        AddError "ide.tpEmis", "tpEmis deve estar entre 1 e 9"
    End If

    If ide.m_idDest < 1 Or ide.m_idDest > 3 Then
        AddError "ide.idDest", "idDest deve ser 1, 2 ou 3"
    End If
End Sub

Private Sub Validate_Emit(emit As hNFe4_Doc_emit, Modulo As String)
    If Len(Trim(emit.m_CNPJ)) = 0 Then
        AddError "emit.CNPJ", "CNPJ do emitente obrigatorio"
    Else
        Dim cnpj As String
        cnpj = OnlyAlphanumeric(emit.m_CNPJ)
        If Len(cnpj) <> 14 Then
            AddError "emit.CNPJ", "CNPJ do emitente deve ter 14 caracteres (alfanumerico, NT 2026.004)"
        End If
    End If

    If Len(Trim(emit.m_xNome)) = 0 Then
        AddError "emit.xNome", "Razao social obrigatoria"
    End If

    If Len(Trim(emit.m_IE)) = 0 Then
        AddError "emit.IE", "IE do emitente obrigatoria"
    End If

    If Len(Trim(emit.m_CRT)) = 0 Then
        AddError "emit.CRT", "CRT obrigatorio"
    ElseIf emit.m_CRT < "1" Or emit.m_CRT > "3" Then
        AddError "emit.CRT", "CRT deve ser 1 (SN), 2 (SN c/ ICMS) ou 3 (Regime Normal)"
    End If

    With emit.m_enderEmit
        If Len(Trim(.m_xLgr)) = 0 Then AddError "emit.enderEmit.xLgr", "Logradouro obrigatorio"
        If Len(Trim(.m_nro)) = 0 Then AddWarning "emit.enderEmit.nro", "Numero nao informado"
        If Len(Trim(.m_xBairro)) = 0 Then AddError "emit.enderEmit.xBairro", "Bairro obrigatorio"
        If Len(Trim(.m_cMun)) = 0 Then AddError "emit.enderEmit.cMun", "cMun obrigatorio"
        If Len(Trim(.m_xMun)) = 0 Then AddError "emit.enderEmit.xMun", "Municipio obrigatorio"
        If Len(Trim(.m_UF)) = 0 Then AddError "emit.enderEmit.UF", "UF obrigatorio"
        If Len(Trim(.m_CEP)) = 0 Then AddWarning "emit.enderEmit.CEP", "CEP nao informado"
    End With
End Sub

Private Sub Validate_Dest(dest As hNFe4_Doc_dest)
    If Len(Trim(dest.m_CNPJ)) = 0 And Len(Trim(dest.m_CPF)) = 0 Then
        AddError "dest.CNPJ/CPF", "CNPJ ou CPF do destinatario obrigatorio"
    Else
        If Len(Trim(dest.m_CNPJ)) > 0 Then
            If Len(OnlyAlphanumeric(dest.m_CNPJ)) <> 14 Then
                AddError "dest.CNPJ", "CNPJ deve ter 14 caracteres (alfanumerico, NT 2026.004)"
            End If
        End If
        If Len(Trim(dest.m_CPF)) > 0 Then
            If Len(OnlyNumbers(dest.m_CPF)) <> 11 Then
                AddError "dest.CPF", "CPF deve ter 11 digitos"
            End If
        End If
    End If

    If Len(Trim(dest.m_xNome)) = 0 Then
        AddWarning "dest.xNome", "Nome do destinatario nao informado"
    End If

    If Len(Trim(dest.m_indIEDest)) = 0 Then
        AddError "dest.indIEDest", "indIEDest obrigatorio"
    ElseIf dest.m_indIEDest < "1" Or dest.m_indIEDest > "9" Then
        AddError "dest.indIEDest", "indIEDest deve ser 1, 2 ou 9"
    End If

    If dest.m_indIEDest = "1" And Len(Trim(dest.m_IE)) = 0 Then
        AddError "dest.IE", "IE obrigatoria para contribuinte (indIEDest=1)"
    End If
End Sub

Private Sub Validate_Itens(Doc As hNFe4_Doc)
    If Doc.ItemCount = 0 Then
        AddError "det", "NF-e deve ter pelo menos 1 item"
        Exit Sub
    End If

    Dim i As Long
    For i = 0 To Doc.ItemCount - 1
        Dim prod As hNFe4_Doc_det_prod
        Set prod = Doc.Item(i).prod

        Dim idx As String
        idx = "det[" & (i + 1) & "].prod."

        If Len(Trim(prod.m_cProd)) = 0 Then AddError idx & "cProd", "cProd obrigatorio"
        If Len(Trim(prod.m_xProd)) = 0 Then AddError idx & "xProd", "xProd obrigatorio"
        If Len(Trim(prod.m_NCM)) = 0 Then AddError idx & "NCM", "NCM obrigatorio"
        If Len(prod.m_NCM) <> 8 Then AddError idx & "NCM", "NCM deve ter 8 digitos"
        If Len(Trim(prod.m_CFOP)) = 0 Then AddError idx & "CFOP", "CFOP obrigatorio"
        If Len(prod.m_CFOP) <> 4 Then AddError idx & "CFOP", "CFOP deve ter 4 digitos"
        If Len(Trim(prod.m_uCom)) = 0 Then AddError idx & "uCom", "uCom obrigatorio"
        If prod.m_qCom <= 0 Then AddError idx & "qCom", "qCom deve ser maior que zero"
        If prod.m_vUnCom <= 0 Then AddError idx & "vUnCom", "vUnCom deve ser maior que zero"
        If prod.m_vProd <= 0 Then AddError idx & "vProd", "vProd deve ser maior que zero"
        If prod.m_indTot <> 1 Then AddError idx & "indTot", "indTot deve ser 1"
    Next
End Sub

Private Sub Validate_Total(total As hNFe4_Doc_total)
    If total.m_vNF <= 0 Then
        AddError "total.ICMSTot.vNF", "vNF deve ser maior que zero"
    End If
End Sub

Private Sub Validate_Pagamentos(Doc As hNFe4_Doc)
    If Doc.PagCount = 0 Then
        AddWarning "detPag", "Nenhuma forma de pagamento informada"
        Exit Sub
    End If

    Dim i As Long
    For i = 0 To Doc.PagCount - 1
        Dim pag As hNFe4_Doc_detPag
        Set pag = Doc.Pag(i)

        Dim idx As String
        idx = "detPag[" & (i + 1) & "]."

        If Len(Trim(pag.tPag)) = 0 Then AddError idx & "tPag", "tPag obrigatorio"
        If pag.vPag <= 0 Then AddError idx & "vPag", "vPag deve ser maior que zero"
    Next
End Sub

Private Sub Validate_Transp(Doc As hNFe4_Doc)
    If Len(Trim(Doc.m_modFrete)) = 0 Then
        AddError "transp.modFrete", "modFrete obrigatorio"
    ElseIf Doc.m_modFrete < "0" Or Doc.m_modFrete > "9" Then
        AddError "transp.modFrete", "modFrete deve estar entre 0 e 9"
    End If
End Sub