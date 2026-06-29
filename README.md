# AxFiscal

ActiveX OCX (DLL) para emissão e gerenciamento de NF-e/NFC-e 4.00 em VB6.

## Funcionalidades

- **Geração de XML** de NF-e/NFC-e (modelo 55/65) conforme leiaute 4.00
- **Envio de lote** (síncrono) para SEFAZ
- **Consulta de Status** do serviço SEFAZ
- **Cancelamento** de NF-e e **Cancelamento por Substituição** (evento 110112)
- **Manifestação do Destinatário** (ciência, confirmação, desconhecimento, não realizada)
- **Carta de Correção** (CC-e)
- **Inutilização** de numeração
- **Roteamento automático** por UF (AM, BA, CE, GO, MG, MS, MT, PE, PR, RJ, SP, SVRS) com suporte aos ambientes de Produção e Homologação
- **Contingências** SVAN, SVC-RS, SVC-AN
- **Assinatura digital XML** com certificado A1/A3 via CAPICOM
- **Validação** automática por XSD antes do envio
- **Validação de regras de negócio** NF-e 4.00 (obrigatoriedade, tamanhos, domínios) integrada ao fluxo de envio
- **Consulta Situação NF-e** (consSitNFe) completa pela chave de acesso
- **Distribuição de DF-e** (nfeDistDFe) — consulta de documentos fiscais pela SEFAZ nacional
- **Download de NF-e** (downloadNFe) — download do XML da NF-e pelo mesmo webservice nacional
- **Consulta de Recibo de Lote** (consReciNFe) — consulta do resultado de envio assíncrono

## Dependências

A pasta `dependencias\` contém os instaladores necessários:

| Arquivo | Descrição |
|---------|-----------|
| `capicom.dll` | Microsoft CAPICOM 2.1 — acesso ao repositório de certificados |
| `msxml5.dll` | Microsoft XML Core Services 5.0 — parsing e validação XSD |
| `msxml5r.dll` | Recursos de idioma do MSXML5 |
| `soapsdk.msi` | Microsoft SOAP SDK — suporte a webservices SOAP |
| `instala.bat` | Script que copia e registra as DLLs no sistema |

Execute `instala.bat` como administrador para instalar as dependências.

## Requisitos

- Windows XP/Vista/7/8/10/11
- Visual Basic 6.0 (para compilar)
- Certificado digital A1 ou A3 (eToken) instalado no repositório do Windows

## Compilação

1. Abra `com_teste.vbg` (VB Group) no VB6
2. Compile a DLL `Ax4096CtrlsFiscal.vbp` (ActiveX DLL → `build\AxFiscal-048.dll`)
3. O projeto de teste `teste\Project1.vbp` já referência a OCX

### 0. Distribuição de DF-e (nfeDistDFe)

```vb
Dim xml_send As String, xml_resp As String

' Consulta por NSU
If nfe.NfeDistDFe("11222333000181", "000000000000001", "", xml_send, xml_resp) Then
    ' retNFe_DistDFe contém: cStat, xMotivo, cUFAutor, dhRecbto, ultNSU
    ' Os docZip estão no XML de resposta (base64 gzipados)
    MsgBox nfe.retNFe_DistDFe.cStat & " - " & nfe.retNFe_DistDFe.xMotivo
End If

' Consulta por chave de NF-e
If nfe.NfeDistDFe("11222333000181", "", "43120922591766000167550010000000081832293328", xml_send, xml_resp) Then
    MsgBox nfe.retNFe_DistDFe.cStat & " - " & nfe.retNFe_DistDFe.xMotivo
End If
```

### 0. Consulta de Recibo de Lote (consReciNFe)

```vb
Dim xml_send As String, xml_resp As String

' nRec = recibo retornado pelo envio assincrono (indSinc=0)
If nfe.NfeConsReciNFe("350000000000001", xml_send, xml_resp) Then
    ' cStat=104 = lote processado
    ' retNFe_ConsReciNFe.protNFe contem chNFe, nProt, digVal, cStat, xMotivo
    MsgBox nfe.retNFe_ConsReciNFe.cStat & " - " & nfe.retNFe_ConsReciNFe.xMotivo
End If
```

### 0. Download de NF-e (downloadNFe)

```vb
Dim xml_send As String, xml_resp As String

' Download pelo webservice nacional (chave da NF-e)
If nfe.NfeDownloadNF("11222333000181", "35", "43120922591766000167550010000000081832293328", xml_send, xml_resp) Then
    ' retNFe_DownloadNFe.docZip cont�m o XML da NF-e em base64 + gzip
    ' retNFe_DownloadNFe.chNFe cont�m a chave de acesso
    MsgBox nfe.retNFe_DownloadNFe.cStat & " - " & nfe.retNFe_DownloadNFe.xMotivo
End If
```

### 0. Consultar Situação da NF-e (consSitNFe)

```vb
Dim xml_send As String, xml_resp As String
If nfe.NfeConsultaProtocolo("43120922591766000167550010000000081832293328", xml_send, xml_resp) Then
    MsgBox nfe.retNFe_ConsSitNFe.cStat & " - " & nfe.retNFe_ConsSitNFe.xMotivo
Else
    MsgBox "Falha: " & nfe.m_LastError
End If
```

## Utilização Básica

### 1. Inicialização

```vb
Dim nfe As New hNFe4

nfe.m_Ambiente = 2          ' 1 = Produção, 2 = Homologação
nfe.m_UF = "SP"
nfe.m_UF_Cod = "35"
nfe.Initialize "Caminho\para\Schema-4.00\", "Caminho\do\app\"

nfe.SetCert "SERIAL_DO_CERTIFICADO"
nfe.Inicia ""
```

### 2. Consultar Status do Serviço

```vb
If nfe.NfeStatusServico Then
    MsgBox nfe.retNfeStatusServico.cStat & " - " & nfe.retNfeStatusServico.xMotivo
Else
    MsgBox "Falha: " & nfe.m_LastError
End If
```

### 3. Gerar XML da NF-e

```vb
With nfe.NFe
    .m_Ambiente = nfe.m_Ambiente
    .Clear_det

    ' --- IDE ---
    .m_ide.m_cUF = nfe.m_UF_Cod
    .m_ide.m_mod = "55"
    .m_ide.m_serie = "1"
    .m_ide.m_nNF = "1"
    .m_ide.m_dhEmi = Now
    .m_ide.m_cMunFG = "3550308"     ' código IBGE São Paulo
    .m_ide.m_tpEmis = "1"
    .m_ide.m_idDest = 1

    ' --- EMITENTE ---
    .m_emit.m_CNPJ = "11222333000181"
    .m_emit.m_xNome = "EMPRESA TESTE LTDA"
    .m_emit.m_xFant = "TESTE"
    .m_emit.m_IE = "111111111111"
    .m_emit.m_CRT = "3"             ' 1=Simples, 2=Presumido, 3=Real
    With .m_emit.m_enderEmit
        .m_xLgr = "Rua Teste"
        .m_nro = "100"
        .m_xBairro = "Centro"
        .m_cMun = "3550308"
        .m_xMun = "SAO PAULO"
        .m_UF = "SP"
        .m_CEP = "01001000"
    End With

    ' --- DESTINATÁRIO ---
    .m_dest.m_CNPJ = "11222333000181"
    .m_dest.m_xNome = "CLIENTE TESTE LTDA"
    .m_dest.m_indIEDest = "9"       ' 9=Não contribuinte
    .m_dest.m_IE = ""
    With .m_dest.m_enderDest
        .m_xLgr = "Av Teste"
        .m_nro = "200"
        .m_xBairro = "Centro"
        .m_cMun = "3550308"
        .m_xMun = "SAO PAULO"
        .m_UF = "SP"
        .m_CEP = "01001000"
    End With

    ' --- PRODUTOS ---
    Dim det As New hNFe4_Doc_det
    det.prod.m_cProd = "001"
    det.prod.m_xProd = "PRODUTO TESTE"
    det.prod.m_NCM = "84713019"
    det.prod.m_CFOP = "5102"
    det.prod.m_uCom = "UN"
    det.prod.m_qCom = 1
    det.prod.m_vUnCom = 100
    det.prod.m_vProd = 100
    det.prod.m_uTrib = "UN"
    det.prod.m_qTrib = 1
    det.prod.m_vUnTrib = 100
    det.prod.m_indTot = 1
    .Add_det det

    ' --- FRETE ---
    .m_modFrete = "9"               ' 9=Sem ocorrência

    ' --- INFORMAÇÕES ADICIONAIS ---
    .m_infAdic.infCpl = "NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL"

    ' --- GERA XML ---
    Dim chave As String, numNota As Long
    Dim xml As String
    xml = .GetNota(chave, numNota)
End With

' --- FORMAS DE PAGAMENTO ---
Dim pag As New hNFe4_Doc_detPag
pag.tPag = "03"             ' 03=Cartao Credito
pag.vPag = 100
pag.indPag = "0"            ' 0=Vista
pag.card.tpIntegra = "2"    ' 2=Nao integrado
pag.card.CNPJ = "11222333000181"
pag.card.tBand = "01"       ' 01=Visa
pag.card.cAut = "123456"
nfe.NFe.Add_pag pag

' Adicione quantos detPag precisar
Dim pag2 As New hNFe4_Doc_detPag
pag2.tPag = "15"            ' 15=Boleto
pag2.vPag = 50
pag2.card.tpIntegra = "2"
nfe.NFe.Add_pag pag2
```

### 4. Enviar Lote

```vb
Dim xml_aut As String
xml_aut = ""

If nfe.EnviaLote2(xmlGerado, xml_aut, serialCertificado) Then
    MsgBox "cStat: " & nfe.retNFe_EnviaLote2.protNFe.cStat
    MsgBox "xMotivo: " & nfe.retNFe_EnviaLote2.protNFe.xMotivo

    ' Gerar XML completo com protocolo (nfeProc)
    Dim xmlProcNFe As String
    xmlProcNFe = nfe.GerarProcNFe(xmlGerado, xml_aut)
    If Len(xmlProcNFe) > 0 Then
        ' Salvar o nfeProc em disco
        SetFileBytes m_path & "\NFe\Proc\" & nfe.NFe.m_Chave_NF & "-proc.xml", xmlProcNFe
    End If
Else
    MsgBox "Falha: " & nfe.m_LastError
End If
```

> **Validacao de regras de negocio:** antes do envio, `EnviaLote2` executa automaticamente `ValidaRegrasNegocio()`, que verifica obrigatoriedade, tamanhos e dominios dos campos. Em caso de erro, `m_LastError` contera as descricoes e o envio e cancelado. Para validacao manual:
> ```vb
> Dim v As New cNFeValidator
> If Not v.Validate(nfe.NFe, nfe.m_Ambiente) Then
>     MsgBox "Erros de validacao: " & v.GetAllErrors
>     MsgBox "Avisos: " & v.GetAllWarnings
> End If
> ```
>
> **NFC-e (modelo 65):** o QR Code com `infNFeSupl` e inserido automaticamente no momento do envio. Basta configurar o CSC antes de chamar `EnviaLote2`:
> ```vb
> nfe.m_CSC = "SEU_CSC_AQUI"       ' C�digo de Seguran�a do Contribuinte
> nfe.NFe.m_Ambiente = 1           ' 1=Produ��o, 2=Homologa��o
> nfe.Inicia                       ' J� define a URL de consulta por UF
> nfe.EnviaLote2 xmlGerado, ...    ' QR Code inserido automaticamente
> ```

### 5. Cancelar NF-e

```vb
Dim xml_send As String, xml_resp As String
If nfe.NfeCancela(serialCert, nfe.m_UF_Cod, chaveNF, cnpjEmitente, justificativa, "0", xml_send, xml_resp) Then
    MsgBox "Cancelado. cStat: " & nfe.retNfeStatusServico.cStat
End If
```

### 6. Carta de Correção

```vb
Dim xml_send As String, xml_resp As String
If nfe.NfeCartaCorrecao(serialCert, nfe.m_UF_Cod, chaveNF, cnpjEmitente, textoCorrecao, xml_send, xml_resp) Then
    MsgBox "CC-e enviada. cStat: " & nfe.retNfeCartaCorrecao.cStat
End If
```

### 7. Cancelamento por Substituição (evento 110112)

Substitui uma NF-e por outra, cancelando a original e referenciando a chave da NF-e substituta.

```vb
Dim xml_send As String, xml_resp As String
If nfe.NfeCancelaPorSubstituicao(serialCert, nfe.m_UF_Cod, chaveNF, cnpjEmitente, justificativa, chaveNFSubstituta, xml_send, xml_resp) Then
    MsgBox "Cancelamento por substituicao enviado. cStat: " & nfe.retNfeStatusServico.cStat
End If
```

### 8. Manifestação do Destinatário (eventos 210200/210210/210220/210240)

```vb
Dim xml_send As String, xml_resp As String
' Ciência da Emissão
If nfe.NfeManifestacaoDestinatario(serialCert, nfe.m_UF_Cod, chaveNF, cnpjEmitente, "210210", "", xml_send, xml_resp) Then
    MsgBox "Ciencia enviada. cStat: " & nfe.retNfeStatusServico.cStat
End If

' Confirmação da Operação
nfe.NfeManifestacaoDestinatario serialCert, nfe.m_UF_Cod, chaveNF, cnpjEmitente, "210200", "", xml_send, xml_resp

' Desconhecimento da Operação
nfe.NfeManifestacaoDestinatario serialCert, nfe.m_UF_Cod, chaveNF, cnpjEmitente, "210220", "", xml_send, xml_resp

' Operação não Realizada (requer justificativa)
nfe.NfeManifestacaoDestinatario serialCert, nfe.m_UF_Cod, chaveNF, cnpjEmitente, "210240", justificativa, xml_send, xml_resp
```

### 9. Sistema de Logging

O projeto possui um sistema centralizado de logging baseado em listeners. Todos os `Debug.Print` e `MsgBox` nos módulos da DLL foram substituídos pelo logger.

```vb
' Grava log informativo
GetLogger.LogInfo "Minha mensagem"

' Grava log de aviso (ex: erro não fatal)
GetLogger.LogWarning "Aviso: algo inesperado"

' Grava log de debug (detalhamento técnico)
GetLogger.Debug "Var1=" & var1 & " Var2=" & var2

' Grava log como erro
GetLogger.LogError "Falha na operação"
```

**Listeners disponíveis:**
| Listener | Descrição |
|----------|-----------|
| `Debug.Print` (built-in) | Sempre ativo, envia para a janela Immediato do VB6 |
| `frmMain` | Exibe logs no formulário principal de teste |
| `FileLogger` (futuro) | Planejado para gravação em arquivo |

Para implementar um listener customizado, implemente a interface `ILogListener`:

```vb
Implements ILogListener

Private Sub ILogListener_LogReceived(ByVal level As String, ByVal msg As String)
    ' level: "INFO", "WARN", "DEBUG", "ERROR"
    ' msg: mensagem de log
    ' Exemplo: gravar em arquivo, enviar para servidor, etc.
End Sub
```

Adicione o listener no `Form_Load`:
```vb
GetLogger.AddListener Me
```

### 7. Inutilizar Numeração

```vb
Dim xml_send As String, xml_resp As String
If nfe.NfeInutilizacao(serialCert, nfe.m_UF_Cod, cnpjEmitente, "55", "1", 1, 1, justificativa, xml_send, xml_resp) Then
    MsgBox "Inutilizado. cStat: " & nfe.retNfeInutilizacao.cStat
End If
```

## Estrutura do Projeto

```
Ax4096CtrlsFiscal.vbp     → ActiveX DLL (OCX)
teste\Project1.vbp         → Aplicação de teste (executável)
teste\frmMain.frm          → Formulário principal de teste
teste\frmCert.frm          → Seleção de certificado digital
teste\modHelpers.bas       → Funções auxiliares (códigos UF, municípios)
teste\Schema-4.00\         → Schemas XSD da NF-e 4.00
_files\
  hNFe4.cls                → Classe principal (webservices + roteamento)
  hNFe4_Doc.cls            → Documento NF-e/NFC-e
  hNFe4_Doc_ide.cls        → Identificação da NF-e
  hNFe4_Doc_emit.cls       → Emitente
  hNFe4_Doc_emit_enderEmit.cls → Endereço do emitente
  hNFe4_Doc_dest.cls       → Destinatário
  hNFe4_Doc_dest_enderDest.cls  → Endereço do destinatário
  hNFe4_Doc_det.cls        → Detalhe (produtos)
  hNFe4_Doc_det_prod.cls   → Produto
  hNFe4_Doc_det_imposto.cls → Impostos
  hNFe4_Doc_detPag.cls      → Forma de pagamento
  hNFe4_Doc_detPag_card.cls → Cartão/PIX/Boletos
  hNFe4_Doc_total.cls      → Totais
  hNFe4_Doc_transp3.cls    → Transportadora
  hNFe4_Doc_veicTransp.cls → Veículo de transporte
  hNFe4_Doc_vol.cls        → Volumes
  hNFe4_Doc_infAdic.cls    → Informações adicionais
  hNFe4_StatusServico.cls  → Retorno de status
  hNFe4_EnviaLote2.cls     → Retorno de envio
  hNFe4_EnviaLote2_protNFe.cls → Protocolo de autorização
  hNFe4_CartaCorrecao.cls  → Retorno de CC-e
  hNFe4_ConsReciNFe.cls    → Retorno de consulta de recibo
  hNFe4_ConsSitNFe.cls     → Retorno de consulta situação
  hNFe4_DistDFe.cls        → Retorno de distribuição de DF-e
  hNFe4_DownloadNFe.cls    → Retorno de download de NF-e
  hNFe4_Inutilizacao.cls   → Retorno de inutilização
  cNFeValidator.cls        → Validador de regras de negócio NF-e 4.00 (obrigatoriedade, tamanhos, domínios)
  hNFe_Signature.cls       → Assinatura digital XML
  hSHA1.cls                → Cálculo SHA-1
  fFuncsBase16.cls         → Base16
  hStringBuilder_private.cls → String builder
  ILogListener.cls         → Interface para listeners de log
  cLogger.cls              → Gerenciador central de logging com suporte a múltiplos listeners
  mLogger.bas              → Módulo público de logging (função `Log` e `GetLogger`)
  mFuncs.bas               → Funções utilitárias
dependencias\              → Instaladores CAPICOM, MSXML5, SOAP SDK
```

## Roadmap — Funcionalidades Pendentes

### Serviços SEFAZ não implementados
| Funcionalidade | Descrição | Prioridade |
|---|---|---|
| ~~**consReciNFe**~~ | ~~Consulta de recibo do lote (necessário para envio assíncrono)~~ | ~~Alta~~ |
| ~~**nfeDistDFe**~~ | ~~Distribuição de DF-e (consulta e download)~~ | ~~Alta~~ |
| ~~**downloadNFe**~~ | ~~Download do XML da NF-e~~ | ~~Alta~~ |

### Cancelamento e Eventos
| Funcionalidade | Descrição | Prioridade |
|---|---|---|
| ~~Cancelamento por substituição (evento 110112)~~ | ~~Substituir NF-e por outra com correções~~ | ~~Média~~ |
| ~~Manifestação do destinatário~~ | ~~Confirmação, ciência, desconhecimento, não realizar~~ | ~~Média~~ |
| EPEC | Evento Prévio de Emissão em Contingência | Média |

### Melhorias na Geração de XML
| Funcionalidade | Descrição | Prioridade |
|---|---|---|
| ~~**ProcNFe**~~ | ~~Geração do XML completo com protocolo (`nfeProc`)~~ | ~~Alta~~ |
| ~~**Múltiplas formas de pagamento**~~ | ~~Suporte a mais de um `detPag` no grupo `pag`~~ | ~~Alta~~ |
| ~~Detalhamento de cartão (`card`)~~ | ~~Bandeira, autorização, CNPJ credenciadora~~ | ~~Média~~ |
| ~~Cobrança / Fatura (`cobr`, `dup`)~~ | ~~Atualmente comentado no código~~ | ~~Média~~ |
| Declaração de Importação (`DI` / `ADI`) | Campos para produtos importados | Média |
| Grupo Combustível (`comb`) | Específico para postos de gasolina | Baixa |
| Grupo Rastreamento (`rastro`) | Rastreamento de produtos | Baixa |
| ISSQN | Imposto sobre serviço | Baixa |
| Exportação (`exporta`) | Venda para exterior | Baixa |

### Infraestrutura & Melhorias Gerais
| Funcionalidade | Descrição | Prioridade |
|---|---|---|
| **Envio assíncrono** | Suporte a `indSinc=0` para lotes com múltiplas NF-e | Alta |
| ~~**Validação de regras de negócio**~~ | ~~Além da validação XSD, validar tamanhos, obrigatoriedade e domínios~~ | ~~Média~~ |
| ~~**QR Code NFC-e**~~ | ~~Mapeamento completo das URLs de consulta por UF + CSC dinâmico~~ | ~~Média~~ |
| Envio de lote com múltiplas NF-e | Suporte a mais de 1 NF-e por lote | Média |
| ~~**Logger configurável**~~ | ~~Substituir `Debug.Print` e `MsgBox` por sistema de log com `ILogListener`~~ | ~~Média~~ |

## Licença

Este é um software open source. Consulte o arquivo `LICENSE` para mais informações (se aplicável).
