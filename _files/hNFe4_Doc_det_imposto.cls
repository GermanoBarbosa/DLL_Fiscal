VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hNFe4_Doc_det_imposto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public m_vTotTrib As Double     'Valor aproximado total de tributos federais, estaduais e municipais.

' === CAMPOS LEGADO (mantidos para compatibilidade) ===
Public m_ICMS_ICMS00_orig
Public m_ICMS_ICMS00_CST  As String
Public m_ICMS_ICMS00_modBC
Public m_ICMS_ICMS00_vBC    As Double
Public m_ICMS_ICMS00_pICMS As Double
Public m_ICMS_ICMS00_vICMS As Double

Public m_pST As Double
Public m_vICMSSTRet As Double
Public m_ICMS_ICMS500_vBCST As Double

Public m_IPI_cEnq
Public m_IPI_IPINT_CST  As String
Public m_PIS_PISAliq_CST  As String
Public m_PIS_PISAliq_vBC As Double
Public m_PIS_PISAliq_pPIS As Double
Public m_PIS_PISAliq_vPIS As Double

Public m_COFINS_COFINSAliq_CST As String
Public m_COFINS_COFINSAliq_vBC As Double
Public m_COFINS_COFINSAliq_pCOFINS As Double
Public m_COFINS_COFINSAliq_vCOFINS As Double

Public m_ICMSUFDest_vBCUFDest
Public m_ICMSUFDest_pFCPUFDest
Public m_ICMSUFDest_pICMSUFDest
Public m_ICMSUFDest_pICMSInter
Public m_ICMSUFDest_pICMSInterPart
Public m_ICMSUFDest_vFCPUFDest
Public m_ICMSUFDest_vICMSUFDest
Public m_ICMSUFDest_vICMSUFRemet

' === FCP (Fundo de Combate a Pobreza) item-level ===
Public m_vBCFCP As Double
Public m_pFCP As Double
Public m_vFCP As Double
Public m_vBCFCPST As Double
Public m_pFCPST As Double
Public m_vFCPST As Double
Public m_vBCFCPSTRet As Double
Public m_pFCPSTRet As Double
Public m_vFCPSTRet As Double

' === NOVOS CAMPOS GENERICOS (ICMS, CSOSN, PIS, COFINS, IPI) ===
' ICMS ST (CST 10, 30, 70, 90)
Public m_ICMS_modBCST As Double
Public m_ICMS_pST As Double
Public m_ICMS_vBCST As Double
Public m_ICMS_vICMSST As Double

' ICMS reducao BC (CST 20, 51)
Public m_ICMS_pRedBC As Double

' ICMS desoneracao (CST 40)
Public m_ICMS_vICMSDeson As Double
Public m_ICMS_motDesICMS As String

' ICMS ST retido (CST 60, 500)
Public m_ICMS_vBCSTRet As Double
Public m_ICMS_pSTRet As Double
Public m_ICMS_vICMSSTRet As Double

' CSOSN Simples Nacional
Public m_CSOSN_orig
Public m_CSOSN As String
Public m_CSOSN_pCredSN As Double
Public m_CSOSN_vCredICMS As Double

' IPI tributado
Public m_IPI_CST  As String
Public m_IPI_vBC As Double
Public m_IPI_pIPI As Double
Public m_IPI_vIPI As Double

' PIS generico
Public m_PIS_CST  As String
Public m_PIS_vBC As Double
Public m_PIS_pPIS As Double
Public m_PIS_vPIS As Double
Public m_PIS_qBCProd As Double
Public m_PIS_vAliqProd As Double

' COFINS generico
Public m_COFINS_CST As String
Public m_COFINS_vBC As Double
Public m_COFINS_pCOFINS As Double
Public m_COFINS_vCOFINS As Double
Public m_COFINS_qBCProd As Double
Public m_COFINS_vAliqProd As Double

' === ICMS - campos genericos (todos CSTs) ===
'  Aplicacao preenche estes campos; Add_det os usa no lugar dos
'  campos legado m_ICMS_ICMS00_vBC / m_ICMS_ICMS00_vICMS quando <> 0.
Public m_ICMS_vBC As Double
Public m_ICMS_vICMS As Double

' === IBS/CBS - Grupo IBSCBS (NT 2025.002, type TTribNFe) ===
Public m_IBSCBS_CST As String               ' CST do IBS/CBS (obrigatorio)
Public m_IBSCBS_cClassTrib As String        ' Codigo de Classificacao Tributaria (1-4)
Public m_IBSCBS_indDoacao As String         ' 0/1 - indica doacao
' Sub-grupo gIBSCBS (type TCIBS) - tributacao normal (alternativo a gIBSCBSMono)
Public m_IBSCBS_vBC As Double               ' BC do IBS/CBS
Public m_IBSCBS_pIBSUF As Double            ' aliquota IBS UF (estadual)
Public m_IBSCBS_vIBSUF As Double            ' valor IBS UF
Public m_IBSCBS_pIBSMun As Double           ' aliquota IBS municipal
Public m_IBSCBS_vIBSMun As Double           ' valor IBS municipal
Public m_IBSCBS_vIBS As Double              ' valor total IBS (= vIBSUF + vIBSMun)
Public m_IBSCBS_pCBS As Double              ' aliquota CBS
Public m_IBSCBS_vCBS As Double              ' valor CBS
Public m_IBSCBS_vCBSST As Double            ' valor CBS ST (opcional)

' === ISSQN ===
Public m_ISSQN_vBC As Double
Public m_ISSQN_vAliq As Double
Public m_ISSQN_vISSQN As Double
Public m_ISSQN_cMunFG As String
Public m_ISSQN_cListServ As String
Public m_ISSQN_vDeducao As Double
Public m_ISSQN_vOutro As Double
Public m_ISSQN_vDescIncond As Double
Public m_ISSQN_vDescCond As Double
Public m_ISSQN_vISSRet As Double
Public m_ISSQN_indISS As Long
Public m_ISSQN_cServico As String
Public m_ISSQN_cMun As String
Public m_ISSQN_cPais As String
Public m_ISSQN_nProcesso As String
Public m_ISSQN_indIncentivo As Long

' === IS - Imposto Seletivo (NT 2025.002, type TIS) ===
Public m_IS_CST As String                   ' CST do IS (obrigatorio)
Public m_IS_cClassTrib As String            ' Codigo de Classificacao Tributaria do IS
Public m_IS_vBC As Double                   ' BC do IS
Public m_IS_pIS As Double                   ' aliquota % do IS
Public m_IS_pISEspec As Double              ' aliquota por valor (especifica)
Public m_IS_uTrib As String                 ' unidade de medida
Public m_IS_qTrib As Double                 ' quantidade tributada
Public m_IS_vIS As Double                   ' valor do IS

