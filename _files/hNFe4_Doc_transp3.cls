VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hNFe4_Doc_transp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public m_CNPJ As String
Public m_CPF As String
Public m_xNome As String
Public m_IE As String
Public m_xEnder As String
Public m_xMun As String
Public m_UF As String

Public m_ret_vServ As Double
Public m_ret_vBCRet As Double
Public m_ret_pICMSRet As Double
Public m_ret_vICMSRet As Double
Public m_ret_CFOP As String
Public m_ret_cMunFG As String

