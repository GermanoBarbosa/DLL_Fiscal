VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hNFe4_Doc_local"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public m_CNPJ As String
Public m_CPF As String
Public m_xNome As String
Public m_xLgr As String
Public m_nro As String
Public m_xCpl As String
Public m_xBairro As String
Public m_cMun As String
Public m_xMun As String
Public m_UF As String
Public m_CEP As String
Public m_cPais As String
Public m_xPais As String
Public m_fone As String
Public m_email As String
Public m_IE As String

Sub Clear()
m_CNPJ = ""
m_CPF = ""
m_xNome = ""
m_xLgr = ""
m_nro = ""
m_xCpl = ""
m_xBairro = ""
m_cMun = ""
m_xMun = ""
m_UF = ""
m_CEP = ""
m_cPais = ""
m_xPais = ""
m_fone = ""
m_email = ""
m_IE = ""
End Sub
