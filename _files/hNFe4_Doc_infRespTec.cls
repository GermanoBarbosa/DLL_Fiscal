VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hNFe4_Doc_infRespTec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public m_CNPJ As String
Public m_xContato As String
Public m_email As String
Public m_fone As String
Public m_idCSRT As String
Public m_hashCSRT As String

Sub Clear()
m_CNPJ = ""
m_xContato = ""
m_email = ""
m_fone = ""
m_idCSRT = ""
m_hashCSRT = ""
End Sub
