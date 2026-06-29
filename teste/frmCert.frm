VERSION 5.00
Begin VB.Form frmCert 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selecionar Certificado"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdSelecionar 
      Caption         =   "Selecionar Certificado..."
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Selecione um certificado digital A1/A3 para os testes de NF-e."
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmCert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mCertificadoSel As String
Public mCertificadoNome As String

Private Sub cmdSelecionar_Click()
    Dim objStore As New Store
    Dim objCerts As ICertificates
    Dim objCert As Object
    
    On Error GoTo ErrHandler
    
    objStore.Open CAPICOM_CURRENT_USER_STORE, "My", CAPICOM_STORE_OPEN_EXISTING_ONLY
    Set objCerts = objStore.Certificates
    
    If objCerts.Count > 0 Then
        Set objCerts = objCerts.Find(CAPICOM_CERTIFICATE_FIND_EXTENDED_PROPERTY, CAPICOM_PROPID_KEY_PROV_INFO)
    End If
    If objCerts.Count > 0 Then
        Set objCerts = objCerts.Find(CAPICOM_CERTIFICATE_FIND_TIME_VALID)
    End If
    
    If objCerts.Count = 0 Then
        MsgBox "Nenhum certificado digital valido com chave privada foi encontrado." & vbCrLf & _
               "Instale um certificado A1 ou conecte um token A3.", vbExclamation, "Certificado"
        Exit Sub
    End If
    
    Set objCert = objCerts.Select("Selecione o certificado digital para testes:", "Certificado A1/A3")
    
    If Not objCert Is Nothing Then
        mCertificadoSel = objCert.Item(1).SerialNumber
        mCertificadoNome = objCert.Item(1).SubjectName
    End If
    
    Unload Me
    Exit Sub
    
ErrHandler:
    MsgBox "Erro ao acessar o repositorio de certificados:" & vbCrLf & Err.Description, vbCritical, "Erro"
    Unload Me
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub

