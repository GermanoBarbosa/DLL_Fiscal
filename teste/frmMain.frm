VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "hNFe4 Test Harness v1.0"
   ClientHeight    =   12480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10800
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   12480
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraResultado 
      Caption         =   "Resultado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   0
      TabIndex        =   54
      Top             =   7400
      Width           =   10750
      Begin VB.TextBox txtResult 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   80
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   55
         Top             =   240
         Width           =   10590
      End
   End
   Begin VB.Frame fraAcoes 
      Caption         =   "Acoes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   0
      TabIndex        =   44
      Top             =   6800
      Width           =   10750
      Begin VB.CommandButton cmdLimparResult 
         Caption         =   "Limpar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9600
         TabIndex        =   53
         Top             =   200
         Width           =   1000
      End
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "Consultar NF-e"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8280
         TabIndex        =   52
         Top             =   200
         Width           =   1200
      End
      Begin VB.CommandButton cmdInutilizar 
         Caption         =   "Inutilizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7120
         TabIndex        =   51
         Top             =   200
         Width           =   1000
      End
      Begin VB.CommandButton cmdCCe 
         Caption         =   "CC-e"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6400
         TabIndex        =   50
         Top             =   200
         Width           =   600
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5480
         TabIndex        =   49
         Top             =   200
         Width           =   800
      End
      Begin VB.CommandButton cmdEnviar 
         Caption         =   "Enviar Lote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4280
         TabIndex        =   48
         Top             =   200
         Width           =   1100
      End
      Begin VB.CommandButton cmdGerarXML 
         Caption         =   "Gerar XML"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         TabIndex        =   47
         Top             =   200
         Width           =   1000
      End
      Begin VB.CommandButton cmdStatus 
         Caption         =   "Status Servico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   46
         Top             =   200
         Width           =   1400
      End
      Begin VB.CommandButton cmdPreencher 
         Caption         =   "Dados Teste"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   45
         Top             =   200
         Width           =   1200
      End
   End
   Begin VB.Frame fraInfAdic 
      Caption         =   "Informacoes Adicionais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1100
      Left            =   5400
      TabIndex        =   40
      Top             =   5650
      Width           =   5350
      Begin VB.TextBox txtInfAdFisco 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   800
         MultiLine       =   -1  'True
         TabIndex        =   43
         Top             =   310
         Width           =   4470
      End
      Begin VB.TextBox txtInfCpl 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   800
         MultiLine       =   -1  'True
         TabIndex        =   42
         Top             =   310
         Width           =   4470
      End
      Begin VB.Label lblInfAdFisco 
         Caption         =   "Fisco:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   310
         Width           =   600
      End
      Begin VB.Label lblInfCpl 
         Caption         =   "Cpl:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   310
         Width           =   600
      End
   End
   Begin VB.Frame fraFrete 
      Caption         =   "Frete / Transportadora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1100
      Left            =   0
      TabIndex        =   34
      Top             =   5650
      Width           =   5350
      Begin VB.TextBox txtFreteNome 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3200
         TabIndex        =   39
         Top             =   680
         Width           =   2050
      End
      Begin VB.TextBox txtFreteCNPJ 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   800
         TabIndex        =   38
         Top             =   680
         Width           =   2200
      End
      Begin VB.ComboBox cboModFrete 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   800
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   240
         Width           =   4470
      End
      Begin VB.Label lblFreteNome 
         Caption         =   "xNome:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2600
         TabIndex        =   36
         Top             =   680
         Width           =   600
      End
      Begin VB.Label lblFreteCNPJ 
         Caption         =   "CNPJ:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   680
         Width           =   600
      End
   End
   Begin VB.Frame fraProdutos 
      Caption         =   "Produtos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2200
      Left            =   0
      TabIndex        =   18
      Top             =   3400
      Width           =   10750
      Begin VB.ListBox lstProdutos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   80
         TabIndex        =   29
         Top             =   1200
         Width           =   10590
      End
      Begin VB.CommandButton cmdLimparProd 
         Caption         =   "Limpar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2000
         TabIndex        =   28
         Top             =   840
         Width           =   800
      End
      Begin VB.CommandButton cmdRemoveProd 
         Caption         =   "Remover"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   27
         Top             =   840
         Width           =   800
      End
      Begin VB.CommandButton cmdAddProd 
         Caption         =   "Adicionar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   900
      End
      Begin VB.TextBox txtProdVProd 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9650
         TabIndex        =   25
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox txtProdVUnCom 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8085
         TabIndex        =   24
         Top             =   480
         Width           =   945
      End
      Begin VB.TextBox txtProdQCom 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6585
         TabIndex        =   23
         Top             =   480
         Width           =   780
      End
      Begin VB.TextBox txtProdCFOP 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5100
         TabIndex        =   22
         Top             =   480
         Width           =   885
      End
      Begin VB.TextBox txtProdNCM 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3255
         TabIndex        =   21
         Top             =   480
         Width           =   1200
      End
      Begin VB.TextBox txtProdXProd 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   20
         Top             =   480
         Width           =   2100
      End
      Begin VB.TextBox txtProdCProd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   19
         Top             =   480
         Width           =   900
      End
      Begin VB.Label lblProdVProd 
         Caption         =   "vProd:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9120
         TabIndex        =   33
         Top             =   480
         Width           =   600
      End
      Begin VB.Label lblProdVUnCom 
         Caption         =   "vUnCom:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7380
         TabIndex        =   32
         Top             =   480
         Width           =   705
      End
      Begin VB.Label lblProdQCom 
         Caption         =   "qCom:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6045
         TabIndex        =   31
         Top             =   480
         Width           =   600
      End
      Begin VB.Label lblProdCFOP 
         Caption         =   "CFOP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4500
         TabIndex        =   30
         Top             =   480
         Width           =   600
      End
      Begin VB.Label lblProdNCM 
         Caption         =   "NCM:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2775
         TabIndex        =   9
         Top             =   480
         Width           =   600
      End
      Begin VB.Label lblProdXProd 
         Caption         =   "xProd:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Top             =   480
         Width           =   600
      End
      Begin VB.Label lblProdCProd 
         Caption         =   "cProd:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   480
         Width           =   600
      End
      Begin VB.Label lblProd 
         Caption         =   "Adicione produtos usando os campos abaixo e o botao Adicionar."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   75
         TabIndex        =   1
         Top             =   240
         Width           =   10440
      End
   End
   Begin VB.Frame fraDestinatario 
      Caption         =   "Destinatario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1900
      Left            =   5400
      TabIndex        =   10
      Top             =   1460
      Width           =   5350
      Begin VB.TextBox txtDestCMun 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4200
         TabIndex        =   13
         Top             =   1320
         Width           =   1000
      End
      Begin VB.TextBox txtDestBairro 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2000
         TabIndex        =   14
         Top             =   1320
         Width           =   1400
      End
      Begin VB.TextBox txtDestNro 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4245
         TabIndex        =   15
         Top             =   960
         Width           =   500
      End
      Begin VB.TextBox txtDestXLgr 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   800
         TabIndex        =   16
         Top             =   960
         Width           =   2400
      End
      Begin VB.ComboBox cboIndIEDest 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2800
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   600
         Width           =   2400
      End
      Begin VB.TextBox txtDestIE 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   800
         TabIndex        =   41
         Top             =   600
         Width           =   1200
      End
      Begin VB.TextBox txtDestNome 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         TabIndex        =   56
         Top             =   240
         Width           =   1600
      End
      Begin VB.TextBox txtDestCNPJ 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   800
         TabIndex        =   92
         Top             =   240
         Width           =   2000
      End
      Begin VB.Label lblDestCMun 
         Caption         =   "cMun:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   57
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label lblDestBairro 
         Caption         =   "xBairro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   58
         Top             =   1320
         Width           =   700
      End
      Begin VB.Label lblDestNro 
         Caption         =   "nro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   59
         Top             =   960
         Width           =   405
      End
      Begin VB.Label lblDestXLgr 
         Caption         =   "xLgr:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   960
         Width           =   600
      End
      Begin VB.Label lblIndIEDest 
         Caption         =   "indIEDest:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2200
         TabIndex        =   61
         Top             =   600
         Width           =   600
      End
      Begin VB.Label lblDestIE 
         Caption         =   "IE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   600
         Width           =   600
      End
      Begin VB.Label lblDestNome 
         Caption         =   "xNome:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   93
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblDestCNPJ 
         Caption         =   "CNPJ:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame fraEmitente 
      Caption         =   "Emitente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1900
      Left            =   0
      TabIndex        =   2
      Top             =   1460
      Width           =   5350
      Begin VB.ComboBox cboEmitCRT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   800
         Style           =   2  'Dropdown List
         TabIndex        =   90
         Top             =   960
         Width           =   1400
      End
      Begin VB.TextBox txtEmitCMun 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4200
         TabIndex        =   89
         Top             =   1320
         Width           =   1000
      End
      Begin VB.TextBox txtEmitBairro 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2000
         TabIndex        =   87
         Top             =   1320
         Width           =   1400
      End
      Begin VB.TextBox txtEmitNro 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4725
         TabIndex        =   85
         Top             =   960
         Width           =   500
      End
      Begin VB.TextBox txtEmitXLgr 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2655
         TabIndex        =   83
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtEmitIE 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         TabIndex        =   81
         Top             =   600
         Width           =   1600
      End
      Begin VB.TextBox txtEmitFant 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         TabIndex        =   79
         Top             =   600
         Width           =   1600
      End
      Begin VB.TextBox txtEmitNome 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   800
         TabIndex        =   6
         Top             =   600
         Width           =   2000
      End
      Begin VB.TextBox txtEmitCNPJ 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   800
         TabIndex        =   3
         Top             =   240
         Width           =   2000
      End
      Begin VB.Label lblEmitCMun 
         Caption         =   "cMun:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   88
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label lblEmitBairro 
         Caption         =   "xBairro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   86
         Top             =   1320
         Width           =   700
      End
      Begin VB.Label lblEmitNro 
         Caption         =   "nro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4365
         TabIndex        =   84
         Top             =   960
         Width           =   405
      End
      Begin VB.Label lblEmitXLgr 
         Caption         =   "xLgr:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2265
         TabIndex        =   82
         Top             =   960
         Width           =   405
      End
      Begin VB.Label lblEmitIE 
         Caption         =   "IE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3200
         TabIndex        =   80
         Top             =   600
         Width           =   400
      End
      Begin VB.Label lblEmitFant 
         Caption         =   "xFant:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   78
         Top             =   600
         Width           =   600
      End
      Begin VB.Label lblEmitCRT 
         Caption         =   "CRT:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   960
         Width           =   600
      End
      Begin VB.Label lblEmitNome 
         Caption         =   "xNome:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   600
      End
      Begin VB.Label lblEmitCNPJ 
         Caption         =   "CNPJ:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame fraConfig 
      Caption         =   "Configuracao"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1400
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10750
      Begin VB.TextBox txtNumNF 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8775
         TabIndex        =   76
         Top             =   600
         Width           =   900
      End
      Begin VB.TextBox txtSerie 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         TabIndex        =   74
         Top             =   600
         Width           =   500
      End
      Begin VB.TextBox txtSchema 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         TabIndex        =   73
         Top             =   960
         Width           =   8800
      End
      Begin VB.CommandButton cmdSelCert 
         Caption         =   "Certificado..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6000
         TabIndex        =   72
         Top             =   240
         Width           =   1400
      End
      Begin VB.ComboBox cboModelo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5505
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   600
         Width           =   700
      End
      Begin VB.ComboBox cboUF 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3435
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   600
         Width           =   800
      End
      Begin VB.ComboBox cboAmbiente 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   600
         Width           =   1500
      End
      Begin VB.TextBox txtCert 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   240
         Width           =   5000
      End
      Begin VB.Label lblNumNF 
         Caption         =   "nNF:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8220
         TabIndex        =   75
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblSerie 
         Caption         =   "Serie:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6765
         TabIndex        =   67
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblModelo 
         Caption         =   "Modelo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   66
         Top             =   600
         Width           =   600
      End
      Begin VB.Label lblUF 
         Caption         =   "UF:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2790
         TabIndex        =   65
         Top             =   600
         Width           =   405
      End
      Begin VB.Label lblAmbiente 
         Caption         =   "Ambiente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   600
         Width           =   800
      End
      Begin VB.Label lblCert 
         Caption         =   "Certificado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   800
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ILogListener

Private ccEx As hNFe4
Private mCertSerial As String
Private mXMLGerado As String

Private Type TProduto
    cProd As String
    xProd As String
    NCM As String
    CFOP As String
    qCom As Double
    vUnCom As Double
    vProd As Double
End Type

Private mProds(1 To 200) As TProduto
Private mProdCount As Long

Private Sub Form_Load()
    Set ccEx = New hNFe4
    GetLogger.AddListener Me
    
    cboAmbiente.AddItem "1 - Producao"
    cboAmbiente.AddItem "2 - Homologacao"
    cboAmbiente.ListIndex = 1
    
    Dim ufs: ufs = Split("AC AL AM AP BA CE DF ES GO MA MG MS MT PA PB PE PI PR RJ RN RO RR RS SC SE SP TO")
    Dim i As Long
    For i = 0 To UBound(ufs)
        cboUF.AddItem ufs(i)
    Next
    cboUF.ListIndex = 25
    
    cboModelo.AddItem "55"
    cboModelo.AddItem "65"
    cboModelo.ListIndex = 0
    
    cboEmitCRT.AddItem "1 - Simples Nacional"
    cboEmitCRT.AddItem "2 - Presumido"
    cboEmitCRT.AddItem "3 - Real"
    cboEmitCRT.ListIndex = 2
    
    cboModFrete.AddItem "0 - Contratacao"
    cboModFrete.AddItem "1 - Proprio"
    cboModFrete.AddItem "2 - Proprio/Contratacao"
    cboModFrete.AddItem "9 - Sem ocorrencia"
    cboModFrete.ListIndex = 3
    
    cboIndIEDest.AddItem "1 - Contribuinte"
    cboIndIEDest.AddItem "2 - Isento"
    cboIndIEDest.AddItem "9 - Nao contribuinte"
    cboIndIEDest.ListIndex = 2
    
    txtSchema = App.Path & "\Schema-4.00\"
    txtSerie = "1"
    txtNumNF = "1"
    txtEmitCMun = GetMunicipioDefault("35")
    txtDestCMun = GetMunicipioDefault("35")
    
    mProdCount = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    GetLogger.RemoveListener Me
    Set ccEx = Nothing
End Sub

Private Sub cmdSelCert_Click()
    frmCert.Show vbModal
    If frmCert.mCertificadoSel <> "" Then
        mCertSerial = frmCert.mCertificadoSel
        txtCert = frmCert.mCertificadoNome & " (" & mCertSerial & ")"
    End If
    Unload frmCert
End Sub

Private Sub cmdPreencher_Click()
    txtEmitCNPJ = "11222333000181"
    txtEmitNome = "EMPRESA TESTE LTDA"
    txtEmitFant = "TESTE"
    txtEmitIE = "111111111111"
    txtEmitXLgr = "Rua Teste"
    txtEmitNro = "100"
    txtEmitBairro = "Centro"
    txtEmitCMun = GetMunicipioDefault(GetUFCode(cboUF))
    
    txtDestCNPJ = "11222333000181"
    txtDestNome = "CLIENTE TESTE LTDA"
    txtDestIE = ""
    txtDestXLgr = "Av Teste"
    txtDestNro = "200"
    txtDestBairro = "Centro"
    txtDestCMun = txtEmitCMun
    
    cboModFrete.ListIndex = 3
    txtInfCpl = "NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL"
    
    cmdLimparProd_Click
    
    AddProduto "001", "PRODUTO TESTE 01", "84713019", "5102", 1, 100
    AddProduto "002", "PRODUTO TESTE 02", "84713019", "5102", 2, 50
    
    txtResult = "Dados de teste preenchidos com sucesso."
End Sub

Private Function AddProduto(cProd As String, xProd As String, NCM As String, CFOP As String, qCom As Double, vUnCom As Double)
    mProdCount = mProdCount + 1
    With mProds(mProdCount)
        .cProd = cProd
        .xProd = xProd
        .NCM = NCM
        .CFOP = CFOP
        .qCom = qCom
        .vUnCom = vUnCom
        .vProd = qCom * vUnCom
    End With
    lstProdutos.AddItem cProd & " - " & xProd & " | " & qCom & " x " & Format(vUnCom, "0.00") & " = " & Format(mProds(mProdCount).vProd, "0.00")
End Function

Private Sub cmdAddProd_Click()
    If txtProdCProd = "" Or txtProdXProd = "" Then
        MsgBox "Preencha cProd e xProd.", vbExclamation
        Exit Sub
    End If
    If mProdCount >= 200 Then
        MsgBox "Maximo de 200 produtos.", vbExclamation
        Exit Sub
    End If
    Dim q As Double, v As Double
    On Error Resume Next
    q = CDbl(txtProdQCom)
    v = CDbl(txtProdVUnCom)
    On Error GoTo 0
    AddProduto txtProdCProd, txtProdXProd, txtProdNCM, txtProdCFOP, q, v
    txtProdCProd = "": txtProdXProd = "": txtProdNCM = "": txtProdCFOP = ""
    txtProdQCom = "": txtProdVUnCom = "": txtProdVProd = ""
End Sub

Private Sub cmdRemoveProd_Click()
    If lstProdutos.ListIndex < 0 Or mProdCount = 0 Then Exit Sub
    Dim idx As Long
    idx = lstProdutos.ListIndex + 1
    If idx < mProdCount Then
        Dim i As Long
        For i = idx To mProdCount - 1
            mProds(i) = mProds(i + 1)
        Next
    End If
    mProdCount = mProdCount - 1
    lstProdutos.RemoveItem lstProdutos.ListIndex
End Sub

Private Sub cmdLimparProd_Click()
    mProdCount = 0
    lstProdutos.Clear
End Sub

Private Sub cmdStatus_Click()
    If Not ConfigOK() Then Exit Sub
    Log "Iniciando Status Servico..."
    
    ccEx.m_Ambiente = Val(Left(cboAmbiente, 1))
    ccEx.m_UF = cboUF
    ccEx.m_UF_Cod = GetUFCode(cboUF)
    ccEx.Initialize txtSchema, App.Path & "\"
    
    Dim n_amb As String
    n_amb = ""
    
    ccEx.SetCert mCertSerial
    If ccEx.Inicia(n_amb) Then
        If ccEx.NfeStatusServico Then
            Log "Status: " & ccEx.retNfeStatusServico.cStat & " - " & ccEx.retNfeStatusServico.xMotivo
        Else
            Log "Falha: " & ccEx.m_LastError
        End If
    Else
        Log "Erro ao inicializar: " & ccEx.m_LastError
    End If
End Sub

Private Sub cmdGerarXML_Click()
    If Not ConfigOK() Then Exit Sub
    Log "Gerando XML da NF-e..."
    
    ccEx.m_Ambiente = Val(Left(cboAmbiente, 1))
    ccEx.m_UF = cboUF
    ccEx.m_UF_Cod = GetUFCode(cboUF)
    ccEx.Initialize txtSchema, App.Path & "\"
    
    Dim n_amb As String
    n_amb = ""
    
    With ccEx.NFe
        .m_Ambiente = ccEx.m_Ambiente
        .Clear_det
        .Clear_reboques
        .Clear_volumes
        
        .m_ide.m_cUF = ccEx.m_UF_Cod
        .m_ide.m_mod = cboModelo
        .m_ide.m_serie = txtSerie
        .m_ide.m_nNF = txtNumNF
        .m_ide.m_dhEmi = Now
        .m_ide.m_cMunFG = txtEmitCMun
        .m_ide.m_tpEmis = "1"
        .m_ide.m_idDest = 1
        
        .m_emit.m_CNPJ = txtEmitCNPJ
        .m_emit.m_xNome = txtEmitNome
        .m_emit.m_xFant = txtEmitFant
        .m_emit.m_IE = txtEmitIE
        .m_emit.m_CRT = Left(cboEmitCRT, 1)
        With .m_emit.m_enderEmit
            .m_xLgr = txtEmitXLgr
            .m_nro = txtEmitNro
            .m_xBairro = txtEmitBairro
            .m_cMun = txtEmitCMun
            .m_xMun = txtEmitNome
            .m_UF = cboUF
            .m_CEP = "01001000"
        End With
        
        .m_dest.m_CNPJ = txtDestCNPJ
        .m_dest.m_xNome = txtDestNome
        .m_dest.m_indIEDest = Left(cboIndIEDest, 1)
        .m_dest.m_IE = txtDestIE
        With .m_dest.m_enderDest
            .m_xLgr = txtDestXLgr
            .m_nro = txtDestNro
            .m_xBairro = txtDestBairro
            .m_cMun = txtDestCMun
            .m_xMun = txtDestNome
            .m_UF = cboUF
            .m_CEP = "01001000"
        End With
        
        Dim i As Long
        For i = 1 To mProdCount
            Dim det As New hNFe4_Doc_det
            det.prod.m_cProd = mProds(i).cProd
            det.prod.m_xProd = mProds(i).xProd
            det.prod.m_NCM = mProds(i).NCM
            det.prod.m_CFOP = mProds(i).CFOP
            det.prod.m_uCom = "UN"
            det.prod.m_qCom = mProds(i).qCom
            det.prod.m_vUnCom = mProds(i).vUnCom
            det.prod.m_vProd = mProds(i).vProd
            det.prod.m_uTrib = "UN"
            det.prod.m_qTrib = mProds(i).qCom
            det.prod.m_vUnTrib = mProds(i).vUnCom
            det.prod.m_indTot = 1
            .Add_det det
        Next
        
        .m_modFrete = Left(cboModFrete, 1)
        
        .m_transp.m_CNPJ = txtFreteCNPJ
        .m_transp.m_xNome = txtFreteNome
        
        .m_infAdic.infCpl = txtInfCpl
        .m_infAdic.m_infAdFisco = txtInfAdFisco
        
        Dim chave As String, numNota As Long
        mXMLGerado = .GetNota(chave, numNota)
        Log "XML gerado (chave: " & chave & ")"
        txtResult = mXMLGerado
    End With
End Sub

Private Sub cmdEnviar_Click()
    If mXMLGerado = "" Then
        MsgBox "Gere o XML primeiro.", vbExclamation
        Exit Sub
    End If
    If Not ConfigOK() Then Exit Sub
    Log "Enviando lote..."
    
    ccEx.m_Ambiente = Val(Left(cboAmbiente, 1))
    ccEx.m_UF = cboUF
    ccEx.m_UF_Cod = GetUFCode(cboUF)
    ccEx.Initialize txtSchema, App.Path & "\"
    
    Dim n_amb As String
    n_amb = ""
    
    ccEx.SetCert mCertSerial
    If ccEx.Inicia(n_amb) Then
        Dim xml_aut As String
        xml_aut = ""
        If ccEx.EnviaLote2(mXMLGerado, xml_aut, mCertSerial) Then
            Log "Enviado com sucesso!"
            Log "cStat: " & ccEx.retNFe_EnviaLote2.protNFe.cStat
            Log "xMotivo: " & ccEx.retNFe_EnviaLote2.protNFe.xMotivo
            txtResult = txtResult & vbCrLf & vbCrLf & "=== XML RETORNO ===" & vbCrLf & xml_aut
        Else
            Log "Falha no envio: " & ccEx.m_LastError
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    If Not ConfigOK() Then Exit Sub
    
    Dim chave As String, just As String
    chave = InputBox("Chave da NF-e (44 digitos):", "Cancelamento")
    If Len(chave) <> 44 Then Exit Sub
    just = InputBox("Justificativa (min 15 caracteres):", "Cancelamento")
    If Len(just) < 15 Then Exit Sub
    
    Log "Cancelando NF-e " & chave & "..."
    
    ccEx.m_Ambiente = Val(Left(cboAmbiente, 1))
    ccEx.m_UF = cboUF
    ccEx.m_UF_Cod = GetUFCode(cboUF)
    ccEx.Initialize txtSchema, App.Path & "\"
    
    Dim n_amb As String
    n_amb = ""
    
    ccEx.SetCert mCertSerial
    If ccEx.Inicia(n_amb) Then
        Dim xml_send As String, xml_resp As String
        If ccEx.NfeCancela(mCertSerial, ccEx.m_UF_Cod, chave, txtEmitCNPJ, just, "0", xml_send, xml_resp) Then
            Log "Cancelamento enviado. cStat: " & ccEx.retNfeStatusServico.cStat
            txtResult = xml_resp
        Else
            Log "Falha no cancelamento: " & ccEx.m_LastError
        End If
    End If
End Sub

Private Sub cmdCCe_Click()
    If Not ConfigOK() Then Exit Sub
    
    Dim chave As String, correcao As String
    chave = InputBox("Chave da NF-e (44 digitos):", "Carta de Correcao")
    If Len(chave) <> 44 Then Exit Sub
    correcao = InputBox("Texto de correcao (max 1000 caracteres):", "CC-e")
    If correcao = "" Then Exit Sub
    
    Log "Enviando CC-e para NF-e " & chave & "..."
    
    ccEx.m_Ambiente = Val(Left(cboAmbiente, 1))
    ccEx.m_UF = cboUF
    ccEx.m_UF_Cod = GetUFCode(cboUF)
    ccEx.Initialize txtSchema, App.Path & "\"
    
    Dim n_amb As String
    n_amb = ""
    
    ccEx.SetCert mCertSerial
    If ccEx.Inicia(n_amb) Then
        Dim xml_send As String, xml_resp As String
        If ccEx.NfeCartaCorrecao(mCertSerial, ccEx.m_UF_Cod, chave, txtEmitCNPJ, correcao, xml_send, xml_resp) Then
            Log "CC-e enviado. cStat: " & ccEx.retNfeCartaCorrecao.cStat
            txtResult = xml_resp
        Else
            Log "Falha no CC-e: " & ccEx.m_LastError
        End If
    End If
End Sub

Private Sub cmdInutilizar_Click()
    If Not ConfigOK() Then Exit Sub
    
    Dim just As String
    just = InputBox("Justificativa (min 15 caracteres):", "Inutilizacao")
    If Len(just) < 15 Then Exit Sub
    
    Log "Inutilizando numeracao..."
    
    ccEx.m_Ambiente = Val(Left(cboAmbiente, 1))
    ccEx.m_UF = cboUF
    ccEx.m_UF_Cod = GetUFCode(cboUF)
    ccEx.Initialize txtSchema, App.Path & "\"
    
    Dim n_amb As String
    n_amb = ""
    
    ccEx.SetCert mCertSerial
    If ccEx.Inicia(n_amb) Then
        Dim xml_send As String, xml_resp As String
        If ccEx.NfeInutilizacao(mCertSerial, ccEx.m_UF_Cod, txtEmitCNPJ, cboModelo, txtSerie, txtNumNF, txtNumNF, just, xml_send, xml_resp) Then
            Log "Inutilizacao enviada. cStat: " & ccEx.retNfeInutilizacao.cStat
            txtResult = xml_resp
        Else
            Log "Falha na inutilizacao: " & ccEx.m_LastError
        End If
    End If
End Sub

Private Sub cmdConsultar_Click()
    If Not ConfigOK() Then Exit Sub
    
    Dim chave As String
    chave = InputBox("Chave da NF-e (44 digitos):", "Consulta Situacao")
    If Len(chave) <> 44 Then Exit Sub
    
    Log "Consultando NF-e " & chave & "..."
    
    ccEx.m_Ambiente = Val(Left(cboAmbiente, 1))
    ccEx.m_UF = cboUF
    ccEx.m_UF_Cod = GetUFCode(cboUF)
    ccEx.Initialize txtSchema, App.Path & "\"
    
    Dim n_amb As String
    n_amb = ""
    
    ccEx.SetCert mCertSerial
    If ccEx.Inicia(n_amb) Then
        Dim xml_send As String, xml_resp As String
        If ccEx.NfeConsultaProtocolo(chave, xml_send, xml_resp) Then
            Log "cStat: " & ccEx.retNFe_ConsSitNFe.cStat & " - " & ccEx.retNFe_ConsSitNFe.xMotivo
        Else
            Log "Falha: " & ccEx.m_LastError
        End If
        txtResult = txtResult & vbCrLf & vbCrLf & "=== XML RESPOSTA ===" & vbCrLf & xml_resp
    Else
        Log "Erro ao inicializar: " & ccEx.m_LastError
    End If
End Sub

Private Sub cmdLimparResult_Click()
    txtResult = ""
End Sub

Private Function ConfigOK() As Boolean
    If mCertSerial = "" Then
        MsgBox "Selecione um certificado digital.", vbExclamation
        Exit Function
    End If
    If Dir(txtSchema & "enviNFe_v4.00.xsd") = "" Then
        MsgBox "Diretorio de schemas invalido: " & txtSchema, vbExclamation
        Exit Function
    End If
    If txtEmitCNPJ = "" Then
        MsgBox "Preencha o CNPJ do emitente.", vbExclamation
        Exit Function
    End If
    ConfigOK = True
End Function

Private Sub Log(Msg As String)
    txtResult = txtResult & vbCrLf & "[" & Now & "] " & Msg
End Sub

Private Sub ILogListener_LogReceived(ByVal Level As Long, ByVal Msg As String)
    Dim prefix As String
    Select Case Level
        Case 0: prefix = "[DEBUG]"
        Case 1: prefix = "[INFO]"
        Case 2: prefix = "[WARN]"
        Case 3: prefix = "[ERROR]"
        Case 4: prefix = "[FATAL]"
    End Select
    txtResult = txtResult & vbCrLf & prefix & " " & Msg
End Sub
