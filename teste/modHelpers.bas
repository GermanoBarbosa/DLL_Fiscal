Attribute VB_Name = "modHelpers"
Option Explicit

Public Function GetUFCode(sUF As String) As String
    Select Case UCase(sUF)
        Case "AC": GetUFCode = "12"
        Case "AL": GetUFCode = "27"
        Case "AM": GetUFCode = "13"
        Case "AP": GetUFCode = "16"
        Case "BA": GetUFCode = "29"
        Case "CE": GetUFCode = "23"
        Case "DF": GetUFCode = "53"
        Case "ES": GetUFCode = "32"
        Case "GO": GetUFCode = "52"
        Case "MA": GetUFCode = "21"
        Case "MG": GetUFCode = "31"
        Case "MS": GetUFCode = "50"
        Case "MT": GetUFCode = "51"
        Case "PA": GetUFCode = "15"
        Case "PB": GetUFCode = "25"
        Case "PE": GetUFCode = "26"
        Case "PI": GetUFCode = "22"
        Case "PR": GetUFCode = "41"
        Case "RJ": GetUFCode = "33"
        Case "RN": GetUFCode = "24"
        Case "RO": GetUFCode = "11"
        Case "RR": GetUFCode = "14"
        Case "RS": GetUFCode = "43"
        Case "SC": GetUFCode = "42"
        Case "SE": GetUFCode = "28"
        Case "SP": GetUFCode = "35"
        Case "TO": GetUFCode = "17"
        Case Else: GetUFCode = "35"
    End Select
End Function

Public Function GetMunicipioDefault(cUF As String) As String
    Select Case cUF
        Case "35": GetMunicipioDefault = "3550308" ' Sao Paulo/SP
        Case "33": GetMunicipioDefault = "3304557" ' Rio de Janeiro/RJ
        Case "31": GetMunicipioDefault = "3106200" ' Belo Horizonte/MG
        Case "41": GetMunicipioDefault = "4106902" ' Curitiba/PR
        Case "43": GetMunicipioDefault = "4314902" ' Porto Alegre/RS
        Case "29": GetMunicipioDefault = "2927408" ' Salvador/BA
        Case "53": GetMunicipioDefault = "5300108" ' Brasilia/DF
        Case "52": GetMunicipioDefault = "5208707" ' Goiania/GO
        Case "26": GetMunicipioDefault = "2611606" ' Recife/PE
        Case "23": GetMunicipioDefault = "2304400" ' Fortaleza/CE
        Case "13": GetMunicipioDefault = "1302603" ' Manaus/AM
        Case "15": GetMunicipioDefault = "1501402" ' Belem/PA
        Case "50": GetMunicipioDefault = "5002704" ' Campo Grande/MS
        Case "51": GetMunicipioDefault = "5103403" ' Cuiaba/MT
        Case "21": GetMunicipioDefault = "2111300" ' Sao Luis/MA
        Case "42": GetMunicipioDefault = "4205407" ' Florianopolis/SC
        Case "32": GetMunicipioDefault = "3205309" ' Vitoria/ES
        Case "22": GetMunicipioDefault = "2211001" ' Teresina/PI
        Case "27": GetMunicipioDefault = "2704302" ' Maceio/AL
        Case "28": GetMunicipioDefault = "2800308" ' Aracaju/SE
        Case "25": GetMunicipioDefault = "2507507" ' Joao Pessoa/PB
        Case "24": GetMunicipioDefault = "2408102" ' Natal/RN
        Case "17": GetMunicipioDefault = "1721000" ' Palmas/TO
        Case "11": GetMunicipioDefault = "1100205" ' Porto Velho/RO
        Case "14": GetMunicipioDefault = "1400100" ' Boa Vista/RR
        Case "12": GetMunicipioDefault = "1200401" ' Rio Branco/AC
        Case "16": GetMunicipioDefault = "1600303" ' Macapa/AP
        Case Else: GetMunicipioDefault = "3550308"
    End Select
End Function
