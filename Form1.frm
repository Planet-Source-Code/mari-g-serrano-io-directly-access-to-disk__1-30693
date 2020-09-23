VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Editor de Disco"
   ClientHeight    =   4530
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar sB 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   4155
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   661
      SimpleText      =   "Valo"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Sectores:"
            TextSave        =   "Sectores:"
            Key             =   "SECTORES"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Offset:"
            TextSave        =   "Offset:"
            Key             =   "OFFSET"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Valor Hex:"
            TextSave        =   "Valor Hex:"
            Key             =   "HEX"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "ESTADO"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3360
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   480
      ScaleHeight     =   375
      ScaleWidth      =   2535
      TabIndex        =   3
      Top             =   0
      Width           =   2535
      Begin VB.TextBox txtLen 
         Height          =   285
         Left            =   600
         TabIndex        =   6
         Text            =   "0"
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdMas 
         Caption         =   "+"
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton cmdMe 
         Caption         =   "-"
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Cluster"
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.PictureBox pCont 
      Height          =   3735
      Left            =   0
      ScaleHeight     =   3675
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   480
      Width           =   7095
      Begin VB.TextBox txtHex 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   0
         Width           =   2415
      End
      Begin VB.VScrollBar vS 
         Height          =   3570
         Left            =   6840
         TabIndex        =   8
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txtData 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   2400
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   0
         Width           =   4455
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Hex"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin VB.Menu mnu 
      Caption         =   "Abrir"
      Begin VB.Menu mnuAbrir 
         Caption         =   "Abrir Archivo"
         Index           =   0
      End
      Begin VB.Menu mnuAbrir 
         Caption         =   "Abrir Disco..."
         Index           =   1
      End
   End
   Begin VB.Menu mnuCerrar 
      Caption         =   "Cerrar"
   End
   Begin VB.Menu mnuLeer 
      Caption         =   "Leer"
   End
   Begin VB.Menu mnuEscribir 
      Caption         =   "Escribir"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Access to all the bytes of a Disk!!
' MaRiØ G.Serrano.Madrid.11/1/02

Private tBytes() As Byte
Private F As CRandom

Private Sub Form_Resize()
    'movemos las cosas...
    pCont.Move 0, 480, Me.ScaleWidth - 25, Me.ScaleHeight - pCont.Top - sB.Height - 25
    txtHex.Move 0, 0, pCont.Width / 3, pCont.Height
    txtData.Move txtHex.Width + 5, 0, (2 * txtHex.Width) - 300, pCont.Height
    vS.Move txtData.Left + txtData.Width, 0, vS.Width, pCont.Height - 55
End Sub

Private Sub mnuAbrir_Click(Index As Integer)
    sB.Panels("ESTADO").Text = "Abriendo..."
    Select Case Index

    Case 0 'File
        cd.Flags = cdlOFNFileMustExist Or cdlOFNLongNames Or cdlOFNExplorer
        cd.ShowOpen
        If cd.FileName = "" Then Exit Sub
        F.OpenFile cd.FileName
    Case 1 'disk (A,C...)
        Dim Drive As String
        Dim Rl As String
        
        Dim b() As Byte
        Drive = InputBox("Drive to Open...? (A: , C:)", "Put a drive", "A:")
        F.OpenFile "\\.\" & Drive
        '\\\\.\\vwin32'para 95-98-me
        F.ReadBytes 512, b
        
        ' buscamos en el sector de arranque de la BIOS el nº de sectores
        ' que tiene el Disco...
        
        If UCase(Drive) <> "A:" Then 'es diskette?
           Rl = "&h" & Hex(b(136)) & Hex(b(135)) & Hex(b(134))
           vS.Max = 32767
        Else
           Rl = "&h" & Hex(b(20)) & Hex(b(19))
           vS.Max = Val(Rl - 1)
        End If
        sB.Panels("SECTORES").Text = "Sectores: " & Val(Rl)
    
    End Select
sB.Panels("ESTADO").Text = "OK"
End Sub

Private Sub mnuCerrar_Click()
    F.CloseFile
    sB.Panels("ESTADO").Text = "Cerrado OK"
End Sub

Private Sub cmdMas_Click()
    txtLen = txtLen + 1
    mnuLeer_Click
End Sub
Private Sub cmdMe_Click()
    txtLen = txtLen - 1
    mnuLeer_Click
End Sub


Private Sub mnuH2A_Click()
    'pasar txthex a txtdata
End Sub

Private Sub mnuLeer_Click()
  Leer
End Sub
Private Sub Leer()
  Dim Temp As Variant
  Dim Bytes As Long
  sB.Panels("ESTADO").Text = "Leyendo..."
  Bytes = txtLen * 512
  If Bytes = 0 Then Bytes = 512
  F.SeekAbsolute 0, txtLen * 512
  txtData = CStr(F.ReadBytes(10 * 512, tBytes))
  txtHex = toHex(tBytes)
  sB.Panels("ESTADO").Text = "OK"
End Sub
Private Function toHex(b() As Byte) As String
Dim i As Long
Dim tmp As String * 3
For i = LBound(b) To UBound(b)

    If Len(Format(CStr(Hex(b(i))), "00")) = 1 Then
        tmp = "0" & Format(CStr(Hex(b(i))), "00") & " "
    Else
        tmp = Format(CStr(Hex(b(i))), "00") & " "
    End If
    toHex = toHex & tmp
    
Next
End Function
Private Sub mnuEscribir_Click()
If MsgBox("Esto puede dañar tu ordenador! " & vbCrLf & "Desea continuar...?", vbYesNo Or vbExclamation, "Escribir en Disco") = vbNo Then Exit Sub
 
 sB.Panels("ESTADO").Text = "Escribiendo..."
 F.WriteBytes tBytes()
 sB.Panels("ESTADO").Text = "OK"
End Sub

Private Sub Form_Load()
  Set F = New CRandom
  'On Error Resume Next
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set F = Nothing
End Sub


Private Sub txtLen_Change()
    If Val(txtLen) < 0 Then txtLen = 0
    
End Sub



'*********


Private Sub txtData_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyHome, vbKeyEnd
        txtData.SelLength = 0
    Case vbKeyPageUp
       txtData.SelLength = 0
       On Error Resume Next
       vS.Value = vS.Value - 10
       'txtLen = txtLen - 1
    Case vbKeyPageDown
        txtData.SelLength = 0
        vS.Value = vS.Value + 10
       'txtLen = txtLen + 1
    End Select
End Sub

Sub txtData_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then txtData.SelLength = 0
End Sub

Private Sub txtData_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    tBytes(txtData.SelStart - 1) = Asc(Mid(txtData, txtData.SelStart, 1))
   txtData.SelLength = 1
   'sB.Panels("ESTADO").Text = txtData.SelStart
   txtHex = toHex(tBytes)
   txtHex.SelStart = 3 * txtData.SelStart - 2
   txtHex.SelLength = 2
   sB.Panels("OFFSET") = "Offset: " & txtData.SelStart * (txtLen + 1)
   sB.Panels("HEX") = "Hex: " & Hex(tBytes(txtData.SelStart - 1)) & " Dec: " & tBytes(txtData.SelStart - 1)
End Sub
Private Sub txtData_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   If txtData.SelLength = 0 Or txtData.SelLength > 1 Then txtData.SelLength = 1
   sB.Panels("OFFSET") = "Offset: " & txtData.SelStart - 1 * (txtLen + 1)
   sB.Panels("HEX") = "Hex: " & Hex(tBytes(txtData.SelStart - 1)) & " Dec: " & tBytes(txtData.SelStart - 1)
End Sub

Private Sub txtLen_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
    Leer
End Sub

Private Sub vS_Change()
    txtLen = vS.Value
    Leer
End Sub

Private Sub vS_Scroll()
    txtLen = vS.Value
    Leer
End Sub
