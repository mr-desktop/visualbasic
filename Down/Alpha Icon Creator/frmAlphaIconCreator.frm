VERSION 5.00
Begin VB.Form frmAlphaIconCreator 
   BackColor       =   &H80000016&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "vbAccelerator Alpha Icon Creator"
   ClientHeight    =   5130
   ClientLeft      =   4980
   ClientTop       =   3825
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAlphaIconCreator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTab 
      BackColor       =   &H80000005&
      Height          =   4635
      Index           =   1
      Left            =   600
      ScaleHeight     =   4575
      ScaleWidth      =   5355
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CommandButton cmdRemoveCustom 
         BackColor       =   &H80000005&
         Caption         =   "&Remove..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   4140
         TabIndex        =   27
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddCustom 
         BackColor       =   &H80000005&
         Caption         =   "&Add..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   26
         Top             =   2640
         Width           =   1095
      End
      Begin VB.ListBox lstCustom 
         Height          =   1230
         Left            =   1500
         TabIndex        =   25
         Top             =   1380
         Width           =   3735
      End
      Begin VB.CheckBox chkSize 
         BackColor       =   &H80000005&
         Caption         =   "&Custom:"
         Height          =   195
         Index           =   4
         Left            =   1260
         TabIndex        =   24
         Top             =   1080
         Width           =   3915
      End
      Begin VB.CheckBox chkSize 
         BackColor       =   &H80000005&
         Caption         =   "24 x 24"
         Height          =   195
         Index           =   1
         Left            =   1260
         TabIndex        =   23
         Top             =   360
         Width           =   3915
      End
      Begin VB.CommandButton cmdPickOutput 
         BackColor       =   &H80000005&
         Caption         =   "&Pick..."
         Height          =   375
         Left            =   1260
         TabIndex        =   11
         Top             =   3540
         Width           =   1095
      End
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   1260
         TabIndex        =   10
         Top             =   3180
         Width           =   3975
      End
      Begin VB.CheckBox chkSize 
         BackColor       =   &H80000005&
         Caption         =   "48 x 48"
         Height          =   195
         Index           =   3
         Left            =   1260
         TabIndex        =   8
         Top             =   840
         Value           =   1  'Checked
         Width           =   3915
      End
      Begin VB.CheckBox chkSize 
         BackColor       =   &H80000005&
         Caption         =   "32 x 32"
         Height          =   195
         Index           =   2
         Left            =   1260
         TabIndex        =   7
         Top             =   600
         Value           =   1  'Checked
         Width           =   3915
      End
      Begin VB.CheckBox chkSize 
         BackColor       =   &H80000005&
         Caption         =   "16 x 16"
         Height          =   195
         Index           =   0
         Left            =   1260
         TabIndex        =   6
         Top             =   120
         Value           =   1  'Checked
         Width           =   3915
      End
      Begin VB.Label lblOutputFile 
         BackStyle       =   0  'Transparent
         Caption         =   "Icon &File Name:"
         Height          =   195
         Left            =   0
         TabIndex        =   9
         Top             =   3240
         Width           =   1155
      End
      Begin VB.Label lblSizes 
         BackStyle       =   0  'Transparent
         Caption         =   "Si&zes to Create:"
         Height          =   195
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox picTab 
      BackColor       =   &H80000005&
      Height          =   4575
      Index           =   0
      Left            =   1260
      ScaleHeight     =   4515
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   60
      Width           =   5415
      Begin VB.CommandButton cmdPickColour 
         BackColor       =   &H80000005&
         Caption         =   "Pic&k..."
         Height          =   375
         Left            =   1320
         TabIndex        =   15
         Top             =   3720
         Width           =   1095
      End
      Begin VB.PictureBox picTransparentColour 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1320
         ScaleHeight     =   285
         ScaleWidth      =   3945
         TabIndex        =   14
         Top             =   3360
         Width           =   3975
      End
      Begin VB.CommandButton cmdPickInput 
         BackColor       =   &H80000005&
         Caption         =   "&Pick..."
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtSourceImage 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   60
         Width           =   3975
      End
      Begin VB.PictureBox picSource 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00F0F0F0&
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   1320
         ScaleHeight     =   2265
         ScaleWidth      =   2265
         TabIndex        =   2
         Top             =   780
         Width           =   2295
      End
      Begin VB.Label lblTransparentColour 
         BackStyle       =   0  'Transparent
         Caption         =   "&Transparent Colour:"
         Height          =   435
         Left            =   60
         TabIndex        =   13
         Top             =   3360
         Width           =   1155
      End
      Begin VB.Label lblSourceImage 
         BackStyle       =   0  'Transparent
         Caption         =   "&Source Image:"
         Height          =   195
         Left            =   60
         TabIndex        =   1
         Top             =   120
         Width           =   1155
      End
   End
   Begin VB.PictureBox picLogo 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   60
      Picture         =   "frmAlphaIconCreator.frx":45A2
      ScaleHeight     =   305
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   19
      Top             =   60
      Width           =   1155
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H80000016&
      Caption         =   "<< &Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   18
      Top             =   4680
      Width           =   1275
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H80000016&
      Caption         =   "&Next >>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   17
      Top             =   4680
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000016&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5340
      TabIndex        =   16
      Top             =   4680
      Width           =   1275
   End
   Begin VB.PictureBox picTab 
      BackColor       =   &H80000005&
      Height          =   4635
      Index           =   2
      Left            =   60
      ScaleHeight     =   4575
      ScaleWidth      =   5355
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CommandButton cmdStartOver 
         BackColor       =   &H80000005&
         Caption         =   "&Start Again"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Your icon has been written out to "
         Height          =   675
         Left            =   120
         TabIndex        =   21
         Top             =   180
         Width           =   5115
      End
   End
End
Attribute VB_Name = "frmAlphaIconCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1
Private Declare Function GetPixelAPI Lib "gdi32" Alias "GetPixel" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private m_cSource As cAlphaDIBSection
Private m_lTransparentColour As OLE_COLOR
Private m_pic As StdPicture
Private m_iTab As Long

Private Sub createIconAtSize( _
      cFI As cFileIcon, _
      ByVal lIndex As Long, _
      ByVal lWidth As Long, _
      ByVal lHeight As Long _
   )
Dim cResampled As cAlphaDIBSection
   ' Resample the input bitmap:
   Set cResampled = m_cSource.AlphaResample(lWidth)
   If (cResampled.Height < lHeight) Then
      ' Need to place the item in a new dib of the
      ' correct size:
      Dim cSized As New cAlphaDIBSection
      cSized.Create lWidth, lHeight
      cSized.SetBackgroundColor m_lTransparentColour
      cSized.SetColourTransparent m_lTransparentColour
      cResampled.CopyTo cSized, (lWidth - cResampled.Width) \ 2, (lHeight - cResampled.Height) \ 2
      Set cResampled = cSized
   End If
   
   ' Set the alpha bits to the result
   cFI.SetImageBits lIndex, cResampled.DIBSectionBitsPtr
   
Dim b() As Byte
Dim lWidthBytes As Long
   lWidthBytes = ((cResampled.Width + 31) \ 32) * 4
   ReDim b(0 To lWidthBytes - 1, 0 To lHeight - 1) As Byte
   
   createMask cResampled, b()
   cFI.SetMaskBits lIndex, VarPtr(b(0, 0))
      
End Sub

Private Sub createMask( _
      cDib As cAlphaDIBSection, _
      b() As Byte _
   )
Dim lWidthBytes As Long
Dim lHeight As Long
Dim lCurVal As Long
Dim lBit As Long
Dim x As Long
Dim y As Long
Dim tSA As SAFEARRAY2D
Dim bDib() As Byte
Dim xOut As Long
Dim yOut As Long
         
   ' Get the bits in the from DIB section:
   With tSA
      .cbElements = 1
      .cDims = 2
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = lHeight
      .Bounds(1).lLbound = 0
      .Bounds(1).cElements = cDib.BytesPerScanLine()
      .pvData = cDib.DIBSectionBitsPtr
   End With
   CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
   
   xOut = 0
   For x = 0 To cDib.BytesPerScanLine() - 4 Step 4
      If (lBit = 8) Then
         lBit = 0
         xOut = xOut + 1
      End If
      For y = 0 To lHeight - 1
         yOut = y
         If (bDib(x + 3, y) = 0) Then
            ' Output = 1
            b(xOut, yOut) = BitSet(b(xOut, yOut), lBit)
         Else
            ' Output = 0
         End If
      Next y
      lBit = lBit + 1
   Next x
   
   ' Clear the temporary array descriptor
   ' (This does not appear to be necessary, but
   ' for safety do it anyway)
   CopyMemory ByVal VarPtrArray(bDib), 0&, 4
   
End Sub
Private Function BitSet(ByVal b As Byte, ByVal lBit As Long) As Byte
   Select Case lBit
   Case 0
      b = b Or &H1
   Case 1
      b = b Or &H2
   Case 2
      b = b Or &H4
   Case 3
      b = b Or &H8
   Case 4
      b = b Or &H10
   Case 5
      b = b Or &H20
   Case 6
      b = b Or &H40
   Case 7
      b = b Or &H80
   End Select
   BitSet = b
End Function

Private Sub createIcon(cFI As cFileIcon)
Dim lIndex As Long
Dim i As Long
Dim iPos As Long
Dim lWidth As Long
Dim lHeight As Long
Dim sWidthHeight As String

   If (chkSize(0).Value = vbChecked) Then
      lIndex = cFI.IconIndex(16, 16, 32)
      If (lIndex = 0) Then
         lIndex = cFI.AddImage(16, 16, 32)
      End If
      createIconAtSize cFI, lIndex, 16, 16
   End If
   If (chkSize(1).Value = vbChecked) Then
      lIndex = cFI.IconIndex(24, 24, 32)
      If (lIndex = 0) Then
         lIndex = cFI.AddImage(24, 24, 32)
      End If
      createIconAtSize cFI, lIndex, 24, 24
   End If
   If (chkSize(2).Value = vbChecked) Then
      lIndex = cFI.IconIndex(32, 32, 32)
      If (lIndex = 0) Then
         lIndex = cFI.AddImage(32, 32, 32)
      End If
      createIconAtSize cFI, lIndex, 32, 32
   End If
   If (chkSize(3).Value = vbChecked) Then
      lIndex = cFI.IconIndex(48, 48, 32)
      If (lIndex = 0) Then
         lIndex = cFI.AddImage(48, 48, 32)
      End If
      createIconAtSize cFI, lIndex, 48, 48
   End If
   If (chkSize(4).Value = vbChecked) Then
      For i = 0 To lstCustom.ListCount
         sWidthHeight = lstCustom.List(i)
         iPos = InStr(sWidthHeight, "x")
         lWidth = CLng(Left(sWidthHeight, iPos - 1))
         lHeight = CLng(Mid(sWidthHeight, iPos + 1))
         lIndex = cFI.IconIndex(lWidth, lHeight, 32)
         If (lIndex = 0) Then
            lIndex = cFI.AddImage(lWidth, lHeight, 32)
         End If
         createIconAtSize cFI, lIndex, lWidth, lHeight
      Next i
   End If
   
End Sub

Private Sub openImage()
   
   Set m_cSource = New cAlphaDIBSection
   m_cSource.CreateFromPicture m_pic
   m_cSource.SetAlpha 255

End Sub

Private Sub setTransparentColour(ByVal lColor As Long)
   
   m_cSource.SetColourTransparent lColor
   
   m_lTransparentColour = lColor

   renderImage
   
End Sub

Private Sub renderImage()
   picSource.Cls
   m_cSource.AlphaPaintPicture picSource.hdc, _
      0, 0, _
      picSource.ScaleWidth \ Screen.TwipsPerPixelX, _
      picSource.ScaleHeight \ Screen.TwipsPerPixelY, _
      0, 0, _
      m_cSource.Width, m_cSource.Height
   picSource.Refresh
End Sub

Private Function fileExists(ByVal sFile As String) As Boolean
Dim sDir As String
   On Error Resume Next
   sDir = Dir(sFile)
   fileExists = (Len(sDir) > 0) And (Err.Number = 0)
End Function

Private Sub chkSize_Click(Index As Integer)
Dim bEnableNext As Boolean
Dim bEnableCustom As Boolean
Dim i As Long

   bEnableNext = (Len(txtFileName.Text) > 0)
   If (bEnableNext) Then
      For i = 0 To 3
         If (chkSize(i).Value = vbChecked) Then
            bEnableNext = True
            Exit For
         End If
      Next i
      If Not (bEnableNext) Then
         If (chkSize(4).Value = vbChecked) Then
            If (lstCustom.ListCount > 0) Then
               bEnableNext = True
            End If
         End If
      End If
   End If
   cmdNext.Enabled = bEnableNext
   
   bEnableCustom = (chkSize(4).Value = vbChecked)
   lstCustom.Enabled = bEnableCustom
   cmdAddCustom.Enabled = bEnableCustom
   cmdRemoveCustom.Enabled = ((lstCustom.ListIndex > -1) And bEnableCustom)
   
End Sub

Private Sub cmdAddCustom_Click()
Dim sR As String
Dim sWidth As String
Dim sHeight As String
Dim bValid As Boolean
Dim lWidth As Long
Dim lHeight As Long
Dim iPos As Long

   sR = InputBox("Enter custom icon size in the form [width] x [height]", App.Title, "")
   If Len(sR) > 0 Then
      iPos = InStr(sR, ",")
      If (iPos = 0) Then
         iPos = InStr(sR, "x")
      End If
      If (iPos > 0) Then
         If (iPos > 1) Then
            sWidth = Trim(Left(sR, iPos - 1))
            If (iPos < Len(sR)) Then
               sHeight = Trim(Mid(sR, iPos + 1))
               If (IsNumeric(sWidth) And IsNumeric(sHeight)) Then
                  On Error Resume Next
                  lWidth = CLng(sWidth)
                  If (Err.Number = 0) Then
                     lHeight = CLng(sHeight)
                     If (Err.Number = 0) Then
                        If (lWidth > 0) And (lHeight > 0) Then
                           If (lWidth = 16) And (lHeight = 16) Then
                           ElseIf (lWidth = 24) And (lHeight = 24) Then
                           ElseIf (lWidth = 32) And (lHeight = 32) Then
                           ElseIf (lWidth = 48) And (lHeight = 48) Then
                           Else
                              lstCustom.AddItem lWidth & "x" & lHeight
                              lstCustom.ListIndex = lstCustom.NewIndex
                              bValid = True
                           End If
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
      If Not (bValid) Then
         MsgBox "The custom icon size " & sR & " is not valid", vbInformation
      End If
   End If
End Sub

Private Sub cmdBack_Click()
   If (m_iTab = 2) Then
      picTab(1).Visible = True
      picTab(0).Visible = False
      picTab(2).Visible = False
      m_iTab = 1
      cmdBack.Enabled = True
      cmdNext.Enabled = (Len(txtFileName.Text) > 0)
   ElseIf (m_iTab = 1) Then
      picTab(0).Visible = True
      picTab(1).Visible = False
      picTab(2).Visible = False
      m_iTab = 0
      cmdBack.Enabled = False
      cmdNext.Enabled = Not (m_pic Is Nothing)
   End If
   
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdNext_Click()
   If (m_iTab = 0) Then
      picTab(1).Visible = True
      picTab(0).Visible = False
      picTab(2).Visible = False
      cmdNext.Enabled = Len(txtFileName.Text) > 0
      cmdBack.Enabled = True
      cmdCancel.Caption = "Cancel"
      m_iTab = 1
   ElseIf (m_iTab = 1) Then
      picTab(2).Visible = True
      picTab(0).Visible = False
      picTab(1).Visible = False
      cmdNext.Enabled = False
      cmdBack.Enabled = True
      cmdStartOver.Enabled = False
      lblInfo.Caption = "Creating Icon..."
      Me.Refresh
         
      Dim cFI As New cFileIcon
      If (fileExists(txtFileName.Text)) Then
         cFI.LoadIcon txtFileName.Text
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         lblInfo.Caption = "Your icon has been written out to " & vbCrLf & _
            txtFileName.Text
      Else
         lblInfo.Caption = "Failed to write the icon to disk."
      End If
      cmdStartOver.Enabled = True
      cmdCancel.Caption = "&Finished"
      m_iTab = 2
   End If
End Sub

Private Sub cmdPickColour_Click()
Dim lColor As Long
Dim cD As New cCommonDialog
   OleTranslateColor picTransparentColour.BackColor, 0, lColor
   If (cD.VBChooseColor(lColor, FullOpen:=True, Owner:=Me.hwnd)) Then
      picTransparentColour.BackColor = lColor
      openImage
      setTransparentColour lColor
   End If
End Sub

Private Sub cmdPickInput_Click()
Dim sFile As String
Dim cD As New cCommonDialog
   sFile = txtSourceImage.Text
   If (cD.VBGetOpenFileName(sFile, _
      Filter:="Picture Files (*.BMP;*.GIF)|*.BMP;*.GIF|Bitmaps (*.BMP)|*.BMP|Graphics Interchange Format (*.GIF)|*.GIF|All Files (*.*)|*.*", _
      DefaultExt:="BMP", _
      Owner:=Me.hwnd)) Then
      Set m_pic = Nothing
      cmdNext.Enabled = False
      txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(sFile)
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      setTransparentColour picTransparentColour.BackColor
      txtSourceImage.Text = sFile
      cmdNext.Enabled = True
   End If
End Sub

Private Sub cmdPickOutput_Click()
Dim sFile As String
Dim cD As New cCommonDialog
   sFile = txtFileName.Text
   If (cD.VBGetSaveFileName(sFile, _
      Filter:="Icon Files (*.ICO)|*.ICO|All Files (*.*)|*.*", _
      DefaultExt:="ico", _
      Owner:=Me.hwnd)) Then
      txtFileName.Text = sFile
      cmdNext.Enabled = Len(txtFileName.Text) > 0 And _
         ((chkSize(0).Value = vbChecked) Or _
         (chkSize(1).Value = vbChecked) Or _
         (chkSize(2).Value = vbChecked))
   End If
End Sub


Private Sub cmdRemoveCustom_Click()
Dim lIndex As Long
   If (lstCustom.ListIndex > -1) Then
      lIndex = lstCustom.ListIndex
      lstCustom.RemoveItem lstCustom.ListIndex
      If (lIndex < lstCustom.ListCount) Then
         lstCustom.ListIndex = lIndex
      Else
         If (lIndex - 1) > -1 Then
            lstCustom.ListIndex = lIndex - 1
         End If
      End If
      
      cmdRemoveCustom.Enabled = (lstCustom.ListIndex > -1)
   End If
End Sub

Private Sub cmdStartOver_Click()
   '
   cmdBack_Click
   cmdBack_Click
   '
End Sub

Private Sub Form_Initialize()

   ' SPM - Form_Initialize fired when object is being created,
   ' i.e. before hWnd created.
   HookAttach
   
   m_lTransparentColour = CLR_INVALID
   
End Sub

Private Sub Form_Load()
   
   ' SPM - Form_Load is fired *after* the window is created.
   ' Therefore we will already have set the style so there
   ' is no need for the hook anymore.
   HookDetach
   
   SetIcon Me.hwnd, "AAA"
   
   picTab(1).Move picTab(0).Left, picTab(0).TOp, picTab(0).Width, picTab(0).Height
   picTab(2).Move picTab(0).Left, picTab(0).TOp, picTab(0).Width, picTab(0).Height
   picTab(0).BorderStyle = 0
   picTab(1).BorderStyle = 0
   picTab(2).BorderStyle = 0
   
End Sub

Private Sub lstCustom_Click()
   If (lstCustom.ListIndex > -1) Then
      cmdRemoveCustom.Enabled = True
   End If
End Sub
