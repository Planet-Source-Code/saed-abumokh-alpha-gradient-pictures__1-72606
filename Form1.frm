VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Alpha gradient picture by Saed Abumokh"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   ScaleHeight     =   523
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   746
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFlip 
      Caption         =   "Flip"
      Height          =   345
      Left            =   7440
      TabIndex        =   10
      Top             =   7350
      Width           =   1245
   End
   Begin VB.ComboBox cmbDirection 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   5850
      List            =   "Form1.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   7380
      Width           =   1515
   End
   Begin VB.CheckBox chkStretch 
      Caption         =   "Stretch"
      Height          =   225
      Left            =   4560
      TabIndex        =   8
      Top             =   7410
      Value           =   1  'Checked
      Width           =   1185
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Copy"
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   7320
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3360
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Choose Picture2"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Choose Picture1"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   7320
      Width           =   1455
   End
   Begin VB.HScrollBar hEnd 
      Height          =   255
      LargeChange     =   20
      Left            =   120
      TabIndex        =   4
      Top             =   6960
      Value           =   32767
      Width           =   10875
   End
   Begin VB.HScrollBar hStart 
      Height          =   255
      LargeChange     =   20
      Left            =   120
      TabIndex        =   3
      Top             =   6600
      Width           =   10935
   End
   Begin VB.PictureBox picDest 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   120
      ScaleHeight     =   425
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   729
      TabIndex        =   2
      Top             =   210
      Width           =   10935
   End
   Begin VB.PictureBox pic2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   120
      Picture         =   "Form1.frx":0024
      ScaleHeight     =   425
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   729
      TabIndex        =   1
      Top             =   120
      Width           =   10935
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   120
      Picture         =   "Form1.frx":72DAA
      ScaleHeight     =   425
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   729
      TabIndex        =   0
      Top             =   120
      Width           =   10935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal blend As Long) As Long

Private Const FilterPictures = "All Picture Files |*.emf;*.wmf;*.jpg;*.jpeg;*.jfif;*.jpe;*.bmp;*.dib;*.gif|" & _
"Enhanced Windows MetaFile(*.emf)|*.emf|" & _
"Windows MetaFile(*.wmf)|*.wmf|" & _
"JPEG Interchange Format(*.jpg;*.jpeg;*.jfif;*.jpe)|*.jpg;*.jpeg;*.jfif;*.jpe|" & _
"Device Independent Bitmap(*.dib)|*.dib|" & _
"Bitmaps(*.bmp;*.rle)|*.bmp|" & _
"GIF Graphics Interchange Format(*.gif)|*.gif|"
Dim Picture1 As IPictureDisp, Picture2 As IPictureDisp
Private Sub AlphaGradientPicture()
        picDest.Cls
        pic1.Picture = Picture1
        If chkStretch.Value = vbChecked Then
            StretchPicture pic1, Picture1
            StretchPicture pic2, Picture2
        End If
        AlphaGradient picDest.hDC, picDest.ScaleWidth, picDest.ScaleHeight, pic2.hDC, pic2.ScaleWidth, pic2.ScaleHeight, hStart.Value, hEnd.Value, cmbDirection.ListIndex
        picDest.Refresh
End Sub

Private Sub chkStretch_Click()
    
    pic1.Picture = Picture1
    pic2.Picture = Picture2
    If chkStretch.Value = vbChecked Then
        StretchPicture pic1, pic1.Picture
        StretchPicture pic2, pic2.Picture
    ElseIf chkStretch.Value = vbUnchecked Then
    End If
    AlphaGradientPicture
End Sub

Private Sub cmbDirection_Change()
    If cmbDirection.ListIndex = 0 Then
        hStart.Max = picDest.ScaleWidth
        hEnd.Max = picDest.ScaleWidth
    ElseIf cmbDirection.ListIndex = 1 Then
        hStart.Max = picDest.ScaleHeight
        hEnd.Max = picDest.ScaleHeight
    End If
End Sub

Private Sub cmbDirection_Click()
    cmbDirection_Change
End Sub

Private Sub cmdFlip_Click()
    Dim PicSwap As StdPicture
    Set PicSwap = pic1.Picture
    Set pic1.Picture = pic2.Picture
    Set pic2.Picture = PicSwap
    Form_Load
End Sub

Private Sub Command1_Click()
    ' choose the first picture and save the selected picture to renew it when the picture is stretched
    cd.ShowOpen
    cd.Filter = FilterPictures
    pic1.Picture = LoadPicture(LCase$(cd.FileName))
    Set Picture1 = pic1.Picture
    Form_Load
End Sub

Private Sub Command2_Click()
    ' choose the second picture and save the selected picture to renew it when the picture is stretched
    cd.ShowOpen
    cd.Filter = FilterPictures
    pic2.Picture = LoadPicture(LCase$(cd.FileName))
    Set Picture2 = pic2.Picture
    Form_Load
End Sub

Private Sub Command3_Click()
    'copy the image
    Clipboard.Clear
    Clipboard.SetData picDest.Image
End Sub

Private Sub Form_Load()
    'save the initial pictures to renew them when the pictures are stretched
    Set Picture1 = pic1.Picture
    Set Picture2 = pic2.Picture
    'stretch the two pictures and the blended image
    StretchPicture pic1, pic1.Picture
    StretchPicture pic2, pic2.Picture
    StretchPicture picDest, pic1.Picture
        
    chkStretch_Click

    If cmbDirection.ListIndex = -1 Then cmbDirection.ListIndex = 0
End Sub
Private Function StretchPicture(PictureBox As PictureBox, Picture As IPictureDisp)
    PictureBox.PaintPicture Picture, 0, 0, PictureBox.ScaleWidth, PictureBox.ScaleHeight, , , , , vbSrcCopy
    PictureBox.Picture = PictureBox.Image
    PictureBox.Refresh
End Function
Private Function BlendValue(BlendVal) As Long
    If BlendVal > 255 Then BlendVal = 255
    BlendValue = RGB(0, 0, CByte(BlendVal))
End Function

Private Sub AlphaGradient(ByVal hDestDC As Long, ByVal DestWidth As Double, ByVal DestHeight As Double, ByVal hSrcDC As Double, ByVal SrcWidth As Double, ByVal SrcHeight As Double, ByVal StartBlend As Double, ByVal EndBlend As Double, ByVal Direction As Double)
    
    Dim i As Double
    Dim DeltaBlend As Double ' the length of the area that drawn the alpha pictures gradually
    Dim StepLength As Double ' to calcualte step length if we want to draw 256 alpha (translucent) parts of the pictures
    
    
    DeltaBlend = EndBlend - StartBlend  '  like   dX = X2 - X1
    StepLength = 256 / DeltaBlend ' calcualte the step length
    
    If Direction = 0 Then ' Horizontal
    
        For i = 1 To 256 Step StepLength ' only 256 steps to draw the alpha picture fast
        
        ' begining form StartBlend and drawing each part (coloumn) of picture,the part length _
          is 1/256 of the DeltaBlend ( start(nXOriginDest) is i * 256 part of the whole _
          picture and the width(nWidthDest) is 1 part of the picture(width / 256)
            AlphaBlend hDestDC, DeltaBlend / 256 * i + StartBlend, 0, DestWidth / 256, DestHeight, _
                        hSrcDC, DeltaBlend / 256 * i + StartBlend, 0, DestWidth / 256, SrcHeight, _
                         BlendValue(i - 1) ' this function requires the blend value is _
                                             between 0*(256^2) and 255*(256^2) (between &H000000 and &HFF0000)
        Next
        'Draw the rest of the alpha picture (after EndBlend with no alpha (alpha=255)
        AlphaBlend hDestDC, EndBlend, 0, DestWidth - EndBlend, DestHeight, hSrcDC, EndBlend, 0, SrcWidth - EndBlend, DestHeight, BlendValue(255)
        
    ElseIf Direction = 1 Then 'Vertical
        For i = 1 To 256 Step StepLength
            AlphaBlend hDestDC, 0, DeltaBlend / 256 * i + StartBlend, DestWidth, DestHeight / 256 + 1, _
                         hSrcDC, 0, (DeltaBlend / 256 * i) + StartBlend, SrcWidth, SrcHeight / 256 + 1, _
                         BlendValue(i - 1)
        Next
        AlphaBlend hDestDC, 0, EndBlend, DestWidth, DestHeight - EndBlend, hSrcDC, 0, EndBlend, SrcWidth, DestHeight - EndBlend, BlendValue(255)
        
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With picDest
        .Width = Me.ScaleWidth - (.Left + 10)
        .Height = Me.ScaleHeight - (.Top + 10 + hStart.Height + 10 + hEnd.Height + 10 + Command1.Height + 10)
        pic1.Move .Left, .Top, .Width, .Height
        pic2.Move .Left, .Top, .Width, .Height
    End With
    
    hStart.Width = Me.ScaleWidth - (hStart.Left + 10)
    hStart.Top = Me.ScaleHeight - (10 + Command1.Height + 10 + hEnd.Height + 10 + hStart.Height)
    
    hEnd.Width = Me.ScaleWidth - (hEnd.Left + 10)
    hEnd.Top = Me.ScaleHeight - (10 + Command1.Height + 10 + hEnd.Height)
    
    Command1.Top = Me.ScaleHeight - (10 + Command1.Height)
    
    Command2.Top = Me.ScaleHeight - (10 + Command2.Height)
    
    Command3.Top = Me.ScaleHeight - (10 + Command3.Height)
    
    chkStretch.Top = Me.ScaleHeight - (chkStretch.Height + 10)
    
        
    chkStretch_Click
End Sub

Private Sub hEnd_Change()
    AlphaGradientPicture
End Sub

Private Sub hEnd_Scroll()
    hEnd_Change
End Sub

Private Sub hStart_Change()
    AlphaGradientPicture
End Sub

Private Sub hStart_Scroll()
    hStart_Change
End Sub
