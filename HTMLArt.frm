VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmHTMLArt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HTMLArt"
   ClientHeight    =   6030
   ClientLeft      =   1620
   ClientTop       =   1530
   ClientWidth     =   8715
   Icon            =   "HTMLArt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6030
   ScaleWidth      =   8715
   Begin VB.CheckBox ckhView 
      Caption         =   "View"
      Height          =   255
      Left            =   6840
      TabIndex        =   11
      ToolTipText     =   "Veiw your art in the default browser"
      Top             =   5400
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox chkAutoH 
      Caption         =   "Keep Aspect"
      Height          =   255
      Left            =   7200
      TabIndex        =   6
      ToolTipText     =   "Keep HTML aspect ratio the same as the original image"
      Top             =   3120
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.PictureBox picBar 
      ForeColor       =   &H8000000D&
      Height          =   200
      Left            =   120
      ScaleHeight     =   135
      ScaleWidth      =   8415
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Shows progress when making HTMLArt"
      Top             =   5760
      Width           =   8475
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "?"
      Height          =   375
      Left            =   8160
      TabIndex        =   13
      ToolTipText     =   "About HTMLArt"
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton cmdTxtOpen 
      Height          =   375
      Left            =   8040
      Picture         =   "HTMLArt.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Open text file"
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8040
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdImgOpen 
      Height          =   375
      Left            =   8040
      Picture         =   "HTMLArt.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Open Image file"
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CheckBox chkUpper 
      Caption         =   "Make text upper case"
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      ToolTipText     =   "Convert text to all upper case?"
      Top             =   2040
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.TextBox txtSaveAs 
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      ToolTipText     =   "File name for your creation"
      Top             =   4680
      Width           =   3615
   End
   Begin VB.TextBox txtTitle 
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      ToolTipText     =   "Title for your creation"
      Top             =   3960
      Width           =   3615
   End
   Begin VB.TextBox txtY 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Text            =   "120"
      ToolTipText     =   "Final HTML height"
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtX 
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Text            =   "180"
      ToolTipText     =   "Final HTML width"
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdMake 
      Caption         =   "Make HTMLArt"
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      ToolTipText     =   "Create and save your HTMLArt"
      Top             =   5280
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000010&
      DrawMode        =   6  'Mask Pen Not
      Height          =   5555
      Left            =   120
      ScaleHeight     =   5490
      ScaleWidth      =   4500
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   4555
   End
   Begin VB.Label lblIT 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(none)"
      Height          =   375
      Left            =   4920
      TabIndex        =   18
      ToolTipText     =   "Selected Text file"
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label lblIF 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(none)"
      Height          =   375
      Left            =   4920
      TabIndex        =   17
      ToolTipText     =   "Selected Image file"
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Save file name:"
      Height          =   255
      Index           =   4
      Left            =   4920
      TabIndex        =   16
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "HTML title:"
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   15
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "HTML Width       HTML Height"
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   14
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Text file:"
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Image file:"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   7
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmHTMLArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''
'''' HTMLArt                             ''''
'''' by Paul Bahlawan March 13, 2001     ''''
'''' Update: Sep 8, 2003                 ''''
'''' Update: July 7, 2005                ''''
'''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Type RGBType
    R As Byte
    G As Byte
    b As Byte
    Filler As Byte
End Type
Private Type RGBLongType
    clr As Long
End Type
Dim iFile As String
Dim tFile As String
Dim warned As Boolean

'Main sub to create ART!
Private Sub cmdMake_Click()
Dim x As Long, y As Long
Dim z As String
Dim oldColor As String

On Error GoTo ErrHandler

    'Some error checking...
    If txtSaveAs.Text = "" Or lblIF.Caption = "(none)" Or tFile = "" Then
        MsgBox "You must select a valid image AND a valid text file before this command is allowed.", vbExclamation, "HTMLArt - Error"
        Exit Sub
    End If
    
    If Val(txtX.Text) < 5 Or Val(txtY.Text) < 5 Then
        MsgBox "The width and/or height setting(s) are too small.", vbExclamation, "HTMLArt - Error"
        Exit Sub
    End If
    
    If Not warned Then
        If Val(txtX.Text) > 300 Or Val(txtY.Text) > 300 Then
            z = MsgBox("Excessivly large width and/or height setting(s) are not recomended." & Chr(13) & "Click OK to continue anyway.", vbOKCancel + vbQuestion, "HTMLArt - Warning")
            If z = vbCancel Then Exit Sub
            warned = True
        End If
    End If
    
    Open tFile For Binary Access Read As #2
    If LOF(2) < 1 Then
        Close #2
        MsgBox "The selected text file is empty.", vbExclamation, "HTMLArt - Error"
        Exit Sub
    End If
    
    
    'Make & write HTML file to disk
    Open App.Path & "\" & txtSaveAs.Text For Output As #1
    Print #1, "<HTML>"
    Print #1, "<TITLE>" & txtTitle.Text & "</TITLE>"
    Print #1, "<!-- This image was generated with HTMLArt by Paul Bahlawan -->"
    Print #1, "<BODY BGCOLOR=""#000000"">"
    Print #1, "<BASEFONT SIZE=""1""><PRE>"
        
        For y = 0 To Picture1.ScaleHeight Step Picture1.ScaleHeight / Val(txtY.Text)
            picBar.Line (0, 0)-(y * (picBar.ScaleWidth / Picture1.ScaleHeight), picBar.ScaleHeight), , BF
            For x = 0 To (Picture1.ScaleWidth - 5) Step Int(Picture1.ScaleWidth / Val(txtX.Text))
                z = Process(x, y)
                If oldColor <> z Then  'same color as last time, then no need to put it again
                    If x + y <> 0 Then Print #1, "</FONT>";
                    Print #1, "<FONT COLOR=" & Chr$(34) & "#" & z & Chr$(34) & ">";
                    oldColor = z
                End If
                Print #1, getChar;
            Next x
            Print #1, Chr$(13);
        Next y
    
    Print #1, "</FONT></PRE></BODY></HTML>"
    Close #2
    Close #1
    
    'View the result in the default browser
    If ckhView.Value Then
        ShellExecute Me.hwnd, "open", App.Path & "\" & txtSaveAs.Text, vbNullString, "", 1
    End If
    
    picBar.Cls
    Exit Sub
    
ErrHandler:
    MsgBox "Unexpected error = " & Err.Description, vbCritical, "HTMLArt - Error"
    Close #2
    Close #1
    picBar.Cls
End Sub

'Get color from original pic and convert it
Function Process(x As Long, y As Long) As String
    Process = ColorToHTML(Picture1.Point(x, y))
End Function

'Convert VB color to HTML color
Function ColorToHTML(clr As Long) As String
Dim R As RGBType, RL As RGBLongType
Dim aR As String, aG As String, aB As String
    RL.clr = clr
    LSet R = RL
    With R
        aR = Trim$(Hex$(.R)): If Len(aR) = 1 Then aR = "0" & aR
        aG = Trim$(Hex$(.G)): If Len(aG) = 1 Then aG = "0" & aG
        aB = Trim$(Hex$(.b)): If Len(aB) = 1 Then aB = "0" & aB
    End With
    
    ColorToHTML = aR & aG & aB
End Function

'Retrieve next byte from text file & convert special characters to HTML
Function getChar() As String
Dim a As String
Dim endless As Long
    a = "X"
    
    'loop until an acceptable character is found (drop spaces, carrage returns, etc)
    Do
        Get #2, , a
        If EOF(2) Then 'When we reach the end... start again!
            Get #2, 1, a
            endless = endless + 1
            If endless > 1 Then 'Prevent an endless loop incase of non-text file
                a = "X"
                Exit Do
            End If
        End If
        If a = Chr$(34) Then a = "&quot;": Exit Do
        If a = Chr$(38) Then a = "&amp;": Exit Do
        If a = Chr$(60) Then a = "&lt;": Exit Do
        If a = Chr$(62) Then a = "&gt;": Exit Do
        If a >= Chr$(33) And a <= Chr$(126) Then Exit Do
    Loop
    
    If chkUpper.Value Then a = UCase(a)
    
    getChar = a
End Function

'Select Image file
Private Sub cmdImgOpen_Click()

On Error GoTo ErrHandler 'Mostly to handle a "Cancel" click

    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
    CommonDialog1.Filter = "Pictures (.bmp .gif .jpg .ico)|*.bmp;*.gif;*.jpg;*.jpeg;*.ico|All (*.*)|*.*"
    If iFile <> "" Then
        CommonDialog1.InitDir = iFile
        CommonDialog1.FileName = lblIF.Caption
    Else
        CommonDialog1.InitDir = ""
        CommonDialog1.FileName = ""
    End If
    CommonDialog1.ShowOpen
    lblIF.Caption = CommonDialog1.FileTitle
    Call ScalePicture(Picture1, CommonDialog1.FileName)
    txtTitle.Text = Left$(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4)
    iFile = CommonDialog1.FileName
    
    If chkAutoH.Value Then
        AutoHeight
    End If
    
ErrHandler:
End Sub

'Select TEXT file
Private Sub cmdTxtOpen_Click()

On Error GoTo ErrHandler 'Mostly to handle a "Cancel" click

    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
    CommonDialog1.Filter = "Text (.txt)|*.txt|All (*.*)|*.*"
    If tFile <> "" Then
        CommonDialog1.InitDir = tFile
        CommonDialog1.FileName = lblIT.Caption
    Else
        CommonDialog1.InitDir = ""
        CommonDialog1.FileName = ""
    End If
    CommonDialog1.ShowOpen
    lblIT.Caption = CommonDialog1.FileTitle
    tFile = CommonDialog1.FileName
  
ErrHandler:
End Sub

'Auto Height (attempt to keep original aspect ratio)
Private Sub AutoHeight()
    If Val(txtX.Text) > 1 Then
        txtY.Text = Trim$(CStr(Int(Picture1.ScaleHeight / (Picture1.ScaleWidth / Val(txtX.Text)) * 0.65)))
    End If
End Sub

'Select/deselect "Keep Aspect"
Private Sub chkAutoH_Click()
    If chkAutoH.Value Then
        txtY.Enabled = False
        AutoHeight
    Else
        txtY.Enabled = True
    End If
End Sub

'Re-calc height when width is changed
Private Sub txtX_Change()
    If chkAutoH.Value Then
        AutoHeight
    End If
End Sub

'ABOUT
Private Sub cmdAbout_Click()
    MsgBox "HTMLArt by Paul Bahlawan  -  Version " & App.Major & "." & App.Minor & " (build" & App.Revision & ")", vbInformation, "HTMLArt - About"
End Sub

'Guess save file name
Private Sub txtTitle_Change()
    txtSaveAs.Text = txtTitle.Text & ".html"
End Sub

'Resize image to best fit
'This sub: Thanks to Frank Adam on comp.lang.basic.visual.misc for his assistance.
Private Sub ScalePicture(pb As PictureBox, strPicturePath As String)
Const xScale As Single = 0.566941199004022
Const yScale As Single = 0.566909725788231
Dim ScaleFactor As Single
Dim pic As StdPicture

    Set pic = LoadPicture(strPicturePath)
    If pic.Width >= pic.Height Then
        Picture1.Width = 4555
        With pb
            ScaleFactor = .Width / (pic.Width * xScale)
            .Height = pic.Height * yScale * ScaleFactor
            .PaintPicture pic, 0, 0, .Width, .Height, , , , , vbSrcCopy
        End With
    Else
        Picture1.Height = 4555
        With pb
            ScaleFactor = .Height / (pic.Height * xScale)
            .Width = pic.Width * yScale * ScaleFactor
            .PaintPicture pic, 0, 0, .Width, .Height, , , , , vbSrcCopy
        End With
    End If
    Set pic = Nothing
End Sub

