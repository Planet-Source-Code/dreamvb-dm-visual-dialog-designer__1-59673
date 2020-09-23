VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "DM++ Visual Dialog Designer"
   ClientHeight    =   5940
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox p1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   5250
      Picture         =   "frmmain.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5940
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox p2 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   557
      TabIndex        =   29
      Top             =   5610
      Width           =   8355
   End
   Begin VB.PictureBox PicProp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2580
      Left            =   0
      ScaleHeight     =   2580
      ScaleWidth      =   2520
      TabIndex        =   12
      Top             =   2580
      Width           =   2520
      Begin VB.TextBox txtProp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "0"
         Top             =   2190
         Width           =   1065
      End
      Begin VB.TextBox txtProp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   1245
         TabIndex        =   25
         Text            =   "0"
         Top             =   1905
         Width           =   1065
      End
      Begin VB.TextBox txtProp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1245
         TabIndex        =   23
         Text            =   "0"
         Top             =   1620
         Width           =   1065
      End
      Begin VB.TextBox txtProp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1245
         TabIndex        =   22
         Text            =   "0"
         Top             =   1335
         Width           =   1065
      End
      Begin VB.TextBox txtProp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1245
         TabIndex        =   21
         Text            =   "0"
         Top             =   1050
         Width           =   1065
      End
      Begin VB.TextBox txtProp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1245
         TabIndex        =   20
         Text            =   "0"
         Top             =   765
         Width           =   1065
      End
      Begin VB.ComboBox cboProp 
         Height          =   315
         Left            =   15
         TabIndex        =   15
         Top             =   345
         Width           =   2430
      End
      Begin VB.PictureBox PicA 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   15
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   13
         Top             =   15
         Width           =   1155
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Properties"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   75
            TabIndex        =   14
            Top             =   45
            Width           =   870
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   945
         X2              =   945
         Y1              =   720
         Y2              =   2475
      End
      Begin VB.Label lbProp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enabled"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   75
         TabIndex        =   26
         Top             =   2205
         Width           =   570
      End
      Begin VB.Label lbProp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   75
         TabIndex        =   24
         Top             =   1905
         Width           =   315
      End
      Begin VB.Label lbProp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   75
         TabIndex        =   19
         Top             =   1620
         Width           =   405
      End
      Begin VB.Label lbProp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   75
         TabIndex        =   18
         Top             =   1335
         Width           =   450
      End
      Begin VB.Label lbProp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   75
         TabIndex        =   17
         Top             =   1050
         Width           =   285
      End
      Begin VB.Label lbProp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Top"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   75
         TabIndex        =   16
         Top             =   765
         Width           =   270
      End
   End
   Begin VB.PictureBox PicBase 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2355
      Left            =   0
      ScaleHeight     =   2355
      ScaleWidth      =   2520
      TabIndex        =   0
      Top             =   15
      Width           =   2520
      Begin VB.PictureBox Picbar 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   15
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   2
         Top             =   15
         Width           =   1155
         Begin VB.Label lblcap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Designer Tools"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   75
            TabIndex        =   3
            Top             =   45
            Width           =   1290
         End
      End
      Begin Project1.DevToolbar DevToolbar1 
         Height          =   1890
         Left            =   15
         TabIndex        =   1
         Top             =   330
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   3334
         Picture         =   "frmmain.frx":0342
      End
   End
   Begin VB.PictureBox PicForm 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3915
      Left            =   2520
      ScaleHeight     =   3915
      ScaleWidth      =   5085
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   15
      Width           =   5085
      Begin VB.PictureBox PicFormHolder 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3780
         Left            =   135
         ScaleHeight     =   3780
         ScaleWidth      =   4965
         TabIndex        =   5
         Top             =   105
         Width           =   4965
         Begin VB.PictureBox FrmHangle 
            BackColor       =   &H00800000&
            BorderStyle     =   0  'None
            Height          =   90
            Left            =   405
            MousePointer    =   8  'Size NW SE
            ScaleHeight     =   90
            ScaleWidth      =   90
            TabIndex        =   6
            Top             =   3105
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.PictureBox PicFrmSrc 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   3660
            Left            =   15
            ScaleHeight     =   3660
            ScaleWidth      =   4800
            TabIndex        =   7
            Tag             =   "1"
            Top             =   30
            Width           =   4800
            Begin VB.CheckBox CheckBox 
               Caption         =   "CheckBox"
               Height          =   225
               Index           =   0
               Left            =   195
               TabIndex        =   31
               Tag             =   "1"
               Top             =   1725
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.ListBox LB 
               Height          =   450
               Index           =   0
               IntegralHeight  =   0   'False
               ItemData        =   "frmmain.frx":1194
               Left            =   45
               List            =   "frmmain.frx":1196
               TabIndex        =   28
               Tag             =   "1"
               Top             =   1005
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.CommandButton TButton 
               Caption         =   "Button"
               Height          =   350
               Index           =   0
               Left            =   60
               TabIndex        =   10
               Tag             =   "1"
               Top             =   30
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.PictureBox Hangle 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   90
               Left            =   2415
               MousePointer    =   8  'Size NW SE
               ScaleHeight     =   90
               ScaleWidth      =   90
               TabIndex        =   9
               Top             =   3150
               Visible         =   0   'False
               Width           =   90
            End
            Begin VB.TextBox tEdit 
               Height          =   300
               Index           =   0
               Left            =   45
               Locked          =   -1  'True
               TabIndex        =   8
               Tag             =   "1"
               Text            =   "Edit"
               Top             =   660
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Shape Selection 
               BorderColor     =   &H00808080&
               Height          =   180
               Left            =   3180
               Top             =   3105
               Visible         =   0   'False
               Width           =   195
            End
            Begin VB.Label tLabel 
               Caption         =   "Static"
               Height          =   240
               Index           =   0
               Left            =   60
               TabIndex        =   11
               Tag             =   "1"
               Top             =   435
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Image ImgTmr 
               Height          =   420
               Index           =   0
               Left            =   2655
               Picture         =   "frmmain.frx":1198
               Tag             =   "0"
               Top             =   1170
               Visible         =   0   'False
               Width           =   420
            End
         End
         Begin VB.Shape sh2 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            BorderWidth     =   3
            Height          =   210
            Left            =   1200
            Top             =   3015
            Visible         =   0   'False
            Width           =   195
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnucpy 
         Caption         =   "&Copy to Clipboard"
      End
      Begin VB.Menu mnublank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuopen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnusave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnublank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu mnugrid 
         Caption         =   "&Grid"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnudesign 
         Caption         =   "&Designer Tools"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "&Properties"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long

' OK this is my DM++ Visual Dialog Designer.

' as you may or maynot be aware I am working on a new version of DM++
' that allow you to add dialogs to your scripts.
' well this has taken me over 2 hours to write from scrach and now my eye are hurting.
' anyway here are some of the features. Please note there maybe one or two bugs
' but for now it seems to do what I want it to.

' at the moment it only supports six controls. but I will be adding more in the final version.

' You can move and resize any control
' Turn on of off the grid support
' added some properties for each control. but they are very basic at the moment' untill I find a better method to use
' Delete controls by selecting a control and pressing the delete key
' Copy a control and paste a control just like in VB
' you can also move controls while selected by pressing CTRL+ArrowKeys and also resize
' You can also save a form and then reloaded it. I have included a small text file

' Well that about it hope you like the code. and what out for the new version of DM++ comming soon
' as aways use the code as you see fit. all I ask is you remmber were it came from.

' Variabes below are for the form designer
Private ObjX As Integer, ObjY As Integer, Form_Object As Object, _
CanObjMove As Boolean, ObjCanResize As Boolean, isControl As Boolean, inPropList As Boolean
Dim m_DialogCaption As String ' used to hold the dialogs caption

Dim CboTmp As String, LastFocus As Integer
Dim PasteCtr As String
Dim clsDialog As New CDialog

Private Sub LoadGUI(lzFileName As String)
' This sub is used for the form and controls loading
' Note this was done in a hurry. as I only wanted to show the data been loaded up as an example/

Dim s As String, DlgBuff As String, StrControls As String
Dim n_cap As String, n_Top As Long
Dim e_pos As Long, n_pos As Long
Dim I As Integer, sLine As String
Dim PropName As String, PropData As Variant
Dim vData As Variant, vControlData As Variant
Dim iCount As Integer
    
    iCount = 0
    
    s = OpenForm(lzFileName)
    'Dialog Code
    e_pos = InStr(1, s, "<dialog>", vbTextCompare)
    n_pos = InStr(e_pos + 1, s, "</dialog>", vbTextCompare)

    If Not CBool((e_pos > 0 And n_pos > 0)) Then
        MsgBox "Can't load form data", vbExclamation, "error"
        Exit Sub
    Else
        DlgBuff = Mid(s, e_pos + 8, n_pos - e_pos - 8)
        vData = Split(DlgBuff, vbCrLf)
        
        For I = 0 To UBound(vData)
            sLine = vData(I)
            e_pos = InStr(1, sLine, "=", vbBinaryCompare)
            If e_pos <> 0 Then
                PropName = UCase(Trim(Mid(sLine, 1, e_pos - 1)))
                PropData = Trim(Mid(sLine, e_pos + 1, Len(sLine)))
                If Left(PropData, 1) = Chr(34) Then PropData = Right(PropData, Len(PropData) - 1)
                If Right(PropData, 1) = Chr(34) Then PropData = Left(PropData, Len(PropData) - 1)
                
                Select Case PropName
                    Case "CAPTION": m_DialogCaption = PropData
                    Case "HEIGHT": PicFrmSrc.Height = PropData
                    Case "WIDTH": PicFrmSrc.Width = PropData
                    Case "ENABLED": PicFrmSrc.Tag = Abs(CBool(PropData))
                End Select
            End If
        Next
        sLine = ""
        I = 0
        Erase vData
    End If
    
    ' next we load the controls
    e_pos = InStr(1, s, "<Controls>", vbTextCompare)
    n_pos = InStr(e_pos + 1, s, "</Controls>", vbTextCompare)
    
    If Not (e_pos > 0 And n_pos > 0) Then
        MsgBox "Can't load controls data", vbExclamation, "error"
        Exit Sub
    Else
        StrControls = Mid(s, e_pos + 10, n_pos - e_pos - 10)
        vData = Split(StrControls, vbCrLf)
        
        For I = 0 To UBound(vData)
            sLine = Trim(vData(I))
            e_pos = InStr(1, sLine, Chr(32), vbBinaryCompare)
            If e_pos <> 0 Then
                If UCase(Mid(sLine, 1, e_pos - 1)) = "ADDCONTROL" Then
                    sLine = Mid(sLine, e_pos + 1, Len(sLine) - 1)
                    vControlData = Split(sLine, ",", , vbBinaryCompare)
                    
                    Select Case UCase(vControlData(0))
                        Case "BUTTON"
                            AddControl CStr(vControlData(0))
                            iCount = TButton.Count - 1
                            Set Form_Object = TButton(iCount)
                        Case "STATIC"
                            AddControl CStr(vControlData(0))
                            iCount = tLabel.Count - 1
                            Set Form_Object = tLabel(iCount)
                        Case "EDIT"
                            AddControl CStr(vControlData(0))
                            iCount = tEdit.Count - 1
                            Set Form_Object = tEdit(iCount)
                        Case "TMR"
                            AddControl CStr(vControlData(0))
                            iCount = ImgTmr.Count - 1
                            Set Form_Object = ImgTmr(iCount)
                        Case "LB"
                            AddControl CStr(vControlData(0))
                            iCount = LB.Count - 1
                            Set Form_Object = LB(iCount)
                        Case "CHECKBOX"
                            AddControl CStr(vControlData(0))
                            iCount = CheckBox.Count - 1
                            Set Form_Object = CheckBox(iCount)
                        End Select
                        
                    If CStr(vControlData(0)) <> "TMR" Then
                        n_cap = vControlData(6)
                        n_cap = Replace(n_cap, Chr(34), "", , , vbBinaryCompare)
        
                        Form_Object.Top = vControlData(1) + 315
                        Form_Object.Left = vControlData(2)
                        Form_Object.Height = vControlData(3)
                        Form_Object.Width = vControlData(4)
                        Form_Object.Tag = Abs(CBool(vControlData(5)))
                        
                        If vControlData(0) = "EDIT" Or vControlData(0) = "LB" Then
                            Form_Object.Text = n_cap
                        Else
                            Form_Object.Caption = n_cap
                        End If
                    Else
                        Form_Object.Tag = Abs(CBool(vControlData(2)))
                    End If
                End If
            End If
        Next
    End If
    
    Call RedrawForm
    PicFrmSrc_MouseDown 1, 0, 0, 0
    
    'Do a quick clean up
    s = ""
    DlgBuff = ""
    StrControls = ""
    n_cap = ""
    PropName = ""
    PropData = ""
    If Not IsEmpty(vControlData) Then Erase vControlData

End Sub
Public Sub UnloadControls()
Dim I As Integer
    ' unload all the controls on the form designer
    For I = 1 To TButton.Count - 1
        Unload TButton(I)
    Next
    
    For I = 1 To tLabel.Count - 1
        Unload tLabel(I)
    Next

    For I = 1 To tEdit.Count - 1
        Unload tEdit(I)
    Next

    For I = 1 To CheckBox.Count - 1
        Unload CheckBox(I)
    Next
    
    For I = 1 To ImgTmr.Count - 1
        Unload ImgTmr(I)
    Next
    
    For I = 1 To LB.Count - 1
        Unload LB(I)
    Next
    
End Sub

Function OpenForm(lzForm As String) As String
Dim iFile As Long
Dim sByte() As Byte

    iFile = FreeFile
    Open lzForm For Binary As #iFile
        ReDim sByte(0 To LOF(iFile))
        Get #iFile, , sByte
    Close #iFile
    
    OpenForm = StrConv(sByte, vbUnicode)
    Erase sByte
    
End Function
Private Sub FixPositions()
    If Form_Object.Left < 0 Then Form_Object.Left = 0
    If Form_Object.Top < 315 Then Form_Object.Top = 315
    If Form_Object.Left > (PicFrmSrc.Width - Form_Object.Width) Then Form_Object.Left = (PicFrmSrc.Width - Form_Object.Width)
    If Form_Object.Top > (PicFrmSrc.Height - Form_Object.Height) Then Form_Object.Top = (PicFrmSrc.Height - Form_Object.Height)
End Sub

Private Sub CheckBox_Click(Index As Integer)
    CheckBox(Index).Value = 0
End Sub

Private Sub checkbox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Designer 1, CheckBox(Index), Button, X, Y, True
End Sub

Private Sub checkbox_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Designer 2, CheckBox, Button, X, Y, False
End Sub

Private Sub checkbox_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Designer 0, CheckBox, Button, X, Y, False
End Sub

Sub DrawStats()
    ' this little sub makes a nise little statusbar at the bottom of the form
    p2.Cls
    p2.Line (0, 2)-(p2.ScaleWidth - 1, p2.ScaleHeight - 1), vbButtonShadow, B
    TransparentBlt p2.hdc, p2.ScaleWidth - 16, p2.ScaleHeight - 16, 16, 16, p1.hdc, 0, 0, 16, 16, RGB(255, 0, 255)
End Sub

Sub RedrawForm()
    PicFrmSrc.Cls
    ' This small sub create a small form
    DrawGrid PicFrmSrc, , , , mnugrid.Checked
    PicFrmSrc.Line (40, 0)-(PicFrmSrc.ScaleWidth - 60, 300), vbHighlight, BF
    PicFrmSrc.Line (0, 0)-(PicFrmSrc.ScaleWidth - 1, 0), vbButtonFace 'top 1
    PicFrmSrc.Line (1, 8)-(PicFrmSrc.ScaleWidth - 8, 8), vbWhite 'top
    PicFrmSrc.Line (0, 0)-(0, PicFrmSrc.ScaleHeight + 1), vbButtonFace 'top left
    PicFrmSrc.Line (8, 8)-(8, PicFrmSrc.ScaleHeight - 8), vbWhite 'top left 1
    PicFrmSrc.Line (8, PicFrmSrc.ScaleHeight - 30)-(PicFrmSrc.ScaleWidth - 30, PicFrmSrc.ScaleHeight - 30), vbButtonShadow     'bottom 1
    PicFrmSrc.Line (0, PicFrmSrc.ScaleHeight - 8)-(PicFrmSrc.ScaleWidth - 1, PicFrmSrc.ScaleHeight - 8), &H404040    'bottom
    PicFrmSrc.Line (PicFrmSrc.ScaleWidth - 30, 8)-(PicFrmSrc.ScaleWidth - 30, PicFrmSrc.ScaleHeight - 20), vbButtonShadow 'right
    PicFrmSrc.Line (PicFrmSrc.ScaleWidth - 8, 0)-(PicFrmSrc.ScaleWidth - 8, PicFrmSrc.ScaleHeight - 8), &H404040 'right
    
    PicFrmSrc.CurrentY = 60
    PicFrmSrc.CurrentX = 60
    PicFrmSrc.FontBold = True
    PicFrmSrc.ForeColor = vbWhite
    PicFrmSrc.Print m_DialogCaption
    PicFrmSrc.Refresh
End Sub

Sub SaveTxt(lpFile As String, lzData As String)
Dim nFile As Long
    nFile = FreeFile
    Open lpFile For Output As #nFile
        Print #nFile, lzData
    Close #nFile
End Sub

Function GenCode() As String
Dim I As Integer, CtrlName As String
Dim StrA As String, StrB As String, StrC As String, StrCaption As String

    'On Error Resume Next
    StrA = ""
    StrC = "<dialog>" & vbCrLf
    StrC = StrC & "height= " & PicFrmSrc.Height & vbCrLf
    StrC = StrC & "width= " & PicFrmSrc.Width & vbCrLf
    StrC = StrC & "backcolor= " & PicFrmSrc.BackColor & vbCrLf
    StrC = StrC & "enabled= " & CBool(PicFrmSrc.Tag) & vbCrLf
    StrC = StrC & "caption= " & Chr(34) & m_DialogCaption & Chr(34) & vbCrLf
    StrC = StrC & "</dialog>" & vbCrLf & vbCrLf
    
    StrC = StrC & "<controls>" & vbCrLf
    
    For I = 0 To frmmain.Controls.Count - 1
        CtrlName = UCase(frmmain.Controls(I).Name)
        
        If CtrlName = "TBUTTON" Or CtrlName = "TLABEL" _
        Or CtrlName = "TEDIT" Or CtrlName = "IMGTMR" Or CtrlName = "LB" _
        Or CtrlName = "CHECKBOX" Then
            
            If frmmain.Controls(I).Index > 0 Then
                ' add the controls
                Select Case CtrlName
                    Case "TBUTTON": StrA = "BUTTON"
                    Case "TLABEL": StrA = "STATIC"
                    Case "TEDIT": StrA = "EDIT"
                    Case "IMGTMR": StrA = "TMR"
                    Case "LB": StrA = "LB"
                    Case "CHECKBOX": StrA = "CHECKBOX"
                End Select
                
                If StrA <> "TMR" Then
                    If Not (StrA = "EDIT" Or StrA = "LB") Then
                        StrCaption = frmmain.Controls(I).Caption
                    Else
                        StrCaption = frmmain.Controls(I).Text
                    End If
                    
                    StrB = frmmain.Controls(I).Top - 315 _
                    & "," & frmmain.Controls(I).Left _
                    & "," & frmmain.Controls(I).Height _
                    & "," & frmmain.Controls(I).Width _
                    & "," & CBool(frmmain.Controls(I).Tag) _
                    & "," & Chr(34) & StrCaption & Chr(34)
                Else
                    StrB = "0," & CBool(frmmain.Controls(I).Tag)
                End If
                
                StrC = StrC & "AddControl " & StrA & "," & StrB & vbCrLf
                StrA = ""
                StrB = ""
                CtrlName = ""
            End If
        End If
    Next
    StrC = StrC & "</controls>"
    I = 0
    GenCode = StrC
    StrC = ""
End Function
Function GetControlAddName() As String

    Select Case UCase(Form_Object.Name)
        Case "TBUTTON": GetControlAddName = "BUTTON"
        Case "TLABEL": GetControlAddName = "STATIC"
        Case "TEDIT": GetControlAddName = "EDIT"
        Case "LB": GetControlAddName = "LB"
        Case "IMGTMR": GetControlAddName = "TMR"
    End Select
    
End Function

Private Sub DoPropWrite()
On Error Resume Next
    
    Select Case UCase(Form_Object.Name)
        Case "TBUTTON", "TLABEL", "TEDIT", "PICFRMSRC", "LB", "CHECKBOX"
            Form_Object.Top = CLng(txtProp(0).Text)
            Form_Object.Left = CLng(txtProp(1).Text)
            Form_Object.Height = CLng(txtProp(2).Text)
            Form_Object.Width = CLng(txtProp(3).Text)
            
            MakeSelection Not UCase(Form_Object.Name) = "PICFRMSRC"
            
            If UCase(Form_Object.Name) = "PICFRMSRC" Then
                m_DialogCaption = txtProp(4).Text
                Exit Sub
            End If

            If UCase(Form_Object.Name) = "TEDIT" Then
                Form_Object.Text = txtProp(4).Text
            Else
                Form_Object.Caption = txtProp(4).Text
            End If
    End Select
    
    
    
End Sub

Private Sub DoPropRead()
Dim I As Integer
On Error Resume Next

    txtProp(LastFocus).SetFocus
    
    For I = 0 To 5
        lbProp(I).Enabled = True
        txtProp(I).Enabled = True
    Next
            
    Select Case UCase(Form_Object.Name)
        Case "TBUTTON", "TLABEL", "TEDIT", "PICFRMSRC", "LB", "CHECKBOX"
            txtProp(0).Text = Form_Object.Top
            txtProp(1).Text = Form_Object.Left
            txtProp(2).Text = Form_Object.Height
            txtProp(3).Text = Form_Object.Width
            
            If UCase(Form_Object.Name) = "TEDIT" Then
                txtProp(4).Text = Form_Object.Text
            ElseIf UCase(Form_Object.Name) = "PICFRMSRC" Then
                txtProp(0).Enabled = False
                txtProp(1).Enabled = False
                lbProp(0).Enabled = False
                lbProp(1).Enabled = False
            ElseIf UCase(Form_Object.Name) = "LB" Then
                lbProp(4).Enabled = False
                txtProp(4).Enabled = False
                txtProp(4).Text = ""
            Else
                txtProp(4).Text = Form_Object.Caption
            End If
            
            txtProp(5).Text = CBool(Form_Object.Tag)
        
        Case "IMGTMR"
            For I = 0 To 4
                lbProp(I).Enabled = False
                txtProp(I).Enabled = False
                txtProp(I).Text = ""
            Next
            txtProp(5).Text = CBool(Form_Object.Tag)
    End Select
    
End Sub

Private Sub Draw3DBorder(PicSrc As PictureBox)
    PicSrc.Cls
    PicSrc.Line (0, 0)-(PicSrc.ScaleWidth - 8, 0), vbWhite
    PicSrc.Line (0, 0)-(0, PicSrc.ScaleHeight - 8), vbWhite
    PicSrc.Line (PicSrc.ScaleWidth - 8, 0)-(PicSrc.ScaleWidth - 8, PicSrc.ScaleHeight - 8), vbButtonShadow
    PicSrc.Line (0, PicSrc.ScaleHeight - 8)-(PicSrc.ScaleWidth - 8, PicSrc.ScaleHeight - 8), vbButtonShadow
End Sub
Private Sub cboProp_Change()
    cboProp.Text = CboTmp
End Sub

Private Sub cboProp_Click()
Dim ObjName As String, ObjIndex As Integer, e_pos As Integer, n_pos As Integer

    CboTmp = cboProp.Text
    
    e_pos = InStr(1, CboTmp, ":", vbBinaryCompare)
    If e_pos > 0 Then n_pos = InStr(e_pos + 1, CboTmp, " ", vbBinaryCompare)
    ' O well I did start on this then gave up I do this next time
    ' as my eyes are starting to hurt after looking at this code for over 2 hours
    If e_pos > 0 And n_pos > 0 Then
        ObjName = UCase(Mid(CboTmp, n_pos + 1, Len(CboTmp) - 1))
        ObjIndex = Mid(CboTmp, e_pos + 1, n_pos - e_pos - 1)
    End If
    
    ObjName = ""
    ObjIndex = -1
    
End Sub

Private Sub DevToolbar1_DevToolBarMouseUp(Button As Integer, Index As Integer, Key As String)
    If Button = 1 Then AddControl Key
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If (isControl And Not inPropList) Then ' see if a control is selected and we are no in the prop list
        FixPositions ' Keep the object within the form area
        ' This code allows you to move the object you know like in VB
        ' select an object and press and CTRL and press an arrow key to move
        If (KeyCode = 40 And Shift) Then
            Form_Object.Top = Form_Object.Top + 15
            MakeSelection True
        ElseIf (KeyCode = 38 And Shift) Then
            Form_Object.Top = Form_Object.Top - 15
            MakeSelection True
        ElseIf (KeyCode = 37 And Shift) Then
            Form_Object.Left = Form_Object.Left - 15
            MakeSelection True
        ElseIf (KeyCode = 39 And Shift) Then
            Form_Object.Left = Form_Object.Left + 15
            MakeSelection True
        End If
        ' Code below is used for copying and pasting a control
        ' same as in VB select a control CTRL+C to copy CTRL+V Paste
        
        If (KeyCode = vbKeyC And Shift) Then
             'copy control name
            PasteCtr = GetControlAddName
        ElseIf (KeyCode = vbKeyV And Shift) Then
            ' add the control using the value of PasteCtr
            AddControl PasteCtr
        End If
        ' delete a control code
        If (KeyCode = 46 Or KeyCode = 8) Then
            Unload Form_Object
            Set Form_Object = Nothing
            MakeSelection False
            Form_MouseDown 1, 0, 0, 0
        End If
        Exit Sub
    Else
        ' if no object is selected and on form and object in PasteCtr then paste control to form
        If (KeyCode = vbKeyV And Shift) And Len(PasteCtr) > 0 Then
            AddControl PasteCtr
        End If
    End If
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicFormHolder_MouseDown 1, 0, 0, 0
End Sub

Private Sub Form_Paint()
    
    DrawStats
End Sub

Private Sub ImgTmr_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Designer 1, ImgTmr(Index), Button, X, Y, True
End Sub

Private Sub ImgTmr_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Designer 2, ImgTmr, Button, X, Y, False
End Sub

Private Sub ImgTmr_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Designer 0, ImgTmr, Button, X, Y, True
End Sub

Private Sub LB_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Designer 1, LB(Index), Button, X, Y, True
End Sub

Private Sub LB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Designer 2, LB, Button, X, Y, False
End Sub

Private Sub LB_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Designer 0, LB, Button, X, Y, False
End Sub

Private Sub lbProp_Click(Index As Integer)
    txtProp_Click Index
End Sub

Private Sub mnuabout_Click()
    MsgBox frmmain.Caption, vbInformation, "About"
End Sub

Private Sub mnucpy_Click()
    Clipboard.Clear
    Clipboard.SetText GenCode
    MsgBox "Code has now been place onto the clipboard.", vbInformation, frmmain.Caption

End Sub

Private Sub mnudesign_Click()
    mnudesign.Checked = Not mnudesign.Checked
    PicBase.Visible = mnudesign.Checked

    If Not PicBase.Visible Then
        PicProp.Top = PicBase.Top
    Else
        PicProp.Top = PicBase.ScaleHeight + 30
    End If
    
End Sub

Private Sub mnuexit_Click()
    Unload frmmain
End Sub

Private Sub mnugrid_Click()
    mnugrid.Checked = Not mnugrid.Checked
    PicFrmSrc_Resize
End Sub

Private Sub mnuopen_Click()
    clsDialog.DlgHwnd = frmmain.hWnd
    clsDialog.DialogTitle = "Open Dialog"
    clsDialog.Filter = "Text Files(*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "Diz Files (*.diz)" + Chr$(0) + "*.diz" + Chr$(0)
    clsDialog.InitialDir = App.Path
    clsDialog.hInst = App.hInstance
    clsDialog.Flags = 0
    clsDialog.ShowOpen
    If clsDialog.CancelError = False Then Exit Sub
    UnloadControls 'unload any controls first
    LoadGUI clsDialog.FileName
End Sub

Private Sub mnuProperties_Click()
    mnuProperties.Checked = Not mnuProperties.Checked
    PicProp.Visible = mnuProperties.Checked
End Sub

Private Sub mnusave_Click()
    clsDialog.DlgHwnd = frmmain.hWnd
    clsDialog.DialogTitle = "Save Dialog"
    clsDialog.Filter = "Text Files(*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "Diz Files (*.diz)" + Chr$(0) + "*.diz" + Chr$(0)
    clsDialog.hInst = App.hInstance
    clsDialog.Flags = 0
    clsDialog.ShowSave
    If clsDialog.CancelError = False Then Exit Sub
    SaveTxt clsDialog.FileName & ".txt", GenCode
End Sub

Private Sub PicFrmSrc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoPropRead
End Sub

Private Sub tEdit_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Designer 1, tEdit(Index), Button, X, Y, True
End Sub

Private Sub tEdit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Designer 2, tEdit, Button, X, Y, False
End Sub

Private Sub tEdit_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Designer 0, tEdit, Button, X, Y, True
End Sub

Private Sub tLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Designer 1, tLabel(Index), Button, X, Y, True
End Sub

Private Sub tLabel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Designer 2, tLabel, Button, X, Y, False
End Sub

Private Sub tLabel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Designer 0, tLabel, Button, X, Y, True
End Sub
Private Sub TButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Designer 1, TButton(Index), Button, X, Y, True
End Sub

Private Sub TButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Designer 2, TButton, Button, X, Y, False
End Sub

Private Sub TButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Designer 0, TButton, Button, X, Y, True
End Sub

Private Sub PicFormHolder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FrmHangle.Visible = False: sh2.Visible = False: MakeSelection False
    For I = 0 To 5
        txtProp(I).Enabled = False
        txtProp(I).Text = ""
        lbProp(I).Enabled = False
    Next
End Sub

Private Sub PicFrmSrc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        MakeSelection False
        Hangle.Left = (PicFrmSrc.Width - Hangle.Width)
        Hangle.Top = (PicFrmSrc.Height - Hangle.Height)
        txtProp(4).Text = m_DialogCaption
        Set Form_Object = PicFrmSrc
        '
        FrmHangle.Visible = True:  sh2.Visible = True
        FrmHangle.Left = (PicFrmSrc.Width + 30): FrmHangle.Top = (PicFrmSrc.Height + 60)
        sh2.Left = PicFrmSrc.Left: sh2.Top = PicFrmSrc.Top: sh2.Width = PicFrmSrc.Width - 16: sh2.Height = PicFrmSrc.Height - 16
    End If
    
End Sub

Private Sub PicFrmSrc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub PicFrmSrc_Resize()
    RedrawForm
End Sub

Private Sub FrmHangle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Hangle_MouseDown Button, Shift, X, Y
End Sub

Private Sub FrmHangle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    If (Button = 1 And ObjCanResize) Then
        FrmHangle.Top = FrmHangle.Top - (ObjY - Y): FrmHangle.Left = FrmHangle.Left - (ObjX - X)
        sh2.Width = FrmHangle.Left - (sh2.Left - 8): sh2.Height = FrmHangle.Top - (sh2.Top - 8)
        Form_Object.Width = (sh2.Width): Form_Object.Height = (sh2.Height)
        If Form_Object.Width <= 1680 Then Form_Object.Width = 1680
        If Form_Object.Height <= 405 Then Form_Object.Height = 405
    End If
End Sub

Private Sub FrmHangle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    If Button = 1 Then
        ObjCanResize = False
        PicFrmSrc_MouseDown Button, Shift, X, Y
    End If
End Sub

Private Sub Hangle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If LCase(Form_Object.Name) = "imgtmr" Then Exit Sub
    If Button = 1 Then ObjCanResize = True: ObjX = X: ObjY = Y
End Sub

Private Sub Hangle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If (Button = 1 And ObjCanResize) Then
        Hangle.Top = Hangle.Top - (ObjY - Y): Hangle.Left = Hangle.Left - (ObjX - X)
        Selection.Width = Hangle.Left - (Selection.Left - 8): Selection.Height = Hangle.Top - (Selection.Top - 8)
    End If
    
End Sub

Private Sub Hangle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
    If Button = 1 Then
        ObjCanResize = False
        Form_Object.Width = Selection.Width - 70
        Form_Object.Height = Selection.Height - 70
        If Form_Object.Width <= 90 Then Form_Object.Width = 90
        If Form_Object.Height <= 90 Then Form_Object.Height = 90
        DoPropRead
    End If
End Sub

Private Sub MakeSelection(mShow As Boolean)

    If mShow Then
        isControl = True
        FrmHangle.Visible = False: sh2.Visible = False
        Selection.Visible = True
        Hangle.Visible = True
        Selection.ZOrder 0
        Hangle.ZOrder 0
        Selection.Top = (Form_Object.Top - 35)
        Selection.Left = (Form_Object.Left - 30)
        Selection.Width = (Form_Object.Width + 70)
        Selection.Height = (Form_Object.Height + 70)
        Hangle.Top = (Form_Object.Top + Selection.Height - 65)
        Hangle.Left = (Form_Object.Left + Selection.Width - 65)
    Else
        isControl = False
        Selection.Visible = False
        Hangle.Visible = False
        DoPropRead
    End If
    
End Sub

Private Sub Designer(Action As Integer, CtrlObj As Object, Button As Integer, X As Single, Y As Single, doSelection As Boolean)

    If Not Button = 1 Then Exit Sub
    
    If Action = 1 Then ' Mouse Down
        If Button = 1 Then
            Set Form_Object = CtrlObj
            Form_Object.ZOrder 0
            ObjX = X: ObjY = Y
            MakeSelection True
            CanObjMove = True
        End If
    ElseIf Action = 2 Then ' MouseMove
        If Button = 1 Then
            Form_Object.Left = Form_Object.Left + (X - ObjX)
            Form_Object.Top = Form_Object.Top + (Y - ObjY)
            MakeSelection True
        End If
    Else ' Mouse up
        DoPropRead
        FixPositions
        inPropList = False
        MakeSelection True
    End If
    
End Sub
Private Sub PicForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub ArangeControl(CtrObj As Object, theForm As PictureBox)
    CtrObj.ZOrder 0
    CtrObj.Visible = True
    CtrObj.Top = (theForm.ScaleHeight - CtrObj.Height) / 2
    CtrObj.Left = (theForm.ScaleWidth - CtrObj.Width) / 2
End Sub

Private Sub AddControl(CtrlName As String)
Dim CtrlCnt As Integer
Dim sName As String

    Select Case CtrlName
        Case "CHECKBOX"
            CtrlCnt = CheckBox.Count
            sName = CheckBox(0).Name
            Load CheckBox(CtrlCnt)
            ArangeControl CheckBox(CtrlCnt), PicFrmSrc
        Case "BUTTON"
            CtrlCnt = TButton.Count
            sName = TButton(0).Name
            Load TButton(CtrlCnt)
            ArangeControl TButton(CtrlCnt), PicFrmSrc
        Case "STATIC"
            CtrlCnt = tLabel.Count
            sName = tLabel(0).Name
            Load tLabel(CtrlCnt)
            ArangeControl tLabel(CtrlCnt), PicFrmSrc
        Case "EDIT"
            CtrlCnt = tEdit.Count
            sName = tEdit(0).Name
            Load tEdit(CtrlCnt)
            ArangeControl tEdit(CtrlCnt), PicFrmSrc
        Case "LB"
            CtrlCnt = LB.Count
            sName = LB(0).Name
            Load LB(CtrlCnt)
            ArangeControl LB(CtrlCnt), PicFrmSrc
        Case "TMR"
            CtrlCnt = ImgTmr.Count
            sName = ImgTmr(0).Name
            Load ImgTmr(CtrlCnt)
            ArangeControl ImgTmr(CtrlCnt), PicFrmSrc
        End Select
        
         cboProp.AddItem CtrlName & ":" & CtrlCnt & "  " & sName
        cboProp.ListIndex = cboProp.ListCount - 1
        
End Sub

Sub Hidebutton()
    DevToolbar1.HideFocus
End Sub

Public Sub DrawGrid(FrmPic As PictureBox, Optional dmScaleX As Integer = 120, Optional dmScaleY As Integer = 120, Optional nGridCol As Long = vbDesktop, Optional mDraw As Boolean = True)
Dim I As Long, j As Long

    If Not mDraw Then PicFrmSrc.Cls: Exit Sub
    
    For I = 40 To FrmPic.ScaleWidth Step dmScaleX
        For j = 310 To FrmPic.ScaleHeight Step dmScaleY
        FrmPic.PSet (I, j), nGridCol
    Next
    
        Next
        FrmPic.Refresh
End Sub

Sub ToolBarAddButtons()

    DevToolbar1.SetupControlBar
    DevToolbar1.AddButton "Command Button", 2, "BUTTON"
    DevToolbar1.AddButton "Static Label", 3, "STATIC"
    DevToolbar1.AddButton "EditBox", 4, "EDIT"
    DevToolbar1.AddButton "Timer", 9, "TMR"
    DevToolbar1.AddButton "ListBox", 1, "LB"
    DevToolbar1.AddButton "CHECKBOX", 8, "CHECKBOX"
    DevToolbar1.DrawToolBar
End Sub

Private Sub DevToolbar1_DevToolBarMouseDown(Button As Integer, Index As Integer, Key As String)
    On Error Resume Next
   ' cboProp.ListIndex = cboProp.ListCount - 1
End Sub

Private Sub Form_Load()

    m_DialogCaption = "Dialog"
    
    Picbar.Width = (PicBase.ScaleWidth - 30)
    PicA.Width = Picbar.Width
    cboProp.Width = PicA.ScaleWidth + 6
    
    DevToolbar1.Width = (PicBase.ScaleWidth - DevToolbar1.Left - 30)
    PicBase.Height = (DevToolbar1.Height + DevToolbar1.Top + 10)
    
    DevToolbar1.ResetButton
    Call ToolBarAddButtons

    mnugrid_Click
    
    PicFrmSrc_MouseDown 1, 0, 0, 0
    DoPropRead
    Draw3DBorder PicBase
    Me.KeyPreview = True
    PicProp.Top = PicBase.ScaleHeight + 30
    PicProp.Height = PicForm.ScaleHeight + PicProp.Top
    Draw3DBorder PicProp
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Hidebutton
End Sub

Private Sub Form_Resize()
    PicForm.Width = (frmmain.ScaleWidth)
    PicForm.Height = (frmmain.ScaleHeight)
    PicFormHolder.Width = PicForm.Width
    PicFormHolder.Height = PicForm.Height
    DrawStats
    PicProp.Height = PicForm.ScaleHeight + PicProp.Top
    Draw3DBorder PicProp
End Sub

Private Sub Picbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Hidebutton
End Sub

Private Sub PicBase_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Hidebutton
End Sub

Private Sub txtProp_Click(Index As Integer)
    LastFocus = Index
    inPropList = True
    If Index = 5 Then
        txtProp(5).Text = Not CBool(txtProp(5).Text)
        txtProp(5).SelStart = 5
        Form_Object.Tag = Abs(CBool(txtProp(5).Text))
    End If
    
    DoPropWrite
End Sub

Private Sub txtProp_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If Index = 4 Then
        If KeyAscii = 34 Then KeyAscii = 0
        If KeyAscii = 13 Then DoPropWrite: RedrawForm: KeyAscii = 0
        Exit Sub
    End If
    
    If Index <> 4 Then
        Select Case KeyAscii
            Case 8, 43, 45, 47 To 57
            Case 13
                DoPropWrite
                KeyAscii = 0
            Case Else
                KeyAscii = 0
        End Select
        Exit Sub
    End If
    
End Sub

Private Sub txtProp_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    FixPositions
End Sub

Private Sub txtProp_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoPropWrite
End Sub
