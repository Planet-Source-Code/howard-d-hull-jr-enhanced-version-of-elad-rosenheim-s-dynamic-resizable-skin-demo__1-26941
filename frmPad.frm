VERSION 5.00
Begin VB.Form frmPad 
   Caption         =   "Dynamic Resizable Skinned Form"
   ClientHeight    =   5295
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6660
   LinkTopic       =   "Form2"
   MouseIcon       =   "frmPad.frx":0000
   ScaleHeight     =   5295
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picClientArea 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4035
      Left            =   360
      ScaleHeight     =   4005
      ScaleWidth      =   5055
      TabIndex        =   14
      Top             =   270
      Width           =   5085
      Begin VB.TextBox txtMaximumFormHeight 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2910
         TabIndex        =   9
         Text            =   "200"
         Top             =   2700
         Width           =   1275
      End
      Begin VB.TextBox txtMaximumFormWidth 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2910
         TabIndex        =   11
         Text            =   "350"
         Top             =   3060
         Width           =   1275
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Apply"
         Enabled         =   0   'False
         Height          =   435
         Left            =   3000
         TabIndex        =   12
         Top             =   3450
         Width           =   1185
      End
      Begin VB.TextBox txtMinimumClientAreaWidth 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2910
         TabIndex        =   7
         Text            =   "350"
         Top             =   2340
         Width           =   1275
      End
      Begin VB.TextBox txtMinimumClientAreaHeight 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2910
         TabIndex        =   5
         Text            =   "260"
         Top             =   1980
         Width           =   1275
      End
      Begin VB.CheckBox chkAlwaysOnTop 
         Caption         =   "Always on top"
         Height          =   225
         Left            =   2670
         TabIndex        =   3
         Top             =   1230
         Width           =   1515
      End
      Begin VB.CheckBox chkResizable 
         Caption         =   "Resizable"
         Height          =   225
         Left            =   2670
         TabIndex        =   2
         Top             =   810
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin VB.CheckBox chkMovable 
         Caption         =   "Movable"
         Height          =   225
         Left            =   2670
         TabIndex        =   1
         Top             =   390
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin VB.ListBox lstSkins 
         Height          =   1230
         ItemData        =   "frmPad.frx":08CA
         Left            =   120
         List            =   "frmPad.frx":08CC
         TabIndex        =   0
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblLoadTime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1650
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "MaximumFormHeight"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   2730
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "MaximumFormWidth"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   3090
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "MinimumClientAreaWidth"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   2370
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "MinimumClientAreaHeight"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   2010
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Select a Skin:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Menu mnuPopUP 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuSelectSkin 
         Caption         =   "&Select Skin"
         Begin VB.Menu mnuSkin 
            Caption         =   "{Skins}"
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "frmPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'// Skin Form Class
Private WithEvents m_SkinnedForm     As clsSkinnedForm
Attribute m_SkinnedForm.VB_VarHelpID = -1

Private Sub chkAlwaysOnTop_Click()
    m_SkinnedForm.AlwaysOnTop = chkAlwaysOnTop.Value
End Sub

Private Sub chkMovable_Click()
    m_SkinnedForm.Movable = chkMovable.Value
End Sub

Private Sub chkResizable_Click()
    m_SkinnedForm.Resizable = chkResizable.Value
End Sub

Private Sub cmdApply_Click()
    m_SkinnedForm.MinimumClientAreaHeight = Val(txtMinimumClientAreaHeight)
    m_SkinnedForm.MinimumClientAreaWidth = Val(txtMinimumClientAreaWidth)
    m_SkinnedForm.MaximumFormHeight = Val(txtMaximumFormHeight)
    m_SkinnedForm.MaximumFormWidth = Val(txtMaximumFormWidth)
    cmdApply.Enabled = False
End Sub

Private Sub Form_Load()
'
    '// Skin Form
    Set m_SkinnedForm = New clsSkinnedForm
    Set m_SkinnedForm.SkinnedForm = Me
    Set m_SkinnedForm.ClientArea = picClientArea
    m_SkinnedForm.MinimumClientAreaHeight = 260
    m_SkinnedForm.MinimumClientAreaWidth = 350
    m_SkinnedForm.MaximumFormHeight = (Screen.Height / 2) / Screen.TwipsPerPixelY
    m_SkinnedForm.MaximumFormWidth = (Screen.Width / 2) / Screen.TwipsPerPixelX
    
    m_SkinnedForm.SkinForm ("Default")
        
    txtMaximumFormHeight = (Screen.Height / 2) / Screen.TwipsPerPixelY
    txtMaximumFormWidth = (Screen.Width / 2) / Screen.TwipsPerPixelX
    
'    m_SkinnedForm.Width = 392
'    m_SkinnedForm.Height = 260
    
    '// Populate list woth available skins
    ListSkins
    
    '// Disable the Apply button
    Me.cmdApply.Enabled = False
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim f       As Form

    '// Clean up
    Set m_SkinnedForm = Nothing
    
    '// Unload forms
    For Each f In Forms
        Unload f
        Set f = Nothing
    Next

End Sub


' Fill the list of skins.
' Actually it's a list of directories under App.Path
Private Sub ListSkins()
Dim CurrSkinName As String, SkinPos As Long
Dim i As Long

    CurrSkinName = Dir(App.Path & "\Skins\", vbDirectory)

    Do While CurrSkinName <> ""

        If CurrSkinName <> "." And CurrSkinName <> ".." Then
            If (GetAttr(App.Path & "\Skins\" & CurrSkinName) And vbDirectory) Then
                lstSkins.AddItem CurrSkinName
                
                If i > 0 Then Load mnuSkin(i)
                mnuSkin(i).Caption = CurrSkinName

                i = i + 1
            End If
        End If

        CurrSkinName = Dir()
    Loop

End Sub


Private Sub lstSkins_Click()
    lblLoadTime.Caption = "Skin loaded in " & Format(m_SkinnedForm.SkinForm(lstSkins.Text) * 1000, "###.000") & " milliseconds"
End Sub

Private Sub m_SkinnedForm_ClientAreaMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuPopUP, , picClientArea.Left + X, picClientArea.Top + Y
    End If
End Sub

Private Sub m_SkinnedForm_Minimize()
'// The Minimize button was clicked and the form was minimized

'// Insert any code you may want to do After the form gets minimized

End Sub

Private Sub m_SkinnedForm_Resize()
'// Event triggered to allow user to rearrange controls in the picClientArea.
    
End Sub

Private Sub m_SkinnedForm_SkinningComplete()
    '// Call routine to change the colors / fonts of the control on the form
    Call UpdateControls
    m_SkinnedForm_Resize

End Sub

Private Sub m_SkinnedForm_Unload()
'// The Close button was clicked
Dim f           As Form

    '// Start Unloading
    For Each f In Forms
        Unload f
        Set f = Nothing
    Next

End Sub
'//---------------------------------------------------------------------------------
'// Edit to suit your need. This is a quick and dirty solution
'// To update the colors and font properties of all controls on the form.
'//---------------------------------------------------------------------------------
Private Sub UpdateControls()
Dim cntrl           As Control

'// By default the ForeColor / BackColor / Font properties
'// of the Form and the Client picture box are updated in the Class
'// So we don't need to do anything with them. The rest of the controls
'// can be changed if wanted.


On Error Resume Next
    
    '// Cycle through each control and set the Colors, etc...
    For Each cntrl In m_SkinnedForm.SkinnedForm.Controls
        cntrl.ForeColor = m_SkinnedForm.FontColor
        cntrl.FontBold = m_SkinnedForm.FontBold
        cntrl.BackColor = m_SkinnedForm.BackColor
        cntrl.Refresh
    Next

End Sub

Private Sub mnuSkin_Click(Index As Integer)
    m_SkinnedForm.SkinForm mnuSkin(Index).Caption
    mnuSkin(Index).Checked = True
End Sub

Private Sub txtMaximumFormHeight_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtMaximumFormWidth_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtMinimumClientAreaHeight_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtMinimumClientAreaWidth_Change()
    cmdApply.Enabled = True
End Sub
