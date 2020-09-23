VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSkinnedForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

 ' Current number of horizontal/vertical segments
Dim NumXSlices As Long
Dim NumYSlices As Long

' Minimum needed number of X slices so we don't mess-up
' the button positions
Dim MinXSlices As Long
Dim MinYSlices As Long
                        
' Maximum possible number of X slices so we don't mess-up
' the button positions or go off screen
Dim MaxXSlices As Long
Dim MaxYSlices As Long
                        
' Width/height of pad w/o any horizontal segment
Dim BaseXSize As Long
Dim BaseYSize As Long

' Used when resizing the window -
' X/Y distance of the mouse pointer from the form's edge
Dim XDistance As Long
Dim YDistance As Long

' Used when resizing to keen the Client Area at least the size of
' the MinimumClientAreaWidth and MinimumClientAreaHeight properties
Dim OffsetXSize     As Long
Dim OffsetYSize     As Long
Dim XEdgeWidth      As Long
Dim YEdgeHeight     As Long

' Boolean flags - the current state of the form
Dim InXDrag As Boolean ' In horizontal resize
Dim InYDrag As Boolean ' In vertical resize
Dim InFormDrag As Boolean ' In window drag

Dim NoRedraw As Boolean

' Set to TRUE when in ListSkins(), to prevent lstSkins_Click()
' events from being handled while the list is created
Dim InListSkins As Boolean

' Size of right/bottom segments
Dim XEdgeSize As Single
Dim YEdgeSize As Single

' Handler for window dragging & docking
Dim DockHandler As New clsDockingHandler

' Holds the actual edge skin bitmaps
Dim EdgeImages(FE_LAST) As clsBitmap

' Holds the region data for each of the skin bitmaps
Dim EdgeRegions(FE_LAST) As RegionDataType

Dim WindowRegion As Long ' Current window region

' Custom Exit/Minimize buttons
Dim MyExitButton As New clsButton
Dim MyMinButton As New clsButton
 
'// Boolean values to determine which button was clicked
Dim bExitButtonClicked              As Boolean
Dim bMinButtonClicked               As Boolean

' Default size of client area. Used to compute the number of
' x/y segments needed when the program is loaded
'Const DEFAULT_CLIENT_SIZE = 250

'// Form to skin and the ClientArea picturebox objects
'// We use the WithEvents keyword to capture all the events within the class
Public WithEvents SkinnedForm       As Form
Attribute SkinnedForm.VB_VarHelpID = -1
Public WithEvents ClientArea        As PictureBox
Attribute ClientArea.VB_VarHelpID = -1

'// Private Variables
Private nMaxFormWidth               As Long             'The SkinnedForm's maximum allowable width
Private nMaxFormHeight              As Long             'The SkinnedForm's maximum allowable height
Private nMinimumClientAreaWidth     As Long             'The ClientArea's minimum allowable width
Private nMinimumClientAreaHeight    As Long             'The ClientArea's minimum allowable height
Private nFormWidth                  As Long             'The SkinnedForm's width
Private nFormHeight                 As Long             'The SkinnedForm's height
Private bFormResizable              As Boolean          'Is the SkinnedForm Resizable?
Private bFormMovable                As Boolean          'Is the SkinnedForm movable?
Private bFormSkinned                As Boolean          'Flag to designate that the Form has been skinned
Private bBorderStyleRemoved         As Boolean          'See if the BorderStyle has been removed

'// public Events
Public Event Resize()
Public Event Unload()
Public Event Minimize()
Public Event SkinningComplete()
Public Event FormMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event FormMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event FormMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ClientAreaMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ClientAreaMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ClientAreaMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


'//------------------------------------------------------------------
'// Set/Return whether the SkinnedForm should be Locked in the TopMost position
'//------------------------------------------------------------------
Public Property Let AlwaysOnTop(vData As Boolean)
    '// Pass to Dockhandler Code
    DockHandler.AlwaysOnTop = vData
End Property

Public Property Get AlwaysOnTop() As Boolean
    '// Get property from Dockhandler class
    AlwaysOnTop = DockHandler.AlwaysOnTop
End Property
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Set/Return BackColor of the Form and Client area
'//------------------------------------------------------------------
Public Property Get BackColor() As OLE_COLOR
    BackColor = SkinPrefs.BackColor
End Property

Public Property Let BackColor(ByVal vData As OLE_COLOR)
    SkinPrefs.BackColor = vData
End Property
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Set/Return whether the Fonts used in the Form/ClientArea should be Bold
'//------------------------------------------------------------------
Public Property Get FontBold() As Boolean
    FontBold = SkinPrefs.FontBold
End Property

Public Property Let FontBold(ByVal vData As Boolean)
    SkinPrefs.FontBold = vData
End Property
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Set/Return Font Color of the Form and Client Area
'//------------------------------------------------------------------
Public Property Get FontColor() As OLE_COLOR
    FontColor = SkinPrefs.FontColor
End Property

Public Property Let FontColor(ByVal vData As OLE_COLOR)
    SkinPrefs.FontColor = vData
End Property
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Set/Return Size of the font on the Form and Client area
'//------------------------------------------------------------------
Public Property Get FontSize() As Long
    FontSize = SkinPrefs.FontSize
End Property

Public Property Let FontSize(ByVal vData As Long)
    SkinPrefs.FontSize = vData
End Property
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Return / Set the Height of the Form
'//------------------------------------------------------------------
Public Property Get Height() As Long
    Height = SkinnedForm.Height
End Property

'//------------------------------------------------------------------
'// Set the form's Height
'//------------------------------------------------------------------
Public Property Let Height(ByVal nHeight As Long)
    SkinnedForm.Height = (nHeight * Screen.TwipsPerPixelY)
    If bFormSkinned Then
        NumYSlices = Round((nHeight - BaseYSize) / EdgeImages(FE_LEFT_V_SEGMENT).Height)
        SetPadSize True
    End If
End Property
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Set the Maximum Height of the Form
'//------------------------------------------------------------------
Public Property Let MaximumFormHeight(ByVal nNewHeight As Long)
    '// Validate that the property is larger than 100 pixels and smaller than the Screen.Height
    If nNewHeight >= 100 And (nNewHeight <= (Screen.Height / Screen.TwipsPerPixelY)) Then
        nMaxFormHeight = nNewHeight
    Else
        '// Invalid property
        Exit Property
    End If
    
    '// Has the class finished apply the Skin
    If bFormSkinned Then
        '// Recalculate then Minumum number of Y Slices based on the new value
        MaxYSlices = FindMaxYSlices()
    
        '// Redraw if the maximum number of Y Slices is greater
        '// than the current number being used
        If MaxYSlices < NumYSlices Then SetPadSize True
    End If
End Property
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Set the Maximum Width of the Form
'//------------------------------------------------------------------
Public Property Let MaximumFormWidth(ByVal nNewWidth As Long)
    '// Validate that the property is larger than 100 pixels and smaller than the Screen.Width
    If nNewWidth >= 100 And (nNewWidth <= (Screen.Width / Screen.TwipsPerPixelX)) Then
        nMaxFormWidth = nNewWidth
    Else
        '// Invalid property
        Exit Property
    End If
    
    '// Has the class finished apply the Skin
    If bFormSkinned Then
        '// Recalculate then Minumum number of X Slices based on the new value
        MaxXSlices = FindMaxXSlices()
    
        '// Redraw if the maximum number of X Slices is less
        '// than the current number being used
        If MaxXSlices < NumXSlices Then SetPadSize True
    End If
End Property
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Set the Minimum Width of the ClientArea. This will resize the
'// Form so that MinimumClientAreaHeight will always be visible
'// regardless of the size of the Edges
'//------------------------------------------------------------------
Public Property Let MinimumClientAreaHeight(ByVal nNewHeight As Long)
    nMinimumClientAreaHeight = nNewHeight

    '// Has the class finished apply the Skin
    If bFormSkinned Then
        '// Recalculate then Minumum number of Y Slices based on the new value
        MinYSlices = FindMinYSlices()

      '// Redraw if the minumum number of Y Slices is greater
      '// than the current number being used
      If MinYSlices > NumYSlices Then SetPadSize True
    End If
   
End Property
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Set the Minimum Width of the ClientArea. This will resize the
'// Form so that MinimumClientAreaWidth will always be visible
'// regardless of the size of the Edges
'//------------------------------------------------------------------
Public Property Let MinimumClientAreaWidth(ByVal nNewWidth As Long)
    nMinimumClientAreaWidth = nNewWidth

    '// Has the class finished apply the Skin
    If bFormSkinned Then
        '// Recalculate then Minumum number of X Slices based on the new value
        MinXSlices = FindMinXSlices()
    
        '// Redraw if the minumum number of X Slices is greater
        '// than the current number being used
        If MinXSlices > NumXSlices Then SetPadSize True
    End If
    
End Property
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Return / Set whether the Form is Movable
'//------------------------------------------------------------------
Public Property Get Movable() As Boolean
    Movable = bFormMovable
End Property

Public Property Let Movable(ByVal bMovable As Boolean)
    bFormMovable = bMovable
End Property
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Return / Set whether the Form is Resizable
'//------------------------------------------------------------------
Public Property Get Resizable() As Boolean
    Resizable = bFormResizable
End Property

Public Property Let Resizable(ByVal bResizable As Boolean)
    bFormResizable = bResizable
End Property
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Return / Set the Width of the Form
'//------------------------------------------------------------------
Public Property Get Width() As Long
    Width = (SkinnedForm.Width / Screen.TwipsPerPixelX)
End Property

Public Property Let Width(ByVal nWidth As Long)
    SkinnedForm.Width = (nWidth * Screen.TwipsPerPixelX)
    If bFormSkinned Then
        NumXSlices = Round((nWidth - BaseXSize) / EdgeImages(FE_TOP_H_SEGMENT).Width)
        SetPadSize True
    End If
End Property
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Public function to start Skinning the Form
'//------------------------------------------------------------------
Public Function SkinForm(ByVal sSkinName As String) As Single
On Error GoTo AttemptToLoadSkin_ErrHandler
    
    '// Validate that a Form and Client Area are set.
    If (SkinnedForm Is Nothing) Then
        Err.Raise -1, "SkinnedForm Class", "The SkinnedForm property needs to be populated!"
    ElseIf (ClientArea Is Nothing) Then
        Err.Raise -2, "SkinnedForm Class", "The ClientArea property needs to be populated!"
    End If
    
    SkinForm = Timer
    
    '// Lock Form
    LockWindowUpdate SkinnedForm.hwnd
    
    '// Disable the Title bar
    DisableTitleBar SkinnedForm.hwnd
    
    '// Disable Docking
    Set DockHandler = Nothing
    
    '// Enable Docking
    Set DockHandler.ParentForm = SkinnedForm
    
    '// Get Settings
    CurrPrefs.SkinName = sSkinName
    ReadSkinPreferences
    
    '// Change Scale Modes
    SkinnedForm.ScaleMode = vbPixels
    ClientArea.ScaleMode = vbPixels
    
    '// Load the skin
    LoadSkin
    
    '// Draw the skin
    SetPadSize True
    
    '// UnLock Form
    LockWindowUpdate 0

    '// Set flag
    bFormSkinned = True
    
    SkinForm = Timer - SkinForm
    
    '// Trigger Event
    RaiseEvent SkinningComplete
    
Exit Function

AttemptToLoadSkin_ErrHandler:
    '// UnLock Form
    LockWindowUpdate 0
    '// Display Error
    MsgBox "Unable to load skin. " & vbCrLf & "Reason: " & Err.Description, vbCritical, "Fatal Error"
    '//Reset Flag
    bFormSkinned = False
    '// Return value
    SkinForm = -1
End Function
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Initialization of the Skin Class
'//------------------------------------------------------------------
Private Sub Class_Initialize()

    '// Load Filenames
    InitEdgeFileNames
    
    '// Setup Initial Variables
    CurrPrefs.SkinName = "Default"
    CurrPrefs.SkinsPath = App.Path + "\Skins\"

    '// Set Default Values
    nMaxFormHeight = (Screen.Height / Screen.TwipsPerPixelY)
    nMaxFormWidth = (Screen.Width / Screen.TwipsPerPixelX)
    
    '// Defaults
    bFormResizable = True
    bFormMovable = True
    
End Sub
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// The user Pressed the mouse down on the client area
'//------------------------------------------------------------------
Private Sub ClientArea_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '// Raise Event
    RaiseEvent ClientAreaMouseDown(Button, Shift, X, Y)
End Sub

'//------------------------------------------------------------------
'// The mouse was moved over the Client Area picturebox.
'// Reset the MousePointer to clear any ResizeCursors
'//------------------------------------------------------------------
Private Sub ClientArea_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SkinnedForm.MousePointer = vbCustom
    
    '// Raise Event
    RaiseEvent ClientAreaMouseMove(Button, Shift, X, Y)
    
End Sub
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// The user released the mouse over the Client Area
'//------------------------------------------------------------------
Private Sub ClientArea_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '// Raise Event
    RaiseEvent ClientAreaMouseUp(Button, Shift, X, Y)
End Sub
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// A mouse button press may initiate form dragging or resizing
'//------------------------------------------------------------------
Private Sub SkinnedForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '// Exit sub if the form is not skinned
    If Not bFormSkinned Then Exit Sub
    
    '// Determine which button is pressed
    If Button = vbLeftButton Then
        
        '// Reset ButtonClick values
        bExitButtonClicked = False
        bMinButtonClicked = False
        
        ' Test whether the user has pressed a "button",
        ' and show the 'down button' image if so
        If MyExitButton.HitTest(CLng(X), CLng(Y)) Then
            MyExitButton.PaintDownImage
            bExitButtonClicked = True
            Exit Sub
        
        ElseIf MyMinButton.HitTest(CLng(X), CLng(Y)) Then
            MyMinButton.PaintDownImage
            bMinButtonClicked = True
            Exit Sub
            
        End If
    
        YDistance = Y - SkinnedForm.ScaleHeight
        XDistance = X - SkinnedForm.ScaleWidth
        
        ' If the mouse pointer is on the the bottom edge,
        ' flag Y (vertical) drag
        If Abs(YDistance) < YEdgeSize Then
            InYDrag = True
        End If
        
        ' If the mouse pointer is on the the right edge,
        ' flag X drag. Don't start drag if wer'e in the window
        ' title area
        If Abs(XDistance) < XEdgeSize And _
           Y > EdgeImages(FE_TOP_RIGHT).Height Then
            InXDrag = True
        End If
        
        ' If we're in the window title area, start form draggin'
        If (Y <= EdgeImages(FE_TOP_H_SEGMENT).Height) Then
            '// See If the Form is Movable
            If bFormMovable Then
                DockHandler.StartDockDrag X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
                InFormDrag = True
            End If
        End If
    
    End If

    '// Passback the Event
    RaiseEvent FormMouseDown(Button, Shift, X, Y)
    
End Sub
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// The mouse is moving over the form. Set appropriate MousePointer.
'// Move form if the mouse is pressed.
'//------------------------------------------------------------------
Private Sub SkinnedForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim NewYSlices As Single
Dim NewXSlices As Single
Dim ShowXResizeCursor As Boolean
Dim ShowYResizeCursor As Boolean
Dim ResizingNeeded As Boolean

    '// Exit sub if the form is not skinned
    If Not bFormSkinned Then Exit Sub
    
    If InFormDrag Then
        ' Continue window draggin'
        DockHandler.UpdateDockDrag X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
        Exit Sub
    End If
    
    '// Exit if Form is not resizable
    If Not bFormResizable Then Exit Sub
    
    ' Determine what kind of cursor should be shown
    If Abs(Y - SkinnedForm.ScaleHeight) < YEdgeSize Or InYDrag Then
        ShowYResizeCursor = True
    End If
    
    If (Abs(X - SkinnedForm.ScaleWidth) < XEdgeSize And Y > EdgeImages(FE_TOP_RIGHT).Height) Or InXDrag Then
        ShowXResizeCursor = True
    End If
    
    If ShowXResizeCursor And ShowYResizeCursor Then
        SkinnedForm.MousePointer = vbSizeNWSE
        
    ElseIf ShowXResizeCursor Then
        SkinnedForm.MousePointer = vbSizeWE
    
    ElseIf ShowYResizeCursor Then
        SkinnedForm.MousePointer = vbSizeNS
    
    Else
        SkinnedForm.MousePointer = vbCustom
        
    End If

    If InXDrag Then
        ' Compute new number of horizontal segments
        NewXSlices = (X - BaseXSize - XDistance) / EdgeImages(FE_TOP_H_SEGMENT).Width
        If NewXSlices < MinXSlices Then NewXSlices = MinXSlices
        
        ' Check if we should actually do the resize. Not every
        ' slightest mouse drag should cause a resize
        If (NewXSlices - NumXSlices >= 0.5) Or _
           (NewXSlices - NumXSlices < -0.5) Then
            
            NumXSlices = NewXSlices
            ResizingNeeded = True
        End If
    End If

    ' Same handling for vertical resize-drag
    If InYDrag Then
        
        NewYSlices = (Y - BaseYSize - YDistance) / EdgeImages(FE_LEFT_V_SEGMENT).Height
        If NewYSlices < 0 Then NewYSlices = 0
        
        If NewYSlices - NumYSlices >= 0.5 Or _
           (NewYSlices - NumYSlices < -0.5) Then
            
            NumYSlices = NewYSlices
            ResizingNeeded = True
        End If
    End If

    If ResizingNeeded Then SetPadSize
    
    '// Raise Event
    RaiseEvent FormMouseMove(Button, Shift, X, Y)
    
End Sub
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// The mouse was released.
'//------------------------------------------------------------------
Private Sub SkinnedForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '// Exit sub if the form is not skinned
    If Not bFormSkinned Then Exit Sub
    
    MyExitButton.PaintUpImage
    MyMinButton.PaintUpImage

    ' Test whether the user has released a "button",
    ' and commit the appropriate action if so
    ' // and that the button was released on the same button as the one clicked
    If (MyExitButton.HitTest(CLng(X), CLng(Y))) And (bExitButtonClicked) Then
        RaiseEvent Unload
        
    ElseIf (MyMinButton.HitTest(CLng(X), CLng(Y))) And (bMinButtonClicked) Then
        '// This will cause our form too to minimize
        SkinnedForm.WindowState = vbMinimized
        RaiseEvent Minimize
        
    End If
   
    ' Clear window dragging/resizing flags
    InXDrag = False
    InYDrag = False
    InFormDrag = False

    '// Raise Event
    RaiseEvent FormMouseUp(Button, Shift, X, Y)
    
End Sub
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Routine to reset/load the Edge bitmap images.
'// Change the BackColor of the Form and ClientArea
'// Move the ClientArea on the form based on the size of the edges
'// Init the Min and Close button images
'// Set the form's MousePointer if there is one
'//------------------------------------------------------------------
Private Sub LoadSkin()
Dim i               As Long
Dim sFileName       As String
Dim PrevXSliceSize  As Long
Dim PrevYSliceSize  As Long

    ' Save for later. You'll see.
    'If Not EdgeImages(0) Is Nothing Then
    '    PrevXSliceSize = EdgeImages(FE_TOP_H_SEGMENT).Width
    '    PrevYSliceSize = EdgeImages(FE_LEFT_V_SEGMENT).Height
    'End If
    
    ' Initialize bitmaps array
    For i = 0 To FE_LAST
        Set EdgeImages(i) = New clsBitmap
    Next
    
    ' Load skin bitmaps. Check that the files actally  exist
    For i = 0 To FE_LAST
        sFileName = CurrPrefs.SkinFullPath & EdgeImageFileNames(i)
        
        If Dir(sFileName) = "" Then
            Err.Raise 1, , "Image file " & sFileName & " not found!"
                        
        ElseIf EdgeImages(i).LoadFile(sFileName) = False Then
            Err.Raise 1, , "Could not load image file: " & sFileName
            
        End If
    Next
    
    ' Set back color according to skin's definition, to match
    ' the skin's "look"
    SkinnedForm.BackColor = SkinPrefs.BackColor
    ClientArea.BackColor = SkinPrefs.BackColor
    
    ' See documentation in start of file for all those variables
    BaseXSize = EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_RIGHT).Width
    BaseYSize = EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_BOTTOM_LEFT).Height

    XEdgeSize = EdgeImages(FE_RIGHT_V_SEGMENT).Width
    YEdgeSize = EdgeImages(FE_BOTTOM_H_SEGMENT).Height

    XEdgeWidth = EdgeImages(FE_TOP_H_SEGMENT).Width
    YEdgeHeight = EdgeImages(FE_LEFT_V_SEGMENT).Height
    
    'OffsetXSize = EdgeImages(FE_RIGHT_V_SEGMENT).Width + EdgeImages(FE_LEFT_V_SEGMENT).Width
    'OffsetYSize = EdgeImages(FE_BOTTOM_H_SEGMENT).Height + EdgeImages(FE_TOP_H_SEGMENT).Height
    
    OffsetXSize = EdgeImages(FE_LEFT_V_SEGMENT).Width + EdgeImages(FE_RIGHT_V_SEGMENT).Width
    OffsetYSize = EdgeImages(FE_TOP_H_SEGMENT).Height + EdgeImages(FE_BOTTOM_H_SEGMENT).Height
    
    ' Here we compute how much horizontal/vertical segments
    ' sould be drawn
    If PrevXSliceSize <> 0 Then
        ' Skin was changed, match number of x/y slices
        ' according to the currect/previous sizes of the slices
        NumXSlices = Round(NumXSlices * PrevXSliceSize / EdgeImages(FE_TOP_H_SEGMENT).Width)
        NumYSlices = Round(NumYSlices * PrevYSliceSize / EdgeImages(FE_LEFT_V_SEGMENT).Height)
    
    Else
        ' Program was just loaded, match number of x/y slices
        ' to the form's current width/height
        NumXSlices = Round(((SkinnedForm.Width / Screen.TwipsPerPixelX) - BaseXSize) / EdgeImages(FE_TOP_H_SEGMENT).Width)
        NumYSlices = Round(((SkinnedForm.Height / Screen.TwipsPerPixelY) - BaseYSize) / EdgeImages(FE_LEFT_V_SEGMENT).Height)
        'NumXSlices = Round((SkinnedForm.Width / Screen.TwipsPerPixelX) / EdgeImages(FE_TOP_H_SEGMENT).Width)
        'NumYSlices = Round((SkinnedForm.Height / Screen.TwipsPerPixelY) / EdgeImages(FE_LEFT_V_SEGMENT).Height)
    
    End If
    
    ' Position Client Area
    ClientArea.Top = EdgeImages(FE_TOP_LEFT).Height
    ClientArea.left = EdgeImages(FE_LEFT_V_SEGMENT).Width

    ' Initialize exit/minimize buttons
    MyExitButton.Init _
       CurrPrefs.SkinFullPath & "exitbutton_up.bmp", _
       CurrPrefs.SkinFullPath & "exitbutton_down.bmp", _
       SkinPrefs.ExitButtonX, SkinPrefs.ExitButtonY, _
       SkinnedForm

    MyMinButton.Init _
       CurrPrefs.SkinFullPath & "minbutton_up.bmp", _
       CurrPrefs.SkinFullPath & "minbutton_down.bmp", _
       SkinPrefs.MinButtonX, SkinPrefs.MinButtonY, _
       SkinnedForm

    ' Limit minimum number of X slices, in order to allow the
    ' buttons to be drawn correctly
    MinXSlices = FindMinXSlices()
    MaxXSlices = FindMaxXSlices()
    NumXSlices = IIf(MinXSlices > NumXSlices, MinXSlices, NumXSlices)
    NumXSlices = IIf(NumXSlices > MaxXSlices, MaxXSlices, NumXSlices)
        
    ' Limit minimum number of Y slices, in order to allow the
    ' Controls on the form to be drawn correctly
    MinYSlices = FindMinYSlices()
    MaxYSlices = FindMaxYSlices()
    NumYSlices = IIf(MinYSlices > NumYSlices, MinYSlices, NumYSlices)
    NumYSlices = IIf(NumYSlices > MaxYSlices, MaxYSlices, NumYSlices)

    ' Create and store region data for each of the skin bitmaps,
    ' for use whenever creating the window region
    Dim LoadedRegionsFromFile As Boolean
    
    ' If the 'load region data from file' box is checked, try loading region data
    ' from a cache file. if the file does not exist yet, we'll create the regions
    ' and save them - for the next time
    If LoadEdgeRegions(EdgeRegions, CurrPrefs.SkinFullPath & "trans.dat") Then
        LoadedRegionsFromFile = True
    End If
    
    If Not LoadedRegionsFromFile Then
        For i = 0 To FE_LAST
            CreateRegionData EdgeImages(i), EdgeRegions(i)
        Next
    
        SaveEdgeRegions EdgeRegions, CurrPrefs.SkinFullPath & "trans.dat"
    End If

   '// Cursor
   sFileName = CurrPrefs.SkinFullPath & "kewl.cur"
   
   '// See if the SkinPrefernces HasCursors is true and we can find a cursor in the Skin folder
   If (SkinPrefs.HasCursors) And Dir(sFileName) <> "" Then
      '// Load Mouse Cursor
      SkinnedForm.MousePointer = vbCustom
      SkinnedForm.MouseIcon = LoadPicture(sFileName)
   Else
      '// Reset to default
      SkinnedForm.MousePointer = vbDefault
      SkinnedForm.MouseIcon = LoadPicture()
   End If
   
End Sub
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Form's paint event was triggered, so redraw the form.
'//------------------------------------------------------------------
Private Sub SkinnedForm_Paint()
    
    '// Exit sub if the form is not skinned
    If Not bFormSkinned Then Exit Sub
    
    If Not NoRedraw Then
        DrawEdges SkinnedForm, EdgeImages, NumXSlices, NumYSlices, False
    
        MyExitButton.PaintUpImage
        MyMinButton.PaintUpImage
    End If

End Sub
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Determine number of X/Y slices it will take to draw the form
'// resize the ClientArea and Form's Height and Width
'//------------------------------------------------------------------
Public Sub SetPadSize(Optional ByVal nForceRebuild As Boolean)
Dim NewSize             As Long
Static nLastXSlices     As Long
Static nLastYSlices     As Long

If NoRedraw Then Exit Sub

    ' We don't want form redraws when in middle of new size
    ' setting, before the new region was set
    NoRedraw = True
    
    '// Validate the Height / Width are greater than the Minimum
    NumYSlices = IIf(MinYSlices > NumYSlices, MinYSlices, NumYSlices)
    NumYSlices = IIf(NumYSlices > MaxYSlices, MaxYSlices, NumYSlices)
    
    '// Validate the Height / Width are less than than the Maximum
    NumXSlices = IIf(NumXSlices > MaxXSlices, MaxXSlices, NumXSlices)
    NumXSlices = IIf(MinXSlices > NumXSlices, MinXSlices, NumXSlices)
    
    ' Compute width/height of form accodring to the number of
    ' x/y slices
    SkinnedForm.Width = ((EdgeImages(FE_TOP_LEFT).Width + (EdgeImages(FE_TOP_H_SEGMENT).Width * NumXSlices) + EdgeImages(FE_TOP_RIGHT).Width)) * Screen.TwipsPerPixelX
    SkinnedForm.Height = ((EdgeImages(FE_TOP_LEFT).Height + (EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices) + EdgeImages(FE_BOTTOM_LEFT).Height)) * Screen.TwipsPerPixelY

    ' Compute size of client area
    NewSize = EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices
    ClientArea.Height = NewSize
    
    '// Form can't be larger than the current screen
    NewSize = (SkinnedForm.Width / Screen.TwipsPerPixelX) - EdgeImages(FE_LEFT_V_SEGMENT).Width - EdgeImages(FE_RIGHT_V_SEGMENT).Width
    ClientArea.Width = NewSize

    NoRedraw = False
    
    '// Create new window region. Also triggers a redraw, now that
    '// wer'e done setting the new form shape
    '// Only redraw if the overall size has changed.
    If (nForceRebuild) Or (NumXSlices <> nLastXSlices) Or (NumYSlices <> nLastYSlices) Then
        nLastXSlices = NumXSlices
        nLastYSlices = NumYSlices
        BuildWindowRegion
        
        '// Refresh Client Area
        'LockWindowUpdate ClientArea.hwnd
        'ClientArea.Refresh
        'LockWindowUpdate 0
        'SkinnedForm_Paint
        
        'SkinnedForm.VScroll1.Move ClientArea.ScaleWidth - SkinnedForm.VScroll1.Width, 0, SkinnedForm.VScroll1.Width, ClientArea.ScaleHeight
        '// Trigger Resize Event in Calling Form
        RaiseEvent Resize
        
    End If
     
End Sub
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Create the region data to tell Windows what parts of the form
'// are visible and what parts are removed
'//------------------------------------------------------------------
Private Sub BuildWindowRegion()
Dim PrevRegion As Long

    PrevRegion = WindowRegion
    
    ' Create initial region that covers the client area
    WindowRegion = CreateRectRgn(ClientArea.left, ClientArea.Top, ClientArea.left + ClientArea.Width, ClientArea.Top + ClientArea.Height)

    ' Add to it the window region of the form edges
    BuildEdgesRegion WindowRegion, EdgeImages, EdgeRegions, NumXSlices, NumYSlices

    ' Finally - set the full region
    SetWindowRgn SkinnedForm.hwnd, WindowRegion, True
    
    ' Don't forget - delete old window region
    DeleteObject PrevRegion
        
End Sub
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Find out the maximum number of horizontal slices
'// that allows the buttons to be drawn correctly
'// And stays within the boundaries of the screen
'//------------------------------------------------------------------
Private Function FindMaxYSlices() As Long
    
    FindMaxYSlices = RoundUp((nMaxFormHeight - BaseYSize) / EdgeImages(FE_LEFT_V_SEGMENT).Height)

End Function
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Find out the maximum number of horizontal slices
'// that allows the buttons to be drawn correctly
'// And stays within the boundaries of the screen
'//------------------------------------------------------------------
Private Function FindMaxXSlices() As Long

    FindMaxXSlices = RoundUp((nMaxFormWidth - BaseXSize) / EdgeImages(FE_TOP_H_SEGMENT).Width)
    
End Function
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Find out the minimum number of vertical slices
'// that allows the buttons to be drawn correctly
'//------------------------------------------------------------------
Private Function FindMinYSlices() As Long
    
    If nMinimumClientAreaHeight > 0 Then
        'FindMinYSlices = RoundUp((nMinimumClientAreaHeight + OffsetYSize) / EdgeImages(FE_LEFT_V_SEGMENT).Height)
        'FindMinYSlices = RoundUp((nMinimumClientAreaHeight) / (OffsetYSize / 2))
        FindMinYSlices = RoundUp(((nMinimumClientAreaHeight + OffsetYSize) - BaseYSize) / YEdgeHeight)
    Else
        FindMinYSlices = RoundUp((BaseYSize) / EdgeImages(FE_LEFT_V_SEGMENT).Height)
    End If

End Function
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Find out the minimum number of horizontal slices
'// that allows the buttons to be drawn correctly
'//------------------------------------------------------------------
Private Function FindMinXSlices() As Long

'    FindMinXSlices = RoundUp((nMinFormWidth - BaseXSize) / EdgeImages(FE_TOP_H_SEGMENT).Width)
    If nMinimumClientAreaWidth > 0 Then
        'FindMinXSlices = RoundUp((nMinimumClientAreaWidth + OffsetXSize) / EdgeImages(FE_TOP_H_SEGMENT).Width)
        FindMinXSlices = RoundUp(((nMinimumClientAreaWidth + OffsetXSize) - BaseXSize) / XEdgeWidth)
    Else
        FindMinXSlices = RoundUp(BaseXSize / EdgeImages(FE_TOP_H_SEGMENT).Width)
    End If
    
End Function
'\\------------------------------------------------------------------


'//------------------------------------------------------------------
'// Given a double number, the function always returns a long
'// number that is the rounding UP of the double value
'//------------------------------------------------------------------
Private Function RoundUp(Number As Double) As Long
    RoundUp = IIf(Number - CLng(Number) <> 0, CLng(Number + 0.5), CLng(Number))
End Function
'\\------------------------------------------------------------------


'//-------------------------------------------------------------------------
'// Private function to remove all the Skinned form's Style settings
'//-------------------------------------------------------------------------
Private Sub DisableTitleBar(ByVal frmhWnd As Long)
Dim nStyle          As Long

    '// Exit if we already disabled the TitleBar and BorderStyle
    If bBorderStyleRemoved Then Exit Sub
    
    '// Get Original
    nStyle = GetWindowLong(frmhWnd, GWL_STYLE)
    
    '// Remove the BORDER, Caption, Resize Frame(THICKFRAME), etc...
    nStyle = nStyle And Not WS_BORDER
    nStyle = nStyle And Not WS_MINIMIZEBOX
    nStyle = nStyle And Not WS_MAXIMIZEBOX
    nStyle = nStyle And Not WS_SYSMENU
    nStyle = nStyle And Not WS_CAPTION
    nStyle = nStyle And Not WS_DLGFRAME
    nStyle = nStyle And Not WS_THICKFRAME
    
    '// Update the Style
    SetWindowLong frmhWnd, GWL_STYLE, nStyle
    
    '// Redraw the form
    SetWindowPos frmhWnd, 0, 0, 0, 0, 0, SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_FRAMECHANGED
    
    '// All done
    bBorderStyleRemoved = True
    
End Sub
'\\------------------------------------------------------------------


