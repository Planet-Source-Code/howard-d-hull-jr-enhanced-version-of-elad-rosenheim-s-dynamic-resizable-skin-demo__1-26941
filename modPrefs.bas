Attribute VB_Name = "modPrefs"
Option Explicit

' Current program preferences, especially not skin-specific
Type CurrentPreferncesType
    SkinName As String
    SkinsPath As String
    SkinFullPath As String
    ' Add more as you like here
End Type

' Skin-specific preferences. Supplied by the skin's creator
' in a skin.ini file
Type SkinPreferencesType
    BackColor       As Long
    MenuColor       As Long
    ExitButtonX     As Long
    ExitButtonY     As Long
    MinButtonX      As Long
    MinButtonY      As Long
    MenuFontColor   As Long
    TitleColor      As Long
    FontName        As String
    FontSize        As Integer
    FontBold        As Boolean
    FontColor       As Long
    ' Add more as you like here
    HasCursors      As Boolean
End Type

Public CurrPrefs As CurrentPreferncesType

Public SkinPrefs As SkinPreferencesType

Private SkinINIFileName As String

' Read skin-specific preferences from skin.ini file
Public Sub ReadSkinPreferences()
    
    CurrPrefs.SkinFullPath = CurrPrefs.SkinsPath + CurrPrefs.SkinName + "\"
    SkinINIFileName = CurrPrefs.SkinFullPath + "skin.ini"
    
    If Dir(SkinINIFileName) = "" Then
        Err.Raise 1, , "Can't find " & SkinINIFileName & "!"
    End If
    
    With SkinPrefs
        .BackColor = ReadColorFromINI("Skin", "BackColor")
        .MenuColor = ReadColorFromINI("Skin", "MenuColor")
        .MenuFontColor = ReadColorFromINI("Skin", "MenuFontColor")
        .TitleColor = ReadColorFromINI("Skin", "TitleColor")
        .FontColor = ReadColorFromINI("Skin", "FontColor")
        .ExitButtonX = INIRead("Skin", "ExitButtonX", SkinINIFileName)
        .ExitButtonY = INIRead("Skin", "ExitButtonY", SkinINIFileName)
        .MinButtonX = INIRead("Skin", "MinButtonX", SkinINIFileName)
        .MinButtonY = INIRead("Skin", "MinButtonY", SkinINIFileName)
        .FontName = INIRead("Skin", "FontName", SkinINIFileName)
        .FontSize = Val(INIRead("Skin", "FontSize", SkinINIFileName))
        .FontBold = ReadBooleanFromINI("Skin", "FontBold")
        .HasCursors = ReadBooleanFromINI("Skin", "HasCursors")
        
        '// Validate
        .FontName = IIf(.FontName <> "", .FontName, "MS Sans Serif")
        .FontSize = IIf(.FontSize >= 6, .FontSize, 8)
        
    End With
    
End Sub


' Reads an True/False string from
' the skin.ini file, and returns it as a Boolean Value
Private Function ReadBooleanFromINI(Section As String, Value As String) As Boolean
Dim bStr As String
    
    bStr = INIRead(Section, Value, SkinINIFileName)
    
    If UCase(bStr) = "TRUE" Then
        ReadBooleanFromINI = True
    
    Else
        ReadBooleanFromINI = False
        
    End If

End Function


' Reads an RGB color string (in the format RRR,GGG,BBB) from
' the skin.ini file, and returns it as long
Private Function ReadColorFromINI(Section As String, Value As String) As Long
Dim ColorStr As String, ColorArr As Variant
    
    ColorStr = INIRead(Section, Value, SkinINIFileName)
    ColorArr = Split(ColorStr, ",")
    
    If ColorStr = "" Then
        ReadColorFromINI = 0
        
    ElseIf UBound(ColorArr) <> 2 Then
        Err.Raise 1, , "Invalid color value for attribute """ & Value & """"
        
    Else
        ReadColorFromINI = RGB(ColorArr(0), _
                               ColorArr(1), _
                               ColorArr(2))
    End If

End Function
