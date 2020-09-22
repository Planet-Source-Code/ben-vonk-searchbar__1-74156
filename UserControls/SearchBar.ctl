VERSION 5.00
Begin VB.UserControl SearchBar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   912
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   912
   ScaleHeight     =   76
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   76
   ToolboxBitmap   =   "SearchBar.ctx":0000
   Begin VB.Timer tmrSearch 
      Enabled         =   0   'False
      Left            =   240
      Top             =   240
   End
End
Attribute VB_Name = "SearchBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'SearchBar Control
'
'Author Ben Vonk
'25-10-2011 First version
'09-12-2011 Second version, Add BarBorder and BarBorderColor properties
Option Explicit

' Public Event
Public Event Cycle(Cycles As Double, Bars As Integer)

' Public Enumerations
Public Enum BarNumbers
   Bars4
   Bars6
   Bars8
   Bars10
   Bars12
End Enum

Public Enum BarStyles
   Small
   Large
End Enum

Public Enum Rotations
   ClockWise
   AntiClockWise
End Enum

Public Enum Speeds
   Slow
   Medium
   Fast
End Enum

' Private Variables
Private m_BarStyle           As BarStyles
Private m_BarBorder          As Boolean
Private IsFirst              As Boolean
Private IsLast               As Boolean
Private m_KeepSearchPointer  As Boolean
Private m_Start              As Boolean
Private StopSearching        As Boolean
Private BarPosition(1, 11)   As Double
Private CyclesCount          As Double
Private BarsCount            As Integer
Private BarStep              As Integer
Private DrawWidthSaved       As Integer
Private m_BarsDisplayed      As Integer
Private m_SearchStartPointer As Integer
Private MaxBars              As Integer
Private SearchPointer(1)     As Integer
Private StartBar             As Integer
Private m_BarColor           As Long
Private m_BarColorSearch(1)  As Long
Private m_BarBorderColor     As OLE_COLOR
Private m_Rotation           As Rotations
Private m_Speed              As Speeds

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."

   BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)

   UserControl.BackColor = NewBackColor
   PropertyChanged "BackColor"
   
   Call DrawAllBars

End Property

Public Property Get BarBorder() As Boolean
Attribute BarBorder.VB_Description = "Returns/sets a value indicating to show or hide the bars border."

   BarBorder = m_BarBorder

End Property

Public Property Let BarBorder(ByVal NewBarBorder As Boolean)

   m_BarBorder = NewBarBorder
   PropertyChanged "BarBorder"
   
   Call SetSettings
   Call MakePicture
   Call DrawAllBars

End Property

Public Property Get BarBorderColor() As OLE_COLOR
Attribute BarBorderColor.VB_Description = "Returns/sets the color of the bars border."

   BarBorderColor = m_BarBorderColor

End Property

Public Property Let BarBorderColor(ByVal NewBarBorderColor As OLE_COLOR)

   m_BarBorderColor = NewBarBorderColor
   PropertyChanged "BarBorderColor"
   
   Call MakePicture
   Call DrawAllBars

End Property

Public Property Get BarColor() As OLE_COLOR
Attribute BarColor.VB_Description = "Returns/sets the color of bars."

   BarColor = m_BarColor

End Property

Public Property Let BarColor(ByVal NewBarColor As OLE_COLOR)

   m_BarColor = NewBarColor
   PropertyChanged "BarColor"
   
   Call DrawAllBars

End Property

Public Property Get BarColorSearch() As OLE_COLOR
Attribute BarColorSearch.VB_Description = "Returns/sets the color of the search bar."

   BarColorSearch = m_BarColorSearch(0)

End Property

Public Property Let BarColorSearch(ByVal NewBarColorSearch As OLE_COLOR)

   m_BarColorSearch(0) = NewBarColorSearch
   PropertyChanged "BarColorSearch"
   
   Call SetSettings
   Call DrawAllBars

End Property

Public Property Get BarsDisplayed() As BarNumbers
Attribute BarsDisplayed.VB_Description = "Returns/sets the number of bars displayed."

   BarsDisplayed = m_BarsDisplayed

End Property

Public Property Let BarsDisplayed(ByVal NewBarsDisplayed As BarNumbers)

Dim sngPointerPercent As Single

   sngPointerPercent = (m_SearchStartPointer + 1) / ((m_BarsDisplayed + 1) * 2 + 2)
   m_BarsDisplayed = NewBarsDisplayed
   PropertyChanged "BarsDisplayed"
   
   Call CreateBars
   Call MakePicture
   Call DrawAllBars
   
   SearchStartPointer = Round((MaxBars + 1) * sngPointerPercent)

End Property

Public Property Get BarStyle() As BarStyles
Attribute BarStyle.VB_Description = "Returns/sets the style of bars."

   BarStyle = m_BarStyle

End Property

Public Property Let BarStyle(ByVal NewBarStyle As BarStyles)

   m_BarStyle = NewBarStyle
   PropertyChanged "BarStyle"
   
   Call SetSettings
   Call MakePicture
   Call DrawAllBars

End Property

Public Property Get KeepSearchPointer() As Boolean
Attribute KeepSearchPointer.VB_Description = "Returns/sets the last search position to use as start for the next search."

   KeepSearchPointer = m_KeepSearchPointer

End Property

Public Property Let KeepSearchPointer(ByVal NewKeepSearchPointer As Boolean)

   m_KeepSearchPointer = NewKeepSearchPointer
   PropertyChanged "KeepSearchPointer"

End Property

Public Property Get Rotation() As Rotations
Attribute Rotation.VB_Description = "Returns/sets the rotation for the search bar."

   Rotation = m_Rotation

End Property

Public Property Let Rotation(ByVal NewRotation As Rotations)

   m_Rotation = NewRotation
   PropertyChanged "Rotation"

End Property

Public Property Get SearchStartPointer() As Integer
Attribute SearchStartPointer.VB_Description = "Returns/sets the first bar to start searching."

   SearchStartPointer = m_SearchStartPointer + 1

End Property

Public Property Let SearchStartPointer(ByVal NewSearchStartPointer As Integer)

   If NewSearchStartPointer < 1 Then NewSearchStartPointer = 1
   If NewSearchStartPointer > (MaxBars + 1) Then NewSearchStartPointer = MaxBars + 1
   
   m_SearchStartPointer = NewSearchStartPointer - 1
   StartBar = m_SearchStartPointer
   PropertyChanged "SearchStartPointer"

End Property

Public Property Get Speed() As Speeds
Attribute Speed.VB_Description = "Returns/sets the speed of the search bar."

   Speed = m_Speed

End Property

Public Property Let Speed(ByVal NewSpeed As Speeds)

   m_Speed = NewSpeed
   PropertyChanged "Speed"
   
   Call SetSettings

End Property

Public Property Get Start() As Boolean
Attribute Start.VB_Description = "Returns/sets the start or stop action of the search bar."

   Start = m_Start

End Property

Public Property Let Start(ByVal NewStart As Boolean)

   If NewStart Then
      If Not m_KeepSearchPointer Or (SearchPointer(0) > MaxBars) Then SearchPointer(0) = m_SearchStartPointer
      
      CyclesCount = 0
      StartBar = SearchPointer(0)
      BarsCount = -1
      StopSearching = False
      IsFirst = True
      IsLast = False
      m_Start = True
      
      Call tmrSearch_Timer
      
   Else
      StopSearching = True
   End If

End Property

Private Function DarkerColor(ByVal Color As OLE_COLOR) As OLE_COLOR

Const COLOR_DARKER As Integer = 36

Dim intBlue        As Integer
Dim intGreen       As Integer
Dim intRed         As Integer

   intRed = Val("&H" & Right(CStr(Hex(Color)), 2))
   intGreen = Val("&H" & Mid(CStr(Hex(Color)), 3, 2))
   intBlue = Val("&H" & Left(CStr(Hex(Color)), 2))
   
   If Len(CStr(Hex(Color))) = 4 Then intGreen = Val("&H" & Left(CStr(Hex(Color)), 2))
   If Len(CStr(Hex(Color))) = 2 Then intGreen = 0
   If Len(CStr(Hex(Color))) < 5 Then intBlue = 0
   
   intRed = intRed - COLOR_DARKER
   intGreen = intGreen - COLOR_DARKER
   intBlue = intBlue - COLOR_DARKER
   
   If intRed < 0 Then intRed = 0
   If intGreen < 0 Then intGreen = 0
   If intBlue < 0 Then intBlue = 0
   
   DarkerColor = RGB(intRed, intGreen, intBlue)

End Function

Private Function DegToRad(ByVal Degrees As Double) As Double

   DegToRad = Degrees * 1.74532925199433E-02

End Function

Private Sub AddValue(ByVal Value As Integer)

   If Not IsFirst Then SearchPointer(0) = SearchPointer(0) + Value
   
   If m_Rotation = ClockWise Then
      If SearchPointer(0) > MaxBars Then
         SearchPointer(0) = 0
         SearchPointer(1) = MaxBars
      End If
      
   ElseIf SearchPointer(0) < 0 Then
      SearchPointer(0) = MaxBars
      SearchPointer(1) = 0
   End If
   
   If Not IsFirst And (SearchPointer(0) = StartBar) Then
      CyclesCount = CyclesCount + 1
      RaiseEvent Cycle(CyclesCount, MaxBars + 1)
      
      If Not StopSearching Then BarsCount = 0
   End If

End Sub

Private Sub CreateBars()

Dim dblRadians As Double
Dim intDegrees As Integer

   BarStep = 360 \ (4 + m_BarsDisplayed * 2)
   MaxBars = 0
   
   For intDegrees = 0 To 359 Step BarStep
      dblRadians = DegToRad(intDegrees)
      BarPosition(0, MaxBars) = Sin(dblRadians)
      BarPosition(1, MaxBars) = Cos(dblRadians)
      MaxBars = MaxBars + 1
   Next 'intDegrees
   
   MaxBars = MaxBars - 1
   
   Call MakePicture

End Sub

Private Sub DrawAllBars()

Dim intIndex As Integer

   Cls
   
   For intIndex = 0 To MaxBars
      Call DrawSingleBar(intIndex, m_BarColor)
   Next 'intIndex
   
   If m_Start Then
      Call DrawSingleBar(SearchPointer(0), m_BarColorSearch(0))
      
      If IsFirst Then
         Call AddValue(-1 + (2 And m_Rotation = ClockWise))
         Call DrawSingleBar(SearchPointer(0), m_BarColorSearch(1))
         
         tmrSearch.Enabled = True
         
      ElseIf SearchPointer(0) <> SearchPointer(1) Then
         Call DrawSingleBar(SearchPointer(1), m_BarColorSearch(1))
      End If
      
   ElseIf IsLast Then
      Call DrawSingleBar(SearchPointer(0), m_BarColorSearch(1))
   End If

End Sub

Private Sub DrawSingleBar(ByVal Index As Integer, ByVal Color As Long)

   Line (BarPosition(0, Index), BarPosition(1, Index))-(BarPosition(0, Index) * 0.6, BarPosition(1, Index) * 0.6), Color

End Sub

Private Sub MakePicture()

Dim intIndex As Integer

   Picture = Nothing
   DrawWidth = DrawWidthSaved
   
   If Not m_BarBorder Then Exit Sub
   
   For intIndex = 0 To MaxBars
      Line (BarPosition(0, intIndex), BarPosition(1, intIndex))-(BarPosition(0, intIndex) * 0.6, BarPosition(1, intIndex) * 0.6), m_BarBorderColor
   Next 'intIndex
   
   Picture = Image
   DrawWidth = DrawWidth - 2

End Sub

Private Sub SetSettings()

   Scale (-1.1, 1.1)-(1.1, -1.1)
   DrawWidth = 2 + (3 * m_BarStyle) + (2 And m_BarBorder)
   DrawWidthSaved = DrawWidth
   m_BarColorSearch(1) = DarkerColor(m_BarColorSearch(0))
   tmrSearch.Interval = 200 - (75 * m_Speed)

End Sub

Private Sub tmrSearch_Timer()

Static intLastPointer As Integer

   If IsLast Then
      IsLast = False
      StopSearching = False
      tmrSearch.Enabled = False
      CyclesCount = CyclesCount + 1
      
      If (StartBar = SearchPointer(0)) Or (BarsCount = 0) Then
         RaiseEvent Cycle(CyclesCount, MaxBars + (1 And (BarsCount <> 0)))
         
      Else
         RaiseEvent Cycle(CyclesCount, BarsCount)
      End If
      
      Call DrawAllBars
      
   ElseIf StopSearching Then
      SearchPointer(0) = intLastPointer
      m_Start = False
      IsLast = True
      
      Call DrawAllBars
      
   Else
      Call DrawAllBars
      
      BarsCount = BarsCount + 1
      SearchPointer(1) = SearchPointer(0)
      intLastPointer = SearchPointer(0)
      
      Call AddValue(-1 + (2 And m_Rotation = ClockWise))
      
      IsFirst = False
   End If

End Sub

Private Sub UserControl_InitProperties()

   m_BarBorderColor = &H404000
   m_BarsDisplayed = Bars8
   m_BarStyle = Large
   m_BarColor = &H808000
   m_Rotation = ClockWise
   m_KeepSearchPointer = False
   m_Speed = Medium
   m_BarColorSearch(0) = &HFFFF00
   UserControl.BackColor = Ambient.BackColor
   
   Call SetSettings
   Call CreateBars
   Call DrawAllBars

End Sub

' load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   With PropBag
      UserControl.BackColor = .ReadProperty("BackColor", Ambient.BackColor)
      m_BarBorder = .ReadProperty("BarBorder", False)
      m_BarBorderColor = .ReadProperty("BarBorderColor", &H404000)
      m_BarColor = .ReadProperty("BarColor", &H808000)
      m_BarColorSearch(0) = .ReadProperty("BarColorSearch", &HFFFF00)
      m_BarsDisplayed = .ReadProperty("BarsDisplayed", Bars8)
      m_BarStyle = .ReadProperty("BarStyle", Large)
      m_KeepSearchPointer = .ReadProperty("KeepSearchPointer", False)
      m_Rotation = .ReadProperty("Rotation", ClockWise)
      m_SearchStartPointer = .ReadProperty("SearchStartPointer", 0)
      m_Speed = .ReadProperty("Speed", Medium)
      SearchPointer(0) = m_SearchStartPointer
   End With
   
   Call SetSettings
   Call CreateBars
   Call DrawAllBars
   
End Sub

Private Sub UserControl_Resize()

   If Height < 61 * Screen.TwipsPerPixelY Then Height = 61 * Screen.TwipsPerPixelY
   If Height > 101 * Screen.TwipsPerPixelY Then Height = 101 * Screen.TwipsPerPixelY
   If Width < 61 * Screen.TwipsPerPixelX Then Width = 61 * Screen.TwipsPerPixelX
   If Width > 101 * Screen.TwipsPerPixelX Then Width = 101 * Screen.TwipsPerPixelX
   
   Call SetSettings
   Call CreateBars
   Call DrawAllBars

End Sub

Private Sub UserControl_Terminate()

   Erase BarPosition, SearchPointer, m_BarColorSearch

End Sub

' write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      .WriteProperty "BackColor", UserControl.BackColor, Ambient.BackColor
      .WriteProperty "BarBorder", m_BarBorder, False
      .WriteProperty "BarBorderColor", m_BarBorderColor, &H404000
      .WriteProperty "BarColor", m_BarColor, &H808000
      .WriteProperty "BarColorSearch", m_BarColorSearch(0), &HFFFF00
      .WriteProperty "BarsDisplayed", m_BarsDisplayed, Bars8
      .WriteProperty "BarStyle", m_BarStyle, Large
      .WriteProperty "KeepSearchPointer", m_KeepSearchPointer, False
      .WriteProperty "Rotation", m_Rotation, ClockWise
      .WriteProperty "SearchStartPointer", m_SearchStartPointer, 0
      .WriteProperty "Speed", m_Speed, Medium
   End With

End Sub

