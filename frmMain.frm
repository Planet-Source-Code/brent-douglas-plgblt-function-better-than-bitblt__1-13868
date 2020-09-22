VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PlgBlt Function Demo"
   ClientHeight    =   5085
   ClientLeft      =   3870
   ClientTop       =   2355
   ClientWidth     =   8220
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   375
      Left            =   645
      TabIndex        =   11
      Top             =   2325
      Width           =   1245
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   2175
      Top             =   4425
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "bmp"
      DialogTitle     =   "Load Graphic"
      Filter          =   "*.bmp;*.jpg"
   End
   Begin VB.Frame Frame2 
      Caption         =   " Destination "
      Height          =   4860
      Left            =   2745
      TabIndex        =   1
      Top             =   105
      Width           =   5385
      Begin VB.PictureBox picDST 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4515
         Left            =   105
         ScaleHeight     =   299
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   341
         TabIndex        =   2
         Top             =   240
         Width           =   5145
         Begin VB.Label Point2 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   135
            Left            =   4995
            MousePointer    =   2  'Cross
            TabIndex        =   9
            Top             =   0
            Width           =   120
         End
         Begin VB.Label Point3 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   135
            Left            =   0
            MousePointer    =   2  'Cross
            TabIndex        =   8
            Top             =   4350
            Width           =   120
         End
         Begin VB.Label Point1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   135
            Left            =   0
            MousePointer    =   2  'Cross
            TabIndex        =   7
            Top             =   0
            Width           =   120
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2130
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   2475
      Begin VB.PictureBox picSRC 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1950
         Left            =   15
         Picture         =   "frmMain.frx":0000
         ScaleHeight     =   130
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   150
         TabIndex        =   10
         Top             =   105
         Width           =   2250
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Drag the different points to any location and watch the picture form. Keep in mind that really big pictures will be jerky."
      ForeColor       =   &H00FF0000&
      Height          =   1125
      Left            =   60
      TabIndex        =   6
      Top             =   3840
      Width           =   2520
   End
   Begin VB.Label Label3 
      Caption         =   "3rd Point"
      Height          =   210
      Left            =   300
      TabIndex        =   5
      Top             =   3465
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "2nd Point"
      Height          =   210
      Left            =   300
      TabIndex        =   4
      Top             =   3225
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "1st Point"
      Height          =   210
      Left            =   300
      TabIndex        =   3
      Top             =   2985
      Width           =   1335
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      Height          =   120
      Left            =   150
      Top             =   3510
      Width           =   105
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   120
      Left            =   150
      Top             =   3270
      Width           =   105
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   120
      Left            =   150
      Top             =   3030
      Width           =   105
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PlgBlt Lib "gdi32" (ByVal hdcDest As Long, _
                                    lpPoint As POINTAPI, _
                                    ByVal hdcSrc As Long, _
                                    ByVal nXSrc As Long, _
                                    ByVal nYSrc As Long, _
                                    ByVal nWidth As Long, _
                                    ByVal nHeight As Long, _
                                    ByVal hbmMask As Long, _
                                    ByVal xMask As Long, _
                                    ByVal yMask As Long) As Long

'PlgBlt Function
'
' hdcDest   Identifies the destination device context.

' lpPoint   Points to an array of three points in logical space that identify three
'           corners of the destination parallelogram. The upper-left corner of the
'           source rectangle is mapped to the first point in this array, the upper-right
'           corner to the second point in this array, and the lower-left corner to the
'           third point. The lower-right corner of the source rectangle is mapped to the
'           implicit fourth point in the parallelogram.

' hdcSrc    Identifies the source device context.

' nXSrc     Specifies the x-coordinate, in logical units, of the upper-left corner of
'           the source rectangle.

' nYSrc     Specifies the y-coordinate, in logical units, of the upper-left corner of
'           the source rectangle.

' nWidth    Specifies the width, in logical units, of the source rectangle.

' nHeight   Specifies the height, in logical units, of the source rectangle.

' hbmMask   Identifies an optional monochrome bitmap that is used to mask the colors of
'           the source rectangle.

' xMask     Specifies the x-coordinate of the upper-left corner of the the monochrome bitmap.

' yMask     Specifies the y-coordinate of the upper-left corner of the the monochrome bitmap.
'
'Return Value
'
'If the function succeeds, the return value is TRUE.
'If the function fails, the return value is FALSE. To get extended error information, call GetLastError.



Private Type POINTAPI
        X As Long
        Y As Long
End Type
' The first point corresponds to the upper-left corner of a parallelogram.
' The second point is the upper-right corner.
' The third point is the lower-left corner.
' The fourth corner is derived from the first three.




'Value       Description
'BLACKNESS   Fills the destination rectangle using the color associated with index 0 in the physical
'            palette. (This color is black for the default physical palette.)
'DSTINVERT   Inverts the destination rectangle.
'MERGECOPY   Merges the colors of the source rectangle with the specified pattern by using the Boolean
'            AND operator.
'MERGEPAINT  Merges the colors of the inverted source rectangle with the colors of the destination
'            rectangle by using the Boolean OR operator.
'NOTSRCCOPY  Copies the inverted source rectangle to the destination.
'NOTSRCERASE Combines the colors of the source and destination rectangles by using the Boolean OR
'            operator and then inverts the resultant color.
'PATCOPY     Copies the specified pattern into the destination bitmap.
'PATINVERT   Combines the colors of the specified pattern with the colors of the destination rectangle
'            by using the Boolean XOR operator.
'PATPAINT    Combines the colors of the pattern with the colors of the inverted source rectangle by
'            using the Boolean OR operator. The result of this operation is combined with the colors
'            of the destination rectangle by using the Boolean OR operator.
'SRCAND      Combines the colors of the source and destination rectangles by using the Boolean AND
'            operator.
'SRCCOPY     Copies the source rectangle directly to the destination rectangle.
'SRCERASE    Combines the inverted colors of the destination rectangle with the colors of the source
'            rectangle by using the Boolean AND operator.
'SRCINVERT   Combines the colors of the source and destination rectangles by using the Boolean XOR operator.
'SRCPAINT    Combines the colors of the source and destination rectangles by using the Boolean OR operator.
'WHITENESS   Fills the destination rectangle using the color associated with index 1 in the physical palette.
'            (This color is white for the default physical palette.)

Const BLACKNESS = &H42          ' (DWORD) dest = BLACK
Const DSTINVERT = &H550009      ' (DWORD) dest = (NOT dest)
Const MERGECOPY = &HC000CA      ' (DWORD) dest = (source AND pattern)
Const MERGEPAINT = &HBB0226     ' (DWORD) dest = (NOT source) OR dest
Const NOTSRCCOPY = &H330008     ' (DWORD) dest = (NOT source)
Const NOTSRCERASE = &H1100A6    ' (DWORD) dest = (NOT src) AND (NOT dest)
Const PATCOPY = &HF00021        ' (DWORD) dest = pattern
Const PATINVERT = &H5A0049      ' (DWORD) dest = pattern XOR dest
Const PATPAINT = &HFB0A09       ' (DWORD) dest = DPSnoo
Const SRCAND = &H8800C6         ' (DWORD) dest = source AND dest
Const SRCCOPY = &HCC0020        ' (DWORD) dest = source
Const SRCERASE = &H440328       ' (DWORD) dest = source AND (NOT dest )
Const SRCINVERT = &H660046      ' (DWORD) dest = source XOR dest
Const SRCPAINT = &HEE0086       ' (DWORD) dest = source OR dest
Const WHITENESS = &HFF0062      ' (DWORD) dest = WHITE


Dim isP1Drag As Boolean ' Is Point1 being dragged?
Dim isP2Drag As Boolean ' Is Point1 being dragged?
Dim isP3Drag As Boolean ' Is Point1 being dragged?

Dim Pts(2) As POINTAPI  ' This holds our points for plotting
Dim ret As Long         ' This is the return value of the PlgBlt function


Private Sub Command1_Click()
On Error Resume Next    ' Ignore any errors
dlgOpen.FileName = ""   ' Clear the dialog filename
dlgOpen.ShowOpen        ' Show the open dialog
If dlgOpen.FileName = "" Then Exit Sub  ' If user selects nothing then do nothing
picSRC.Picture = LoadPicture(dlgOpen.FileName)  ' Load the picture
Call DrawPic    ' Draw new picture on the screen
End Sub

Private Sub Form_Load()
Call DrawPic    ' Draw the picture
End Sub

Private Sub picDST_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
If isP1Drag = True Then     ' If you are dragging a point, update while dragging
    Point1.Left = X - 4     ' Keep the point centered horizontally
    Point1.Top = Y - 4      ' Keep the point centered vertically
    Call DrawPic            ' Update
    Exit Sub
End If
If isP2Drag = True Then     ' Same
    Point2.Left = X - 4
    Point2.Top = Y - 4
    Call DrawPic
    Exit Sub
End If
If isP3Drag = True Then     ' Same
    Point3.Left = X - 4
    Point3.Top = Y - 4
    Call DrawPic
    Exit Sub
End If
End Sub


Private Sub picDST_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Button Then      ' Bring all points out of drag mode
    isP1Drag = False
    isP2Drag = False
    isP3Drag = False
End If

End Sub

Private Sub Point1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
isP1Drag = True     ' Is a mouse button is down, start dragging
Point1.Drag
End Sub



Private Sub Point2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button Then      ' Is a mouse button is down, start dragging
    isP2Drag = True
    Point2.Drag
End If

End Sub


Private Sub Point3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button Then      ' Is a mouse button is down, start dragging
    isP3Drag = True
    Point3.Drag
End If

End Sub



Public Sub DrawPic()

Pts(0).X = Point1.Left  ' Put first point into the point array
Pts(0).Y = Point1.Top

Pts(1).X = Point2.Left  ' Put second point into the point array
Pts(1).Y = Point2.Top

Pts(2).X = Point3.Left  ' Put third point into the point array
Pts(2).Y = Point3.Top


picDST.Cls  ' Clear the destination surface

' picDST.hDC            - Destination
' Pts(0)                - Point Array
' PicSRC.hDC            - Source Image
' 0, 0                  - Upper Left Corner of Source
' picSRC.ScaleWidth     - Width of Source
' picSRC.ScaleHeight    - Height of Source
' 0, 0, 0               - We are not using a mask
ret = PlgBlt(picDST.hDC, Pts(0), picSRC.hDC, 0, 0, picSRC.ScaleWidth, picSRC.ScaleHeight, 0, 0, 0)

picDST.Refresh  ' Refresh the destination surface
End Sub
