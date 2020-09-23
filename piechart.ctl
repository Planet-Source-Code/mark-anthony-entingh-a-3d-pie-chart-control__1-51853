VERSION 5.00
Begin VB.UserControl Pie 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7080
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ScaleHeight     =   375
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   472
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2190
      Left            =   1455
      ScaleHeight     =   146
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   183
      TabIndex        =   9
      Top             =   1260
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2940
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   2
      Top             =   2565
      Visible         =   0   'False
      Width           =   1215
      Begin VB.Label LblInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Creating chart..."
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   720
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   1785
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   1215
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   165
      Left            =   210
      TabIndex        =   8
      Top             =   4380
      Width           =   3150
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   165
      Left            =   165
      TabIndex        =   7
      Top             =   4200
      Width           =   3315
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   165
      Left            =   150
      TabIndex        =   6
      Top             =   4035
      Width           =   3105
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   150
      TabIndex        =   5
      Top             =   3780
      Width           =   3105
   End
   Begin VB.Label Label1 
      Height          =   180
      Left            =   150
      TabIndex        =   4
      Top             =   3540
      Width           =   150
   End
End
Attribute VB_Name = "Pie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'
'
'
'
'                       PieChart 1.0
'                       copyright Mark Entingh, Beta3 Software Inc.
'                       Tuesday, October 29th, 2002
'
'
'
'
'---------------------------------------------------------------------------------
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Pie Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long

Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long '$USED
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Event UpdateComplete()
Event MouseOverPiePiece(index, button, X, Y)
Event MouseDownPiePiece(index, button, X, Y)
Event MouseUpPiePiece(index, button, X, Y)
Event MouseOutPiePiece(index, button, X, Y)

Private ColorX(), ColorY()
Private DefaultColors As Boolean
Private WaitMsgX
Private OldXYColor
Private TmpChart As Integer


Private PiePiece() As TypPiePiece
Private Type TypPiePiece
    PercentStart As Integer
    PercentEnd As Integer
    FaceColor As String
    ShadowColor As String
End Type

Private Type typPieChart
    Width As Integer
    Height As Integer
    Length As Integer
End Type

Private PieChart As typPieChart


Public Sub CreatePie(perc As Variant, Width, Height, Length, Optional noshow = 0)
On Error Resume Next
Randomize

'create temp pie
    If TmpChart < 2 Then
        TmpChart = 2
        Dim varperc(99)
        For X = 0 To 99
            varperc(X) = 1
        Next
        WidthX = Width
        HeightX = Height
        Lengthx = Length
        CreatePie varperc, WidthX, HeightX, Lengthx, 1
        u = 0
        Do Until u = 2000
            u = u + 1
            DoEvents
        Loop
    End If
    
If Length < 15 Then Length = 15
If Width < 200 Then Width = 200
If Height < 100 Then Height = 100
'load the default colors if needed
colx = ColorX(0)
If colx = "" Then
    LoadDefaultColors
End If

width2 = Width
height2 = Height
Length = height2 + (Length)

PieChart.Width = width2
PieChart.Height = height2
PieChart.Length = Length - height2

UserControl.Cls
UserControl.Picture = Picture3.Picture
Picture1.Cls
Picture2.Cls
Picture1.Picture = UserControl.Picture
Picture2.Picture = UserControl.Picture

UserControl.Width = (width2 + 1) * Screen.TwipsPerPixelX
UserControl.Height = (Length) * Screen.TwipsPerPixelY

Picture1.Visible = False
Picture2.Visible = False

Picture1.Width = width2 * 2
Picture1.Height = Length * 2
Picture1.Move 0, 0

Picture2.Width = width2 * 2
Picture2.Height = Length * 2
Picture2.Move 0, 0

Picture4.Width = width2 * 2
Picture4.Height = Length * 2
'create mask

SelectObject Picture1.hdc, CreateSolidBrush(RGB(0, 0, 0))
Ellipse Picture1.hdc, 1, 1, Width, Height
For X = (height2 / 2) - 4 To (height2 / 2) + (Length - height2) + 4
    Picture1.Line (1, X)-(width2 + 10, X), RGB(0, 0, 0)
Next
Ellipse Picture1.hdc, 1, (Length) - Height, Width + 1, Length

Dim Xstart, Xend, colorIndex, colorDir, colorDirType, X1, X2, X3, X4, X5, Y1, Y2, Y3, Y4, Y5
Xend = 0
Xstart = 0
For X = LBound(perc) To UBound(perc)
    If perc(X) = "" Then Exit For
Next

ReDim PiePiece(X)

For X = LBound(perc) To UBound(perc)
'-----------------------------------------------------------
    'creating the percent start and end x-y positions
'-----------------------------------------------------------
    If perc(X) > 0 Then
        If X = 0 Then
            Xstart = 0
            Xend = Int(perc(X))
            'Xend = Xend + 1
        Else
            Xstart = Int(Xend)
            Xend = Int(Xstart) + Int(perc(X))
        End If
        
        If Xend > 100 Then
            MsgBox "The percentage level exceeds 100% in array #" & X, vbCritical, "Error"
            LblInfo.Caption = "error creating pie chart"
            Exit Sub
        End If
        
        If X > 0 Then
            If perc(X) = 0 Then
                'colorX = RGB(225, 225, 225)
            End If
        End If
'-----------------------------------------------------------
        'drawing lines
'-----------------------------------------------------------
        
            Dim percx
            percx = Xstart
            If percx = 38 Then
                w = 0
            Else
                w = (percx - 38) / 50
            End If
            
            z1 = 3.14159265358979 * (w)
            
            percx = Xend
            If percx = 38 Then
                w = 0
            Else
                w = (percx - 38) / 50
            End If
            
            z2 = 3.14159265358979 * (w)
            
            
            'top of the chart color
            For w = z1 To z2 Step 0.01
                X2 = (Cos(w) * (width2 / 2) + (width2 / 2))
                Y2 = (Sin(w) * (height2 / 2) + (height2 / 2))
                'Picture2.Line ((width2 / 2), (height2 / 2))-(X2, Y2), colorX
            Next
            
            
        
        'create black lines
            If Xend = 100 And X = 0 Then
            Else
                X2 = (Cos(z1) * (width2 / 2) + (width2 / 2))
                Y2 = (Sin(z1) * (height2 / 2) + (height2 / 2))
                Picture2.Line ((width2 / 2), (height2 / 2))-(X2, Y2), vbBlack
                X3 = (Cos(z2) * (width2 / 2) + (width2 / 2))
                Y3 = (Sin(z2) * (height2 / 2) + (height2 / 2))
                Picture2.Line ((width2 / 2), (height2 / 2))-(X3, Y3), vbBlack
            End If
        'create line on pie side
        If Xend > 37 And Xend < 89 Then
            Call GetPixelDegreeOfPercent(Xend, X2, Y2, width2, height2)
            
            Picture2.Line (X2, Y2)-(X2, Y2 + (Length - height2)), vbBlack
        End If
    End If
Next
    
'-----------------------------------------------------------
    'create side lines
'-----------------------------------------------------------
    
    
    If Xend < 100 Then
        'create cut piece end going down
        Dim nomore As Boolean
        nomore = False
        Call GetPixelDegreeOfPercent(0, X2, Y2, width2, height2)
        Call GetPixelDegreeOfPercent(Xend, X3, Y3, width2, height2)
        Y4 = Y2 + (Length - height2) 'the length of the complete line
        X2 = X2
                For Y = Y2 + 1 To Y4 Step 1
                    Picture2.Line (X2, Y2)-(X2, Y), vbBlack
                    If Picture2.Point(X2, Y) = 0 Then
                        Picture2.Line (X2, Y2)-(X2, Y), vbBlack
                        nomore = True
                        Exit For
                    End If
                Next
                'MsgBox X & "," & Y3
                
                Y4 = Y

        
        'create cut piece end going across
        Debug.Print nomore
        If nomore = False Then
            'create first line
            Dim endFor As Boolean
            endofr = False
            Call GetPixelDegreeOfPercent(0, X2, Y2, width2, height2)
            Call GetPixelDegreeOfPercent(Xend, X3, Y3, width2, height2)
            oldx = X2 + 1
            oldy = Y4 - 1
            Y = Y4 - 1
            For X = X2 + 1 To (width2 / 2) + 1 Step 1
                    'MsgBox Picture2.Point(X + 1, Y + 1)
                    If Picture2.Point(X + 1, Y + 1) = 0 Then endFor = True
                    If Picture2.Point(X, Y + 1) = 0 Then endFor = True
                    If Picture2.Point(X + 1, Y) = 0 Then endFor = True
                    Y = Y + ((((Height / 2) + (Length - height2)) - Y4) / ((width2 / 2) - X2))
                    Picture2.Line (oldx, oldy)-(X, Y), vbBlack
                    
                    'If Picture2.Point(X + 1, Y) = 0 Then endFor = True
                    'If Picture2.Point(X, Y + 1) = 0 Then endFor = True
                    'If Picture2.Point(X + 1, Y + 1) = 0 Then endFor = True
                    'Picture2.Circle (X, Y), 1, vbRed
                    oldx = X
                    oldy = Y
                If endFor = True Then
                    Y = Y + ((((height2 / 2) + (Length - height2)) - (height2 / 2)) / ((width2 / 2) - X2))
                    Picture2.Line (oldx, oldy)-(X + 1, Y), vbBlack
                    Exit For
                End If
            Next
                    
            'create second line
                If Xend < 64 And Xend > 13 Then
                Call GetPixelDegreeOfPercent(Xend, X2, Y2, width2, height2)
                Picture2.Line (X2, Y2 + Length - height2)-((width2 / 2) - 1, (height2 / 2) + Length - height2 - 1), vbBlack
                End If
        End If
    End If
    
    '-----------------------------------------------------------
    'APPLY MASK
    '-----------------------------------------------------------
    

    
    hDCDest = UserControl.hdc
    Picture2.Picture = Picture2.Image
    'BitBlt hDCDest, 0, 0, width2, length, Picture2.hdc, 0, 0, vbSrcAnd
    'BitBlt hDCDest, 0, 0, width2, length, Picture1.hdc, 0, 0, vbSrcPaint
    UserControl.Picture = Picture2.Picture
    UserControl.Picture = UserControl.Image
    
    
    '--------------------------------------------------------------
    'circle border
    '--------------------------------------------------------------
    percx = Xend
    If percx = 38 Then
        w = 0
    Else
        w = (percx - 38) / 50
    End If
            
    z2 = 3.14159265358979 * (w)
    z1 = 3.14159265358979 * (-38 / 50)
    
    oldx = -1
    oldy = -1
    width3 = width2 - 1
    height3 = height2 - 1
    For X = z1 To z2 + 0.018 Step 0.05
        X2 = (Cos(X) * (width3 / 2) + (width2 / 2))
        Y2 = (Sin(X) * (height3 / 2) + (height2 / 2))
        If oldx = -1 Then
            UserControl.Line (X2, Y2)-(X2, Y2), vbBlack
        Else
            UserControl.Line (oldx, oldy)-(X2, Y2), vbBlack
        End If
        oldx = X2
        oldy = Y2
    Next
    X2 = (Cos(z2 + 0.018) * (width3 / 2) + (width2 / 2))
    Y2 = (Sin(z2 + 0.018) * (height3 / 2) + (height2 / 2))
    UserControl.Line (oldx, oldy)-(X2, Y2), vbBlack
    
    'bottom circle border
    oldx = -1
    oldy = -1
    If z2 > 3.14 Then z2 = 3.14
    width3 = width2 - 1
    height3 = height2 - 1
    For X = 0 To z2 + 0.018 Step 0.05
        X2 = (Cos(X) * (width3 / 2) + (width2 / 2))
        Y2 = (Sin(X) * (height3 / 2) + (Length - (height2 / 2)))
        Y2 = Y2 - 1
        If oldx = -1 Then
            UserControl.Line (X2, Y2)-(X2, Y2), vbBlack
        Else
            UserControl.Line (oldx, oldy)-(X2, Y2), vbBlack
        End If
        oldx = X2
        oldy = Y2
    Next
    If Xend > 38 Then
        If Xend > 82 Then
            xi = 0.08
        Else
            xi = 0.018
        End If
        X2 = (Cos(z2 + xi) * (width3 / 2) + (width2 / 2))
        Y2 = (Sin(z2 + xi) * (height3 / 2) + (Length - (height2 / 2)))
        UserControl.Line (oldx, oldy)-(X2, Y2), vbBlack
    End If
    
    UserControl.Picture = UserControl.Image
    
    'create side lines
    Call GetPixelDegreeOfPercent(0, X2, Y2, width2, height2)
    Call GetPixelDegreeOfPercent(Xend, X3, Y3, width2, height2)
    X4 = Int(Cos(3.14) * (width2 / 2) + (width2 / 2))
    Y4 = Int(Sin(3.14) * (height2 / 2) + (height2 / 2))
    If Xend > 87 Then
        UserControl.Line (X4 + 1, Y4)-(X4 + 1, Y4 + (Length - height2))
        UserControl.Line (width2, height2 / 2)-(width2, (height2 / 2) + (Length - height2))
    ElseIf Xend = 87 Then
        UserControl.Line (X4 + 1, Y4 + 3)-(X4 + 1, Y4 + (Length - height2) + 2)
        UserControl.Line (width2, height2 / 2)-(width2, (height2 / 2) + (Length - height2))
    ElseIf Xend = 37 Then
        UserControl.Line (width2 / 2, height2 / 2)-(width2 / 2, (height2 / 2) + (Length - height2))
        UserControl.Line (X3, Y3)-(X3, Y3 + (Length - height2))
    ElseIf Xend > 37 And Xend < 87 Then
        UserControl.Line (width2, height2 / 2)-(width2, (height2 / 2) + (Length - height2))
        If Xend < 64 Then
            UserControl.Line (width2 / 2, height2 / 2)-(width2 / 2, (height2 / 2) + (Length - height2))
        End If
    ElseIf Xend < 37 Then
        UserControl.Line (width2 / 2, height2 / 2)-(width2 / 2, (height2 / 2) + (Length - height2))
        If Xend > 13 Then
            If Xend > 36 Then
                UserControl.Line (X3, Y3 - 2)-(X3, Y3 + (Length - height2))
            Else
                UserControl.Line (X3, Y3)-(X3, Y3 + (Length - height2))
            End If
        End If
    ElseIf Xend < 38 And Xend > 13 Then
        UserControl.Line (width2 / 2, height2 / 2)-(width2 / 2, (height2 / 2) + (Length - height2))
    End If
    
'-------------------------------------------
'FILL IN THE CHART
'-------------------------------------------
colorDir = 2 '1=down, 2=up
colorIndex = -1
colorDirType = 2 '1=up and down, 2=up over and over
For X = LBound(perc) To UBound(perc)

    'selecting the color
        
    If colorDirType = 1 Then
    If colorIndex = UBound(ColorX) And colorDir = 2 Then
        colorDir = 1
        colorIndex = UBound(ColorX) - 1
    ElseIf colorIndex = 0 And colorDir = 1 Then
        colorDir = 2
        colorIndex = 1
    Else
        If colorDir = 2 Then
            colorIndex = colorIndex + 1
        Else
            colorIndex = colorIndex - 1
        End If
    End If
    Else
        If colorIndex = UBound(ColorX) Then
            colorIndex = 0
        Else
            colorIndex = colorIndex + 1
        End If
    End If
    If TmpChart = 2 Then colorIndex = X

    If perc(X) > 0 Then
        If X = 0 Then
            Xstart = 0
            Xend = Int(perc(X))
            'Xend = Xend + 1
        Else
            Xstart = Int(Xend)
            Xend = Int(Xstart) + Int(perc(X))
        End If
        
        '----------------
        'Create PiePiece in Array
        '----------------
        
        If TmpChart <> 2 Then
            PiePiece(X).FaceColor = ColorX(colorIndex)
            PiePiece(X).ShadowColor = ColorY(colorIndex)
            PiePiece(X).PercentStart = Xstart
            PiePiece(X).PercentEnd = Xend
        End If
        
        If TmpChart = 2 Then
            linecolors = RGB(X, X, X)
        Else
            linecolors = vbBlack
        End If
        
            'create fill
            UserControl.FillColor = ColorA(colorIndex)
                
                'fill in top of chart
                percx = Xstart
                If percx = 38 Then
                    w = 0
                Else
                    w = (percx - 38) / 50
                End If
                        
                z1 = 3.14159265358979 * (w)
                
                percx = Xend
                If percx = 38 Then
                    w = 0
                Else
                    w = (percx - 38) / 50
                End If
                        
                z2 = 3.14159265358979 * (w)
                width3 = width2 / 2
                height3 = height2 / 2
                If z1 > z2 Then
                    z3 = ((z1 - z2) / 2) + z2
                Else
                    z3 = ((z2 - z1) / 2) + z1
                End If
                X5 = (Cos(z3) * (width3 / 2) + (width2 / 2))
                Y5 = (Sin(z3) * (height3 / 2) + (height2 / 2))
                'UserControl.Line (width2, 0)-(X5, Y5), vbBlack
                'UserControl.Circle (X5, Y5), 5, vbBlack
                ExtFloodFill UserControl.hdc, X5, Y5, UserControl.Point(X5, Y5), 1
                
                'fill in side of chart
                If Xend > 40 And Xstart < 86 Then
                    If Xstart < 40 Then
                        percx = 40
                    Else
                        percx = Xstart
                    End If
                    If percx = 38 Then
                        w = 0
                    Else
                        w = (percx - 38) / 50
                    End If
                    z1 = 3.14159265358979 * (w)
                    
                    If Xend > 86 Then
                        percx = 86
                    Else
                        percx = Xend
                    End If
                    If percx = 38 Then
                        w = 0
                    Else
                        w = (percx - 38) / 50
                    End If
                    z2 = 3.14159265358979 * (w)
                    If z1 > z2 Then
                        z3 = ((z1 - z2) / 2) + z2
                    Else
                        z3 = ((z2 - z1) / 2) + z1
                    End If
                    X5 = (Cos(z3) * (width2 / 2) + (width2 / 2))
                    Y5 = (Sin(z3) * (height2 / 2) + (height2 / 2) + ((Length - height2) / 2))
                    UserControl.FillColor = ColorB(colorIndex)
                    'UserControl.Line (width2, 0)-(X5, Y5), vbBlack
                    'UserControl.Circle (X5, Y5), 5, vbBlack
                    ExtFloodFill UserControl.hdc, X5, Y5, UserControl.Point(X5, Y5), 1
                End If
        
                percx = Xstart
                If percx = 38 Then
                    w = 0
                Else
                    w = (percx - 38) / 50
                End If
                        
                z1 = 3.14159265358979 * (w)
                
                percx = Xend
                If percx = 38 Then
                    w = 0
                Else
                    w = (percx - 38) / 50
                End If
                        
                z2 = 3.14159265358979 * (w)
                
        'top of the chart color slow way(to fill in any spaces left out)
            For w = z1 To z2 Step 0.01
                X2 = (Cos(w) * (width2 / 2) + (width2 / 2))
                Y2 = (Sin(w) * (height2 / 2) + (height2 / 2))
                UserControl.Line ((width2 / 2), (height2 / 2))-(X2, Y2), ColorA(colorIndex)
            Next
            
            
            'create black lines
            If Xend = 100 And X = 0 Then
            Else
                X2 = (Cos(z1) * (width2 / 2) + (width2 / 2))
                Y2 = (Sin(z1) * (height2 / 2) + (height2 / 2))
                UserControl.Line ((width2 / 2), (height2 / 2))-(X2, Y2), linecolors
                X3 = (Cos(z2) * (width2 / 2) + (width2 / 2))
                Y3 = (Sin(z2) * (height2 / 2) + (height2 / 2))
                UserControl.Line ((width2 / 2), (height2 / 2))-(X3, Y3), linecolors
            End If
            'create line on pie side
            
        If Xend > 37 And Xend < 89 Then
            Call GetPixelDegreeOfPercent(Xend, X2, Y2, width2, height2)
            UserControl.Line (X2, Y2)-(X2, Y2 + (Length - height2)), linecolors
        End If
            
        End If
Next
            'fill in side for first pie piece
            
            If Xend < 100 Then
                percx = 0
                If percx = 38 Then
                    w = 0
                Else
                    w = (percx - 38) / 50
                End If
                
                z2 = 3.14159265358979 * (w)
                X2 = (Cos(z2) * (width2 / 2) + (width2 / 2))
                Y2 = (Sin(z2) * (height2 / 2) + (height2 / 2))
                
                'X2 = (width2 / 2) - (((width2 / 2) - X2) / 2)
                'Y2 = ((Y2 - (height2 / 2)) / 2) + (height2 / 2)
                UserControl.FillColor = RGB(254, 176, 16)
                'UserControl.Circle (X2 + 1, Y2 + 1), 5, vbBlack
                UserControl.FillColor = ColorB(0)
                If TmpChart <> 2 Then ExtFloodFill UserControl.hdc, X2 + 1, Y2 + 1, UserControl.Point(X2 + 1, Y2 + 1), 1
            
            End If
            
            'fill in side for end pie piece
            
            If Xend < 63 And Xend > 13 Then
            
                percx = Xend
                If percx = 38 Then
                    w = 0
                Else
                    w = (percx - 38) / 50
                End If
                
                z2 = 3.14159265358979 * (w)
                X2 = (Cos(z2) * (width2 / 2) + (width2 / 2))
                Y2 = (Sin(z2) * (height2 / 2) + (height2 / 2))
                
                
    
                If Xend < 59 Then
                UserControl.FillColor = ColorB(colorIndex)
                Else
                UserControl.FillColor = ColorA(colorIndex)
                End If
                X2 = ((X2 - (width2 / 2)) / 2) + (width2 / 2)
                Y2 = (Y2 + ((Length - height2) + (height2 / 2))) / 2
                'UserControl.Circle (X2, Y2), 5, vbBlack
                ExtFloodFill UserControl.hdc, X2, Y2, UserControl.Point(X2, Y2), 1
                
            
            End If
    '--------------------------------------------------------------
    'circle border repeated
    '--------------------------------------------------------------
    percx = Xend
    If percx = 38 Then
        w = 0
    Else
        w = (percx - 38) / 50
    End If
            
    z2 = 3.14159265358979 * (w)
    z1 = 3.14159265358979 * (-38 / 50)
    
    oldx = -1
    oldy = -1
    width3 = width2 - 1
    height3 = height2 - 1
    For X = z1 To z2 + 0.018 Step 0.05
        X2 = (Cos(X) * (width3 / 2) + (width2 / 2))
        Y2 = (Sin(X) * (height3 / 2) + (height2 / 2))
        If oldx = -1 Then
            UserControl.Line (X2, Y2)-(X2, Y2), vbBlack
        Else
            UserControl.Line (oldx, oldy)-(X2, Y2), vbBlack
        End If
        oldx = X2
        oldy = Y2
    Next
    X2 = (Cos(z2 + 0.018) * (width3 / 2) + (width2 / 2))
    Y2 = (Sin(z2 + 0.018) * (height3 / 2) + (height2 / 2))
    UserControl.Line (oldx, oldy)-(X2, Y2), vbBlack
    
    'bottom circle border
    oldx = -1
    oldy = -1
    If z2 > 3.14 Then z2 = 3.14
    width3 = width2 - 1
    height3 = height2 - 1
    For X = 0 To z2 + 0.018 Step 0.05
        X2 = (Cos(X) * (width3 / 2) + (width2 / 2))
        Y2 = (Sin(X) * (height3 / 2) + (Length - (height2 / 2)))
        Y2 = Y2 - 1
        If oldx = -1 Then
            UserControl.Line (X2, Y2)-(X2, Y2), vbBlack
        Else
            UserControl.Line (oldx, oldy)-(X2, Y2), vbBlack
        End If
        oldx = X2
        oldy = Y2
    Next
    If Xend > 38 Then
        If Xend > 83 Then
            xi = 0.022
        Else
            xi = 0.018
        End If
        X2 = (Cos(z2 + xi) * (width3 / 2) + (width2 / 2))
        Y2 = (Sin(z2 + xi) * (height3 / 2) + (Length - (height2 / 2)))
        UserControl.Line (oldx, oldy)-(X2, Y2), vbBlack
    End If
    If noshow = 0 Then
        UserControl.Picture = UserControl.Image
    Else
        Picture4.Picture = UserControl.Image
        UserControl.Cls
        UserControl.Picture = Picture3.Picture
    End If
    If TmpChart = 2 Then
        TmpChart = 0
    End If
End Sub


Private Sub UserControl_Initialize()
On Error Resume Next
DefaultColors = True
LoadDefaultColors
WaitMsgX = "creating pie..."

End Sub

Private Sub UserControl_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    pieindex = GetPiePieceFromXY(X, Y)
    If pieindex > -1 Then RaiseEvent MouseDownPiePiece(pieindex, button, X, Y)
End Sub

Private Sub UserControl_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    pieindex = GetPiePieceFromXY(X, Y)
    If pieindex > -1 Then
        RaiseEvent MouseOverPiePiece(pieindex, button, X, Y)
    Else
        RaiseEvent MouseOutPiePiece(pieindex, button, X, Y)
    End If
End Sub

Private Sub UserControl_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    pieindex = GetPiePieceFromXY(X, Y)
    If pieindex > -1 Then RaiseEvent MouseUpPiePiece(pieindex, button, X, Y)
End Sub

Private Function GetPixelDegreeOfPercent(perc, X, Y, WidthX, HeightX)
On Error Resume Next
    'retrieves the x and y position of where the line starts
    'at the border of the pie chart
    'for a percent of a pie chart
    
    'find the percent of 360 degrees
    '180 degrees = Pi
    'figure out how to take the percent
    
    

    If perc = 38 Then
        w = 0
    Else
        w = (perc - 38) / 50
    End If
    
    z = 3.14159265358979 * (w)
    X = (Cos(z) * (WidthX / 2) + (WidthX / 2))
    Y = (Sin(z) * (HeightX / 2) + (HeightX / 2))
    'If X Mod 1 > 4 Then
    '    X = Int(X) + 1
    'Else
    '    X = Int(X)
    'End If
    'If Y Mod 1 > 4 Then
    '    Y = Int(Y) + 1
    'Else
    '    Y = Int(Y)
    'End If
End Function


Private Sub LoadDefaultColors()
    On Error Resume Next
    DefaultColors = True
    ReDim ColorX(5)
    ReDim ColorY(5)
    ColorX(0) = RGB(254, 176, 16)
    ColorY(0) = RGB(214, 145, 1)
    ColorX(1) = RGB(0, 145, 215)
    ColorY(1) = RGB(0, 115, 165)
    ColorX(2) = RGB(1, 169, 44)
    ColorY(2) = RGB(1, 122, 32)
    ColorX(3) = RGB(210, 44, 2)
    ColorY(3) = RGB(155, 33, 2)
    ColorX(4) = RGB(250, 206, 1)
    ColorY(4) = RGB(205, 169, 1)
    ColorX(5) = RGB(193, 2, 170)
    ColorY(5) = RGB(135, 1, 122)
End Sub

Public Sub NewColor(FaceColor, ShadowColor)
    On Error Resume Next
    If DefaultColors = True Then
        DefaultColors = False
        ReDim ColorX(0)
        ReDim ColorY(0)
    Else
        ReDim Preserve ColorX(UBound(ColorX) + 1)
        ReDim Preserve ColorY(UBound(ColorY) + 1)
    End If
    
    ColorX(UBound(ColorX)) = FaceColor
    ColorY(UBound(ColorY)) = ShadowColor
End Sub

Public Sub ClearColors()
    DefaultColors = True
    LoadDefaultColors
End Sub

Public Property Get Backcolor()
    Backcolor = UserControl.Backcolor
End Property

Public Property Let Backcolor(bgcolor)
        UserControl.Backcolor = bgcolor
        Picture1.Backcolor = bgcolor
        Picture2.Backcolor = bgcolor
        Picture3.Backcolor = bgcolor
End Property

Public Property Get PieChartWidth()
    PieChartWidth = PieChart.Width
End Property

Public Property Get PieChartHeight()
    PieChartHeight = PieChart.Height
End Property

Public Property Get PieChartLength()
    PieChartLength = PieChart.Length
End Property
Private Function GetPiePieceFromXY(X, Y)
    On Error Resume Next
        For z = 1 To 100
            piebg = Picture4.Point(X, Y)
            If piebg = RGB(z, z, z) Then
                    For w = LBound(PiePiece) To UBound(PiePiece)
                        If z >= PiePiece(w).PercentStart And z < PiePiece(w).PercentEnd Then
                            GetPiePieceFromXY = w
                            Exit Function
                        End If
                    Next
                Exit For
            End If
        Next
        GetPiePieceFromXY = -1
End Function

Private Function ColorA(index)
    On Error Resume Next
    If TmpChart = 2 Then
    ColorA = RGB(index, index, index)
    Else
    ColorA = ColorX(index)
    End If
End Function

Private Function ColorB(index)
    On Error Resume Next
    If TmpChart = 2 Then
    ColorB = RGB(index, index, index)
    Else
    ColorB = ColorY(index)
    End If
End Function
