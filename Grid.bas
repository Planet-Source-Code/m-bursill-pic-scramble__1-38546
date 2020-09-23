Attribute VB_Name = "Grid"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)

Public YScale As Double
Public XScale As Double

Public CurrentBox As cord


Public Sub GetScale(NumOfRow As Integer, NumOfCol As Integer, GridContainer As PictureBox)
    
    'Figure out the scale based on the size of the box
    YScale = (GridContainer.Height) / NumOfRow
    XScale = (GridContainer.Width) / NumOfCol

End Sub

Public Sub DrawGrid(NumOfRow As Integer, NumOfCol As Integer, C, GridContainer As PictureBox)
    
    'Draw the lines
    For y = 0 To GridContainer.Height Step RoundDown(YScale)
        GridContainer.Line (1, y)-(GridContainer.Width, y), C
    Next y
    For x = 0 To GridContainer.Width Step RoundDown(XScale)
        GridContainer.Line (x, 1)-(x, GridContainer.Height), C
    Next x

End Sub

Public Sub FillBox(SourceX As Integer, SourceY As Integer, DestX As Integer, DestY As Integer, NumOfCol As Integer, AddToArray As Boolean, CopyMethod As Long, SourceGridContainer As PictureBox, GridContainer As PictureBox, RefreshDest As Boolean)

    Dim StartY As Integer
    Dim StartX As Integer
    
    SourceStartY = RoundDown(SourceY * YScale)
    SourceStartX = RoundDown(SourceX * XScale)
    SourceEndX = ((SourceX * XScale) + XScale)
    SourceEndY = ((SourceY * YScale) + YScale)
        
    DestStartY = RoundDown(DestY * YScale)
    DestStartX = RoundDown(DestX * XScale)
    DestEndX = ((DestX * XScale) + XScale)
    DestEndY = ((DestY * YScale) + YScale)
    
    R = BitBlt(GridContainer.hDC, DestStartX, DestStartY, (DestEndX - DestStartX), (DestEndY - DestStartY), SourceGridContainer.hDC, SourceStartX, SourceStartY, CopyMethod)
    
    If RefreshDest = True Then
        GridContainer.Refresh
    End If
End Sub


Public Function FindBoxNum(NumOfCol As Integer, CounterX As Integer, CounterY As Integer)
    
    FindBoxNum = (NumOfCol * CounterY) + CounterX

End Function
Public Sub Point2Box(x As Single, y As Single)
    
    CurrentBox.x = RoundDown((x / XScale))
    CurrentBox.y = RoundDown((y / YScale))
    
End Sub

Function RoundDown(Value As Double)
   
   RoundDown = Fix(Value)

End Function
