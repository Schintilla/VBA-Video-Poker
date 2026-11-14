Attribute VB_Name = "Module1"
Option Explicit

Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
    x As Long
    Y As Long
End Type

'GetDeviceCaps function
'http://msdn.microsoft.com/en-us/library/windows/desktop/dd144877%28v=vs.85%29.aspx
Private Const LOGPIXELSX As Long = 88&
Private Const LOGPIXELSY As Long = 90&

'Size of shape's bounding box in points
Private Const SHAPE_WIDTH = 100
Private Const SHAPE_HEIGHT = 100


Public Sub Add_Shape_At_Cursor_Position()

    Dim hdc As Long
    Dim PointsPerPixelX As Double, PointsPerPixelY As Double
    Dim CursorPos As POINTAPI
    Dim ExcelPos As POINTAPI
    Dim ShapePos As POINTAPI
    
    'Get number of points per screen pixel, depending on screen device size
    
    hdc = GetDC(0)
    PointsPerPixelX = 72 / GetDeviceCaps(hdc, LOGPIXELSX)
    PointsPerPixelY = 72 / GetDeviceCaps(hdc, LOGPIXELSY)
    ReleaseDC 0, hdc
    
    'Scale points per pixel according to current window zoom. The smaller the zoom, the higher the number of points per pixel
    
    PointsPerPixelX = PointsPerPixelX * 100 / ActiveWindow.Zoom
    PointsPerPixelY = PointsPerPixelY * 100 / ActiveWindow.Zoom
    
    'Get position of Excel window in screen pixels
    
    ExcelPos.x = ActiveWindow.PointsToScreenPixelsX(0)
    ExcelPos.Y = ActiveWindow.PointsToScreenPixelsY(0)
     
    'Get mouse cursor position in screen pixels
    'Cells(r, c).Select
    GetCursorPos CursorPos
    
    'Set shape position according to mouse position relative to Excel window position, scaled to the
    'number of points per pixel.  Since the AutoShape's position is defined by the top left corner
    'of its bounding box, subtract half the shape's size to centre it over the mouse
    
    ShapePos.x = (CursorPos.x - ExcelPos.x) * PointsPerPixelX - SHAPE_WIDTH / 2
    ShapePos.Y = (CursorPos.Y - ExcelPos.Y) * PointsPerPixelY - SHAPE_HEIGHT / 2
       
    ActiveSheet.Shapes.AddShape msoShapeOval, ShapePos.x, ShapePos.Y, SHAPE_WIDTH, SHAPE_HEIGHT

End Sub
