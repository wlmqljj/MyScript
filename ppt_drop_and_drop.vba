Option Explicit

Private Const SM_SCREENX = 1
Private Const SM_SCREENY = 0
Private Const msgCancel = "."
Private Const msgNoXlInstance = "."
Private Const sigProc = "Drag & Drop"
Private Const VK_SHIFT = &H10
Private Const VK_CTRL = &H11
Private Const VK_ALT = &H12

Public Type PointAPI
  X As Long
  Y As Long
End Type
  
Public Type RECT
  lLeft As Long
  lTop As Long
  lRight As Long
  lBottom As Long
End Type

Public Type EndPosition
  X As Long
  Y As Long
End Type
Public Type StartPosition
  X As Long
  Y As Long
End Type

#If VBA7 Then
  Public Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As LongPtr) As Integer
  Public Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal xPoint As LongPtr, ByVal yPoint As LongPtr) As LongPtr
  Public Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As LongPtr
  Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As LongPtr
  Public Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As LongPtr) As LongPtr
#Else
  Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
  Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
  Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
  Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
  Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If
 
Public mPoint As PointAPI
Private ActiveShape As Shape
Private dragMode As Boolean
Private dx As Double, dy As Double
Private objEnd As EndPosition
Private objStart As StartPosition
Dim obj_end As String
Private TotalPoints As Integer

Sub start()
    TotalPoints = 0
    ActivePresentation.SlideShowWindow.View.Next
End Sub

Sub Next_Page()
    ActivePresentation.SlideShowWindow.View.Next
End Sub

Sub Last_Page()
    ActivePresentation.SlideShowWindow.View.Previous
End Sub

Sub DragAndDrop(selectedShape As Shape)
  obj_end = left$(selectedShape.Name, 3) + "_end"
  
  dragMode = Not dragMode
  DoEvents
  ' If the shape has text and we're starting to drag, copy it with its formatting to the clipboard
  If selectedShape.HasTextFrame And dragMode Then selectedShape.TextFrame.TextRange.Copy
  
  dx = GetSystemMetrics(SM_SCREENX)
  dy = GetSystemMetrics(SM_SCREENY)
  
  'Start Position
  objStart.X = selectedShape.left
  objStart.Y = selectedShape.top
  
  'End Position
  objEnd.X = selectedShape.Parent.Shapes(obj_end).left
  objEnd.Y = selectedShape.Parent.Shapes(obj_end).top
  
  Drag selectedShape

  ' Paste the original text while maintaining its formatting, back to the shape
  If selectedShape.HasTextFrame Then selectedShape.TextFrame.TextRange.Paste
  DoEvents
End Sub
 
Private Sub Drag(selectedShape As Shape)
  #If VBA7 Then
    Dim mWnd As LongPtr
  #Else
    Dim mWnd As Long
  #End If
  Dim sx As Long, sy As Long
  Dim WR As RECT ' Slide Show Window rectangle
  Dim StartTime As Single
  ' Change this value to change the timer to automatically drop the shape (can by integer or decimal)
  Const DropInSeconds = 3
  
  ' Get the system cursor coordinates
  GetCursorPos mPoint
  
  ' Find a handle to the window that the cursor is over
  mWnd = WindowFromPoint(mPoint.X, mPoint.Y)
  
  ' Get the dimensions of the window
  GetWindowRect mWnd, WR
  sx = WR.lLeft
  sy = WR.lTop
  Debug.Print sx, sy
  
  With ActivePresentation.PageSetup
    dx = (WR.lRight - WR.lLeft) / .SlideWidth
    dy = (WR.lBottom - WR.lTop) / .SlideHeight
    Select Case True
      Case dx > dy
        sx = sx + (dx - dy) * .SlideWidth / 2
        dx = dy
      Case dy > dx
        sy = sy + (dy - dx) * .SlideHeight / 2
        dy = dx
    End Select
  End With
 
  StartTime = Timer
  
  While dragMode
    GetCursorPos mPoint
    selectedShape.left = (mPoint.X - sx) / dx - selectedShape.Width / 2
    selectedShape.top = (mPoint.Y - sy) / dy - selectedShape.Height / 2
    
    Dim left As Integer
    Dim top As Integer
    left = selectedShape.left
    top = selectedShape.top
    
    'Comment out line below and add shape and name it "position" to get original position of shapes
    selectedShape.Parent.Shapes("position").TextFrame.TextRange = "X: " + CStr(left) + " Y:" + CStr(top)

    
    ' Comment out the next line if you do NOT want to show the countdown text within the shape
    ' If selectedShape.HasTextFrame Then selectedShape.TextFrame.TextRange.Text = CInt(DropInSeconds - (Timer - StartTime))
     
        selectedShape.Parent.Shapes("countdown").TextFrame.TextRange.Text = CInt(DropInSeconds - (Timer - StartTime))
    
    DoEvents
    If Timer > StartTime + DropInSeconds Then
     dragMode = False

        With selectedShape.Parent.Shapes(obj_end) ' EXAMPLE:square_end is where you want the square to land
             If selectedShape.left >= .left And selectedShape.top >= .top And (selectedShape.left + selectedShape.Width) <= (.left + .Width) And (selectedShape.top + selectedShape.Height) <= (.top + .Height) Then
               selectedShape.Parent.Shapes("message").TextFrame.TextRange = "做得好!"
               TotalPoints = TotalPoints + 1
            Else
                selectedShape.Parent.Shapes("message").TextFrame.TextRange = "考虑考虑，再试一下吧"
                selectedShape.left = objStart.X
                selectedShape.top = objStart.Y
            End If
         End With
    Else
        With selectedShape.Parent.Shapes(obj_end) ' EXAMPLE:square_end is where you want the square to land
             If selectedShape.left >= .left And selectedShape.top >= .top And (selectedShape.left + selectedShape.Width) <= (.left + .Width) And (selectedShape.top + selectedShape.Height) <= (.top + .Height) Then
               selectedShape.Parent.Shapes("message").TextFrame.TextRange = "做得好!"
               TotalPoints = TotalPoints + 1
            Else
                selectedShape.Parent.Shapes("message").TextFrame.TextRange = "请拖动到正确位置"
            End If
         End With

    End If
  Wend
  DoEvents
End Sub

Public Sub ResetObjects(selectedShape As Shape)
    selectedShape.Parent.Shapes("nou1").left = 450
    selectedShape.Parent.Shapes("nou1").top = 80 + 0
    
    selectedShape.Parent.Shapes("nou2").left = 450
    selectedShape.Parent.Shapes("nou2").top = 80 + 40 * 1
    
    selectedShape.Parent.Shapes("ver1").left = 450
    selectedShape.Parent.Shapes("ver1").top = 80 + 40 * 2
    
    selectedShape.Parent.Shapes("ver2").left = 450
    selectedShape.Parent.Shapes("ver2").top = 80 + 40 * 3
    
    selectedShape.Parent.Shapes("adj1").left = 450
    selectedShape.Parent.Shapes("adj1").top = 80 + 40 * 4
    
    selectedShape.Parent.Shapes("adj2").left = 450
    selectedShape.Parent.Shapes("adj2").top = 80 + 40 * 5
    
    selectedShape.Parent.Shapes("adv1").left = 450
    selectedShape.Parent.Shapes("adv1").top = 80 + 40 * 6
    
    selectedShape.Parent.Shapes("adv2").left = 450
    selectedShape.Parent.Shapes("adv2").top = 80 + 40 * 7
    
    selectedShape.Parent.Shapes("pre1").left = 450
    selectedShape.Parent.Shapes("pre1").top = 80 + 40 * 8
    
    selectedShape.Parent.Shapes("pre2").left = 450
    selectedShape.Parent.Shapes("pre2").top = 80 + 40 * 9
End Sub
