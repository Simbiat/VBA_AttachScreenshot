Attribute VB_Name = "Screenshot"
Option Explicit
#If VBA7 Then
    Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Public Declare PtrSafe Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
    Public Declare PtrSafe Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
#Else
    Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function CloseClipboard Lib "user32" () As Long
    Public Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
    Public Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
#End If
Private Sub AttachScreenshot()
    Dim CellToPaste As Range
    Set CellToPaste = Range("B2")
    
    Dim TempFilePath As String
    TempFilePath = Environ("temp") & "\temporaryfile.jpg"
    Dim ScreenshotName As String
    ScreenshotName = "ScreenShot" & CellToPaste.Address

    Dim Screenshot As Shape
    Dim FinalObject As OLEObject
    Dim ChartTemp As ChartObject
    Dim ChartAreaTemp As Chart
    'Check that we have a picture in clipboard
    If ScreenShotInClipBoard = True Then
        'Select cell to paste object into it
        CellToPaste.Select
        'Remove previous screenshot if any
        For Each Screenshot In Selection.Parent.Shapes
            Debug.Print Screenshot.Name
            If Screenshot.Name = ScreenshotName Then
                Screenshot.Delete
            End If
        Next Screenshot
        'Paste screenshot. Doing this to, essentially get dimensions of the image required for export
        ActiveSheet.Paste
        Set Screenshot = Selection.Parent.Shapes(Selection.Name)
        If Screenshot.Type <> msoPicture Then
            Screenshot.Delete
            Exit Sub
        End If
        'Create chart to export image through it
        Set ChartTemp = ActiveSheet.ChartObjects.Add(0, 0, Screenshot.Width, Screenshot.Height)
        Set ChartAreaTemp = ChartTemp.Chart
        'Export to image
        With ChartAreaTemp
            .ChartArea.Select
            .Paste
            .Export Filename:=TempFilePath, FilterName:="JPEG", Interactive:=False
        End With
        'Delete chart
        ChartTemp.Delete
        'Delete screenshot
        Screenshot.Delete
        'Attach file
        Set FinalObject = ActiveSheet.OLEObjects.Add(Filename:=TempFilePath, Link:=False, DisplayAsIcon:=True, IconLabel:=ScreenshotName)
        Kill TempFilePath
        'Set shape name to allow auto-removal tp avoid duplicates
        FinalObject.Name = ScreenshotName
        With Selection.Parent.Shapes(ScreenshotName)
            .LockAspectRatio = msoFalse
            .Top = CellToPaste.Top
            .Left = CellToPaste.Left
            .Width = CellToPaste.Width
            .Height = CellToPaste.Height
        End With
    End If
End Sub
Private Function ScreenShotInClipBoard() As Boolean
    Dim sClipboardFormatName As String, sBuffer As String
    Dim CF_Format As Long, i As Long
    Dim bDtataInClipBoard As Boolean
    If OpenClipboard(0) Then
        CF_Format = EnumClipboardFormats(0&)
        Do While CF_Format <> 0
            sClipboardFormatName = String(255, vbNullChar)
            i = GetClipboardFormatName(CF_Format, sClipboardFormatName, 255)
            sBuffer = sBuffer & Left(sClipboardFormatName, i)
            bDtataInClipBoard = True
            CF_Format = EnumClipboardFormats(CF_Format)
        Loop
        CloseClipboard
    End If
    ScreenShotInClipBoard = bDtataInClipBoard And Len(sBuffer) = 0
End Function
