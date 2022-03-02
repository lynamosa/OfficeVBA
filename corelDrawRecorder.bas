Attribute VB_Name = "RecordedMacros"
Option Explicit

Private Const ssAppName As String = "Myprint"
Private Const ssSection As String = "X7_cr"

Sub OutLineC()
    ' Recorded 4/19/2015
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    OrigSelection.SetOutlineProperties BehindFill:=cdrTrue, LineCaps:=cdrOutlineRoundLineCaps, LineJoin:=cdrOutlineRoundLineJoin
End Sub
Sub Macro1()
    'MsgBox Application.FrameWork.MainMenu.Controls(12).DescriptionText
    MsgBox Application.FrameWork.FrameWindows.First.Caption
    
End Sub
Sub BehindFill()
    ' Recorded 7/19/2015
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    'OrigSelection.SetOutlineProperties BehindFill:=cdrTrue, LineJoin:=cdrOutlineRoundLineJoin
    OrigSelection.SetOutlineProperties LineJoin:=cdrOutlineRoundLineJoin
End Sub
Sub sizeSS()
    ' Recorded 7/4/2015
    ActiveDocument.BeginCommandGroup "SizeSS"
    ActiveDocument.Unit = cdrMillimeter
    'ActiveDocument.ReferencePoint = cdrMiddleLeft
    ActiveDocument.ReferencePoint = cdrCenter
    Dim sz, sz_2 As Double
    Dim indexS, index2 As Integer
    indexS = 1
    sz = 200
    index2 = 2
    sz_2 = 100
    Dim cC As Page
    For Each cC In ActiveDocument.Pages
        If cC.Shapes(indexS).SizeWidth > sz Then cC.Shapes(indexS).SizeWidth = sz
        'If cc.Shapes(index2).SizeWidth > sz_2 Then cc.Shapes(index2).SizeWidth = sz_2
        'cc.Shapes(indexS).Move -1.25, 0
    Next cC
    ActiveDocument.EndCommandGroup
End Sub
Sub sizeSS_Selected()
    ' Fix Shape size selected
    ActiveDocument.BeginCommandGroup "SizeInSelect"
    ActiveDocument.Unit = cdrMillimeter
    'ActiveDocument.ReferencePoint = cdrMiddleLeft
    ActiveDocument.ReferencePoint = cdrCenter
    Dim s As Shape
    Dim sz, sz_2 As Double
    Dim indexS, index2 As Integer
    sz = 38
    For Each s In ActiveSelection.Shapes
        If s.SizeWidth > sz Then s.SizeWidth = sz
    Next
    ActiveDocument.EndCommandGroup
End Sub
Sub sizeFix(ShapeIndex, ShapeSize, Condition, Direction, newSize)
    ' Thu tu - Kich thuoc - Dieu kien - huong - KT moi
    ActiveDocument.Unit = cdrMillimeter
    'ActiveDocument.ReferencePoint = cdrCenter
    ActiveDocument.ReferencePoint = Direction
    Dim cC As Page
    For Each cC In ActiveDocument.Pages
        tmp = cC.Shapes(ShapeIndex).SizeWidth
        If Evaluate(tmp & Condition & ShapeSize) Then cC.Shapes(ShapeIndex).SizeWidth = newSize
    Next cC
End Sub

Sub groupLikeMe()
    ActiveDocument.BeginCommandGroup "Group Like Me"
    Optimization = True
    Dim x As Double, y As Double, w As Double, h As Double
    ActiveDocument.Unit = cdrMillimeter
    ActiveSelection.GetBoundingBox x, y, w, h
    Dim sr As Shape
    For Each sr In ActivePage.Shapes.FindShapes(Query:="@width = {" & w & " mm} & @height = {" & h & " mm}").Shapes
        sr.GetBoundingBox x, y, w, h
        ActivePage.SelectShapesFromRectangle(x, y, x + w, y + h, False).Group
        'sr.SizeWidth = 20
        'sr.SizeHeight = 60
    Next
    Optimization = False
    ActiveDocument.EndCommandGroup
    Refresh
    MsgBox "Group Finish"
End Sub

Sub ExportLikeMe()
    ActiveDocument.BeginCommandGroup "Export Like Me"
    Optimization = True
    Dim x As Double, y As Double, w As Double, h As Double
    Dim i As Integer
    i = 1
    ActiveDocument.Unit = cdrMillimeter
    ActiveSelection.GetBoundingBox x, y, w, h
    Dim OrigSelection As ExportFilter
    Dim fName As String
    Dim wT As Long
    Dim hT As Long
    Dim res As Long
    Dim Resolution As Long
    res = 1600 'Max Size (pixels)
    Resolution = 600
    If ActiveSelection.Shapes.count < 1 Then
        MsgBox "Vui long chon doi tuong de Export", vbCritical, "Alert!"
        Exit Sub
    End If
    If ActiveSelection.SizeWidth < ActiveSelection.SizeHeight Then
        wT = res
        hT = res * ActiveSelection.SizeHeight / ActiveSelection.SizeWidth
    Else
        hT = res
        wT = res * ActiveSelection.SizeWidth / ActiveSelection.SizeHeight
    End If
    fName = InputBox("nhap ten file", "File name to export", Replace(ActiveDocument.FileName, ".cdr", ""))
    If fName = "" Then Exit Sub
    
    Dim sr As Shape
    For Each sr In ActivePage.Shapes.FindShapes(Query:="@width = {" & w & " mm} & @height = {" & h & " mm}").Shapes
        sr.GetBoundingBox x, y, w, h
        ActivePage.SelectShapesFromRectangle x, y, x + w, y + h, False
        Set OrigSelection = ActiveDocument.ExportBitmap("E:\User\desktop\File maket\" & fName & "_" & i & ".jpg", cdrJPEG, cdrSelection, cdrRGBColorImage, wT, hT, Resolution, Resolution, cdrNormalAntiAliasing, True, False)
        OrigSelection.Finish
        i = i + 1
    Next
    Optimization = False
    ActiveDocument.EndCommandGroup
    Refresh
    MsgBox "Export Finish"
    
End Sub
Sub SelectLikeMe()
    Dim x As Double, y As Double, w As Double, h As Double
    ActiveDocument.Unit = cdrMillimeter
    ActiveSelection.GetBoundingBox x, y, w, h
    ActivePage.Shapes.FindShapes(Query:="@width = {" & w & " mm} & @height = {" & h & " mm}").CreateSelection
End Sub
Sub SelectLikeMe_ss()
    ActiveDocument.ReferencePoint = cdrCenter
    Dim cC As Page
    For Each cC In ActiveDocument.Pages
        cC.Activate
        ActivePage.Shapes.FindShapes(Query:="@width < {" & 2 & " mm} & @height < {" & 2 & " mm}").Delete
    Next
End Sub
Sub SelectLikeMe_small()
    ActiveDocument.ReferencePoint = cdrCenter
    ActiveDocument.Unit = cdrMillimeter
    Dim cC As Page
    'For Each cC In ActiveDocument.Pages
        ActivePage.Shapes.FindShapes(Query:="@width < {1 mm} and @height < {1 mm}").Move 200, 0
    'Next
End Sub
Sub SelectShapesWithRedColor()
    ActivePage.Shapes.FindShapes(Query:="@colors.find('Red')").CreateSelection
End Sub
Sub selec_color_mod()
    MsgBox ActiveSelection.Shapes(1).Fill.UniformColor.IsCMYK
End Sub
Sub SelectLikeMe_color()
    Dim x As Double, y As Double, w As Double, h As Double
    Dim RGBmode As Boolean
    RGBmode = False
    If ActiveSelection.Shapes(1).Fill.UniformColor.IsCMYK = False Then
        Dim cR, cB, cG As Integer
        cR = ActiveSelection.Shapes(1).Fill.UniformColor.RGBRed
        cB = ActiveSelection.Shapes(1).Fill.UniformColor.RGBBlue
        cG = ActiveSelection.Shapes(1).Fill.UniformColor.RGBGreen
        ActivePage.Shapes.FindShapes(Query:="@fill.color = rgb(" & cR & ", " & cG & ", " & cB & ")").CreateSelection
    Else
        Dim cC, cM, cY, cK As Integer
        cC = ActiveSelection.Shapes(1).Fill.UniformColor.CMYKCyan
        cM = ActiveSelection.Shapes(1).Fill.UniformColor.CMYKMagenta
        cY = ActiveSelection.Shapes(1).Fill.UniformColor.CMYKYellow
        cK = ActiveSelection.Shapes(1).Fill.UniformColor.CMYKBlack
        ActivePage.Shapes.FindShapes(Query:="@Outline.color = cmyk(" & cC & ", " & cM & ", " & cY & ", " & cK & ")").CreateSelection
    End If
    'ActivePage.Shapes.FindShapes(Query:="@width > {30 mm} & @height > {8 mm}").CreateSelection
    
End Sub
Sub BarCode2Vector()
    ' Recorded 12/16/2015
    Dim OrigSelection As ShapeRange
    Dim x, y As Double
    Set OrigSelection = ActiveSelectionRange
    x = ActiveSelection.PositionX
    y = ActiveSelection.PositionY
    OrigSelection.Cut
    ActiveLayer.PasteSpecial "Metafile"
    ActiveSelection.PositionX = x
    ActiveSelection.PositionY = y
End Sub
Sub EditText()
    Dim og As ShapeRange
    Set og = ActiveSelectionRange
    og.Shapes(1).Text.Story.Font = "Vni-Helve"
    og.Shapes(1).Text.Story.Size = "18"
    og.Shapes(1).Text.Story.Italic = False
    og.Shapes(1).Text.Story.Bold = False
    og.Shapes(1).Text.Story.ChangeCase cdrTextUpperCase
End Sub
Sub IMG_Export()
    ' Recorded 25/01/2016
    Dim OrigSelection As ExportFilter
    Dim fName As String
    Dim w As Long
    Dim h As Long
    Dim res As Long
    Dim Resolution As Long
    res = 1200 'Max Size (pixels)
    Resolution = 300
    If ActiveSelection.Shapes.count < 1 Then
        MsgBox "Vui long chon doi tuong de Export", vbCritical, "Alert!"
        Exit Sub
    End If
    If ActiveSelection.SizeWidth < ActiveSelection.SizeHeight Then
        w = res
        h = res * ActiveSelection.SizeHeight / ActiveSelection.SizeWidth
    Else
        h = res
        w = res * ActiveSelection.SizeWidth / ActiveSelection.SizeHeight
    End If
    fName = InputBox("nhap ten file", "File nameSSSS to export", Replace(ActiveDocument.FileName, ".cdr", ""))
    If fName = "" Then Exit Sub
    'Set OrigSelection = ActiveDocument.ExportBitmap("E:\Users\Desktop\File maket\" & fName & ".jpg", cdrJPEG, cdrSelection, cdrCMYKColorImage, w, h, Resolution, Resolution, cdrNormalAntiAliasing, True, False)
    Set OrigSelection = ActiveDocument.ExportBitmap("E:\User\desktop\File maket\" & fName & ".jpg", cdrJPEG, cdrSelection, cdrRGBColorImage, w, h, Resolution, Resolution, cdrNormalAntiAliasing, True, False)
    'Set OrigSelection = ActiveDocument.ExportBitmap("E:\User\desktop\Giao trinh phun theu tham my-My Tram\hinh new\" & fName & ".jpg", cdrJPEG, cdrSelection, cdrRGBColorImage, w, h, Resolution, Resolution, cdrNormalAntiAliasing, True, False)
    OrigSelection.Finish
End Sub
Sub IMG_multi_Export()
    ' Recorded 25/01/2016
    Dim OrigSelection As ExportFilter
    Dim fName As String
    Dim w As Long
    Dim h As Long
    Dim res As Long
    Dim Resolution As Long
    Dim x As Shape
    Dim Orig As ShapeRange
    Dim j As Integer
    j = 1
    res = 1200
    Resolution = 300
    If ActiveSelection.Shapes.count < 1 Then
        MsgBox "Vui long chon doi tuong de Export", vbCritical, "Alert!"
        Exit Sub
    End If
    fName = InputBox("nhap ten file", "File name to export", Replace(ActiveDocument.FileName, ".cdr", ""))
    If fName = "" Then Exit Sub
    Set Orig = ActiveSelectionRange
    For Each x In Orig
        x.CreateSelection
        If ActiveSelection.SizeWidth < ActiveSelection.SizeHeight Then
            w = res
            h = res * ActiveSelection.SizeHeight / ActiveSelection.SizeWidth
        Else
            h = res
            w = res * ActiveSelection.SizeWidth / ActiveSelection.SizeHeight
        End If
        Set OrigSelection = ActiveDocument.ExportBitmap("E:\User\desktop\File maket\" & fName & j & ".jpg", cdrJPEG, cdrSelection, cdrRGBColorImage, w, h, Resolution, Resolution, cdrNormalAntiAliasing, True, False)
        OrigSelection.Finish
        j = j + 1
    Next
    'Set OrigSelection = ActiveDocument.ExportBitmap("E:\Users\Desktop\File maket\" & fName & ".jpg", cdrJPEG, cdrSelection, cdrCMYKColorImage, w, h, 300, 300, cdrNormalAntiAliasing, True, False)
    
End Sub
Sub IMG_Export_page()
    ' Recorded 25/01/2016
    Dim OrigSelection As ExportFilter
    Dim pg As Page
    Dim fName As String
    Dim w As Long
    Dim h As Long
    Dim res As Long
    Dim Resolution As Long
    res = 1600  'Max Size (pixels)
    Resolution = 600
    If ActivePage.SizeWidth < ActivePage.SizeHeight Then
        w = res
        h = res * ActivePage.SizeHeight / ActivePage.SizeWidth
    Else
        h = res
        w = res * ActivePage.SizeWidth / ActivePage.SizeHeight
    End If
    fName = InputBox("nhap ten file", "File name to export", Replace(ActiveDocument.FileName, ".cdr", ""))
    If fName = "" Then Exit Sub
    'Set OrigSelection = ActiveDocument.ExportBitmap("E:\Users\Desktop\File maket\" & fName & ".jpg", cdrJPEG, cdrSelection, cdrCMYKColorImage, w, h, Resolution, Resolution, cdrNormalAntiAliasing, True, False)
    For Each pg In ActiveDocument.Pages
        pg.Activate
        Set OrigSelection = ActiveDocument.ExportBitmap("E:\User\desktop\File maket\" & fName & "_" & pg.Name & ".jpg", cdrJPEG, cdrCurrentPage, cdrRGBColorImage, w, h, Resolution, Resolution, cdrNormalAntiAliasing, True, False)
        OrigSelection.Finish
    Next
End Sub

Sub My_Print()
    ' Print N page continous, for large page
    ' select print before Print
    ' Shortcut: Ctrl + Shift + P
    Optimization = True
    Dim stepPage As Integer
    stepPage = 25
    If stepPage = 0 Then
        With ActiveDocument.PrintSettings
            .PrintRange = prnCurrentPage
        End With
        ActiveDocument.PrintOut
        Exit Sub
    Else
        Dim a As Integer
        a = GetSetting(ssAppName, ssSection, "page", 1)
        
        With ActiveDocument.PrintSettings
            .PrintRange = prnPageRange
            .PageRange = a & "-" & (a + stepPage - 1)
        End With
        'MsgBox a & "-" & (a + 49)
        'If ActiveDocument.PrintSettings.Printer.Name <> "RICOH Aficio MP C6501" Then
        '    MsgBox "Ban co thuc su muon in vao may" & ActiveDocument.PrintSettings.Printer.Name, vbCritical
        '    Exit Sub
        'End If
        ActiveDocument.PrintOut
        SaveSetting ssAppName, ssSection, "page", a + stepPage
        ActiveDocument.Pages(a + stepPage).Activate
    End If
    Optimization = False
    Application.Refresh
End Sub
Sub print_info()
    MsgBox ActiveDocument.PrintSettings.Printer.Type
End Sub
Sub Show_current_page()
    Dim a As Integer
    a = GetSetting(ssAppName, ssSection, "page", 1)
    MsgBox "Current page is:" & a, vbOKOnly, "My Print auto"
End Sub
Sub My_TestPrint()
    ' 1. Set firstpage for my_print()
    ' 2. Print current page
    
    SaveSetting ssAppName, ssSection, "page", 1
    Exit Sub
    Dim a As Integer
    With ActiveDocument.PrintSettings
        .PrintRange = prnCurrentPage
    End With
    ActiveDocument.PrintOut
End Sub
Sub copy_shape_pos()
    ActiveDocument.Unit = cdrMillimeter
    ActiveSelection.Shapes(2).Stretch ActiveSelection.Shapes(2).SizeHeight / ActiveSelection.Shapes(1).SizeHeight
    ActiveSelectionRange.AlignAndDistribute 3, 3, 0, 0, False, 2
    ActiveSelection.Shapes(1).Move 0#, 220#
End Sub

Sub CopyText_without_format()
    ' Recorded 29/03/2016
    ' Ctrl + Numpad_9
    ActiveDocument.Unit = cdrMillimeter
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    If OrigSelection.Shapes(1).Shapes.count > 1 Then
        OrigSelection.Shapes(1).Shapes(1).Text.Story.Text = OrigSelection.Shapes(2).Text.Story.Text
        'OrigSelection.Shapes(1).Shapes(1).Text.Story.ChangeCase cdrTextUpperCase
    Else
        OrigSelection.Shapes(1).Text.Story.Text = OrigSelection.Shapes(2).Text.Story.Text
        'OrigSelection.Shapes(1).Fill.UniformColor.CMYKAssign 0, 100, 100, 0
        'OrigSelection.Shapes(1).Text.Story.Text = Replace(OrigSelection.Shapes(2).Text.Story.Text, " ", vbCrLf)

    End If
    
    OrigSelection.Shapes(2).Selected = False
    OrigSelection.Shapes(2).Delete
    'If OrigSelection.Shapes(1).SizeWidth > 54 Then OrigSelection.Shapes(1).SizeWidth = 54
    
'    Dim s1 As Shape
 '   Dim p As Page
  '  For Each p In ActiveDocument.Pages
   ''     p.Activate
     '   Set s1 = ActiveLayer.CreateRectangle(2.112248, 9.162661, 6.246106, 3.316205)
      '  s1.Rectangle.CornerType = cdrCornerTypeRound
       ' s1.Rectangle.RelativeCornerScaling = True
        's1.Fill.ApplyNoFill
       ' s1.Outline.SetPropertiesEx 0.003, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 75), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#, Justification:=cdrOutlineJustificationMiddle
    'Next
End Sub
Sub CopyText_Replace()
    ' Recorded 29/03/2016
    ' Ctrl + Numpad_9
    ActiveDocument.Unit = cdrMillimeter
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    If OrigSelection.Shapes(1).Text.Story.Text = "AAAA" Then
        OrigSelection.Shapes(1).Text.Story.Text = OrigSelection.Shapes(2).Text.Story.Text
        OrigSelection.Shapes(1).Text.Story.ChangeCase cdrTextUpperCase
    Else
        OrigSelection.Shapes(1).Text.Replace "Abc", OrigSelection.Shapes(2).Text.Story.Text, True
    End If
    OrigSelection.Shapes(2).Selected = False
    OrigSelection.Shapes(2).Delete
End Sub
Sub Num_Page()
    Dim p As Page, s As Shape
    For Each s In ActivePage.Shapes.FindShapes(, cdrTextShape)
        If s.Text.Story.Font = "VNI-Avo" Then
            s.Text.Story.Text = Right("00" & ActivePage.index - 2, 2)
        End If
    Next
End Sub
Sub check_range()
    ' Recorded 7/4/2015
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrCenter
    Dim cC As Page
    For Each cC In ActiveDocument.Pages
        If cC.Shapes.count <> 1 Then MsgBox cC.Name & " have " & cC.Shapes.count & " object", vbOKOnly, "Failed!"
        'If cc.Shapes.All.SizeHeight <> 225 Or cc.Shapes.All.SizeWidth <> 150 Then
            'cc.Activate
            'MsgBox cc.Shapes.All.SizeHeight & vbCrLf & cc.Shapes.All.SizeWidth, vbOKOnly, cc.Name
            'cc.Shapes.All.SizeHeight = 225
            'cc.Shapes.All.SizeWidth = 150
        'End If
    Next cC
    
End Sub
Sub Edit_PageNumber()
    Dim p As Page, s As Shape
    For Each s In ActivePage.Shapes.FindShapes(, cdrTextShape)
        If s.Text.Story.Font = "VNI-Avo" Then
            s.Text.Story.Text = ActivePage.index - 2
        End If
    Next
End Sub
Public Function Arccos(x) As Double
    If Round(x, 8) = 1# Then Arccos = 0#: Exit Function
    If Round(x, 8) = -1# Then Arccos = PI: Exit Function
    Arccos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
End Function
Sub Macro11()
    ' Recorded 16/04/2016
    Dim OrigSelection As ShapeRange
    ActiveDocument.Unit = cdrMillimeter
    Dim a, b, c As Double
    
    Set OrigSelection = ActiveSelectionRange
    a = OrigSelection.Shapes(1).Curve.Length
    MsgBox a
    b = OrigSelection.SizeWidth
    c = OrigSelection.SizeHeight
    
End Sub

Sub Crop_to_A4()
    ' Recorded 23/04/2016
    Dim OrigSelection As Shape
    
    Set OrigSelection = ActiveSelectionRange.Shapes(1)
    Dim s1 As Shape
    Set s1 = ActiveLayer.CreateRectangle(0#, 11.680551, 8.26389, 0#)
    s1.Rectangle.CornerType = cdrCornerTypeRound
    s1.Rectangle.RelativeCornerScaling = True
    s1.Fill.ApplyNoFill
    s1.Outline.SetPropertiesEx 0.003, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#, Justification:=cdrOutlineJustificationMiddle
    Dim s2 As Shape
    Set s2 = s1.Intersect(OrigSelection, True, True)
    OrigSelection.Delete
    s1.Delete
    s2.OrderToFront
End Sub
Sub ssts()
    ActiveDocument.ReferencePoint
    MsgBox cdrMiddleLeft & vbCrLf & _
        cdrMiddleRight & vbCrLf
    
End Sub
Sub watermark_remove()
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrCenter
    Dim cC As Page
    For Each cC In ActiveDocument.Pages
        'If cc.Layers("Layer 1").Shapes(1).SizeHeight < 150 Then cc.Layers("Layer 1").Shapes(1).Delete
        If cC.Layers(2).Name = "Warning: Latest changes must be saved before extracting." Then cC.Layers(2).Delete
    Next cC
End Sub
Sub page_trim()
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrCenter
    Dim cC As Page
    Dim s1 As Shape
    For Each cC In ActiveDocument.Pages
        cC.Activate
        cC.Shapes.All.Move 0#, 7#
        s1 = ActiveDocument.MasterPage.DesktopLayer.Shapes(1).Trim(ActiveLayer.Shapes(1), True, True)
        ActiveLayer.Shapes(2).Delete
        Exit Sub
    Next cC
End Sub
Sub cell_connect()
    If ActiveSelection.Shapes.count <> 2 Then Exit Sub
    Dim s1 As Shape
    Dim s2 As Shape
    Dim canh As Integer
    Set s2 = ActiveSelection.Shapes(2)
    If ActiveSelectionRange(2).SizeWidth < ActiveSelectionRange(2).SizeHeight Then
        'MsgBox "1: " & ActiveSelectionRange(1).SizeWidth & vbCrLf & ActiveSelectionRange(1).SizeHeight
        Set s1 = ActiveLayer.CreateRightAngleConnector(ActiveSelectionRange(2).SnapPoints.Object(cdrObjectPointBottom), ActiveSelectionRange(1).SnapPoints.Edge(4, 0.5))
    Else
        'MsgBox "2: " & ActiveSelectionRange(1).SizeWidth & vbCrLf & ActiveSelectionRange(1).SizeHeight
        Set s1 = ActiveLayer.CreateRightAngleConnector(ActiveSelectionRange(2).SnapPoints.Object(cdrObjectPointRight), ActiveSelectionRange(1).SnapPoints.Edge(4, 0.5))
    End If
    s2.CreateSelection
End Sub


Sub Macro9()
    ' Recorded 14/07/2016
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    Dim s1 As Shape
    Set s1 = ActiveLayer.CreateRightAngleConnector(OrigSelection(2).SnapPoints.Object(cdrObjectPointLeft), OrigSelection(1).SnapPoints.Edge(2, 0.499995))
End Sub
Sub Macro10()
    ' Recorded 14/07/2016
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    Dim s1 As Shape
    Set s1 = ActiveLayer.CreateRightAngleConnector(OrigSelection(1).SnapPoints.Object(cdrObjectPointBottom), ActiveLayer.Shapes(3).SnapPoints.Edge(2, 0.499995))
End Sub
Sub Macro12()
    ' Recorded 14/07/2016
    ActiveDocument.Unit = cdrMillimeter
    Dim s As Shape
    For Each s In ActivePage.Shapes
        If s.Shapes(2).SizeWidth > 44 Then
            s.Shapes(2).SizeWidth = 44
        End If
        s.Move 56.5 * InStrRev(s.Shapes(2).Text.Story.Text, vbTab), 0#
    Next
End Sub
Sub Num_Page_2()
    Dim p As Page, s As Shape
    ActiveDocument.BeginCommandGroup "Page Number"
    Dim cC As Page
    For Each cC In ActiveDocument.Pages
        cC.Activate
        For Each s In ActivePage.Shapes.FindShapes(, cdrTextShape)
            If s.Text.Story.Text = "01" Then
                s.Text.Story.Text = Right("00" & ActivePage.index, 2)
            End If
        Next s
    Next cC
    ActiveDocument.EndCommandGroup
End Sub
Sub CenterPage()
    ' Recorded 31/08/2016
    Dim p As Page, s As Shape
    ActiveDocument.BeginCommandGroup "Page Center"
    Dim cC As Page
    For Each cC In ActiveDocument.Pages
        cC.Activate
        cC.Shapes.All.AlignAndDistribute 3, 3, 2, 0, False, 2
    Next cC
    ActiveDocument.EndCommandGroup
End Sub
Sub bb()
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    OrigSelection.SetOutlineProperties BehindFill:=cdrTrue, LineCaps:=cdrOutlineRoundLineCaps, LineJoin:=cdrOutlineRoundLineJoin
End Sub

Sub SizeCDNTNDT()
    Dim p As Page, s As Shape
    ActiveDocument.BeginCommandGroup "QDS"
    Dim cC As Page
    For Each cC In ActiveDocument.Pages
        If cC.Shapes(1).Type = cdrTextShape Then
            cC.Shapes(1).Text.Story.Size = 12
        End If
    Next
    ActiveDocument.EndCommandGroup
End Sub
Public Sub Table()
    Dim i As Integer
    Dim t_row As Integer
    Dim chon As ShapeRange
    Set chon = ActiveSelectionRange
    t_row = chon(1).Custom.rows.count
    For i = 1 To 8
        chon(1).Custom.Columns(i).Width = chon(2).Custom.Columns(i).Width
    Next
    For i = 1 To 8
        chon(1).Custom.cell(i, 1).TextShape.Text.Story.Bold = True
        chon(1).Custom.cell(i, 1).TextShape.Text.Story.Alignment = cdrCenterAlignment
        chon(1).Custom.cell(i, t_row).TextShape.Text.Story.Bold = True
        chon(1).Custom.cell(i, t_row).TextShape.Text.Story.Alignment = cdrCenterAlignment
    Next
    For i = 2 To t_row
        chon(1).Custom.cell(1, i).TextShape.Text.Story.Alignment = cdrCenterAlignment
        chon(1).Custom.cell(3, i).TextShape.Text.Story.Alignment = cdrCenterAlignment
        chon(1).Custom.cell(4, i).TextShape.Text.Story.Alignment = cdrCenterAlignment
        chon(1).Custom.cell(5, i).TextShape.Text.Story.Alignment = cdrCenterAlignment
        chon(1).Custom.cell(6, i).TextShape.Text.Story.Alignment = cdrCenterAlignment
        chon(1).Custom.cell(7, i).TextShape.Text.Story.Alignment = cdrRightAlignment
        chon(1).Custom.cell(8, i).TextShape.Text.Story.Alignment = cdrRightAlignment
    Next
    For i = 1 To t_row
        chon(1).Custom.rows(i).Height = chon(2).Custom.rows(i).Height
    Next
    'Dim t As Shape
    't.Text.Story.Alignment = cdrRightAlignment
End Sub
Sub Num_Page_Dau_Thu_Y()
    Dim p As Page, s As Shape, stt As Integer, tinh As String
    tinh = InputBox("Ma tinh", "ma tinh")
    stt = 1
    For Each s In ActiveSelection.Shapes
        s.Text.Story.Text = "MAÕ SOÁ: 40." & Right("00" & tinh, 2) & "." & Right("00" & stt, 2)
        stt = stt + 1
    Next
End Sub
Sub auto_number()
    Optimization = True
    ActiveDocument.BeginCommandGroup "Auto Number"
    Dim p As Page, s As Shape, stt As Integer, tinh As String
    stt = 1
    For Each s In ActiveSelection.Shapes
        s.Text.Story.Text = Right("000" & stt, 2)
        stt = stt + 1
    Next
    ActiveDocument.EndCommandGroup
    Optimization = False
    Refresh
End Sub
Sub Macro14()
    ' Recorded 11/15/2016
    ActiveLayer.Editable = False
End Sub
Sub text_Counter()
    ' Recorded 11/16/2016
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    Dim st As String
    Dim i As Integer
    Dim html As String
    st = CInt(InputBox("So bat dau", "STT cho danh sach"))
    html = st
    For i = st + 1 To st + 17
        html = html & vbCrLf & i
    Next
    OrigSelection.Shapes(1).Text.Story.Text = html
End Sub
Public Function FindReplace(ByVal str As String, ByVal toFind As String, ByVal toReplace As String) As String
    Dim i As Integer
    For i = 1 To Len(str)
        If Mid(str, i, Len(toFind)) = toFind Then   ' does the string match?
            FindReplace = FindReplace & toReplace               ' add the new replacement to the final result
            i = i + (Len(toFind) - 1)               ' move to the character after the toFind
        Else
            FindReplace = FindReplace & Mid(str, i, 1)        ' add a character
        End If
    Next i
End Function
Public Sub TextTranslate()
    Dim huyen As Integer
    Dim s As Shape
    huyen = InputBox("Ma so huyen", "Huyen")
    ActiveDocument.BeginCommandGroup "Text Translate" & Right("00" & huyen, 2)
    For Each s In ActiveSelectionRange.Shapes
        If s.Type = cdrTextShape Then
            If s.Text.Story = "55." Then s.Text.Story = Right("00" & huyen, 2)
        End If
    Next s
    ActiveDocument.EndCommandGroup
End Sub
Sub Copy_Header()
    ' Recorded 12/23/2016
    Dim s1 As Shape
    ActiveDocument.Unit = cdrMillimeter
    If ActivePage.index Mod 2 = 0 Then
        Set s1 = ActiveDocument.MasterPage.DesktopLayer.Shapes(1).Duplicate
        s1.Move -200, 0#
    Else
        Set s1 = ActiveDocument.MasterPage.DesktopLayer.Shapes(2).Duplicate
        s1.Move 200, 0#
    End If
    
End Sub
Sub Page_number()
    On Error Resume Next
    Dim p As Page, s As Shape
    ActiveDocument.BeginCommandGroup "SoTrang"
    Dim cC As Page
    For Each cC In ActiveDocument.Pages
        cC.Activate
        If ActiveLayer.Shapes(1).Shapes(1).Text.Story.Text = ActivePage.index Then
            ActiveLayer.Shapes(1).Shapes(1).Text.Story.Size = 17
        End If
        'ActiveLayer.Shapes(1).Shapes(1).Text.Story.Text = ActivePage.Index
    Next
    ActiveDocument.EndCommandGroup
End Sub
Sub move_desktop_layer()
    ' Recorded 1/5/2017
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    OrigSelection.MoveToLayer
    ActiveDocument.Pages(2).Activate
End Sub

Sub save_as_X7()
    ' Recorded 09/03/2017
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    Dim SaveOptions As StructSaveAsOptions
    Set SaveOptions = CreateStructSaveAsOptions
    With SaveOptions
        .EmbedVBAProject = True
        .Filter = cdrCDR
        .IncludeCMXData = False
        .Range = cdrAllPages
        .EmbedICCProfile = True
        .Version = cdrVersion17
        .KeepAppearance = True
    End With
    ActiveDocument.SaveAs "D:\GDVH 2001-29sss83.cdr", SaveOptions
    MsgBox "Hello world"
End Sub
Sub groupLikeMeS()
    ActiveDocument.BeginCommandGroup "ss"
    Optimization = True
    ActiveDocument.Unit = cdrMillimeter
    Dim sr As Shapes
    Set sr = ActivePage.Shapes.FindShapes(Query:="@outline = {1.5 mm} ").Shapes
    sr.All.SetOutlineProperties Width:=4

    Optimization = False
    ActiveDocument.EndCommandGroup
    Refresh
    MsgBox "Group Finish"
End Sub
Sub RemoveTrans()
    ' Recorded 13/04/2017
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    OrigSelection(1).Shapes(2).Style.StringAssign "{""fill"":{""type"":""1"",""overprint"":""0"",""primaryColor"":""CMYK255,USER,0,0,255,0,100,cccd19cb-4675-4a5e-8bda-d0bbbaab8af0"",""secondaryColor"":""CMYK,USER,0,0,0,0,100,00000000-0000-0000-0000-000000000000""},""outline"":{""width"":""199"",""color"":""CMYK255,USER,255,0,255,0,100,cccd19cb-4675-4a5e-8bda-d0bbbaab8af0""},""transparency"":{""fill"":{""type"":""0"",""overprint"":""0"",""fillName"":null},""uniformTransparency"":""1"",""appliesTo"":""0""}}"
End Sub

Sub page_move()
    ActiveDocument.BeginCommandGroup "Page Move"
    Optimization = True
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrCenter
    Dim cC As Page
    Dim s1 As Shape
    For Each cC In ActiveDocument.Pages
        cC.Activate
        cC.Shapes.All.Move 0#, 0.5
    Next cC
    Optimization = False
    ActiveDocument.EndCommandGroup
    MsgBox "ok"
End Sub

Sub CreateA4_Border()
Attribute CreateA4_Border.VB_Description = "Tao Khung A4 cho trang"
    ' Recorded 10/05/2017
    '
    ' Description:
    '     Tao Khung A4 cho trang
    ActiveDocument.BeginCommandGroup "Page Border"
    Optimization = True
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrCenter
    Dim cC As Page
    Dim s1 As Shape
    For Each cC In ActiveDocument.Pages
        cC.Activate
        Set s1 = ActiveLayer.CreateRectangle(0#, 210, 297, 0#)
        s1.Rectangle.CornerType = cdrCornerTypeRound
        s1.Rectangle.RelativeCornerScaling = True
        s1.Fill.ApplyNoFill
        s1.Outline.SetPropertiesEx 0.15, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 80), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=45#, Justification:=cdrOutlineJustificationMiddle
    Next cC
    Optimization = False
    ActiveDocument.EndCommandGroup
    MsgBox "ok"
End Sub
Sub Import_eps()
    ' Recorded 18/05/2017
    ActiveDocument.BeginCommandGroup "Page import"
    Optimization = True
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrCenter
    Dim cC As Page
    Dim s1 As Shape
    Dim c As Integer
    Dim impflt As ImportFilter
        
    c = 1
    For Each cC In ActiveDocument.Pages
        cC.Activate
        Set s1 = ActiveLayer.CreateRectangle(15.137795, 22.401579, 31.673224, 10.708661)
        s1.Rectangle.CornerType = cdrCornerTypeRound
        s1.Rectangle.RelativeCornerScaling = True
        s1.Fill.ApplyNoFill
        s1.Outline.SetPropertiesEx 0.003, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#, Justification:=cdrOutlineJustificationMiddle
        Dim impopt As StructImportOptions
        Set impopt = CreateStructImportOptions
        With impopt
            .MaintainLayers = True
            With .ColorConversionOptions
                .SourceColorProfileList = "sRGB IEC61966-2.1,U.S. Web Coated (SWOP) v2,Dot Gain 20%"
                .TargetColorProfileList = "sRGB IEC61966-2.1,U.S. Web Coated (SWOP) v2,Dot Gain 20%"
            End With
        End With
        Set impflt = ActiveLayer.ImportEx("E:\Users\Desktop\so diem dong gay\Sach\Full" & c & ".eps", cdrPSInterpreted, impopt)
        impflt.Finish
        'Dim s2 As Shape
        'Set s2 = ActiveShape
        's2.Move 0.021146, 0.131917
        c = c + 1
    Next cC
    Optimization = False
    ActiveDocument.EndCommandGroup
    MsgBox "ok"
End Sub
Sub HalfTone_bg()
    ' Recorded 8/21/2017
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    OrigSelection.ApplyUniformFill CreateCMYKColor(0, 0, 0, 20)
    OrigSelection(1).Style.StringAssign "{""fill"":{""type"":""1"",""overprint"":""0"",""primaryColor"":""CMYK255,USER,0,0,0,51,100,cccd19cb-4675-4a5e-8bda-d0bbbaab8af0"",""screenSpec"":""0,0,45000000,60,0"",""winding"":""1""},""outline"":{""overprint"":""0"",""angle"":""0"",""screenSpec"":""0,0,45000000,60,0"",""behindFill"":""0"",""scaleWithObject"":""1"",""overlapArrow"":""0"",""shareArrow"":""0"",""endCaps"":""0"",""joinType"":""0"",""width"":""0"",""aspect"":""100"",""matrix"":""1,0,0,0,1,0"",""color"":""GRAY255,USER,0,100,00000000-0000-0000-0000-000000000000"",""dashDotSpec"":""0"",""leftArrow"":""|0"",""leftArrowAttributes"":""0|0|0|0|0|0|0"",""rightArrow"":""|0"",""rightArrowAttributes"":""0|0|0|0|0|0|0"",""dotLength"":""0"",""miterLimit"":""11.478339999999999"",""justification"":""0""},""transparency"":{""fill"":{""type""" & _
                ":""1"",""overprint"":""0"",""primaryColor"":""RGB255,USER,0,0,0,100,00000000-0000-0000-0000-000000000000"",""screenSpec"":""0,0,45000000,60,0""},""mode"":""0"",""uniformTransparency"":""0.3"",""startTransparency"":""0.5"",""endTransparency"":""0"",""appliesTo"":""2""}}"
    Dim s1 As Shape
    Set s1 = OrigSelection.ConvertToBitmapEx(5, False, True, 400, 1, True, False, 95)
    s1.Selected = True
    s1.Bitmap.ConvertToBW cdrRenderHalftone, 45, 90, cdrHalftoneRound, 135, 0
End Sub
Sub Mirror_PageNumber()
    ActiveDocument.BeginCommandGroup "Page mirror"
    Optimization = True
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrCenter
    Dim p As Page, s As Shape
    Dim i As Integer
    For i = 2 To ActiveDocument.Pages.count Step 2
        ActiveDocument.Pages(i).Shapes(1).Move 32, 0#
    Next
    Optimization = False
    ActiveDocument.EndCommandGroup
    MsgBox "ok"
End Sub
Sub Macro16_web_img_import()
    ' Recorded 11/13/2017
    Dim impopt As StructImportOptions
    Set impopt = CreateStructImportOptions
    With impopt
        .Mode = cdrImportFull
        .MaintainLayers = True
        With .ColorConversionOptions
            .SourceColorProfileList = "sRGB IEC61966-2.1,U.S. Web Coated (SWOP) v2,Dot Gain 20%"
            .TargetColorProfileList = "sRGB IEC61966-2.1,U.S. Web Coated (SWOP) v2,Dot Gain 20%"
        End With
    End With
    Dim impflt As ImportFilter
    Set impflt = ActiveLayer.ImportEx("C:\Users\Hung\AppData\Local\Microsoft\Windows\Temporary Internet Files\Content.IE5\PGGE3WD5\staticmap[1]", cdrPNG, impopt)
    impflt.Finish
    Dim s1 As Shape
    Set s1 = ActiveShape
    ActivePage.Shapes.All.CreateSelection
    ActiveSelection.Move 701.9581, 438.5907
End Sub
Sub Macro16()
    ' Recorded 11/13/2017
    MsgBox ActiveDocument.SourceFileVersion
End Sub
Sub Add_giaykhen_2_bg()
    ' Recorded 11/16/2017
    ActiveLayer.Paste
    Dim Paste1 As ShapeRange
    Set Paste1 = ActiveSelectionRange
    Dim lr1 As Layer
    Set lr1 = ActivePage.CreateLayer("bg")
    lr1.Master = True
    ActivePage.Layers("bg").Activate
    lr1.Name = "bg"
    Paste1.MoveToLayer ActiveLayer
    ActiveLayer.Editable = False
End Sub
Sub Macro17_cmyk()
    ' Recorded 12/5/2017
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    OrigSelection(1).Bitmap.ConvertTo 5
    ActiveLayer.Shapes(8).Bitmap.ConvertTo cdrCMYKColorImage
End Sub
Sub fttp()
    ' Fit text to path
    Dim ot As ShapeRange
    Dim ef As Effect
    ActiveDocument.Unit = cdrMillimeter
    Set ot = ActiveSelectionRange
    Set ef = ot(2).Text.FitToPath(ot(1))
    ef.TextOnPath.DistanceFromPath = 1.4
    ef.TextOnPath.Offset = 4
End Sub

Sub page_move_1()
    ' Recorded 7/4/2015
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrCenter
    Dim cC As Page
    For Each cC In ActiveDocument.Pages
        cC.Shapes.All.Move 0#, -2
    Next cC
End Sub
Sub change_tone()
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.BeginCommandGroup "Recolor"
    Optimization = True
    
    Dim x As Shape
    Dim tmp As Integer
    For Each x In ActiveSelectionRange.Shapes
        If x.Fill.UniformColor.IsCMYK Then
            tmp = x.Fill.UniformColor.CMYKBlack
            x.Fill.UniformColor.CMYKMagenta = tmp
            x.Fill.UniformColor.CMYKBlack = 0
        End If
    Next
    
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
    Application.Refresh
End Sub
Sub change_tone_2()
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.BeginCommandGroup "Recolor"
    Optimization = True
    
    Dim x As Shape
    Dim tmp As Integer
    For Each x In ActiveSelectionRange.Shapes
        If x.Fill.UniformColor.IsCMYK Then
            If x.Fill.UniformColor.CMYKCyan > 0 Then
                tmp = x.Fill.UniformColor.CMYKBlack
                x.Fill.UniformColor.CMYKMagenta = x.Fill.UniformColor.CMYKCyan
                x.Fill.UniformColor.CMYKCyan = 0
            End If
        End If
    Next
    
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
    Application.Refresh
End Sub
Sub resample1()
    Dim sr As ShapeRange, p As Page, s As Shape
    For Each p In ActiveDocument.Pages
        p.Activate
        Set sr = ActivePage.Shapes.FindShapes(, cdrBitmapShape)
        For Each s In sr
              s.Bitmap.Resample , , , 600, 600
        Next s
        'sr.Move 0, 0.4
    Next p
End Sub
Sub resample2()
'    ActiveSelection.Shapes(1).Bitmap.Resample , , , 300, 300
    Call resample3
End Sub
Sub resample3(Optional d As Integer = 300)
    'Giam do phan giai hinh anh qua 300dpi
    'Resample big image resolution
    ActiveDocument.BeginCommandGroup "Resample"
    Optimization = True
    Dim sr As ShapeRange, p As Page, s As Shape, t As Shape
    For Each p In ActiveDocument.Pages
        p.Activate
        Set sr = ActivePage.Shapes.FindShapes(, cdrBitmapShape)
        For Each s In sr
            If s.Bitmap.ResolutionX > d Then
                s.Bitmap.Crop
                s.Bitmap.Resample , , , d, d
            End If
        Next s
        For Each t In ActivePage.Shapes.FindShapes(Query:="!@com.powerclip.IsNull")
            Set sr = t.PowerClip.Shapes.FindShapes(, cdrBitmapShape)
            For Each s In sr
                If s.Bitmap.ResolutionX > d Then
                    s.Bitmap.Crop
                    s.Bitmap.Resample , , , d, d
                End If
            Next s
        Next t
    Next p
    
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
    Application.Refresh
    MsgBox "Resample image finish", vbOKOnly, "Resample"
End Sub
Sub resampleSelection()
    'Giam do phan giai hinh anh qua 400dpi
    'Resample big image resolution
    ActiveDocument.BeginCommandGroup "Resample"
    Optimization = True
    Dim sr As ShapeRange, p As Page, s As Shape
    Dim d As Integer
    Set sr = ActiveSelection.Shapes.FindShapes(, cdrBitmapShape)
    d = 400
    For Each s In sr
        If s.Bitmap.ResolutionX > d Then
            s.Bitmap.Resample , , , d, d
        End If
    Next s
    
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
    Application.Refresh
    MsgBox "Resample image finish", vbOKOnly, "Resample"
End Sub
Sub CMYK_Pallete_test()
    ' Recorded 6/24/2016
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.BeginCommandGroup "Pallet test"
    Optimization = True
    Dim s1 As Shape
    Dim c, M, y, k, CMY, black
    Dim pX, pY As Double
    CMY = Array(0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100)
    'CMY = Array(20, 40, 60, 80, 100)
    black = Array(0, 10, 20, 30, 40, 50)
    pX = 5
    pY = 5
    For Each c In CMY
        For Each M In CMY
            For Each y In CMY
                For Each k In black
                   Set s1 = ActiveLayer.CreateRectangle(pX, pY, pX + 3, pY + 2.3)
                   s1.Fill.ApplyUniformFill CreateCMYKColor(c, M, y, k)
                   pX = pX + 3
                Next
            Next
            pY = pY + 2.3
            pX = 5
        Next
    Next
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveLayer.Shapes.All.SetOutlineProperties Width:=0
    ActiveWindow.Refresh
    Application.Refresh
End Sub
Sub reduceNode()
    ' Recorded 3/27/2019
    Dim OrigSelection As ShapeRange
    Dim a As Shape
    Set OrigSelection = ActiveSelectionRange
    For Each a In ActiveSelectionRange
        a.Curve.AutoReduceNodes
    Next
End Sub
Sub Delete_all_lock()
    ' Fix locked object before import svg
    Dim x As Shape
    For Each x In ActiveLayer.Shapes
        If x.Type <> 7 Then
            x.Locked = False
            x.Delete
        End If
    Next
    MsgBox ActiveLayer.Shapes.count
End Sub

Sub MatchColorsToPalette()
    Dim s As Shape
    For Each s In ActiveSelectionRange
        s.Fill.UniformColor.ConvertToRGB
    Next s
End Sub
Sub LenPath()
    ' Recorded 16/04/2016
    ActiveDocument.Unit = cdrMillimeter
    MsgBox ActiveSelectionRange.Shapes(1).Curve.Length & " mm", , "Length path"
End Sub

Sub pal_get()
    Dim i, j, x As Integer
    x = 1
    For i = 1 To Application.Palettes.count
        If Application.Palettes(i).Name <> "Document Palette" Then
            For j = 1 To Application.Palettes(i).ColorCount
                ActiveLayer.Shapes(x).Fill.ApplyUniformFill Application.Palettes(i).Color(j)
                x = x + 1
            Next
        End If
    Next
End Sub
Sub shape_line()
    ActiveDocument.Unit = cdrMillimeter
    ActivePage.Shapes.FindShapes(Query:="@outline.width < {0.1 mm}").CreateSelection
    ActiveSelection.Outline.Width = 0.1
End Sub

Sub anc()
    MsgBox ActiveLayer.Shapes.count
End Sub

Sub Convert_color()
    ' Recorded 4/4/2019
    'MsgBox ActiveLayer.Shapes(3).Type
    Optimization = False
    ActiveDocument.EndCommandGroup
    Refresh
    Exit Sub
    Dim x As Shape
    For Each x In ActiveSelection.Shapes
        x.Fill.UniformColor.ConvertToGray
    Next
End Sub
Sub Macro23()
    ' Recorded 10/18/2019
    Dim file
    Dim i As Integer
    file = Split("trang_1.psd,trang 2.psd,trang_3.psd,trang_4.psd,tang_5.psd,trang_6.psd,trang_7.psd,trang_8.psd,trang_9.psd,trang_10.psd,trang_11.psd,trang_12.psd,trang_13.psd,tang_14.psd,tang_15.psd,tang_16.psd,tang_17.psd,tang_18.psd,tang_19.psd,tang_20.psd,tang_21.psd,trang 22.psd,tang_24.psd", ",")
    Dim impopt As StructImportOptions
    Set impopt = CreateStructImportOptions
    With impopt
        .Mode = cdrImportFull
        .MaintainLayers = True
        With .ColorConversionOptions
            .SourceColorProfileList = "sRGB IEC61966-2.1 (Linear RGB Profile),U.S. Web Coated (SWOP) v2,Dot Gain 20%"
            .TargetColorProfileList = "sRGB IEC61966-2.1,U.S. Web Coated (SWOP) v2,Dot Gain 20%"
        End With
    End With
    Dim impflt As ImportFilter
    For i = 1 To UBound(file)
        Set impflt = ActiveLayer.ImportEx("E:\User\desktop\Hoang Thy 2\" & file(i), cdrPSD, impopt)
        impflt.Finish
    Next
End Sub
Sub Macro24()
    ' Recorded 10/19/2019
    Dim impopt As StructImportOptions
    Set impopt = CreateStructImportOptions
    With impopt
        .Mode = cdrImportFull
        .MaintainLayers = True
        With .ColorConversionOptions
            .SourceColorProfileList = "sRGB IEC61966-2.1 (Linear RGB Profile),U.S. Web Coated (SWOP) v2,Dot Gain 20%"
            .TargetColorProfileList = "sRGB IEC61966-2.1 (Linear RGB Profile),U.S. Web Coated (SWOP) v2,Dot Gain 20%"
            '.TargetColorProfileList = "sRGB IEC61966-2.1,U.S. Web Coated (SWOP) v2,Dot Gain 20%"
        End With
    End With
    Dim impflt As ImportFilter
    Set impflt = ActiveLayer.ImportEx("E:\User\desktop\Hoang Thy 2\tang_20.psd", cdrPSD, impopt)
    impflt.Finish

End Sub
Sub find_Arial()

    Dim sr As ShapeRange

    Set sr = ActivePage.Shapes.FindShapes(, cdrTextShape, True, "@com.text.story.size < '9'")
    sr.Shapes(1).Text.Story.Size = 18
    sr.Shapes(1).Text.Story.Name = 18
End Sub
Sub txt_unicode()
    Dim objStream
    Dim i, fi, mStep As Integer
    fi = GetSetting(ssAppName, ssSection, "Merge_page", 1)
    mStep = 200
    
    Set objStream = CreateObject("ADODB.Stream")
    objStream.CharSet = "utf-16le"
    objStream.Open
    objStream.WriteText ChrW(&HFEFF)
    objStream.WriteText "2" & vbCrLf
    objStream.WriteText "\s1\\s2\" & vbCrLf
    For i = fi To fi + mStep - 1
        objStream.WriteText "\" & Right("00000" & i, 6) & "\\" & Right("00000" & i + 2500, 6) & "\" & vbCrLf
    Next
    objStream.SaveToFile "E:\User\desktop\DS Eadrong_print.txt", 2
    SaveSetting ssAppName, ssSection, "Merge_page", fi + mStep
End Sub
Sub set_ss()
    SaveSetting ssAppName, ssSection, "Merge_page", 1101
End Sub
Sub name_by_me()
    ' Recorded 11/11/2019
    Dim i As Integer
    Dim x As String
    For i = 1 To 3
        x = ActiveSelection.Shapes(i).ObjectData("Name").Value
        ActiveSelection.Shapes(i).ObjectData("Name").Value = Replace(x, "1", 4)
    Next
End Sub
Sub num_by_me()
    Dim i As Integer
    ActiveDocument.BeginCommandGroup "Auto Num for merge"
    Optimization = True
    Dim p As Page
    For Each p In ActiveDocument.Pages
        p.Activate
        For i = 1 To 4
            ActiveLayer.Shapes("fr" & i).Text.Story = (ActiveLayer.Shapes("so" & i).Text.Story - 1) * 50 + 1
            ActiveLayer.Shapes("to" & i).Text.Story = (ActiveLayer.Shapes("so" & i).Text.Story) * 50
        Next
    Next
    ActiveDocument.EndCommandGroup
    Optimization = False
    Refresh
End Sub
Sub change_tone_3()
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.BeginCommandGroup "Recolor"
    Optimization = True
    
    Dim x As Shape
    Dim t As Integer
    Dim tmp As Integer
    t = 30
    For Each x In ActiveSelectionRange.Shapes
        'x.Fill.UniformColor.CMYKMagenta = t
        x.Fill.UniformColor.CMYKBlack = t
        t = t + 5
    Next
    
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
    Application.Refresh
End Sub
Sub Tran_Test()
    ' Recorded 11/26/2019
    Dim OrigSelection As ShapeRange
    Dim i As Integer
    Set OrigSelection = ActiveSelectionRange
    For i = 1 To 19
        OrigSelection(i).Style.StringAssign "{""transparency"":{""fill"":{""type"":""1"",""fillName"":null},""uniformTransparency"":""" & (i * 5 / 100) & """}}"
    Next
End Sub
Sub Edit_font()
    ' Change Font name to: Myriad Pro Cond
    
    Dim oFont As ShapeRange
    Set oFont = ActiveSelectionRange
    oFont.Shapes(1).Text.Story.Font = "Myriad Pro Cond"
    oFont.Shapes(1).Text.Story.Bold = False
End Sub
Sub Macro2ss()
    Dim opp As Shape
    Set opp = ActiveSelectionRange(1)
    opp.Fill.UniformColor
End Sub
Sub Change_Color()
    ActiveDocument.BeginCommandGroup "[05]Change Color"
    ActiveDocument.Unit = cdrMillimeter
    Dim sr As Shape
    For Each sr In ActivePage.Shapes.FindShapes(Query:="@width = {420 mm} & @height = {297 mm}").Shapes
        sr.Fill.UniformColor.CMYKAssign 10, 80, 20, 0
    Next
    ActiveDocument.EndCommandGroup
    Refresh
End Sub
Sub Change_Color_RBG()
    ActiveDocument.BeginCommandGroup "[05]Change Color"
    ActiveDocument.Unit = cdrMillimeter
    Dim cRed, k, c, j As Integer
    cRed = 100
    Dim sr As Shape
    For Each sr In ActiveSelection.Shapes
        sr.Fill.UniformColor.RGBBlue = 10
        cRed = cRed + 5
    Next
    ActiveDocument.EndCommandGroup
    Refresh
End Sub

Sub copyCDR_ID()
    MsgBox Application.Clipboard.Parent
    Dim oFont As ShapeRange
    Set oFont = ActiveSelectionRange
    oFont.Shapes(1).Text.Story.Font = "Myriad Pro Cond"
    oFont.Shapes(1).Text.Story.Bold = False
End Sub
Sub txt_ID()
    Dim objStream
    Dim i, fi, mStep As Integer
    Set objStream = CreateObject("ADODB.Stream")
    objStream.CharSet = "utf-8"
    objStream.Open
    objStream.WriteText ActiveSelectionRange.Shapes(1).Text.Story.Text & vbCrLf
    objStream.WriteText ActiveSelectionRange.Shapes(1).Text.Story.Size & ":" & ActiveSelectionRange.Shapes(1).Text.Story.Font & vbCrLf
    objStream.SaveToFile "C:\Program Files\Adobe\Adobe InDesign CC 2018\Scripts\Scripts Panel\Samples\JavaScript\clip.txt", 2
End Sub
Sub svg_unlock()
    For Each sr In ActivePage.Shapes.FindShapes(Query:="@name = 'SVG data'").Shapes
        sr.GetBoundingBox x, y, w, h
        ActivePage.SelectShapesFromRectangle(x, y, x + w, y + h, False).Group
        'sr.SizeWidth = 20
        'sr.SizeHeight = 60
    Next
End Sub
Sub tststs()
    'MsgBox ActivePage.Shapes.FindShapes(Query:="@name = 'SVG data'").Count
    'MsgBox Asc(ActiveSelection.Shapes(1).Text.Story)
    'ActiveDocument.MasterPage.DesktopLayer.Shapes("chan").Duplicate 0#, 74, 25
    Optimization = False
    ActiveDocument.EndCommandGroup
    Refresh
End Sub
Sub pg_move()
    Dim s1 As Shape
    Set s1 = ActiveSelection.Duplicate
    s1.Move 0#, 2.923228
    s1.OrderToBack
    Exit Sub
    Dim a As String
    a = InputBox("Trang", "Di chuyen")
    ActiveSelectionRange.MoveToLayer ActiveDocument.Pages(a).Layers("Layer 1")
    Refresh
End Sub
Sub Macro2()
    ' Recorded 5/9/2020
    ActiveDocument.BeginCommandGroup "Page Number"
    Optimization = True
    
    Dim s1 As Shape
    Dim i As Integer
    For i = 2 To 44
        ActiveDocument.Pages(i).Activate
        If i Mod 2 = 0 Then
            Set s1 = ActiveDocument.MasterPage.DesktopLayer.Shapes("chan").Duplicate(0#, 74.25)
            's1.Move 0#, 74.25
        Else
            Set s1 = ActiveDocument.MasterPage.DesktopLayer.Shapes("le").Duplicate(0#, 148.5)
            's1.Move 0#, 148.5
        End If
        s1.Shapes(1).Text.Story = i
    Next
    Optimization = False
    ActiveDocument.EndCommandGroup
    Refresh
End Sub
Sub poscripot()
    Dim s1 As Shape
    Dim x As Shape
    Dim w, h, i, j, k As Integer
    w = 10
    h = 10
    k = 1
    Set s1 = ActiveSelection
    For i = 0 To 13
        For j = 0 To 9
            ActiveSelection.Shapes(k).Fill.ApplyPostscriptFill(3).SetProperties w, h, 100, 100
            h = h + 5
            k = k + 1
        Next
        h = 10
        w = w + 2
    Next
End Sub

Sub Macro3()
    ' Recorded 6/26/2020
    Dim impopt As StructImportOptions
    Set impopt = CreateStructImportOptions
    With impopt
        .Mode = cdrImportFull
        .LinkBitmapExternally = True
        .MaintainLayers = True
        With .ColorConversionOptions
            .SourceColorProfileList = "sRGB IEC61966-2.1,U.S. Web Coated (SWOP) v2,Dot Gain 20%"
            .TargetColorProfileList = "sRGB IEC61966-2.1,U.S. Web Coated (SWOP) v2,Dot Gain 20%"
        End With
    End With
    Dim impflt As ImportFilter
    
    Dim sname As String
    Dim wo As Document
    Dim objStartFolder, objFolder, colFiles, objFSO, objFile
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objStartFolder = "E:\User\download\Caxinh\png\"
    Set objFolder = objFSO.GetFolder(objStartFolder)
    Set colFiles = objFolder.Files
    For Each objFile In colFiles
        Set impflt = ActiveLayer.ImportEx(objStartFolder & objFile.Name, cdrPNG, impopt)
        impflt.Finish
    Next
End Sub
Sub Macro4()
    ' Recorded 7/10/2020
    ActiveDocument.EndCommandGroup
    Optimization = False
    Refresh
End Sub
Sub Macro5()
    ' Recorded 8/13/2020
    ActiveDocument.ReferencePoint = cdrCenter
    ActivePage.Shapes.All.CreateSelection
    ActiveSelection.SetSize 346.1538, 250#
    
    ActiveDocument.BeginCommandGroup "Size S2"
    ActiveDocument.Unit = cdrMillimeter
    Dim cC As Page
    For Each cC In ActiveDocument.Pages
        cC.Activate
        ActivePage.Shapes.All.SetSize 346.1538, 250#
    Next cC
    ActiveDocument.EndCommandGroup
End Sub
Sub Macro7()
    ' Recorded 8/17/2020
    ActiveDocument.BeginCommandGroup "Fix MisPos"
    ActiveDocument.Unit = cdrMillimeter
    Dim cC As Page
    For Each cC In ActiveDocument.Pages
        cC.Activate
        If cC.index Mod 2 = 0 Then
            ActivePage.Shapes.All.Move 6, 1    'Trang chan: 2,4,6 (left/right, top/bottom)
        Else
            ActivePage.Shapes.All.Move -2, -2   'Trang le: 1,3,5..(left/right, top/bottom)
        End If
    Next cC
    ActiveDocument.EndCommandGroup
End Sub

Sub ab()
    'MsgBox ActiveSelection.Shapes(1).Shapes(1).Shapes(1).Shapes(1).Shapes(1).Outline.Color.ToString
    ActiveSelection.Shapes(2).Fill.ApplyUniformFill ActiveSelection.Shapes(1).Shapes(1).Shapes(1).Shapes(1).Shapes(1).Outline.Color
End Sub

Sub text_border()
    ' Recorded 8/24/2020
    ActiveDocument.BeginCommandGroup "Hairline"
    ActiveDocument.Unit = cdrMillimeter
    Dim cC As Page
    For Each cC In ActiveDocument.Pages
        cC.Activate
        ActivePage.Shapes.FindShapes(, cdrTextShape).All.SetOutlineProperties 0
    Next cC
    ActiveDocument.EndCommandGroup
End Sub
Sub text_Del()
    ' Recorded 8/24/2020
    ActiveDocument.BeginCommandGroup "Del text"
    ActiveDocument.Unit = cdrMillimeter
    Dim i
    For i = 1 To 21
        ActiveDocument.Pages(i).Activate
        ActivePage.Shapes.FindShapes(, cdrTextShape).All.Delete
        'ActivePage.Shapes.FindShapes(, cdrBitmapShape).All.Delete
    Next
    ActiveDocument.EndCommandGroup
End Sub
Sub page_delete()
    ActiveDocument.BeginCommandGroup "Del page"
    Dim pRange
    Dim i
    pRange = Split("1,2,3,4,9,10,13,15,17,18,21,27,28,29,30,32,33,34,36,38,41,42,43,45,46", ",")
    For i = 1 To UBound(pRange)
        ActiveDocument.Pages(pRange(UBound(pRange) - i)).Delete
    Next
    ActiveDocument.EndCommandGroup
End Sub
Sub GetRotate()
    Dim x As Double, y As Double, x1 As Double, y1 As Double, shift As Long
    Dim w As Double, h As Double, goc As Double, c As Double, count As Integer
    Dim PI
    PI = 3.14159265358979
    Dim fRadToDeg
    fRadToDeg = 180 / PI
    On Error Resume Next
    count = ActiveSelection.Shapes.count
    ActiveDocument.GetUserClick x, y, shift, 100, False, 351
    ActiveDocument.GetUserClick x1, y1, shift, 100, True, 351
    w = x1 - x
    h = y - y1
    c = Sqr(w * w + h * h)
    
    goc = fRadToDeg * ASin(h / c)
    If count >= 1 Then
        If goc < 0 Then
            goc = goc + 0.35
        Else
            goc = goc - 0.35
        End If
        ActiveSelection.Rotate goc
        MsgBox "goc"
    Else
        ActivePage.Shapes(1).Rotate goc
    End If
End Sub
Function ASin(val As Double) As Double
    ASin = 2 * Atn(val / (1 + Sqr(1 - (val * val))))
End Function

Sub Fix_pos()
    ' Recorded 8/17/2020
    ActiveDocument.BeginCommandGroup "Fix MisPos"
    ActiveDocument.Unit = cdrMillimeter
    Dim cC As Page
    For Each cC In ActiveDocument.Pages
        cC.Activate
        ActivePage.Shapes.All.Move -2, 2
        
'        If cC.Index Mod 2 = 0 Then
'            ActivePage.Shapes.All.Move 1, -4   'Trang chan: 2,4,6 (left/right, top/bottom)
'        Else
'            ActivePage.Shapes.All.Move 1, -4  'Trang le: 1,3,5..(left/right, top/bottom)
'        End If
    Next cC
    ActiveDocument.EndCommandGroup
End Sub
Sub Page_Sumary()
    ActiveDocument.BeginCommandGroup "Resample"
    Optimization = True
    
    Dim sr As ShapeRange, p As Page, s As Shape
    Dim d As Integer
    Dim html, html2 As String
    html = ""
    html2 = ""
    
    d = 300
    For Each p In ActiveDocument.Pages
        p.Activate
        Set sr = ActivePage.Shapes.FindShapes(, cdrBitmapShape)
        If sr.count > 0 Then
            html = html & "," & p.index
        Else
            html2 = html2 & "," & p.index
        End If
        ActiveLayer.CreateArtisticText 200, 12, p.index, Font:="Times New Roman", Size:=14, Alignment:=cdrRightAlignment
        For Each s In sr
            If s.Bitmap.ResolutionX > d Then
                s.Bitmap.Crop
                s.Bitmap.Resample , , , d, d
            End If
        Next s
    Next p
    ActiveLayer.CreateParagraphText 0, -20, 200, -60, html & vbCrLf & html2
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
    Application.Refresh
    MsgBox "Resample image finish", vbOKOnly, "Resample"
End Sub
Sub Page_Del_image()
    ActiveDocument.BeginCommandGroup "Del image"
    Optimization = True
    
    Dim sr As ShapeRange, p As Page, s As Shape
    Dim d As Integer
    
    d = 300
    For Each p In ActiveDocument.Pages
        p.Activate
        Set sr = ActivePage.Shapes.FindShapes(, cdrBitmapShape)
        If sr.count > 1 Then
            ActivePage.Shapes.FindShapes(, cdrTextShape).All.Delete
            ActivePage.Shapes.All.Move 1, -5
        End If
    Next p
    Optimization = False
    ActiveDocument.EndCommandGroup
    ActiveWindow.Refresh
    Application.Refresh
    MsgBox "Resample image finish", vbOKOnly, "Resample"
End Sub
Sub text_formating_export()
    Dim i As Integer
    For i = 1 To 14
        MsgBox ActiveSelection.Shapes(1).Text.Story.Paragraphs(14).TextFormatter & vbCrLf & i
    Next
End Sub
Sub text_301_export()
    On Error Resume Next
    Dim x As Page
    Dim objStream
    Dim i, fi, mStep As Integer
    Set objStream = CreateObject("ADODB.Stream")
    objStream.CharSet = "utf-8"
    objStream.Open
    For Each x In ActiveDocument.Pages
        x.Activate
        If ActivePage.Shapes(1).Shapes(1).Shapes.count = 8 Then
            objStream.WriteText ActivePage.Shapes(1).Shapes(1).Shapes(6).Text.Story & vbTab
            objStream.WriteText ActivePage.Shapes(1).Shapes(1).Shapes(5).Text.Story & vbTab
            objStream.WriteText Replace(ActivePage.Shapes(1).Shapes(1).Shapes(2).Text.Story, vbCrLf, " ") & vbCrLf
        End If
    Next
    objStream.SaveToFile "D:\ds.txt", 2
End Sub
Sub saveX7()
    'Function: Save file force version X7
    'Author: Bo Phi Yen
    Dim x As StructSaveAsOptions
    Dim fpath As String
    Set x = CreateStructSaveAsOptions
    With x
        .EmbedVBAProject = False
        .Filter = cdrCDR
        .IncludeCMXData = False
        .Range = cdrAllPages
        .EmbedICCProfile = False
        .Version = cdrVersion17
    End With
    fpath = ActiveDocument.FilePath & ActiveDocument.FileName
    ActiveDocument.SaveAs fpath, x
End Sub
Sub Macro15()
    ' Recorded 20/12/2020
    ActiveDocument.BeginCommandGroup "Crop"
    Optimization = True
    Dim sr As ShapeRange
    Dim pNext As Page
    Dim sDuplicate As Shape
    
    Dim x As Shape
    ActiveDocument.ReferencePoint = cdrMiddleRight
    For Each x In ActiveSelectionRange
        'x.Stretch 0.5, 1#
        x.Bitmap.Crop
    Next
    
    Optimization = False
    ActiveDocument.EndCommandGroup
    Refresh
End Sub
Sub DupPage()
    ActiveDocument.BeginCommandGroup "Duplicate Page"
    Optimization = True
    Dim sr As ShapeRange
    Dim pNext As Page
    Dim sDuplicate As Shape
    
    Set sr = ActivePage.Shapes.All.Duplicate
    Set pNext = ActiveDocument.InsertPages(1, False, ActivePage.index)
    sr.MoveToLayer pNext.Layers("Layer 1")
    pNext.Activate
    Optimization = False
    ActiveDocument.EndCommandGroup
    Refresh
End Sub
Sub RemoveGuides()
    On Error Resume Next
    ActiveDocument.MasterPage.Guides(cdrAllGuides).Delete
    If VersionMajor >= 14 Then ActivePage.Shapes.FindShapes(Query:="@type='guideline'").Delete
End Sub
Sub cccc()
    ' Recorded 14/01/2021
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    OrigSelection(1).Bitmap.Crop
End Sub
Sub Macro17()
    ' Recorded 17/01/2021
    Dim x As Page
    Dim s2 As Shape
    For Each x In ActiveDocument.Pages
        x.Activate
        ActiveLayer.CreateRectangle 3.823264, 5.705319, 3.846701, 5.668862
        Set s2 = ActiveLayer.Shapes.All.Group
        s2.Outline.SetProperties Color:=CreateCMYKColor(20, 0, 60, 0)
    Next
End Sub
Sub rgb2cmyk()
    'On Error Resume Next
    'ActivePage.Shapes.FindShapes(Query:="@type='guideline'").Delete
    'MsgBox ActiveSelection.Shapes.First.Fill.UniformColor.Type & vbCrLf & ActiveSelection.Shapes.Last.Fill.UniformColor.Type
    MsgBox ActiveSelection.Shapes.First.Fill.UniformColor
    'ActivePage.Shapes.FindShapes(Query:="@fill.color.Type = 5").CreateSelection
End Sub
Sub bold_font()
    ActiveDocument.BeginCommandGroup "ss Page"
    Optimization = True
    Dim p As Page, s As Shape
    For Each s In ActivePage.Shapes.FindShapes(, cdrTextShape)
        If s.Text.Story.Font = "UTM Avo" Then
            s.Text.Story.Bold = True
        End If
    Next
    Optimization = False
    ActiveDocument.EndCommandGroup
    Refresh
End Sub
Sub oriSize()
    'On Error Resume Next
    'ActivePage.Shapes.FindShapes(Query:="@type='guideline'").Delete
    'MsgBox ActiveSelection.Shapes.First.Fill.UniformColor.Type & vbCrLf & ActiveSelection.Shapes.Last.Fill.UniformColor.Type
    MsgBox ActiveSelection.Shapes(1).OriginalWidth & vbCrLf & ActiveSelection.Shapes(1).SizeWidth
    'ActivePage.Shapes.FindShapes(Query:="@fill.color.Type = 5").CreateSelection
End Sub
Sub GUID_list()
    Dim objStream
    Dim i, fi, mStep As Integer
    Dim ctrl As Control
    Dim cmb As CommandBar
    Set objStream = CreateObject("ADODB.Stream")
    objStream.CharSet = "utf-16le"
    objStream.Open
    objStream.WriteText ChrW(&HFEFF)
    For Each cmb In Application.CommandBars
        objStream.WriteText cmb.Name & vbTab & cmb.Type & vbTab & cmb.Controls.count & vbTab & "++" & vbCrLf
    Next
    For Each ctrl In Application.CommandBars(1).Controls
        objStream.WriteText ctrl.Caption & vbTab & ctrl.ID & vbTab & ctrl.ToolTipText & vbTab & "__" & vbCrLf
    Next
    objStream.SaveToFile "E:\User\desktop\guid2.txt", 2
End Sub

Private Sub copySize_Click()
    Dim sh As New Shape, x As Double, y As Double, shift As Long, sr As New ShapeRange
    
    On Error Resume Next
    If ActiveDocument.GetUserClick(x, y, shift, 100, True, cdrCursorPick) <> 0 Then
        Exit Sub
    End If
    Set sh = ActivePage.SelectShapesAtPoint(x, y, False, 0.01)
    If sh Is Nothing Then Exit Sub
    If sh.Shapes.count = 0 Then Exit Sub
    Set sh = sh.Shapes(sh.Shapes.count):
    ActiveDocument.ClearSelection: sr.CreateSelection
    If (shift And 1) = 0 Then ' fill
        If sh.Fill.Type <> cdrUniformFill Then Beep: Exit Sub
        cfg.clrFind.CopyAssign sh.Fill.UniformColor
    Else
        If sh.Outline.Type <> cdrOutline Then Beep: Exit Sub
        cfg.clrFind.CopyAssign sh.Outline.Color
    End If

End Sub
Sub eps_export()
    ' Recorded 25/08/2021
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    OrigSelection.CreateSelection
    Dim expopt As StructExportOptions
    Set expopt = CreateStructExportOptions
    expopt.UseColorProfile = False
    Dim expflt As ExportFilter
    Set expflt = ActiveDocument.ExportEx("E:\User\desktop\viettel.eps", cdrEPS, cdrSelection, expopt)
    With expflt
        .Header = 1 ' FilterEPSLib.epsTIFFHeader
        .TIFFHeaderType = 4 ' FilterEPSLib.epsTIFFColor8
        .Resolution = 96
        .Transparent = False
        .TextAsCurves = True
        .IncludeFonts = False
        .MaintainOPILinks = False
        .AdjustFountainSteps = False
        .FountainSteps = 128
        .PSLevel = 2 ' FilterEPSLib.epsPSLevel2
        .UserName = "bophiyen"
        .UseJPEGCompression = False
        .JPEGCompression = 2
        .PreserveOverprints = True
        .OverprintBlack = False
        .AutoSpread = False
        .FixedWidth = False
        .MaxSpread = 0.000142
        .TextSpread = 0.005102
        .BoundingBox = 0 ' FilterEPSLib.epsObjects
        .UseBleed = False
        .Bleed = 31750#
        .CropMarks = False
        .UseFloatNumbers = False
        .ConvertSpot = False
        .OutputObjectColorMode = 3 ' FilterEPSLib.epsObjectsNative
        .ConvertSpotColorsTo = 0 ' FilterEPSLib.epsColorCMYK
        .Finish
    End With
End Sub

Sub Maket_tui_Giay_Kraft()
    ' Recorded 11/10/2021
    Dim OrigSelection As Shape
    Set OrigSelection = ActiveSelection.Shapes(1)
    Dim s1 As Shape
    Set s1 = OrigSelection.Duplicate()
    s1.Style.StringAssign "{""fill"":{""type"":""9"",""overprint"":""0"",""primaryColor"":""CMYK,USER,0,0,0,100,100,00000000-0000-0000-0000-000000000000"",""secondaryColor"":""CMYK,USER,0,0,0,0,100,00000000-0000-0000-0000-000000000000"",""fillName"":""1ace86eb-bc6d-401a-8441-ebbe794bd4bf.4.colorbitmap"",""angle"":""0"",""skew"":""0"",""tilingWidth"":""200000"",""tilingHeight"":""200000"",""tilingXOffset"":""0"",""tilingYOffset"":""0"",""tilingInterTileOffset"":""0"",""tilingFlagsColumnOffset"":""0"",""tilingFlagsNoSeams"":""0"",""tilingFlagsScale"":""0"",""tilingFlagsScaleToObject"":""0"",""tilingFlagsMirrorHorizontal"":""0"",""tilingFlagsMirrorVertical"":""0""},""outline"":{""width"":""2000"",""color"":""CMYK,USER,0,0,0,100,100,00000000-0000-0000-0000-000000000000""},""transparency"":{""mode"":""9""}}"
    OrigSelection.CreateLens 10, 0, CreateRGBColor(154, 44, 41), CreateRGBColor(255, 255, 255), 0
End Sub
Sub sign()
    ' Recorded 2021-10-25
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    Dim s1 As Shape
    Set s1 = OrigSelection.ConvertToBitmapEx(4, False, True, 600, 1, True, False, 95)
    's1.Bitmap.ApplyBitmapEffect "Gaussian Blur", "GaussianBlurEffect GaussianBlurRadius=200,GaussianBlurResampled=0"
    s1.Bitmap.ApplyBitmapEffect "Crystalize", "CrystalizeEffect Size=2"
    s1.Bitmap.ApplyBitmapEffect "Diffuse", "DiffuseEffect DiffuseLevel=22"
End Sub
Sub auto_text_column()
    Dim sh As ShapeRange
    Set sh = ActiveSelectionRange
    x = sh.Shapes.FindShapes(, cdrCurveShape)
End Sub
Sub fix_end_error()
    Optimization = False
    ActiveDocument.EndCommandGroup
    Refresh
    MsgBox "Group Finish"
End Sub
Sub text_fix_drag()
    ActiveSelection.Shapes(1).Text.Story.Style = 1
    MsgBox ActiveSelection.Shapes(1).Text.Story.Style
End Sub
Sub Macro6()
    Dim tr As TextRange
    Dim para As TextRange
    Dim line As TextRange
    Dim w As TextRange
    For Each para In ActiveLayer.Shapes(1).Text.Story.Paragraphs
    With para
        Debug.Print "para", .Font, .Size, .Bold, .Italic, .Fill.Type
    End With
        For Each line In para.Lines
            For Each w In line.Words
                With w
                    Debug.Print "Sent", .Font, .Size, .Bold, .Italic, .Fill.Type, .Text
                End With
            Next
        Next
    Next
End Sub
Sub text_add()
    ' Recorded 7/4/2015
    Dim pg As Page
    ActiveDocument.BeginCommandGroup "Group Like Me"
    Optimization = True
    Dim x As Double, y As Double, w As Double, h As Double
    ActiveDocument.Unit = cdrMillimeter
    For Each pg In ActiveDocument.Pages
        pg.Activate
        ActivePage.Shapes.FindShapes(Query:="@width = {200 mm} & @height = {190 mm}").Delete
    Next
    Optimization = False
    ActiveDocument.EndCommandGroup
    Refresh
    MsgBox "Group Finish"
End Sub

Sub Macro8()
    ' Recorded 28/02/2022
    ActiveDocument.BeginCommandGroup "Group Like Me"
    Optimization = True
    ActiveDocument.Unit = cdrMillimeter
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    Dim x, y As Integer
    For x = 1 To 10
        For y = 0 To 7
            OrigSelection.Clone 100 + (36 * y), (20 * x)
        Next
    Next
    Optimization = False
    ActiveDocument.EndCommandGroup
    Refresh
    MsgBox "Group Finish"
End Sub
