Attribute VB_Name = "Module1"
Sub fix_textbox()
    'Fix text after CAT-Translate
    Dim sh As Shape
    Dim sl As Slide
    Dim pr As TextRange
    For Each sl In ActivePresentation.Slides
        Debug.Print sl.Name, sl.Shapes.Count
        For Each sh In sl.Shapes
            If sh.Type = msoTextBox Or sh.Type = msoPlaceholder Then
                If sh.Width < 400 Then
                    sh.TextFrame.WordWrap = msoFalse
                    sh.TextFrame.AutoSize = ppAutoSizeShapeToFitText
                End If
                If sh.TextFrame.TextRange.Paragraphs.Count > 1 Then
                    'Fix Long paragraphs make overflow page
                    For Each pr In sh.TextFrame.TextRange.Paragraphs
                        If pr.ParagraphFormat.SpaceAfter > 3 Then pr.ParagraphFormat.SpaceAfter = 3
                        If pr.ParagraphFormat.SpaceBefore > 3 Then pr.ParagraphFormat.SpaceBefore = 3
                    Next
                End If
                If sh.Width + sh.Left > sl.Design.SlideMaster.Width Then
                    'Fix single line so long overflow width page
                    sh.TextFrame.WordWrap = msoTrue
                    sh.Width = sl.Design.SlideMaster.Width - sh.Left - 50
                    sh.TextFrame.AutoSize = ppAutoSizeShapeToFitText
                End If
            End If
        Next
    Next
End Sub
Sub fix_textbox_i()
    Dim sh As Shape
    Dim sl As Slide
    For Each sh In ActivePresentation.Slides(37).Shapes
        Debug.Print sh.Type, sh.HasTextFrame, sh.TextFrame.TextRange.Words.Count, sh.Width
        If (sh.Type = 17 Or sh.Type = 14) And sh.Width < 400 Then
            sh.TextFrame.WordWrap = msoFalse
            If sh.Width + sh.Left > sl.Design.SlideMaster.Width Then
                sh.TextFrame.WordWrap = msoTrue
                sh.Width = sl.Design.SlideMaster.Width - sh.Left - 50
                sh.TextFrame.AutoSize = ppAutoSizeShapeToFitText
            End If
        End If
    Next
End Sub
Sub fix_para()
    Dim sh As Shape
    Dim sl As Slide
    Dim pr As TextRange
    For Each sh In ActivePresentation.Slides(41).Shapes
        Debug.Print sh.Name, sh.Type, sh.TextFrame.TextRange.Words.Count
        If sh.Type = 17 Or sh.Type = 14 Then
            For Each pr In sh.TextFrame.TextRange.Paragraphs
                Debug.Print "---> ", pr.ParagraphFormat.SpaceAfter, pr.ParagraphFormat.SpaceBefore
                If pr.ParagraphFormat.SpaceAfter > 3 Then pr.ParagraphFormat.SpaceAfter = 3
                If pr.ParagraphFormat.SpaceBefore > 3 Then pr.ParagraphFormat.SpaceBefore = 3
            Next
        End If
    Next
End Sub

