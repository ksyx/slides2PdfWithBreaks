Attribute VB_Name = "addBreak"
Option Explicit
Sub addAnimationBreaks()
    On Error Resume Next
    Dim i As Integer, j As Integer, t As Integer, cnt As Integer
    Dim v As MsoTriState
    t = 1
    With ActivePresentation
        .SaveAs "backup.pptx", ppSaveAsOpenXMLPresentation
        While t <= .Slides.Count
            With .Slides(t)
                Dim seq As Sequence
                cnt = 1
                Set seq = .TimeLine.MainSequence
                Dim last As Integer
                last = seq.Count
                For i = seq.Count To 1 Step -1
                    Dim item As Effect
                    Set item = seq.item(i)
                    If seq.item(i).Timing.TriggerType = msoAnimTriggerOnPageClick Then
                        .Duplicate
                        cnt = cnt + 1
                        For j = last To i Step -1
                            seq.item(j).Shape.Visible = msoFalse
                        Next
                    End If
                    last = i
                Next
            End With
            t = t + cnt
        Wend
        .SaveAs "export.pdf", ppSaveAsPDF
    End With
End Sub
