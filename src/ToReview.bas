Attribute VB_Name = "ToReview"
Option Explicit

Dim footnote_id As Long
Dim in_footnote As Boolean


Public Sub ToReview()
    Dim doc As Document
    Dim para As Paragraph
    Dim line As String
    Dim is_prev_empty As Boolean
    
    Dim footnote_coll As New Collection
    Dim footnote_text As Variant
    
    Dim fs As New FileSystemObject
    Dim ts As TextStream
    
    Set doc = ActiveDocument
    
    Set ts = fs.CreateTextFile("D:\Work\review.txt", True)
    
    is_prev_empty = True
    footnote_id = 0
    in_footnote = False
    For Each para In doc.Paragraphs
        line = ""
        Set footnote_coll = New Collection
        
        '見出し
        Call ToReviewOutline(para, line)
        
        '箇条書き
        Call ToReviewListFormat(para, line)
        
        '段落
        Call ToReviewParagraph(para, footnote_coll, line)
        
        If line = vbCr And Len(line) = 1 Then
            line = ""
        End If
        
        If is_prev_empty And Len(line) = 0 Then
            '連続の空は出力しない
        Else
            ts.Write (line)
            If Len(line) > 0 Then
                ts.WriteBlankLines (1)
            End If
            
            '脚注
            For Each footnote_text In footnote_coll
                ts.WriteLine (footnote_text)
            Next
            If footnote_coll.Count >= 1 Then
                ts.WriteBlankLines (1)
            End If
        End If
        
        
        If Len(line) = 0 Then
            is_prev_empty = True
        Else
            is_prev_empty = False
        End If
        
        Set footnote_coll = Nothing
    Next
    
    ts.Close

    Set doc = Nothing
    
    Debug.Print "finish"
End Sub

'段落
Private Sub ToReviewParagraph(ByRef para As Paragraph _
                                , ByRef footnote_coll As Collection _
                                , ByRef line As String)
    Dim r As Range
    Dim in_hyperlink As Boolean
    Dim hyperlink_text As String
    Dim hyperlink_address As String
    
    in_hyperlink = False
    For Each r In para.Range.Words
        If r.Hyperlinks.Count > 0 And Not in_hyperlink Then
            'ハイパーリンクに入った
            in_hyperlink = True
            hyperlink_address = r.Hyperlinks.Item(1).Address
            Call ToReviewRange(r, footnote_coll, hyperlink_text)
            
        ElseIf r.Hyperlinks.Count > 0 And in_hyperlink Then
            'ハイパーリンク中
            Call ToReviewRange(r, footnote_coll, hyperlink_text)
            
        ElseIf r.Hyperlinks.Count = 0 And in_hyperlink Then
            'ハイパーリンクを出た
            If Len(hyperlink_address) > 0 Then
                If hyperlink_address = hyperlink_text Then
                    line = line & "@<href>{" & hyperlink_address & "}"
                Else
                    line = line & "@<href>{" & hyperlink_address & "," & hyperlink_text & "}"
                End If
                hyperlink_address = ""
            End If
            Call ToReviewRange(r, footnote_coll, line)
        Else
            Call ToReviewRange(r, footnote_coll, line)
        End If
    Next
    
End Sub

'単語？
Private Sub ToReviewRange(ByRef r As Range, ByRef footnote_coll As Collection, ByRef line As String)
    If r.Footnotes.Count > 0 And Not in_footnote Then
        Call ToReviewFootNote(r, footnote_coll, line)
        
    Else
        If InStr(r.Text, Chr(2)) > 0 And in_footnote Then
            '脚注の頭の番号のところは無視
        ElseIf r.Text = vbFormFeed Then
        Else
            line = line & r.Text
        End If
    End If
End Sub

'脚注
Private Sub ToReviewFootNote(ByRef r As Range, ByRef footnote_coll As Collection, ByRef line As String)
    Dim i As Long
    Dim f_line As String
    Dim p As Paragraph
    
    in_footnote = True
    
    For i = 1 To r.Footnotes.Count
        line = line & "@<fn>{fnid_" & Format(footnote_id, "0##") & "}"
        
        For Each p In r.Footnotes.Item(i).Range.Paragraphs
            Call ToReviewParagraph(p, footnote_coll, f_line)
        Next
        
        Call RemoveEndCr(f_line)
        
        Call footnote_coll.Add("//footnote[fnid_" & Format(footnote_id, "0##") & "][" & f_line & "]")
        footnote_id = footnote_id + 1
    Next
    
    in_footnote = False
End Sub

'箇条書き
Private Sub ToReviewListFormat(ByRef para As Paragraph, ByRef line As String)

    If Len(line) > 0 Then
        Exit Sub
    End If

    Select Case para.Range.ListFormat.ListType
    Case wdListBullet, wdListPictureBullet
        line = line & " * "
    
    Case wdListListNumOnly, wdListSimpleNumbering, wdListOutlineNumbering, wdListMixedNumbering
        line = line & " " & para.Range.ListFormat.ListValue & ". "
    
    Case Else
    End Select
End Sub

'見出し
Private Sub ToReviewOutline(ByRef para As Paragraph, ByRef line As String)
    Dim i As Long
    
    If para.OutlineLevel <> wdOutlineLevelBodyText Then
        For i = 1 To para.OutlineLevel
            line = line & "="
        Next
        line = line & " "
    End If
End Sub

Private Sub RemoveEndCr(ByRef line As String)
    If InStrRev(line, vbCr) = Len(line) And Len(line) > 0 Then
        line = Left(line, Len(line) - 1)
    End If
End Sub

