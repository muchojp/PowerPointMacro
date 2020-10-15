Sub MakeProgressBar()
    Const r As String = "42"    '色・RGB値のR
    Const g As String = "86"    '色・RGB値のG
    Const b As String = "f5"    '色・RGB値のB
    Const pbH As Long = 10     '高さ
    Const pbBG As Single = 0.6  '背景の透過性
    Const rb As String = "3F"    '色・黒線のRGB値のR
    Const gb As String = "38"    '色・黒線のRGB値のG
    Const bb As String = "38"    '色・黒線のRGB値のB
    Const barH As Long = 3     '黒線の高さ

    Dim i As Long
    Dim s As Shape
    Dim j As Long

    Dim wTop As Long 'プログレスバー位置
    Dim barTop As Long '黒線の位置
    Dim rc As Integer

    On Error Resume Next

    rc = MsgBox("プログレスバー位置はどこにしますか？" & vbCrLf & "上部（はい）　下部（いいえ）", vbYesNo + vbQuestion, "確認")
    If rc = vbYes Then
        wTop = 0
        barTop = pbH ' + wTop
    Else
        wTop = ActivePresentation.PageSetup.SlideHeight - pbH
        barTop = -barH + wTop
    End If

    rc = MsgBox("スライド1枚目にもプログレスバーを付けますか？" & vbCrLf & "付ける（はい）　付けない（いいえ）", vbYesNo + vbQuestion, "確認")
    If rc = vbYes Then
        j = 1
    Else
        j = 2
    End If

    With ActivePresentation
        '背景 ProgressBarBG の設定
        .SlideMaster.Shapes("ProgressBarBG").Delete
        Set s = .SlideMaster.Shapes.AddShape( _
        Type:=msoShapeRectangle, _
        Left:=0, _
        Height:=pbH, _
        Top:=wTop, _
        Width:=.PageSetup.SlideWidth)
        With s
            .Fill.ForeColor.RGB = _
            RGB(CInt("&H" & r), CInt("&H" & g), CInt("&H" & b))
            .Fill.Transparency = pbBG
            .Line.Visible = msoFalse
            .Name = "ProgressBarBG"
        End With
        

        .Slides(1).Shapes("ProgressBarBGShadow").Delete
        'プログレスバー ProgressBar の設定
        For i = j To .Slides.Count
            .Slides(i).Shapes("ProgressBar").Delete
            Set s = .Slides(i).Shapes.AddShape( _
            Type:=msoShapeRectangle, _
            Left:=0, _
            Height:=pbH, _
            Top:=wTop, _
            Width:=(i - 1) * .PageSetup.SlideWidth / (.Slides.Count - 1))
            With s
                .Fill.ForeColor.RGB = _
                RGB(CInt("&H" & r), CInt("&H" & g), CInt("&H" & b))
                .Line.Visible = msoFalse
                .Name = "ProgressBar"
            End With
            
            '背景 ProgressBarBGShadow の設定
            .Slides(i).Shapes("ProgressBarBGShadow").Delete
            Set s = .Slides(i).Shapes.AddShape( _
            Type:=msoShapeRectangle, _
            Left:=0, _
            Height:=barH, _
            Top:=barTop, _
            Width:=.PageSetup.SlideWidth)
            With s
                .Fill.ForeColor.RGB = _
                RGB(CInt("&H" & rb), CInt("&H" & gb), CInt("&H" & bb))
                .Line.Visible = msoFalse
                .Name = "ProgressBarBGShadow"
            End With
        Next i
    End With

End Sub
