Imports Microsoft.Office.Interop

Public Class frmButterfly

    Private Sub showProgress(ByVal no As Integer)
        ToolStripProgressBar1.Value = no
        If no = 100 Then
            ToolStripProgressBar1.Visible = False
        Else
            ToolStripProgressBar1.Visible = True
        End If
        Application.DoEvents()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPDF.Click
        frmMain.ShowDialog()
    End Sub

    Sub reportH1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnH1.Click
        If Not checkDir("H1") Then
            MsgBox("目錄內沒有H1目錄或目錄內無檔案, 請重新選擇!", MsgBoxStyle.Critical, "錯誤的目錄")
            Return
        End If
        Dim oWord As Word.Application
        Me.Enabled = False
        oWord = CreateObject("Word.Application")
        'oWord.Visible = True 
        Call openFile(oWord, "Standard.dot")
        Dim fileNo As Integer = getFilesNumber(txtImport.Text & "\H1\")
        Call h1(oWord, fileNo, fileNo)
        Result(oWord)
        Me.Enabled = True
    End Sub

    Sub reportH99(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnC.Click
        If Not checkDir("H99") Then
            MsgBox("目錄內沒有H99目錄或目錄內無檔案, 請重新選擇!", MsgBoxStyle.Critical, "錯誤的目錄")
            Return
        End If
        Dim oWord As Word.Application
        Me.Enabled = False
        oWord = CreateObject("Word.Application")
        Call openFile(oWord, "Standard.dot")
        Dim fileNo As Integer = getFilesNumber(txtImport.Text & "\H99\")
        Call h99(oWord, 0, fileNo)
        Result(oWord)
        Me.Enabled = True
    End Sub

    Sub reportH999(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnB.Click
        If Not checkDir("H99") Then
            MsgBox("目錄內沒有H99目錄或目錄內無檔案, 請重新選擇!", MsgBoxStyle.Critical, "錯誤的目錄")
            Return
        End If
        Dim oWord As Word.Application
        Me.Enabled = False
        oWord = CreateObject("Word.Application")
        Call openFile(oWord, "H999.dot")
        Dim fileNo As Integer = getFilesNumber(txtImport.Text & "\H99\")
        Call H999(oWord, fileNo)
        Result(oWord)
        Me.Enabled = True
    End Sub

    Sub reportH1H99(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnA.Click
        If Not checkDir("H99") Then
            MsgBox("目錄內沒有H99目錄或目錄內無檔案, 請重新選擇!", MsgBoxStyle.Critical, "錯誤的目錄")
            Return
        ElseIf Not checkDir("H1") Then
            MsgBox("目錄內沒有H1目錄或目錄內無檔案, 請重新選擇!", MsgBoxStyle.Critical, "錯誤的目錄")
            Return
        End If
        Dim oWord As Word.Application
        Me.Enabled = False
        oWord = CreateObject("Word.Application")
        Call openFile(oWord, "Standard.dot")
        Dim nextID As Integer
        Dim fileNoH1 As Integer = getFilesNumber(txtImport.Text & "\H1\")
        Dim fileNoH99 As Integer = getFilesNumber(txtImport.Text & "\H99\")
        nextID = h1(oWord, fileNoH1, fileNoH1 + fileNoH99)
        oWord.Selection.MoveDown(5, 1)
        oWord.Selection.TypeParagraph()
        Call h99(oWord, nextID, fileNoH1 + fileNoH99)
        Result(oWord)
        Me.Enabled = True
    End Sub


    Sub resizeImage2(ByVal oWord As Word.Application)
        With oWord
            .Selection.SelectCell()
            .Selection.InlineShapes.Item(1).Fill.Visible = 0
            .Selection.InlineShapes.Item(1).Fill.Solid()
            .Selection.InlineShapes.Item(1).Fill.Transparency = 0.0#
            .Selection.InlineShapes.Item(1).Line.Weight = 0.75
            .Selection.InlineShapes.Item(1).Line.Transparency = 0.0#
            .Selection.InlineShapes.Item(1).Line.Visible = 0
            .Selection.InlineShapes.Item(1).LockAspectRatio = 0
            .Selection.InlineShapes.Item(1).Height = 195.0#
            .Selection.InlineShapes.Item(1).Width = 266.15
            .Selection.InlineShapes.Item(1).PictureFormat.Brightness = 0.5
            .Selection.InlineShapes.Item(1).PictureFormat.Contrast = 0.5
            .Selection.InlineShapes.Item(1).PictureFormat.ColorType = 1
            .Selection.InlineShapes.Item(1).PictureFormat.CropLeft = 0.0#
            .Selection.InlineShapes.Item(1).PictureFormat.CropRight = 0.0#
            .Selection.InlineShapes.Item(1).PictureFormat.CropTop = 0.0#
            .Selection.InlineShapes.Item(1).PictureFormat.CropBottom = 0.0#
        End With
    End Sub

    Sub resizeImage3(ByVal oword As Word.Application)
        With oWord
            .Selection.SelectCell()
            .Selection.InlineShapes.Item(1).Fill.Visible = 0
            .Selection.InlineShapes.Item(1).Fill.Solid()
            .Selection.InlineShapes.Item(1).Fill.Transparency = 0.0#
            .Selection.InlineShapes.Item(1).Line.Weight = 0.75
            .Selection.InlineShapes.Item(1).Line.Transparency = 0.0#
            .Selection.InlineShapes.Item(1).Line.Visible = 0
            .Selection.InlineShapes.Item(1).LockAspectRatio = 0
            .Selection.InlineShapes.Item(1).Height = 197.3
            .Selection.InlineShapes.Item(1).Width = 267.0#
            .Selection.InlineShapes.Item(1).PictureFormat.Brightness = 0.5
            .Selection.InlineShapes.Item(1).PictureFormat.Contrast = 0.5
            .Selection.InlineShapes.Item(1).PictureFormat.ColorType = 1
            .Selection.InlineShapes.Item(1).PictureFormat.CropLeft = 0.0#
            .Selection.InlineShapes.Item(1).PictureFormat.CropRight = 0.0#
            .Selection.InlineShapes.Item(1).PictureFormat.CropTop = 0.0#
            .Selection.InlineShapes.Item(1).PictureFormat.CropBottom = 0.0#
        End With
    End Sub

    Sub resizeImage(ByVal oWord As Word.Application)
        With oWord
            .Selection.SelectCell()
            .Selection.InlineShapes.Item(1).Fill.Visible = 0
            .Selection.InlineShapes.Item(1).Fill.Solid()
            .Selection.InlineShapes.Item(1).Fill.Transparency = 0.0#
            .Selection.InlineShapes.Item(1).Line.Weight = 0.75
            .Selection.InlineShapes.Item(1).Line.Transparency = 0.0#
            .Selection.InlineShapes.Item(1).Line.Visible = 0
            .Selection.InlineShapes.Item(1).LockAspectRatio = 0
            '.Selection.InlineShapes.Item(1).Height = 283.45
            .Selection.InlineShapes.Item(1).Height = 300.45
            .Selection.InlineShapes.Item(1).Width = 393.75
            .Selection.InlineShapes.Item(1).PictureFormat.Brightness = 0.5
            .Selection.InlineShapes.Item(1).PictureFormat.Contrast = 0.5
            .Selection.InlineShapes.Item(1).PictureFormat.ColorType = 1
            .Selection.InlineShapes.Item(1).PictureFormat.CropLeft = 0.0#
            .Selection.InlineShapes.Item(1).PictureFormat.CropRight = 0.0#
            .Selection.InlineShapes.Item(1).PictureFormat.CropTop = 0.0#
            .Selection.InlineShapes.Item(1).PictureFormat.CropBottom = 0.0#
            .Selection.Tables.Item(1).Rows.HeightRule = 0
            .Selection.Tables.Item(1).Rows.Height = .CentimetersToPoints(0)
        End With
    End Sub

    Private Sub frmButterfly_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If My.Settings.txtImport = "" Then
            txtImport.Text = System.Windows.Forms.Application.StartupPath()
        Else
            txtImport.Text = My.Settings.txtImport
        End If
    End Sub

    Private Function checkDir(ByVal check As String) As Boolean
        'If System.IO.Directory.Exists(My.Settings.txtImport & "\" & check) Then
        If getFilesNumber(My.Settings.txtImport & "\" & check, True) > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        FolderBrowserDialog1.SelectedPath = Me.txtImport.Text
        If FolderBrowserDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            'If System.IO.Directory.Exists(FolderBrowserDialog1.SelectedPath & "\H1") And System.IO.Directory.Exists(FolderBrowserDialog1.SelectedPath & "\H99") Then
            Me.txtImport.Text = FolderBrowserDialog1.SelectedPath
            My.Settings.txtImport = txtImport.Text
            'Else
            'MsgBox("指定的目錄沒有包含H1及H99目錄, 所以放棄變更," & vbCrLf & vbCrLf & "你可以先行在該目錄產生H1、H99目錄後, 再進行選擇。", MsgBoxStyle.Critical, "放棄變更目錄")
            ' End If
        End If
    End Sub

    Function h1(ByVal oWord As Word.Application, ByVal fileNo As Integer, ByVal total As Integer) As Integer
        If fileNo = 0 Then
            Return 0
        End If
        Dim tableLine As Integer = Math.Round(fileNo / 2)
        Try
            With oWord
                .ActiveDocument.Tables.Add(.Selection.Range, tableLine * 2, 2, 1, 0)
                With .Selection.Tables.Item(1)
                    .ApplyStyleHeadingRows = True
                    .ApplyStyleLastRow = True
                    .ApplyStyleFirstColumn = True
                    .ApplyStyleLastColumn = True
                End With
                With .Selection.Tables.Item(1)
                    .TopPadding = oWord.CentimetersToPoints(0)
                    .BottomPadding = oWord.CentimetersToPoints(0)
                    .LeftPadding = oWord.CentimetersToPoints(0.05)
                    .RightPadding = oWord.CentimetersToPoints(0.05)
                    .Spacing = 0
                    .AllowPageBreaks = True
                    .AllowAutoFit = False
                End With

                Dim sFileName As String = Dir(txtImport.Text & "\H1\*.*", vbHidden Or vbReadOnly Or vbSystem)
                Dim no As Integer = 0
                Dim pos As Integer = 0
                Do While sFileName <> ""
                    If sFileName <> "." And sFileName <> ".." And Mid(sFileName, 1, 1) <> "~" And _
                            (sFileName.ToLower.EndsWith(".jpg") Or sFileName.ToLower.EndsWith(".jpeg")) Then
                        If (no Mod 2) = 0 And no <> 0 Then
                            .Selection.MoveRight(12)
                        End If
                        no += 1
                        'lblStatus.Text = "處理 " & txtImport.Text & "\H1\" & sFileName
                        .Selection.InlineShapes.AddPicture(FileName:= _
                            txtImport.Text & "\H1\" & sFileName, LinkToFile:=False, _
                             SaveWithDocument:=True)
                        'lblStatus.Text = "完成 " & txtImport.Text & "\H1\" & sFileName
                        Call resizeImage3(oWord)
                        .Selection.MoveRight(12)
                        If (no Mod 2) = 0 Then
                            pos += 1
                            .Selection.TypeText(Text:="NO." & pos)
                            pos += 1
                            .Selection.MoveRight(12)
                            .Selection.TypeText(Text:="NO." & pos)
                        End If
                    End If
                    sFileName = Dir()
                    Me.showProgress(CInt(no / total * 100))
                Loop

                If (no <> pos) Then
                    pos += 1
                    .Selection.MoveRight(12)
                    .Selection.TypeText(Text:="NO." & pos)
                End If
                .ActiveWindow.ActivePane.VerticalPercentScrolled = 3
                h1 = pos
            End With
        Catch ex As Exception
            lblStatus.Text = "轉換過程發生異常(" & ex.Message & ")"
            MsgBox("轉換過程發生異常(" & ex.Message & ")")
            Debug.Print(ex.StackTrace)
        End Try

    End Function

    Sub h99(ByVal oWord As Word.Application, ByVal start As Integer, ByVal total As Integer)
        Try
            Dim fileNo As Integer = total - start
            If fileNo = 0 Then
                Return
            End If
            With oWord
                .Selection.TypeText(Text:="異常狀況")
                .Selection.TypeParagraph()
                .Selection.TypeParagraph()
                .ActiveDocument.Tables.Add(.Selection.Range, fileNo * 2, 1, 1, 0)
                With .Selection.Tables.Item(1)
                    .ApplyStyleHeadingRows = True
                    .ApplyStyleLastRow = True
                    .ApplyStyleFirstColumn = True
                    .ApplyStyleLastColumn = True
                End With
                With .Selection.Tables.Item(1)
                    .TopPadding = oWord.CentimetersToPoints(0)
                    .BottomPadding = oWord.CentimetersToPoints(0)
                    .LeftPadding = oWord.CentimetersToPoints(0.04)
                    .RightPadding = oWord.CentimetersToPoints(0.04)
                    .Spacing = 0
                    .AllowPageBreaks = True
                    .AllowAutoFit = False
                End With
                .Selection.Tables.Item(1).Rows.Alignment = 1
                .Selection.Tables.Item(1).PreferredWidthType = 3
                .Selection.Tables.Item(1).PreferredWidth = .CentimetersToPoints(14)

                Dim sFileName As String = Dir(txtImport.Text & "\H99\*.*", vbHidden Or vbReadOnly Or vbSystem)
                Dim no As Integer = start
                Do While sFileName <> ""
                    If sFileName <> "." And sFileName <> ".." And Mid(sFileName, 1, 1) <> "~" And _
                            (sFileName.ToLower.EndsWith(".jpg") Or sFileName.ToLower.EndsWith(".jpeg")) Then
                        If no <> start Then
                            .Selection.MoveRight(12)
                        End If
                        'lblStatus.Text = "處理 " & txtImport.Text & "\H1\" & sFileName
                        .Selection.InlineShapes.AddPicture(txtImport.Text & "\H99\" & sFileName, False, True)
                        Call resizeImage(oWord)
                        'lblStatus.Text = "處理 " & txtImport.Text & "\H1\" & sFileName & "(完成)"
                        no += 1
                        .Selection.MoveRight(12)
                        .Selection.TypeText("NO." & no)
                    End If
                    sFileName = Dir()
                    Me.showProgress(CInt(no / total * 100))
                Loop
            End With
        Catch ex As Exception
            lblStatus.Text = ex.Message
        End Try
    End Sub

    Sub H999(ByVal oWord As Word.Application, ByVal total As Integer)
        Try
            With oWord
                Dim sFileName As String = Dir(txtImport.Text & "\H99\*.*", vbHidden Or vbReadOnly Or vbSystem)
                Dim no As Integer = 0
                Do While sFileName <> ""
                    If sFileName <> "." And sFileName <> ".." And Mid(sFileName, 1, 1) <> "~" And _
                            (sFileName.ToLower.EndsWith(".jpg") Or sFileName.ToLower.EndsWith(".jpeg")) Then
                        no += 1
                        .Selection.MoveRight(12)
                        If no = 1 Then
                            'lblStatus.Text = " 處理" & txtImport.Text & "\H1\" & sFileName
                            .Selection.InlineShapes.AddPicture(txtImport.Text & "\H99\" & sFileName, False, True)
                            'lblStatus.Text = " 處理" & txtImport.Text & "\H1\" & sFileName & "(完成)"
                            Call resizeImage2(oWord)
                            .Selection.MoveRight(12)
                            .Selection.Copy()
                        Else
                            .Selection.MoveRight(12)
                            'lblStatus.Text = " 處理" & txtImport.Text & "\H1\" & sFileName
                            .Selection.InlineShapes.AddPicture(txtImport.Text & "\H99\" & sFileName, False, True)
                            'lblStatus.Text = " 處理" & txtImport.Text & "\H1\" & sFileName & "(完成)"
                            Call resizeImage2(oWord)
                            .Selection.MoveRight(12)
                            .Selection.PasteAndFormat(0)
                        End If
                        .Selection.MoveRight(12)
                        .Selection.TypeText(Text:="NO." & no)
                    End If
                    sFileName = Dir()
                    Me.showProgress(CInt(no / total * 100))
                Loop
            End With
        Catch ex As Exception
            lblStatus.Text = ex.Message
        End Try
    End Sub

    Function getFilesNumber(ByVal currentdir As String, Optional ByVal checkDir As Boolean = False) As Integer
        Dim sFileName As String = Dir(currentdir & "\*.*", vbHidden Or vbReadOnly Or vbSystem)
        Dim no As Integer = 0
        Do While sFileName <> ""
            lblStatus.Text = "檔案統計中"
            If sFileName <> "." And sFileName <> ".." And Mid(sFileName, 1, 1) <> "~" And _
                    (sFileName.ToLower.EndsWith(".jpg") Or sFileName.ToLower.EndsWith(".jpeg")) Then
                no += 1
                If checkDir Then
                    Return no
                End If
            End If
            sFileName = Dir()
        Loop
        Return no
    End Function

    Sub openFile(ByVal oWord As Word.Application, ByVal filename As String)
        Dim oMissing As Object = System.Reflection.Missing.Value
        Try
            With oWord
                .Application.WindowState = Word.WdWindowState.wdWindowStateMinimize
                .Visible = True
                If RadioButton1.Checked Then
                    .Documents.Add(System.Windows.Forms.Application.StartupPath() & "\template\" & filename)
                ElseIf RadioButton2.Checked Then
                    .Documents.Add(System.Windows.Forms.Application.StartupPath() & "\template\" & "2-" & filename)
                ElseIf RadioButton3.Checked Then
                    .Documents.Add(System.Windows.Forms.Application.StartupPath() & "\template\" & "3-" & filename)
                End If
                oWord.Visible = False
            End With
        Catch ex As Exception
            MsgBox("異常, 請檢查工作環境是否正常, 參考:" & ex.Message)
            Debug.Print(ex.StackTrace)
        End Try
    End Sub

    Sub Result(ByVal oWord As Word.Application)
        oWord.Application.NormalTemplate.Saved = False
        oWord.Visible = True
        oWord.WindowState = Word.WdWindowState.wdWindowStateMaximize
        'oWord.Application.Quit(0)
        lblStatus.Text = "完成表格。"
    End Sub

End Class



