
Public Class frmMain

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.txtFolder.Text = CreateObject("WScript.Shell").Specialfolders("Desktop")
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Cursor = Cursors.WaitCursor
        Button3.Enabled = False
        Button1.Enabled = False
        Dim oWord As Object

        Try
            'Dim oWord As Word.Application = New Word.Application
            oWord = CreateObject("Word.Application")
            Dim currentDir As String = Me.txtFolder.Text
            Dim sFileName As String
            sFileName = Dir(currentDir & "\*.*", vbHidden Or vbReadOnly Or vbSystem)
            If currentDir = "False" Then
                MsgBox("沒有選擇一個目錄, 放棄作業!!", vbInformation, "轉檔程式")
                Exit Sub
            End If
            With oWord
                .ChangeFileOpenDirectory(currentDir)
                Do While sFileName <> ""
                    lblStatus.Text = "列印" & sFileName
                    If sFileName <> "." And sFileName <> ".." And Mid(sFileName, 1, 1) <> "~" And _
                            (sFileName.ToLower.EndsWith(".doc") Or sFileName.ToLower.EndsWith(".docx")) Then
                        .Documents.Open(sFileName)
                        .ActivePrinter = "FinePrint"
                        .Application.PrintOut(False, , , , , , , , , , , , , , , , , 11907, 16839)
                        .ActiveDocument.Close()
                    End If
                    sFileName = Dir()
                    Application.DoEvents()
                Loop
            End With
            lblStatus.Text = "轉換完畢。"
        Catch ex As Exception
            lblStatus.Text = "轉換過程發生異常(" & ex.Message & ")"
            MsgBox("轉換過程發生異常(" & ex.Message & ")")
        End Try
        oWord = Nothing
        Button3.Enabled = True
        Button1.Enabled = True
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Dispose(True)
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        FolderBrowserDialog1.SelectedPath = Me.txtFolder.Text
        If FolderBrowserDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Me.txtFolder.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub
End Class
