Public Class Help

    Private Sub Help_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Left = Form1.Left + 100
        Me.Top = Form1.Top + 100

        If My.Computer.FileSystem.FileExists(Application.StartupPath + "\Help.rtf") = False Then
            RichTextBox1.Text = "Help file not found"
            Return
        End If

        Try
            RichTextBox1.LoadFile(Application.StartupPath & "\Help.rtf", RichTextBoxStreamType.RichText)
        Catch ex As Exception
            Mymessages = "Failed to load - Help.txt" : MyMessageBox()
        End Try


    End Sub
End Class