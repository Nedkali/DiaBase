Public Class YesNoD2Style
    'YES BUTTONPRESS RETURN RESULT
    Private Sub YesNoCONFIRMBUTTON_Click(sender As Object, e As EventArgs) Handles YesNoCONFIRMBUTTON.Click
        DialogResult = Windows.Forms.DialogResult.Yes
    End Sub

    'NO BUTTONPRESS RETURN RESULT
    Private Sub YesNoCANCELBUTTON_Click(sender As Object, e As EventArgs) Handles YesNoCANCELBUTTON.Click
        DialogResult = Windows.Forms.DialogResult.No
    End Sub

    'PLAYS DIABLO2 DONG SOUND AND POSITIONS FORM DISPLAY LOCATION
    Private Sub YesNoD2Style_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background) 'plays d2 thunk alert sound
        Me.Left = Form1.Left + 150
        Me.Top = Form1.Top + 100
    End Sub
End Class