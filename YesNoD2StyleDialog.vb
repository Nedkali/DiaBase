﻿Public Class YesNoD2Style

    Private Sub YesNoCONFIRMBUTTON_Click(sender As Object, e As EventArgs) Handles YesNoCONFIRMBUTTON.Click
        DialogResult = Windows.Forms.DialogResult.Yes

    End Sub

    Private Sub YesNoCANCELBUTTON_Click(sender As Object, e As EventArgs) Handles YesNoCANCELBUTTON.Click
        DialogResult = Windows.Forms.DialogResult.No

    End Sub

   
   
    Private Sub YesNoD2Style_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background) 'plays d2 thunk alert sound
    End Sub
End Class