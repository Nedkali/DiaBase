﻿Public Class YesNoD2Style

    Private Sub YesNoCONFIRMBUTTON_Click(sender As Object, e As EventArgs) Handles YesNoCONFIRMBUTTON.Click
        DialogResult = Windows.Forms.DialogResult.Yes
        Me.Close()
    End Sub

    Private Sub YesNoCANCELBUTTON_Click(sender As Object, e As EventArgs) Handles YesNoCANCELBUTTON.Click
        DialogResult = Windows.Forms.DialogResult.No
        Me.Close()
    End Sub
End Class