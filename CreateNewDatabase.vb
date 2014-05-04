Imports System.IO

Public Class CreateNewDatabase

    Private Sub NewDatabaseCreateBUTTON_Click(sender As Object, e As EventArgs) Handles NewDatabaseCreateBUTTON.Click

        If NewDatabaseTEXTBOX.Text <> Nothing Or NewDatabaseTEXTBOX.Text <> "" Then
            If My.Computer.FileSystem.FileExists(Application.StartupPath & "\DataBase\" & NewDatabaseTEXTBOX.Text & ".txt") = True Then
                'IF THE FILE NAME EXIST THE ASK FOR CONFIRMATION
                YesNoD2Style.Text = "Create Database Error"
                YesNoD2Style.YesNoHeaderLABEL.Text = "Database File Name Already Exists Error"
                YesNoD2Style.YesNoMessageLABEL.Text = "The file name you have entered is already being used by another database." & vbCrLf & vbCrLf & _
                                                      "If you continue this new database will replace the old one. All items in the old database will be lost." & vbCrLf & vbCrLf & _
                                                      "Select " & Chr(34) & "Confirm" & Chr(34) & " to replace the old database or" & Chr(34) & "Cancel" & Chr(34) & " to abort.."
                DialogResult = YesNoD2Style.ShowDialog()
            Else
                DialogResult = Windows.Forms.DialogResult.Yes
            End If

            'DELETE IF NESSICARY AND CREATE NEW DATABASE FILE
            YesNoD2Style.Close()
            If DialogResult = Windows.Forms.DialogResult.Yes Then
                'If My.Computer.FileSystem.FileExists(Application.StartupPath & "\DataBase\" & NewDatabaseTEXTBOX.Text & ".txt") = True Then My.Computer.FileSystem.DeleteFile(Application.StartupPath & "\DataBase\" & NewDatabaseTEXTBOX.Text & ".txt")

                Dim CreateFile = My.Computer.FileSystem.OpenTextFileWriter(Application.StartupPath & "\DataBase\" & NewDatabaseTEXTBOX.Text & ".txt", False)
                CreateFile.Close()

                'CHECK THE AUTO SAVE AND OPEN NEW DATABASE CHECKBOX

                If NewDatabaseAutoOpenCHECKBOX.Checked = True Then

                    'SAVE CURRENT DATABASE - MUST ALSO REFRESH BACKUP HERE I THINK SO REPLACE ACTION CAN BE UNDONE IF NESSICARY WITH RESTORE BACKUP
                    BackupDatabase()
                    Form1.SaveItems()
                    Databasefile = Application.StartupPath & "\DataBase\" & NewDatabaseTEXTBOX.Text & ".txt"

                    'Clear form 1 old database stats
                    Form1.SearchLISTBOX.Items.Clear() '                                                   clean out old search matches
                    Form1.PictureBox1.BackgroundImage = DiaBase.My.Resources.Resources.ImageBackground '  clean out old image
                    Form1.ClearStats()
                    OpenDatabaseRoutine(Databasefile)
                    Form1.Display_Items()
                    Form1.CurrentDatabaseLABEL.Text = Replace(My.Computer.FileSystem.GetName(Databasefile), ".txt", "")
                    Me.Close()

                End If

            End If
            If DialogResult = Windows.Forms.DialogResult.No Then
                'CANCEL REPLACE - Do Nothing
            End If
        Else
            NewDatabaseTEXTBOX.Select()
        End If
    End Sub
    'CANCEL CREATE DATABASE
    Private Sub NewDatabaseCancelBUTTON_Click(sender As Object, e As EventArgs) Handles NewDatabaseCancelBUTTON.Click
        Me.Close()
    End Sub

    Private Sub CreateNewDatabase_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        NewDatabaseTEXTBOX.Text = Nothing
        NewDatabaseTEXTBOX.Select()
    End Sub
End Class