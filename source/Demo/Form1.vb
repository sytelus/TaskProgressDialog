Imports Sytel.Presentation.TaskProgress

Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        SharedTaskProgress.Show("Processing Go Home Request", "Finishing Windows shutdown...", 0, 100)
        For i As Integer = 1 To 100
            System.Threading.Thread.Sleep(50)
            SharedTaskProgress.UpdateProgress("Doing thing #" + i.ToString, i)
            If i = 50 Then
                SubMethod()
            End If
        Next
        SharedTaskProgress.UpdateAsFinishedAndUnload()
    End Sub

    Private Sub SubMethod()
        SharedTaskProgress.Show("Now Sub Method is updating progress", "Item #50 has sub progress...", 0, 100)
        For i As Integer = 1 To 40
            System.Threading.Thread.Sleep(50)
            SharedTaskProgress.UpdateProgress("Doing thing #50." + i.ToString, i)
        Next
        SharedTaskProgress.UpdateAsFinishedAndUnload()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        SubMethod()
    End Sub
End Class
