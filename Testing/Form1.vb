Public Class Form1
    Private WithEvents H As New MultiThreading.WorkerHandler

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CheckForIllegalCrossThreadCalls = False


        H.Include_Worker(New Custom)


        H.Run_Process()
    End Sub

    Sub AddReport(ByVal Message) Handles H.ReportChange
        ListBox1.Items.Add(Message)
    End Sub

    Sub TheEnd() Handles H.ProcessFinished
        End
    End Sub

End Class

Friend Class Custom
    Inherits MultiThreading.Worker

    Public Overrides Sub MyWork()
        'Work to do:

    End Sub
End Class