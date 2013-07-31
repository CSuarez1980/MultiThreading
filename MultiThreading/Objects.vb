Public Enum EventType
    ReportProgress = 1
    WorkCompleted = 2
    test = 3
End Enum

Public Class WorkerHandler
#Region "Events"
    Public Event ReportChange(ByVal Message As String)
    Public Event ProcessFinished()
#End Region
#Region "Variables"
    Private WithEvents _BGW As New System.ComponentModel.BackgroundWorker
    Private _Workers As New List(Of Worker)
    Private _Completed As Integer
#End Region
#Region "Properties"
    Public ReadOnly Property Workers() As List(Of Worker)
        Get
            Return _Workers
        End Get
    End Property
#End Region
#Region "Methods"
    Public Sub Clear_Workers()
        _Workers.Clear()
    End Sub
    Public Sub Run_Process()
        _BGW.RunWorkerAsync()
    End Sub
    Private Sub Start() Handles _BGW.DoWork
        For Each W In _Workers
            W.Start_Thread()
        Next
        RaiseEvent ReportChange(_Completed & " of " & _Workers.Count & " workers completed. ")
    End Sub
    Private Sub BGW_Finished() Handles _BGW.RunWorkerCompleted
        RaiseEvent ReportChange("All workers are running")
    End Sub
    Public Sub Get_Event(ByVal e As Worker_Event_Parameters)
        Select Case e.EventType
            Case MultiThreading.EventType.WorkCompleted
                _Completed += 1
                RaiseEvent ReportChange(_Completed & " of " & _Workers.Count & " workers completed. " & e.Message)

            Case MultiThreading.EventType.ReportProgress
                RaiseEvent ReportChange(e.Message)
        End Select

        If _Completed = _Workers.Count Then
            RaiseEvent ProcessFinished()
        End If
    End Sub
    Public Function Include_Worker(ByVal Worker As Worker) As Boolean
        Dim _Status As Boolean
        Try
            Worker.PublishMyEvent = [Delegate].Combine(Worker.PublishMyEvent, New Worker.EventFirm(AddressOf Me.Get_Event))
            _Workers.Add(Worker)
            _Status = True
        Catch ex As Exception
            _Status = False
        End Try

        Return _Status
    End Function
#End Region
End Class

Public MustInherit Class Worker
#Region "Events"
    Public Delegate Sub EventFirm(ByVal e As Worker_Event_Parameters)
    Public PublishMyEvent As EventFirm
#End Region
#Region "Variables"
    Private WithEvents _BGW As New System.ComponentModel.BackgroundWorker
#End Region
#Region "Properties"
    Public ReadOnly Property IsBusy() As Boolean
        Get
            Return _BGW.IsBusy
        End Get
    End Property
#End Region
#Region "Methods"
    Public MustOverride Sub MyWork() Handles _BGW.DoWork
    Public Sub New()
        AddHandler _BGW.ProgressChanged, AddressOf RaiseMyEvent
        AddHandler _BGW.RunWorkerCompleted, AddressOf RaiseMyEvent

        _BGW.WorkerReportsProgress = True
        _BGW.WorkerSupportsCancellation = True
    End Sub
    Public Sub RaiseMyEvent(ByVal Sender As Object, ByVal e As Object)
        Dim R As New Worker_Event_Parameters

        Select Case e.GetType.Name
            Case "RunWorkerCompletedEventArgs"
                R.EventType = EventType.WorkCompleted
                R.Message = "Work completed."

            Case "String"
                R.EventType = EventType.ReportProgress
                R.Message = e
        End Select

        PublishMyEvent(R)
    End Sub
    Public Sub Start_Thread()
        _BGW.RunWorkerAsync()
    End Sub
    Public Sub ReportChange(ByVal Message As String)
        RaiseMyEvent(Me, Message)
    End Sub
#End Region
End Class

Public Class Worker_Event_Parameters
#Region "Variables"
    Private _EventType As EventType
    Private _Message As String
#End Region
#Region "Properties"
    Public Property EventType() As EventType
        Get
            Return _EventType
        End Get
        Set(ByVal value As EventType)
            _EventType = value
        End Set
    End Property
    Public Property Message() As String
        Get
            Return _Message
        End Get
        Set(ByVal value As String)
            _Message = value
        End Set
    End Property
#End Region
End Class
