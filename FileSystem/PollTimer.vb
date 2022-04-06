Imports System.Threading
''' <summary>
''' Raises events at specified intervals.
''' </summary>
''' <remarks>
''' Wrapped Threading.Timer
''' - to enable using VB.NET declarative event handlers
''' - to supply a timer that does not fire while the elapsed handler runs
''' - to offer a Start/Stop API
''' Uses Threading.Timer because Timers.Timer suppresses all exceptions thrown by event handlers for the Elapsed event.
''' </remarks>
Public Class PollTimer
    Public Property DueTime As Integer
    Public Property Period As Integer
    Public Event DueElapsed(ByVal sender As Object, ByVal e As EventArgs)
    Public Event PeriodElapsed(ByVal sender As Object, ByVal e As EventArgs)

    Private _timer As New Threading.Timer(Sub() _timer_Elapsed())

    Public Sub New()
    End Sub

    Public Sub New(ByVal dueTime As Integer, ByVal period As Integer)
        Me.DueTime = dueTime
        Me.Period = period
    End Sub

    Public Sub Start()
        _timer.Change(DueTime, Timeout.Infinite)
    End Sub

    Private Sub _timer_Elapsed()
        Static _isDueTime As Boolean = True
        If _isDueTime Then
            _isDueTime = False
            RaiseEvent DueElapsed(Me, New EventArgs)
        Else
            RaiseEvent PeriodElapsed(Me, New EventArgs)
        End If
        _timer.Change(Period, Timeout.Infinite)
    End Sub

    Public Sub [Stop]()
        _timer.Dispose()
    End Sub
End Class
