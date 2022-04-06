Imports System.Threading
Imports System.IO
Imports NLog

''' <summary>
''' 
''' </summary>
''' <remarks>
''' Uses timers to not block the calling thread. Even for the first poll!
''' First poll runs immediately.
''' Next polls are scheduled when no more files exist.
''' Files are sorted by LastWriteTime.
''' May throw IOException
''' May raise one NewFileDetected event after stop has been called.
''' The consumer must take care of stopping gracefully (i.e. wait until he has finished processing the NotifyFile event)
''' Does not keep raising events when the watched directory is renamed.
''' Does not automatically resumes raising event after a watch directory accessability problem has been resolved.
''' </remarks>
Public Class FilePoller : Implements IFileWatcher
    'TODO check is this usage of properties makes sense
    Public Property WatchPath As String = "" Implements IFileWatcher.WatchPath
    Public Property FileNameFilter As String = "*.*" Implements IFileWatcher.FileNameFilter
    Public Property IncludeSubDirectories As IO.SearchOption = IO.SearchOption.TopDirectoryOnly Implements IFileWatcher.IncludeSubDirectories
    Public Property PollIntervalMsec As Integer = 5 * 1000 Implements IFileWatcher.PollIntervalMsec
    Public Event NewFileDetected(ByVal sender As Object, ByVal e As FileDetectedEventArgs) Implements IFileWatcher.NewFileDetected
    Public Event [Error](ByVal sender As Object, ByVal e As FileWatcherErrorEventArgs) Implements IFileWatcher.Error

    Const Immediately = 0
    Private WithEvents _pollTimer As New PollTimer(Immediately, PollIntervalMsec)
    Private _isStopRequested As New ManualResetEvent(False) 'We do not use a boolean, because there is not "Volatile" keyword in VB.NET

    Private Shared _trace As Logger = NLog.LogManager.GetLogger(Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString)

    Public Sub New(ByVal watchPath As String, ByVal fileNameFilter As String)
        Me.WatchPath = watchPath
        Me.FileNameFilter = fileNameFilter
    End Sub

    Public Sub Start() Implements IFileWatcher.Start
        _trace.Trace("")
        _pollTimer.Start()
    End Sub

    Private Sub _pollTimer_Elapsed(ByVal sender As Object, ByVal e As EventArgs
                                   ) Handles _pollTimer.DueElapsed, _pollTimer.PeriodElapsed
        _trace.Trace("")
        If _isStopRequested.WaitOne(0) Then Exit Sub
        DetectFiles()
    End Sub

    Private Sub DetectFiles()
        _trace.Trace("")
        Try
            Dim files = GetAllFilesSortedByLastAccess(WatchPath, FileNameFilter, IncludeSubDirectories)
            Do While files.Count > 0 And Not _isStopRequested.WaitOne(0)
                For Each fileinfo In files
                    _trace.Trace("Existing file detected={0}", fileinfo.Name)
                    NotifyFile(fileinfo.FullName, DateTime.Now)
                Next
                files = GetAllFilesSortedByLastAccess(WatchPath, FileNameFilter, IncludeSubDirectories)
            Loop
        Catch ex As IOException
            If Directory.Exists(WatchPath) Then
                Throw
            Else
                Throw (New WatchPathAvailabilityException("WatchPath is not accessible. Will not automatically recover!" & WatchPath, False))
            End If
        End Try

    End Sub

    Private Sub NotifyFile(ByVal fileFullName As String, ByVal detectedTime As DateTime)
        _trace.Trace("")
        If _isStopRequested.WaitOne(0) Then Exit Sub
        RaiseEvent NewFileDetected(Me, New FileDetectedEventArgs(fileFullName, FileDetectedReason.Existed, detectedTime))
    End Sub

    Public Sub [Stop]() Implements IFileWatcher.Stop
        _trace.Trace("")
        _isStopRequested.Set()
        _pollTimer.Stop()
    End Sub
End Class



