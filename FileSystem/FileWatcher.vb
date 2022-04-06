Imports System.Threading
Imports System.Threading.Tasks
Imports System.IO
Imports NLog

''' <summary>
''' Watches for existing, new and renamed files in a directory and reaises events for them. Uses an initial GetFiles and listening to filesystem events.
''' </summary>
''' <remarks>
''' Uses timers to not block the calling thread. Even for the first poll.
''' First poll runs immediately.
''' Files are sorted by LastWriteTime.
''' May throw WatchPathAvailabilityException, IO.InternalBufferOverflowException
''' May raise one NewFileDetected event after stop has been called.
''' The consumer must take care of stopping gracefully (i.e. wait until he has finished processing the NotifyFile event)
''' Keeps raising events when the watched directory is renamed.
''' Automatically resumes raising event after a watch directory accessability problem has been resolved.
''' </remarks>
Public Class FileWatcher : Implements IFileWatcher
    Public Property WatchPath As String = "" Implements IFileWatcher.WatchPath
    Public Property FileNameFilter As String = "*.*" Implements IFileWatcher.FileNameFilter
    Public Property IncludeSubDirectories As IO.SearchOption = IO.SearchOption.TopDirectoryOnly Implements IFileWatcher.IncludeSubDirectories
    Public Property CheckIntervalMsec As Integer = 5 * 1000 Implements IFileWatcher.PollIntervalMsec
    Public Property InternalBufferSizeKB As Integer = Me.OptimalDefaultInternalBuffersizeKB

    Public Event NewFileDetected(ByVal sender As Object, ByVal e As FileDetectedEventArgs) Implements IFileWatcher.NewFileDetected
    Public Event [Error](ByVal sender As Object, ByVal e As FileWatcherErrorEventArgs) Implements IFileWatcher.Error

    Private Const Immediately = 0
    Private WithEvents _timer As New PollTimer(Immediately, CheckIntervalMsec)
    Private WithEvents _FSWatcher As FileSystemWatcher

    Private _stopRequestedEvent As New ManualResetEvent(False)  'We do not use a boolean isStopRequested, because there is not "Volatile" keyword in VB.NET
    Private _initialDetectEvent As New Threading.ManualResetEvent(False)

    Private Shared _trace As Logger = NLog.LogManager.GetLogger(Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString)

    Public Sub New()
    End Sub

    Public Sub New(ByVal watchPath As String, ByVal fileNameFilter As String)
        Me.WatchPath = watchPath
        Me.FileNameFilter = fileNameFilter
    End Sub

    Public Sub Start() Implements IFileWatcher.Start
        'Throw New ApplicationException("Test")
        _trace.Trace("")
        ConfigureFSWatcher()

        'We start watching before we initially detect existing files.
        'FileSystemWatcher event handlers are blocked until the initial dectection has finished.  
        _FSWatcher.EnableRaisingEvents = True
        _timer.Start()
    End Sub

    Private Sub ConfigureFSWatcher()
        _FSWatcher = New FileSystemWatcher
        _FSWatcher.Path = WatchPath
        _FSWatcher.Filter = FileNameFilter
        _FSWatcher.IncludeSubdirectories = IncludeSubDirectories
        _FSWatcher.InternalBufferSize = InternalBufferSizeKB * 1024
    End Sub

    Private Function OptimalDefaultInternalBuffersizeKB() As Integer
        'http://msdn.microsoft.com/en-us/library/ded0dc5s.aspx
        'http://msdn.microsoft.com/en-us/library/aa366778(VS.85).aspx
        Const maxInternalBufferSize = 16 * 4
        If Environment.Is64BitOperatingSystem Then
            Return maxInternalBufferSize
        Else
            Return 2 * 4
        End If
    End Function

    Private Sub _timer_DueElapsed(ByVal sender As Object, ByVal e As EventArgs) Handles _timer.DueElapsed
        'Throw New ApplicationException("Test")
        _trace.Trace("")
        DetectFiles()
        _initialDetectEvent.Set()
    End Sub

    Private Sub _timer_PeriodElapsed(ByVal sender As Object, ByVal e As EventArgs) Handles _timer.PeriodElapsed
        Static _wasWatchPathAccessible As Boolean = True

        _trace.Trace("")
        Try
            'Throw (New ApplicationException("Test"))
            If _stopRequestedEvent.WaitOne(0) Then Exit Sub
            Dim isWatchPathAccessible = Directory.Exists(WatchPath)
            If isWatchPathAccessible <> _wasWatchPathAccessible Then
                Dim message = If(isWatchPathAccessible, "<-WatchPath is accessible again: ", "->WatchPath is not accessible. Will recover automatically!: ")
                _wasWatchPathAccessible = isWatchPathAccessible
                Throw (New WatchPathAvailabilityException(message & WatchPath, _wasWatchPathAccessible))
            End If
        Catch ex As WatchPathAvailabilityException
            If Not IsErrorHandled(ex) Then
                Throw
            End If
        End Try
    End Sub

    Private Sub DetectFiles()
        _trace.Trace("")
        'In contrast to FilePoller no Loop While files.Count > 0 becauss FSWatcher is detecting new files
        Dim files = GetAllFilesSortedByLastAccess(WatchPath, FileNameFilter, IncludeSubDirectories)
        For Each fileinfo In files
            _trace.Trace("Existing file detected={0}", fileinfo.Name)
            NotifyFile(fileinfo.FullName, DateTime.Now, FileDetectedReason.Existed)
        Next
    End Sub

    Private Sub NotifyFile(ByVal fileFullName As String, ByVal detectedTime As DateTime, ByVal detectedReason As FileDetectedReason, Optional ByVal oldFullFileName As String = "")
        _trace.Trace("")
        If _stopRequestedEvent.WaitOne(0) Then Exit Sub
        RaiseEvent NewFileDetected(Me, New FileDetectedEventArgs(fileFullName, detectedReason, detectedTime, oldFullFileName))
    End Sub

    Private Sub _FSWatcher_Created(ByVal sender As Object, ByVal e As System.IO.FileSystemEventArgs) Handles _FSWatcher.Created
        Dim detectedFileInfo As New FileDetectedEventArgs(e.FullPath, FileDetectedReason.Created, DateTime.Now)
        _trace.Trace(detectedFileInfo)
        _trace.Trace("Waiting for DetecFiles={0}", _initialDetectEvent.WaitOne(0))
        _initialDetectEvent.WaitOne()
        Try
            NotifyFile(e.FullPath, DateTime.Now, FileDetectedReason.Created)
        Catch ex As Exception
            If Not IsErrorHandled(ex) Then Throw
        End Try
    End Sub

    Private Sub _FSWatcher_Renamed(ByVal sender As Object, ByVal e As System.IO.RenamedEventArgs) Handles _FSWatcher.Renamed
        _trace.Trace("Waiting until initial DetectFiles finishes...")
        _initialDetectEvent.WaitOne() 'Wait until initial polling has finished
        _trace.Trace("Renamed file detected={0}", e.Name)
        Try
            NotifyFile(e.FullPath, DateTime.Now, FileDetectedReason.Renamed, e.OldFullPath)
        Catch ex As Exception
            If Not IsErrorHandled(ex) Then Throw
        End Try
    End Sub

    Private Sub _FSWatcher_Error(ByVal sender As Object, ByVal e As System.IO.ErrorEventArgs) Handles _FSWatcher.Error
        Dim ex = e.GetException
        If Not IsErrorHandled(ex) Then Throw ex
    End Sub

    Private Function IsErrorHandled(ByVal ex As Exception)
        'Allow consumer to handle error
        Dim fileDetectorErrorEventArgs As New FileWatcherErrorEventArgs(ex)
        RaiseEvent Error(Me, fileDetectorErrorEventArgs)
        Return fileDetectorErrorEventArgs.Handled
    End Function

    Public Sub [Stop]() Implements IFileWatcher.Stop
        _trace.Trace("")
        _stopRequestedEvent.Set()
        _timer.Stop()
        _FSWatcher.Dispose()
    End Sub

End Class




