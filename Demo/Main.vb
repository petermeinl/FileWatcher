Imports Meinl.LeanWork.FileSystem
Imports System.Threading
Imports System.Threading.Tasks

Module Main
    Property WatchPath = "X:\temp\WatchPath\in"
    Property FileNameFilter = "*.XML"

    Private WithEvents _fileWatcher As IFileWatcher
    Private _filesQueue As New Concurrent.BlockingCollection(Of FileDetectedEventArgs)

    Sub Main()
        AddHandler System.AppDomain.CurrentDomain.UnhandledException, AddressOf UnhandledException

        Console.Title = String.Format("Watching {0} for {1} ", WatchPath, FileNameFilter)

        _fileWatcher = New FileWatcher(WatchPath, FileNameFilter)
        '_fileWatcher = New FilePoller(WatchPath, FileNameFilter) 
        _fileWatcher.PollIntervalMsec = 3 * 1000
        _fileWatcher.Start()

        'TODO Add graceful cancelling using CancellationTokenSource
        Task.Factory.StartNew(Sub()
                                  Do
                                      Console.WriteLine(_filesQueue.Take())
                                  Loop
                              End Sub)
        WriteLineInColor("Watching for new Files...", ConsoleColor.Green)
        PromptUser("Press <Esc> to Stop:", ConsoleColor.White)
    End Sub

    Private Sub _fileWatcher_NewFileDetected(ByVal sender As Object, ByVal e As Meinl.LeanWork.FileSystem.FileDetectedEventArgs) Handles _fileWatcher.NewFileDetected
        'Console.WriteLine(e)
        _filesQueue.Add(e)
    End Sub

    'Filewatcher:
    '-Keeps raising events when the watched directory is renamed.
    '-Automatically resumes raising event after a watch directory accessability problem has been resolved.
    'Using the Error Event you can observe and handle directory problems
    Private Sub _fileWatcher_Error(ByVal sender As Object, ByVal e As Meinl.LeanWork.FileSystem.FileWatcherErrorEventArgs) Handles _fileWatcher.Error
        Select Case True
            Case TypeOf e.Exception Is IO.InternalBufferOverflowException
                Abort(e.Exception)
            Case TypeOf e.Exception Is WatchPathAvailabilityException
                Dim ex = CType(e.Exception, WatchPathAvailabilityException)
                WriteLineInColor(ex.Message, ConsoleColor.Yellow)
                WriteLineInColor("Continuing...", ConsoleColor.Green)
                e.Handled = True
            Case Else
                Abort(e.Exception)
        End Select
    End Sub

    Sub UnhandledException(ByVal sender As Object, ByVal e As UnhandledExceptionEventArgs)
        Abort(CType(e.ExceptionObject, Exception))
    End Sub

    Sub Abort(ByVal ex As Exception)
        '_trace.FatalException("Unexpected exception:", ex)
        PromptUser("Unexpected exception: " & ex.Message, ConsoleColor.Red)
        Const failed = -1
        Environment.Exit(failed)
    End Sub

#Region "Console helpers"
    Public Sub PromptUser(ByVal message As String, ByVal foregroundColor As ConsoleColor)
        WriteLineInColor(message, foregroundColor)
        Do Until (Console.ReadKey(True).Key = ConsoleKey.Escape) : Loop
    End Sub

    Public Sub WriteLineInColor(ByVal message As String, ByVal foregroundColor As ConsoleColor)
        Console.ForegroundColor = foregroundColor
        Console.WriteLine(message)
        Console.ResetColor()
    End Sub
#End Region
End Module
