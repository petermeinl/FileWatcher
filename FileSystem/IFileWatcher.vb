Imports System.ComponentModel
Imports System.Runtime.Serialization

Public Interface IFileWatcher
    Property WatchPath As String
    Property FileNameFilter As String
    'Property FileDetectionMethod As FileDetectionMethod
    Property IncludeSubDirectories As IO.SearchOption
    Property PollIntervalMsec As Integer
    Event NewFileDetected(ByVal sender As Object, ByVal e As FileDetectedEventArgs)
    Event [Error](ByVal sender As Object, ByVal e As FileWatcherErrorEventArgs)

    Sub Start()
    Sub [Stop]()
End Interface

<DataContract()>
Public Class FileDetectedEventArgs : Inherits EventArgs
    <DataMember()>
    Public ReadOnly DetectedTime As DateTime
    <DataMember()>
    Public ReadOnly DetectedReason As FileDetectedReason
    <DataMember()>
    Public ReadOnly FileFullName As String
    Public ReadOnly OldFileFullName As String

    Public Sub New(ByVal fileFullName As String, ByVal detectedReason As FileDetectedReason, ByVal detectedTime As DateTime, Optional ByVal oldFullFileName As String = Nothing)
        Me.FileFullName = fileFullName
        Me.DetectedReason = detectedReason
        Me.DetectedTime = detectedTime
        Me.OldFileFullName = oldFullFileName
    End Sub

    Public Overrides Function ToString() As String
        'Return String.Format("{0}={1}", Me.GetType.Name, LogHelper.ToJson(Me)) 'Does not serialize the text of enums when uses standalone
        Return String.Format("{0} {1}: {2}", Me.DetectedTime, Me.DetectedReason, IO.Path.GetFileName(Me.FileFullName))
    End Function
End Class

<DataContract()>
Public Enum FileDetectedReason
    Existed
    Created
    'Deleted
    Renamed
    'Changed
End Enum

Public Class FileWatcherErrorEventArgs : Inherits HandledEventArgs
    Public ReadOnly Exception As Exception

    Public Sub New(ByVal exception As Exception)
        Me.Exception = exception
    End Sub
End Class
Public Enum FileDetectionMethod
    Polling
    FileSystemWatcher
End Enum

Public Class WatchPathAvailabilityException : Inherits Exception
    Public ReadOnly IsWatchPathAvailable As Boolean

    Public Sub New(ByVal message As String, ByVal isWatchPathAvailable As Boolean)
        MyBase.New(message)
        Me.IsWatchPathAvailable = isWatchPathAvailable
    End Sub

End Class




