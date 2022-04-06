Imports System.IO

Public Module Functions

    Public Function IsFileLocked(ByVal ex As Exception) As Boolean
        Const sharingViolation = 32
        Const lockViolation = 33
        Dim errorCode = Runtime.InteropServices.Marshal.GetHRForException(ex) And ((1 << 16) - 1)
        Return (errorCode = sharingViolation OrElse errorCode = lockViolation)
    End Function

    Public Function IsNetworkPath(ByVal path As String) As Boolean
        If Not IO.Path.IsPathRooted(path) Then Throw New ArgumentException("Path is not rooted=" & path)
        Dim di As New DirectoryInfo(path)
        Dim drive As New DriveInfo(System.IO.Path.GetPathRoot(di.FullName))
        Return (drive.DriveType = System.IO.DriveType.Network)
    End Function

    Public Function GetAllFilesSortedByLastAccess(ByVal path As String, ByVal fileNameFilter As String, ByVal includeSubdirectories As IO.SearchOption) As IO.FileInfo()
        Dim pathInfo As New IO.DirectoryInfo(path)
        'FileSystemWatcher watches for all files if Filter=""
        Dim fileList = pathInfo.GetFiles(If(fileNameFilter = String.Empty, "*.*", fileNameFilter), includeSubdirectories)
        Return (From fileInfo In fileList Order By fileInfo.LastWriteTime Ascending Select fileInfo).ToArray
    End Function
End Module

'TODO add Sort Configuration
Public Enum FileSortOrder
    'ByName
    'ByCreationTime
    ByLastWriteTime
    '...
End Enum
Public Enum FileSortDirection
    Ascending
    'Descending
End Enum
