Imports Meinl.LeanWork.FileSystem

Module Main
    Private WithEvents _fileWatcher As New FileWatcher("X:\temp\WatchPath\in", "*.xml")

    Sub Main()
        _fileWatcher.Start()

        Console.WriteLine("Press <Enter> to stop")
        Console.ReadLine()
    End Sub

    Private Sub _fileWatcher_NewFileDetected(ByVal sender As Object, ByVal e As Meinl.LeanWork.FileSystem.FileDetectedEventArgs) Handles _fileWatcher.NewFileDetected
        Console.WriteLine(e)
    End Sub
End Module
