Imports ionic.zip
Imports Upload_Class_Library

''' <summary>
''' Class for creating multiple smaller zip files from a large zip file.
''' </summary>
Public Class SplitZip
    Public Class SplitZipStatus
        Implements IStatus
        Public Property SourceFile As String
        Public Property DestinationFile As String
        Public Property InnerFileCount As Integer
        Public Property TotalInnerFileCount As Integer
        ' ReSharper disable once UnusedMember.Global
        Public ReadOnly Property ProgressRatio As Single
            Get
                If TotalInnerFileCount = 0 Then
                    Return -1
                End If
                Return InnerFileCount / TotalInnerFileCount
            End Get
        End Property

        Public ReadOnly Property Message As String Implements IStatus.message
            Get
                Return $"Unzipping {IO.Path.GetFileName(SourceFile)}: {InnerFileCount} of {TotalInnerFileCount} remaining."
            End Get
        End Property
    End Class

    Public Class SavingFileStatus
        Implements IStatus
        Public Property DestinationFile As String
        Public ReadOnly Property Message As String Implements IStatus.Message
            Get
                Return $"Saving {DestinationFile} ..."
            End Get
        End Property
    End Class

    Private ReadOnly _status As New SplitZipStatus

    Public Event Progress(status As IStatus)

    ''' <summary>
    ''' Destination folder for the new smaller zip files
    ''' </summary>
    ''' <returns></returns>
    Public Property OutputFolder As String = String.Empty

    ''' <summary>
    ''' Default value of 100 is used as this is the (undocumented) limit of the SharePoint extract folder action for Microsoft Power Automate
    ''' https://docs.microsoft.com/en-us/connectors/sharepointonline/#extract-folder
    ''' </summary>
    Private _maxZipFiles As Integer = 100
    ''' <summary>
    ''' Maximum number of files in each new zip file
    ''' </summary>
    ''' <returns></returns>
    Public Property MaxZipFiles As Integer
        Get
            Return _maxZipFiles
        End Get
        Set
            If Value <> _maxZipFiles Then
                _maxZipFiles = Value
                _destinationFilePaths = Nothing
            End If
        End Set
    End Property

    ''' <summary>
    ''' Default value of is used as this is the (undocumented) limit of the SharePoint extract folder action for Microsoft Power Automate
    ''' https://docs.microsoft.com/en-us/connectors/sharepointonline/#extract-folder
    ''' </summary>
    Private _maxZipBytes As Integer = 314572800
    ''' <summary>
    ''' Maximum number of bytes in each new zip file
    ''' </summary>
    ''' <returns></returns>
    Public Property MaxZipBytes As Integer
        Get
            Return _maxZipBytes
        End Get
        Set
            If Value <> _maxZipBytes Then
                _maxZipBytes = Value
                _destinationFilePaths = Nothing
            End If
        End Set
    End Property

    Private _sourceFilePath As String = String.Empty
    ''' <summary>
    ''' The source file to split
    ''' </summary>
    ''' <returns></returns>
    Public Property SourceFilePath As String
        Get
            Return _sourceFilePath
        End Get
        Set
            If _sourceFilePath <> Value Then
                If IO.File.Exists(Value) Then
                    _sourceFilePath = Value
                    _destinationFilePaths = Nothing
                    If OutputFolder = String.Empty Then
                        OutputFolder = IO.Path.GetDirectoryName(Value)
                    End If
                    _status.SourceFile = Value
                Else
                    Throw New IO.FileNotFoundException("File not found.", Value)
                End If
            End If
        End Set
    End Property

    Private _destinationFilePaths As List(Of String) = Nothing
    ''' <summary>
    ''' Splits large zip files into smaller zip files.
    ''' Copies zip files that meet the file and byte size limits.
    ''' Creates a zip file for non-zip files.
    ''' WARNING: Overwrites any destination files with extreme prejudice.
    ''' </summary>
    ''' <returns>Destination paths of the new smaller zip files</returns>
    Public Function DestinationFilePaths() As IEnumerable(Of String)
        If _destinationFilePaths Is Nothing And IO.File.Exists(_sourceFilePath) Then
            If ZipFile.IsZipFile(_sourceFilePath) Then
                Using zip = ZipFile.Read(_sourceFilePath)
                    Dim zipFile As New IO.FileInfo(_sourceFilePath)
                    If zipFile.Length <= MaxZipBytes And zip.Count <= MaxZipFiles Then
                        _destinationFilePaths = DoNotSplitFile()
                    Else
                        Dim files = zip.EntriesSorted.Where(Function(e) Not e.IsDirectory).ToList
                        _status.TotalInnerFileCount = files.Count
                        _status.InnerFileCount = files.Count
                        _destinationFilePaths = SplitFile(files)
                    End If
                End Using
            Else
                _destinationFilePaths = NonZipFile()
            End If
        End If
        Return _destinationFilePaths
    End Function

    ''' <summary>
    ''' For files that do not need to be split, copies them to the destination folder
    ''' </summary>
    ''' <returns>Destination paths of the output files</returns>
    Private Function DoNotSplitFile() As IEnumerable(Of String)
        If IO.File.Exists(SourceFilePath) Then
            Dim filename = IO.Path.Combine(OutputFolder, IO.Path.GetFileNameWithoutExtension(SourceFilePath) + "_0.zip")
            _status.DestinationFile = filename
            _status.TotalInnerFileCount = 1
            _status.InnerFileCount = 0
            RaiseEvent Progress(_status)
            RaiseEvent Progress(New SavingFileStatus With {.DestinationFile = _status.DestinationFile})
            IO.File.Copy(SourceFilePath, filename, True)
            Return New List(Of String) From {filename}
        Else
            Throw New IO.FileNotFoundException("File not found.", SourceFilePath)
        End If
    End Function

    ''' <summary>
    ''' For files thar are not a zip file, creates a zip file containing only that file
    ''' </summary>
    ''' <returns>Destination paths of the output files</returns>
    Private Function NonZipFile() As IEnumerable(Of String)
        If IO.File.Exists(SourceFilePath) Then
            Dim filename = IO.Path.Combine(OutputFolder, IO.Path.GetFileNameWithoutExtension(SourceFilePath) + "_0.zip")
            _status.DestinationFile = filename
            _status.TotalInnerFileCount = 1
            _status.InnerFileCount = 0
            RaiseEvent Progress(_status)
            Using target = New ZipFile
                target.AddEntry(IO.Path.GetFileName(SourceFilePath), IO.File.ReadAllBytes(SourceFilePath))
                RaiseEvent Progress(New SavingFileStatus With {.DestinationFile = _status.DestinationFile})
                target.Save(filename)
            End Using
            Return New List(Of String) From {filename}
        Else
            Throw New IO.FileNotFoundException("File not found.", SourceFilePath)
        End If
    End Function

    ''' <summary>
    ''' Creates a new zip file from the first MaxZipFiles or MaxZipBytes in the entries list
    ''' Called recursively it works through the entire source zip file
    ''' If any single file in the archive is larger than MaxZipBytes it will create a zip file containing only that file
    ''' </summary>
    ''' <param name="entries">The list to process</param>
    ''' <param name="depth">How far into the original file we are - used to create the output file names</param>
    ''' <returns></returns>
    Private Function SplitFile(entries As IList(Of ZipEntry), Optional depth As Integer = 0) As IEnumerable(Of String)
        If IO.File.Exists(SourceFilePath) Then
            Dim i = 0
            Dim currentSize As Long = 0
            Dim retVal = New List(Of String)
            Using unused = ZipFile.Read(SourceFilePath)
                Using target = New ZipFile()
                    IO.Directory.CreateDirectory(OutputFolder)
                    Dim filename = IO.Path.Combine(OutputFolder, IO.Path.GetFileNameWithoutExtension(SourceFilePath) + "_" + CStr(depth) + ".zip")
                    _status.DestinationFile = filename
                    Do
                        _status.InnerFileCount -= 1
                        RaiseEvent Progress(_status)
                        Dim entry = entries(i)
                        target.AddEntry(entry.FileName, entry.OpenReader)
                        i += 1
                        currentSize += entry.CompressedSize
                    Loop Until i = entries.Count - 1 Or target.Count >= MaxZipFiles Or currentSize >= MaxZipBytes
                    RaiseEvent Progress(New SavingFileStatus With {.DestinationFile = _status.DestinationFile})
                    target.Save(filename)
                    retval.Add(filename)
                    If i < entries.Count - 1 Then
                        retval.AddRange(SplitFile(entries.Skip(i - 1).ToList, depth + 1))
                    End If
                End Using
            End Using
            Return retval
        Else
            Throw New IO.FileNotFoundException("File not found.", SourceFilePath)
        End If
    End Function
End Class
