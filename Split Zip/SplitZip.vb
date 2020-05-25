Imports Ionic.Zip

''' <summary>
''' Class for creating multiple smaller zip files from a large zip file.
''' </summary>
Public Class SplitZip
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
    Public Property MaxZipFiles() As Integer
        Get
            Return _maxZipFiles
        End Get
        Set(ByVal value As Integer)
            If value <> _maxZipFiles Then
                _maxZipFiles = value
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
    Public Property MaxZipBytes() As Integer
        Get
            Return _maxZipBytes
        End Get
        Set(ByVal value As Integer)
            If value <> _maxZipBytes Then
                _maxZipBytes = value
                _destinationFilePaths = Nothing
            End If
        End Set
    End Property

    Private _sourceFilePath As String = String.Empty
    ''' <summary>
    ''' The source file to split
    ''' </summary>
    ''' <returns></returns>
    Public Property SourceFilePath() As String
        Get
            Return _sourceFilePath
        End Get
        Set(ByVal value As String)
            If _sourceFilePath <> value Then
                If IO.File.Exists(value) Then
                    _sourceFilePath = value
                    _destinationFilePaths = Nothing
                    If OutputFolder = String.Empty Then
                        OutputFolder = IO.Path.GetDirectoryName(value)
                    End If
                Else
                    Throw New IO.FileNotFoundException("File not found.", value)
                End If
            End If
        End Set
    End Property

    Private _destinationFilePaths As List(Of String) = Nothing
    ''' <summary>
    ''' Splits large zip files into smaller zip files.
    ''' Coppies zip files that meet the file and byte size limits.
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
                        _destinationFilePaths = DoNotSplitFile(_sourceFilePath)
                    Else
                        _destinationFilePaths = SplitFile(zip.EntriesSorted, _sourceFilePath)
                    End If
                End Using
            Else
                _destinationFilePaths = NonZipFile(_sourceFilePath)
            End If
        End If
        Return _destinationFilePaths
    End Function

    ''' <summary>
    ''' For files that do not need to be split, copies them to the destination folder
    ''' </summary>
    ''' <returns>Destination paths of the output files</returns>
    Private Function DoNotSplitFile(sourcefilepath As String) As IEnumerable(Of String)
        If IO.File.Exists(sourcefilepath) Then
            Dim filename = IO.Path.Combine(OutputFolder, IO.Path.GetFileNameWithoutExtension(sourcefilepath) + "_0.zip")
            IO.File.Copy(sourcefilepath, filename, True)
            Return New List(Of String) From {filename}
        Else
            Throw New IO.FileNotFoundException("File not found.", sourcefilepath)
        End If
    End Function

    ''' <summary>
    ''' For files thar are not a zip file, creates a zip file containing only that file
    ''' </summary>
    ''' <returns>Destination paths of the output files</returns>
    Private Function NonZipFile(sourcefilepath As String) As IEnumerable(Of String)
        If IO.File.Exists(sourcefilepath) Then
            Dim filename = IO.Path.Combine(OutputFolder, IO.Path.GetFileNameWithoutExtension(sourcefilepath) + "_0.zip")
            Using target = New ZipFile
                target.AddEntry(IO.Path.GetFileName(sourcefilepath), IO.File.ReadAllBytes(sourcefilepath))
                target.Save()
            End Using
            Return New List(Of String) From {filename}
        Else
            Throw New IO.FileNotFoundException("File not found.", sourcefilepath)
        End If
    End Function

    ''' <summary>
    ''' Creates a new zip file from the first MaxZipFiles or MaxZipBytes in the entries list
    ''' Called recursively it works through the entire source zip file
    ''' If any single file in the archieve is larger than MaxZipBytes it will create a zip file containing only that file
    ''' </summary>
    ''' <param name="entries">The list to process</param>
    ''' <param name="sourcefilepath"></param>
    ''' <param name="depth">How far into the original file we are - used to create the output file names</param>
    ''' <returns></returns>
    Private Function SplitFile(entries As IList(Of ZipEntry), sourcefilepath As String, Optional depth As Integer = 0) As IEnumerable(Of String)
        If IO.File.Exists(sourcefilepath) Then
            Dim i As Integer = 0
            Dim currentSize As Long = 0
            Dim retval = New List(Of String)
            Using source = ZipFile.Read(sourcefilepath)
                Using target = New ZipFile()
                    Do
                        Dim entry = entries(i)
                        target.AddEntry(entry.FileName, entry.OpenReader)
                        i += 1
                        currentSize += entry.CompressedSize
                    Loop Until i = entries.Count - 1 Or target.Count >= MaxZipFiles Or currentSize >= MaxZipBytes
                    IO.Directory.CreateDirectory(OutputFolder)
                    Dim filename = IO.Path.Combine(OutputFolder, IO.Path.GetFileNameWithoutExtension(sourcefilepath) + "_" + CStr(depth) + ".zip")
                    target.Save(filename)
                    retval.Add(filename)
                    If i < entries.Count - 1 Then
                        retval.AddRange(SplitFile(entries.Skip(i).ToList, sourcefilepath, depth + 1))
                    End If
                End Using
            End Using
            Return retval
        Else
            Throw New IO.FileNotFoundException("File not found.", sourcefilepath)
        End If
    End Function
End Class
