Imports Ionic.Zip

Public Class SplitZip
    Public Property OutputFolder As String = String.Empty

    Private _maxZipFiles As Integer = 100
    Public Property MaxZipFiles() As Integer
        Get
            Return _maxZipFiles
        End Get
        Set(ByVal value As Integer)
            If value <> _maxZipFiles Then
                _maxZipFiles = value
                _destinationFilePaths = Nothing
                'DeleteTempFiles()
            End If
        End Set
    End Property

    Private _maxZipBytes As Integer = 300000000
    Public Property MaxZipBytes() As Integer
        Get
            Return _maxZipBytes
        End Get
        Set(ByVal value As Integer)
            If value <> _maxZipBytes Then
                _maxZipBytes = value
                _destinationFilePaths = Nothing
                'DeleteTempFiles()
            End If
        End Set
    End Property

    Private _sourceFilePath As String = String.Empty
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
                    'DeleteTempFiles()
                Else
                    Throw New IO.FileNotFoundException("File not found.", value)
                End If
            End If
        End Set
    End Property

    Private _destinationFilePaths As List(Of String) = Nothing
    Public Function DestinationFilePaths() As IEnumerable(Of String)
        If _destinationFilePaths Is Nothing And IO.File.Exists(_sourceFilePath) Then
            If ZipFile.IsZipFile(_sourceFilePath) Then
                Using zip = ZipFile.Read(_sourceFilePath)
                    Dim zipFile As New IO.FileInfo(_sourceFilePath)
                    If zipFile.Length <= MaxZipBytes And zip.Count <= MaxZipFiles Then
                        _destinationFilePaths = New List(Of String) From {_sourceFilePath}
                    Else
                        _destinationFilePaths = SplitFile(zip.EntriesSorted.ToList)
                    End If
                End Using
            Else
                _destinationFilePaths = New List(Of String) From {_sourceFilePath}
            End If
        End If
        Return _destinationFilePaths
    End Function

    Private Function SplitFile(entries As IList(Of ZipEntry), Optional depth As Integer = 0) As IEnumerable(Of String)
        Dim i As Integer = 0
        Dim currentSize As Long = 0
        Dim retval = New List(Of String)
        Using source = ZipFile.Read(_sourceFilePath)
            Using target = New ZipFile()
                Do
                    Dim entry = entries(i)
                    target.AddEntry(entry.FileName, entry.OpenReader)
                    i += 1
                    currentSize += entry.CompressedSize
                Loop Until i = entries.Count - 1 Or target.Count >= MaxZipFiles Or currentSize >= MaxZipBytes
                IO.Directory.CreateDirectory(OutputFolder)
                Dim filename = IO.Path.Combine(OutputFolder, IO.Path.GetFileNameWithoutExtension(_sourceFilePath) + "_" + CStr(depth) + ".zip")
                target.Save(filename)
                retval.Add(filename)
                If i < entries.Count - 1 Then
                    retval.AddRange(SplitFile(entries.Skip(i).ToList, depth + 1))
                End If
            End Using
        End Using
        Return retval
    End Function

    'Private Sub DeleteTempFiles()
    '    If _destinationFilePaths IsNot Nothing Then
    '        For Each filepath In _destinationFilePaths
    '            If filepath <> _sourceFilePath And IO.File.Exists(filepath) Then
    '                Try
    '                    IO.File.Delete(filepath)
    '                Catch
    '                End Try
    '            End If
    '        Next
    '    End If
    'End Sub

    '#Region "IDisposable Support"
    '    Private disposedValue As Boolean ' To detect redundant calls

    '    ' IDisposable
    '    Protected Overridable Sub Dispose(disposing As Boolean)
    '        If Not disposedValue Then
    '            If disposing Then
    '                ' TODO: dispose managed state (managed objects).
    '            End If

    '            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
    '            DeleteTempFiles()
    '            ' TODO: set large fields to null.
    '        End If
    '        disposedValue = True
    '    End Sub

    '    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    '    Protected Overrides Sub Finalize()
    '        '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '        Dispose(False)
    '        MyBase.Finalize()
    '    End Sub

    '    ' This code added by Visual Basic to correctly implement the disposable pattern.
    '    Public Sub Dispose() Implements IDisposable.Dispose
    '        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '        Dispose(True)
    '        ' TODO: uncomment the following line if Finalize() is overridden above.
    '        GC.SuppressFinalize(Me)
    '    End Sub
    '#End Region
End Class
