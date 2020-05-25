Imports System.Runtime.CompilerServices
Imports CommandLine

Module Module1
    ''' <summary>
    ''' Command line options class
    ''' </summary>
    Class Options
        <[Option]("f", "files", Required:=False, HelpText:="The maximum number of files to be placed in each output zip file.")>
        Public Property MaxFiles As Integer
            Get
                Return My.Settings.MaxFiles
            End Get
            Set(value As Integer)
                If value > 0 Then
                    My.Settings.MaxFiles = value
                    My.Settings.Save()
                End If
            End Set
        End Property

        <[Option]("b", "bytes", Required:=False, HelpText:="The maximum size in bytes iof each output zip file.")>
        Public Property MaxBytes As Integer
            Get
                Return My.Settings.MaxBytes
            End Get
            Set(value As Integer)
                If value > 0 Then
                    My.Settings.MaxBytes = value
                    My.Settings.Save()
                End If
            End Set
        End Property

        <[Option]("d", "deletesource", Required:=False, HelpText:="Delete the input file on successfull completion.")>
        Public Property DeleteSource As Boolean = False

        <[Option]("w", "wait", Required:=False, HelpText:="Wait for user keypress at end.")>
        Public Property Wait As Boolean = False

        <[Option]("o", "outputFolder", Required:=False, HelpText:="Output folder for the output zip files. If ommitted it uses the source file's folder.")>
        Public Property OutputFolder As String = String.Empty

        <[Option]("v", "verbose", Required:=False, HelpText:="Set output to verbose messages.")>
        Public Property Verbose As Boolean = False

        <Value(0, Required:=True, HelpText:="Files to be split into output zip files.")>
        Public Property InputFiles As IEnumerable(Of String)
    End Class

    Private Verbose As Boolean
    Private Const Hyphens = "---------------------------------------------------------------------------"

    ''' <summary>
    ''' Prints to console if the verbose flag is set
    ''' </summary>
    ''' <param name="aString"></param>
    <Extension>
    Private Sub Echo(ByVal aString As String)
        If Verbose Then
            Console.WriteLine(aString)
        End If
    End Sub

    Private Sub ConsoleOpeningScreed(opts As Options)
        Console.WriteLine(Hyphens)
        Echo("Verbose output")
        Hyphens.Echo
        String.Format("Maximum files: {0}", opts.MaxFiles).Echo
        String.Format("Maximum size:  {0:#,##0} bytes", opts.MaxBytes).Echo
        String.Format("Delete source: {0}", opts.DeleteSource.ToString).Echo
        If String.IsNullOrEmpty(opts.OutputFolder) Then
            String.Format("Output Folder: {0}", "Same as source file").Echo
        Else
            String.Format("Output Folder: {0}", opts.OutputFolder).Echo
        End If
        Hyphens.Echo
        Console.WriteLine("{0} file{1} to process", opts.InputFiles.Count, IIf(opts.InputFiles.Count = 1, "", "s"))
        Console.WriteLine(Hyphens)
    End Sub

    ''' <summary>
    ''' Splits the source file into zip files that meet the MAxFiles and MaxBytes limits
    ''' </summary>
    ''' <param name="opts"></param>
    ''' <returns></returns>
    Private Function SplitFiles(opts As Options) As Integer
        Dim retVal As Integer = 0
        Verbose = opts.Verbose

        ConsoleOpeningScreed(opts)
        For Each fileName In opts.InputFiles
            Console.WriteLine("Processing {0}", IO.Path.GetFileName(fileName))
            If IO.File.Exists(fileName) Then
                Dim rezip = New Split_Zip.SplitZip With {
                    .MaxZipFiles = opts.MaxFiles,
                    .MaxZipBytes = opts.MaxBytes,
                    .OutputFolder = opts.OutputFolder,
                    .SourceFilePath = fileName
                }
                Dim outputfiles = rezip.DestinationFilePaths
                Console.WriteLine("{0} output file{1} created", outputfiles.Count, IIf(outputfiles.Count = 1, "", "s"))
                For Each output In outputfiles
                    output.Echo
                Next
                If opts.DeleteSource Then
                    IO.File.Delete(fileName)
                End If
            Else
                Console.WriteLine("File not found.")
                retVal += 1
            End If
            Console.WriteLine(Hyphens)
        Next
        If opts.Wait Then
            Console.Write("Press any key ... ")
            Console.ReadKey()
        End If
        Return retVal
    End Function

    ''' <summary>
    ''' Writes command line errors to the console.
    ''' </summary>
    ''' <param name="errs"></param>
    ''' <returns></returns>
    Private Function CommandLineErrors(errs As IEnumerable(Of [Error])) As Integer
        Console.WriteLine("{0} error{1} in command line arguments", errs.Count, IIf(errs.Count = 1, "", "s"))
        For Each e In errs
            Console.WriteLine([Enum].GetName(e.Tag.GetType, e.Tag))
        Next
        Console.Write("Press any key ... ")
        Console.ReadKey()
        Return errs.Count
    End Function

    <STAThread>
    Sub Main(args As String())
        Parser.Default.ParseArguments(Of Options)(args).WithParsed(Function(opts As Options) SplitFiles(opts)).WithNotParsed(Function(errs As IEnumerable(Of [Error])) CommandLineErrors(errs))
    End Sub
End Module
