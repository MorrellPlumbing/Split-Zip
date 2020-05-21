Imports System.Runtime.CompilerServices
Imports CommandLine

Module Module1
    Class Options
        <[Option]("f", "files", Required:=False, [Default]:=100, HelpText:="The maximum number of files to be placed in each output zip file.")>
        Public Property MaxFiles As Integer

        <[Option]("b", "bytes", Required:=False, [Default]:=300000000, HelpText:="The maximum size in bytes iof each output zip file.")>
        Public Property MaxSize As Integer

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

    <Extension>
    Private Sub Echo(ByVal aString As String)
        If Verbose Then
            Console.WriteLine(aString)
        End If
    End Sub

    Private Function SplitFiles(opts As options) As Integer
        Dim retVal As Integer = 0
        Verbose = opts.Verbose

        Console.WriteLine(Hyphens)
        Echo("Verbose output")
        Hyphens.Echo
        String.Format("Maximum files: {0}", opts.MaxFiles).Echo
        String.Format("Maximum size:  {0:#,##0} bytes", opts.MaxSize).Echo
        String.Format("Delete source: {0}", opts.DeleteSource.ToString).Echo
        If String.IsNullOrEmpty(opts.OutputFolder) Then
            String.Format("Output Folder: {0}", "Same as source file").Echo
        Else
            String.Format("Output Folder: {0}", opts.OutputFolder).Echo
        End If
        Hyphens.Echo
        Console.WriteLine("{0} file{1} to process", opts.InputFiles.Count, IIf(opts.InputFiles.Count = 1, "", "s"))
        Console.WriteLine(Hyphens)
        For Each fileName In opts.InputFiles
            Console.WriteLine("Processing {0}", IO.Path.GetFileName(fileName))
            If IO.File.Exists(fileName) Then
                Dim rezip = New Split_Zip.SplitZip With {
                    .MaxZipFiles = opts.MaxFiles,
                    .MaxZipBytes = opts.MaxSize,
                    .OutputFolder = opts.OutputFolder,
                    .SourceFilePath = fileName
                }
                Dim outputfiles = rezip.DestinationFilePaths
                If outputfiles(0) = fileName Then
                    Console.WriteLine("Original file is not a zip file or already meets limits. Use original file.")
                Else
                    Console.WriteLine("{0} output file{1} created", outputfiles.Count, IIf(outputfiles.Count = 1, "", "s"))
                    For Each output In outputfiles
                        output.Echo
                    Next
                    If opts.DeleteSource Then
                        IO.File.Delete(fileName)
                    End If
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
        Parser.Default.ParseArguments(Of options)(args).WithParsed(Function(opts As options) SplitFiles(opts)).WithNotParsed(Function(errs As IEnumerable(Of [Error])) CommandLineErrors(errs))
    End Sub
End Module
