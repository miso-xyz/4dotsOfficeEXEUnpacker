Imports _4depack
Module Module1

    Sub Main(ByVal args As String())
        Console.Title = "4dotsOfficeToExeUnpacker v1.0 (Rewritten) - sinister.ly <3"
        Dim mp As New _4depack.MainProcess(args(0))
        Dim prj As ProjectFile = mp.GetProjectFile()
        Dim resources As ResourceFiles() = mp.GetResourceFiles
        Console.ForegroundColor = ConsoleColor.White
        Console.Write("Packer: ")
        Console.ForegroundColor = ConsoleColor.Yellow
        Console.WriteLine(mp.PackerName)
        Console.ForegroundColor = ConsoleColor.White
        Console.Write("Password: ")
        Console.ForegroundColor = ConsoleColor.Magenta
        Console.WriteLine(prj.Password)
        Console.ForegroundColor = ConsoleColor.White
        Console.WriteLine()
        If Not IO.Directory.Exists("4dotsOfficeUnpacker") Then
            IO.Directory.CreateDirectory("4dotsOfficeUnpacker")
        End If
        IO.Directory.CreateDirectory("4dotsOfficeUnpacker\" & IO.Path.GetFileName(args(0)))
        For x = 0 To resources.Count - 1
            Console.WriteLine("Dumping """ & resources(x).Filename & """....")
            IO.File.WriteAllBytes(IO.Path.GetFullPath("4dotsOfficeUnpacker\" & IO.Path.GetFileName(args(0)) & "\" & resources(x).Filename), resources(x).Extract(prj.EncryptImages))
        Next
        Console.WriteLine("done")
        Console.ReadKey()
    End Sub



End Module