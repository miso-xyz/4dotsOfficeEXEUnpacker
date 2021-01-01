Module Module1
    Dim Unpacker As New Unpacker
    Sub Main(ByVal args() As String)
        Console.Title = "4dots Excel to EXE Unpacker v1.0.1"
        If args.Contains("--help") Then
            Console.WriteLine("4dots Excel to EXE Unpacker v1.0.1 by misonothx | sinister.ly | miyako <3")
            Console.WriteLine("Tested with Version 2.0")
            Console.WriteLine()
            Console.WriteLine("--help | Lists available arguments")
            Console.WriteLine("--only_audio | only dumps the background audio")
            Console.WriteLine("--only_images | only dumps images")
            Console.WriteLine()
            Console.WriteLine("PS: The actual Excel file cannot be dumped due to 4dots's packer taking screenshots of the file instead of actual giving the file")
            Console.WriteLine()
            Console.WriteLine("PS N°2: The version that the tool has been tested with actually isn't even able to include the screenshots it took of the file, the unpacker however can extract them yet isn't used")
            Console.ReadKey()
            End
        End If
        Console.WriteLine("4dots Excel to EXE Unpacker v1.0.1 by misonothx | sinister.ly | miyako <3")
        Console.Write("arguments given: ")
        If args.Length = 0 Then
            Console.Write("nothing")
        Else
            For x = 0 To args.Count - 1
                If IO.File.Exists(args(x)) Then
                    Console.Write("%input_file%")
                End If
            Next
        End If
        Console.WriteLine()
        Console.WriteLine()
        For x = 0 To args.Count - 1
            If IO.File.Exists(args(x)) Then
                Try
                    Unpacker.asm = System.Reflection.Assembly.LoadFile(args(x))
                Catch ex As Exception
                    Console.ForegroundColor = ConsoleColor.Red
                    Console.WriteLine("The input file isn't a valid assembly file!")
                    Console.ReadKey()
                    End
                End Try
                Unpacker.asm_name = IO.Path.GetFileName(args(x))
                If args.Contains("--only_audio") Then
                    Unpacker.only_audio = True
                ElseIf args.Contains("--only_img") Then
                    Unpacker.only_img = True
                End If
                Unpacker.Extract()
            End If
        Next
        Console.ForegroundColor = ConsoleColor.Red
        Console.WriteLine("No input files given!")
        Console.ReadKey()
        End
    End Sub

End Module
