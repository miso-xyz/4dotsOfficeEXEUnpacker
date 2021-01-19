Module Module1
    Dim Unpacker As New Unpacker
    Sub Main(ByVal args() As String)
        Console.Title = "4dots Office to EXE Converter Unpacker v1.0.6"
        If args.Contains("--help") Then
            Console.WriteLine("4dots Office to EXE Converter Unpacker v1.0.6 by misonothx | sinister.ly")
            Console.WriteLine()
            Console.WriteLine("--help | Lists available arguments")
            Console.WriteLine("--only_audio | only dumps the background audio")
            Console.WriteLine("--only_images | only dumps images")
            Console.WriteLine()
            Console.ReadKey()
            End
        End If
        Console.WriteLine("4dots Office to EXE Converter Unpacker v1.0.6 by misonothx | sinister.ly")
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
