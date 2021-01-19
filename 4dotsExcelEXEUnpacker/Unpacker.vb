Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Reflection
Imports System.Xml
Imports System.Drawing

Public Class Unpacker
    Private prj As New XmlDocument()
    Private prj_node As XmlNode()
    Public asm As Assembly
    Public asm_name As String
    Private ext_path As String
    Public only_audio As Boolean = False
    Public only_img As Boolean = False

    Public Function CryptedImages() As Boolean
        Return (prj.SelectSingleNode("//Misc").Attributes.GetNamedItem("EncryptImages").Value.ToString() = Boolean.TrueString)
    End Function

    Public Shared Function DecryptString(ByVal Message As String, ByVal Passphrase As String) As String
        Dim utf8Encoding As UTF8Encoding = New UTF8Encoding()
        Dim md5CryptoServiceProvider As MD5CryptoServiceProvider = New MD5CryptoServiceProvider()
        Dim key As Byte() = md5CryptoServiceProvider.ComputeHash(utf8Encoding.GetBytes(Passphrase))
        Dim tripleDESCryptoServiceProvider As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()
        tripleDESCryptoServiceProvider.Key = key
        tripleDESCryptoServiceProvider.Mode = CipherMode.ECB
        tripleDESCryptoServiceProvider.Padding = PaddingMode.PKCS7
        Dim array As Byte() = Convert.FromBase64String(Message)
        Dim bytes As Byte()
        Try
            Dim cryptoTransform As ICryptoTransform = tripleDESCryptoServiceProvider.CreateDecryptor()
            bytes = cryptoTransform.TransformFinalBlock(array, 0, array.Length)
        Finally
            tripleDESCryptoServiceProvider.Clear()
            md5CryptoServiceProvider.Clear()
        End Try
        Return utf8Encoding.GetString(bytes)
    End Function

    Public Shared Function IsImage(ByVal fp As String) As Boolean
        fp = fp.ToLower()
        Dim text As String = "*.png;*.bmp;*.gif;*.jpg;*.jpeg;*.bmp;*.tiff;"
        Dim num As Integer = fp.LastIndexOf(".")
        Dim result As Boolean
        If num < 0 Then
            result = False
        Else
            Dim str As String = fp.Substring(num)
            result = (text.IndexOf("*" + str + ";") >= 0)
        End If
        Return result
    End Function

    Public Shared Function DecryptBytes(ByVal bMessage As Byte(), ByVal Passphrase As String) As Byte()
        Dim utf8Encoding As UTF8Encoding = New UTF8Encoding()
        Dim md5CryptoServiceProvider As MD5CryptoServiceProvider = New MD5CryptoServiceProvider()
        Dim key As Byte() = md5CryptoServiceProvider.ComputeHash(utf8Encoding.GetBytes(Passphrase))
        Dim tripleDESCryptoServiceProvider As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()
        tripleDESCryptoServiceProvider.Key = key
        tripleDESCryptoServiceProvider.Mode = CipherMode.ECB
        tripleDESCryptoServiceProvider.Padding = PaddingMode.PKCS7
        Dim result As Byte()
        Try
            Dim cryptoTransform As ICryptoTransform = tripleDESCryptoServiceProvider.CreateDecryptor()
            result = cryptoTransform.TransformFinalBlock(bMessage, 0, bMessage.Length)
        Finally
            tripleDESCryptoServiceProvider.Clear()
            md5CryptoServiceProvider.Clear()
        End Try
        Return result
    End Function
    Function GetPacker(ByVal printOutput As Boolean) As String
        Dim manifestResourceNames As String() = asm.GetManifestResourceNames()
        Dim aa_ As New List(Of String)
        Dim ab__ As New List(Of String)
        For Each Type In asm.GetTypes()
            aa_.Add(Type.Namespace)
            ab__.Add(Type.Name)
        Next
        Dim ab_ As String
        For x = 0 To aa_.Count - 1
            ab_ += aa_(x) & "|"
        Next
        Console.ForegroundColor = ConsoleColor.Yellow
        Dim name_
        If aa_.Contains("ConvertExcelToEXE4dots") Then
            name_ = "Excel To EXE Converter"
        ElseIf aa_.Contains("ConvertWordToEXE4dots") Then
            name_ = "Word To EXE Converter"
        ElseIf aa_.Contains("ConvertPowerpointToEXE4dots") Then
            name_ = "Powerpoint To EXE Converter"
        ElseIf aa_.Contains("PDFToEXEConverter") Then
            name_ = "PDF To EXE Converter"
        ElseIf aa_.Contains("ZIPSelfExtractor") Then
            name_ = "ZIP Self Extractor Maker"
        ElseIf aa_.Contains("LockedDocument") Then
            name_ = "EXE (Document) Locker"
        Else
            name_ = "???"
        End If
        If name_ = "???" AndAlso ab__.Contains("EXESlideshowProject") Then
            name_ = "EXE Slideshow Maker"
        End If
        If printOutput Then
            Console.WriteLine("Packer: " & name_)
            Console.WriteLine()
            Return Nothing
        Else
            Return name_
        End If
    End Function

    Private Function MoveWithinArray(ByVal array As Array, ByVal source As Integer, ByVal dest As Integer) As Array
        Dim temp As Object = array.GetValue(source)
        System.Array.Copy(array, dest, array, dest + 1, source - dest)
        array.SetValue(temp, dest)
        Return array
    End Function

    Function getEncAlg(ByVal int As Integer) As String ' The packer only contains 1 algorithm, made a function incase they add some in the future
        Select Case int
            Case 0
                Return "Triple DES"
            Case Else
                Return "???"
        End Select
    End Function

    Public Sub Extract()
        GetPacker(True)
        Dim prj_dmp As Boolean = False
        Dim zip_sem As Boolean = False
        Dim zip_sem_s As Boolean = False
        Dim audio_dmp As Boolean = False
        Dim prj_load As Boolean = False
        Dim loc_doc As Boolean = False
        ext_path = My.Application.Info.DirectoryPath & "/4dotsOfficeEXEConverterUnpacker/" & asm_name & "/"
        IO.Directory.CreateDirectory(ext_path)
        Dim img_int As Integer = 0
        Dim manifestResourceNames As String() = asm.GetManifestResourceNames()
        For x = 0 To manifestResourceNames.Count - 1
            If manifestResourceNames(x).Contains("project.xml") Then
                manifestResourceNames = MoveWithinArray(manifestResourceNames, x, 0)
                Exit For
            ElseIf manifestResourceNames(x).Contains("project.zsp") Then
                zip_sem = True
                manifestResourceNames = MoveWithinArray(manifestResourceNames, x, 0)
                Exit For
            End If
        Next
        For i As Integer = 0 To manifestResourceNames.Length - 1
            If If(zip_sem, manifestResourceNames(i).IndexOf("project.zsp") >= 0, manifestResourceNames(i).IndexOf("project.xml") >= 0) Then
                Using temp_str As New IO.MemoryStream()
                    Using FileStream As FileStream = File.Create(ext_path & If(zip_sem, "project.zsp", "project.xml"))
                        asm.GetManifestResourceStream(manifestResourceNames(i)).CopyTo(FileStream)
                    End Using
                    prj_dmp = True
                    Try
                        prj.Load(ext_path & If(zip_sem, "project.zsp", "project.xml"))
                        Console.ForegroundColor = ConsoleColor.Green
                        Console.WriteLine("""" & If(zip_sem, "project.zsp", "project.xml") & """ (Not Crypted) found & loaded!")
                        prj_load = True
                    Catch ex As Exception
                        Try
                            IO.File.WriteAllText(ext_path & "decrypted_project.xml", DecryptString(IO.File.ReadAllText(ext_path & If(zip_sem, "project.zsp", "project.xml")), "4dotsSoftware012301230123"))
                            prj.Load(ext_path & "decrypted_project.xml")
                            Console.ForegroundColor = ConsoleColor.Green
                            Console.WriteLine("""" & If(zip_sem, "project.zsp", "project.xml") & """ (Crypted) found & loaded!")
                            prj_load = True
                        Catch ex2 As Exception
                            Console.ForegroundColor = ConsoleColor.Red
                            Console.WriteLine("Failed to load project file!")
                        End Try
                    End Try
                    Console.WriteLine()
                    Try
                        If zip_sem AndAlso prj_load Then
                            If GetPacker(False) = "EXE Slideshow Maker" Then
                                Console.ForegroundColor = ConsoleColor.Yellow
                                Console.WriteLine("No password can be set for these packed files.")
                                Console.WriteLine()
                                Continue For
                            End If
                            Dim cn As New List(Of String)
                            For x = 0 To prj.SelectSingleNode("//Project").ChildNodes.Count - 1
                                cn.Add(prj.SelectSingleNode("//Project").ChildNodes(x).InnerText)
                            Next
                            Console.ForegroundColor = ConsoleColor.Yellow
                            Console.WriteLine("Encrypted: " & cn(0))
                            Console.WriteLine("Encryption Algorithm: " & getEncAlg(cn(1)))
                            Console.ForegroundColor = ConsoleColor.Green
                            Console.WriteLine("Encryption Password: " & DecryptString(cn(2), "BAD0B46C-EDB8-4BAE-B538-F7C99556A023"))
                            'Console.WriteLine("Ask For Password: " & cn(3))
                            Console.WriteLine()
                            Console.ForegroundColor = ConsoleColor.Magenta
                            Console.WriteLine("Password: " & If(cn(4) = "", "Nothing set", DecryptString(cn(4), "BAD0B46C-EDB8-4BAE-B538-F7C99556A023")))
                        Else
                            If prj_load Then
                                Try
                                    Console.ForegroundColor = ConsoleColor.Magenta
                                    If prj.SelectSingleNode("//Project").Attributes(1).Name = "Password" Then
                                        Console.WriteLine("Password: " & prj.SelectSingleNode("//Project").Attributes.GetNamedItem("Password").Value.ToString)
                                        Console.ForegroundColor = ConsoleColor.Yellow
                                        Console.WriteLine("Note: You can also use ""@#%$%DDGCS@#$%$$#%$%##%@#%$@#%fgsfgfdg"" to unlock")
                                    Else
                                        Console.WriteLine("Password: " & DecryptString(prj.SelectSingleNode("//Misc").Attributes.GetNamedItem("AskForPasswordValue").Value.ToString(), "493589549485043859430889230823"))
                                    End If
                                Catch ex As Exception
                                    Try
                                        Console.WriteLine("Password: " & DecryptString(prj.SelectSingleNode("//Misc").Attributes.GetNamedItem("AskForPasswordValue").Value.ToString(), "493589549485043859430889230823"))
                                    Catch ex2 As Exception
                                        If IO.File.ReadAllText(ext_path & "project.xml").Contains("AskForPasswordValue") = True Then
                                            Console.ForegroundColor = ConsoleColor.Red
                                            Console.WriteLine("Failed to retrieve password!")
                                        Else
                                            Console.ForegroundColor = ConsoleColor.Green
                                            Console.WriteLine("Password: Nothing set")
                                        End If
                                    End Try
                                End Try
                            Else
                                Console.ForegroundColor = ConsoleColor.Red
                                Console.WriteLine("Failed to retrieve password & encryption information!")
                            End If
                        End If
                    Catch ex As Exception
                        Console.ForegroundColor = ConsoleColor.Red
                        Console.WriteLine("Cannot get information! (project.xml cannot be processed)")
                    End Try
                    Console.WriteLine()
                End Using
            Else
                Using binaryReader As BinaryReader = New BinaryReader(asm.GetManifestResourceStream(manifestResourceNames(i)))
                    Using memoryStream As MemoryStream = New MemoryStream()
                        While True
                            Dim num As Long = 32768L
                            Dim buffer As Byte() = New Byte(num - 1) {}
                            Dim num2 As Integer = binaryReader.Read(buffer, 0, CInt(num))
                            If num2 <= 0 Then
                                Exit While
                            End If
                            memoryStream.Write(buffer, 0, num2)
                        End While
                        If zip_sem Then
                            If manifestResourceNames(i).IndexOf("zipexe.zip") >= 0 Then
                                Console.ForegroundColor = ConsoleColor.Yellow
                                Console.WriteLine("Found Main ZIP Archive!")
                                Try
                                    Using ms As New MemoryStream
                                        asm.GetManifestResourceStream(manifestResourceNames(i)).CopyTo(ms)
                                        File.WriteAllBytes(ext_path & "zipexe.zip", DecryptBytes(ms.ToArray, DecryptString(prj.SelectSingleNode("//Project").ChildNodes(2).InnerText, "BAD0B46C-EDB8-4BAE-B538-F7C99556A023")))
                                    End Using
                                    Console.ForegroundColor = ConsoleColor.Green
                                    Console.WriteLine("Main ZIP Archive Successfully dumped!")
                                    zip_sem_s = True
                                Catch ex As Exception
                                    Console.ForegroundColor = ConsoleColor.Red
                                    Console.WriteLine("Failed to dump Main ZIP Archive")
                                    Console.ResetColor()
                                End Try
                            End If
                        Else
                            If manifestResourceNames(i).IndexOf("LockedDocument.rtf") >= 0 Then
                                Console.ForegroundColor = ConsoleColor.Yellow
                                Try
                                    If prj.SelectSingleNode("//Project").Attributes.GetNamedItem("EncryptFiles").Value.ToString = "True" Then
                                        Console.WriteLine("Found LockedDocument.rtf (Crypted)!")
                                        Using ms As New MemoryStream
                                            asm.GetManifestResourceStream(manifestResourceNames(i)).CopyTo(ms)
                                            Try
                                                File.WriteAllBytes(ext_path & "LockedDocument.rtf.exe", DecryptBytes(ms.ToArray, prj.SelectSingleNode("//Project").Attributes.GetNamedItem("Password").Value.ToString))
                                            Catch ex As Exception
                                                Try
                                                    File.WriteAllBytes(ext_path & "LockedDocument.rtf", DecryptBytes(ms.ToArray, "4dotsSoftware012301230123"))
                                                Catch ex2 As Exception
                                                    Exit For
                                                    Console.ForegroundColor = ConsoleColor.Red
                                                    Console.WriteLine("Failed to dump LockedDocument.rtf")
                                                End Try
                                            End Try
                                        End Using
                                    Else
                                        Console.WriteLine("Found LockedDocument.rtf (Not Crypted)!")
                                        Using ms As New MemoryStream
                                            asm.GetManifestResourceStream(manifestResourceNames(i)).CopyTo(ms)
                                            File.WriteAllBytes(ext_path & "LockedDocument.rtf.exe", ms.ToArray)
                                        End Using
                                    End If
                                    Console.ForegroundColor = ConsoleColor.Green
                                    Console.WriteLine("LockedDocument.rtf Successfully dumped!")
                                    loc_doc = True
                                Catch ex As Exception
                                    Console.ForegroundColor = ConsoleColor.Red
                                    Console.WriteLine("Failed to dump LockedDocument.rtf")
                                    Console.ResetColor()
                                End Try
                            End If
                            If manifestResourceNames(i).IndexOf("4dotsAudio") >= 0 Then
                                If only_img Then
                                    Console.ForegroundColor = ConsoleColor.Yellow
                                    Console.WriteLine("Found Background Audio, skipping...")
                                    Console.ResetColor()
                                    Continue For
                                End If
                                Console.ForegroundColor = ConsoleColor.Green
                                Console.WriteLine("Found Background Audio File, setting it in memory to dump!")
                                Dim text As String = Path.GetTempFileName() + ".wav"
                                File.WriteAllBytes(text, memoryStream.ToArray())
                                Dim text2 As String = manifestResourceNames(i).Substring(manifestResourceNames(i).IndexOf("4dotsAudio"))
                                If text2.IndexOf("4dotsAudioBackgroundMusic") >= 0 Then
                                    Console.ForegroundColor = ConsoleColor.Yellow
                                    Console.WriteLine("dumping Background Audio File...")
                                    Using FileStream As FileStream = File.Create(ext_path & "background-music.wav")
                                        asm.GetManifestResourceStream(manifestResourceNames(i)).CopyTo(FileStream)
                                    End Using
                                    Console.ForegroundColor = ConsoleColor.Green
                                    Console.WriteLine("Background Audio File dumped!")
                                    Console.ResetColor()
                                    audio_dmp = True
                                    If only_audio Then
                                        Return
                                    End If
                                End If
                            ElseIf IsImage(manifestResourceNames(i)) Then
                                If only_audio Then
                                    Console.ForegroundColor = ConsoleColor.Yellow
                                    Console.WriteLine("Found Image, skipping...")
                                    Console.ResetColor()
                                    Continue For
                                End If
                                If GetPacker(False) = "EXE Slideshow Maker" Then
                                    GoTo img_nocr
                                End If
                                If prj.SelectSingleNode("//Misc").Attributes.GetNamedItem("EncryptImages").Value.ToString = "True" Then
                                    Try
                                        Console.ForegroundColor = ConsoleColor.Green
                                        Console.WriteLine("Found Image N°" & img_int + 1 & "! (Crypted), decrypting...")
                                        Dim buffer2 As Byte() = DecryptBytes(memoryStream.ToArray(), "433424234234-93435849839453")
                                        'File.WriteAllBytes(manifestResourceNames(i), buffer2)
                                        Dim stream As MemoryStream = New MemoryStream(buffer2)
                                        Image.FromStream(stream).Save(ext_path & manifestResourceNames(i), System.Drawing.Imaging.ImageFormat.Png)
                                        Console.ForegroundColor = ConsoleColor.Yellow
                                        Console.WriteLine("Dumping image n°" & img_int + 1 & "...")
                                        'Using fs As New FileStream(ext_path & "\image_" & i & ".png", FileMode.Create)
                                        '    memoryStream.CopyTo(fs)
                                        'End Using
                                        'item.Save(ext_path & manifestResourceNames(i), System.Drawing.Imaging.ImageFormat.Png)
                                        Console.WriteLine("Image N°" & img_int + 1 & " successfully dumped!")
                                        Console.ResetColor()
                                        img_int += 1
                                    Catch ex As Exception
                                        Console.ForegroundColor = ConsoleColor.Red
                                        Console.WriteLine("Failed to extract crypted image n°" & img_int + 1)
                                        Console.ResetColor()
                                    End Try

                                Else
img_nocr:
                                    Try
                                        Console.ForegroundColor = ConsoleColor.Green
                                        Console.WriteLine("Found Image N°" & img_int + 1 & "! (Not Crypted), decrypting...")
                                        Image.FromStream(memoryStream).Save(ext_path & manifestResourceNames(i), System.Drawing.Imaging.ImageFormat.Png)
                                        Console.ForegroundColor = ConsoleColor.Yellow
                                        Console.WriteLine("Dumping image n°" & img_int + 1 & "...")
                                        Console.ForegroundColor = ConsoleColor.Green
                                        Console.WriteLine("Image N°" & img_int + 1 & " successfully dumped!")
                                        Console.ResetColor()
                                        img_int += 1
                                    Catch ex As Exception
                                        Console.ForegroundColor = ConsoleColor.Red
                                        Console.WriteLine("Failed to extract non-crypted image n°" & img_int + 1)
                                        Console.ResetColor()
                                    End Try
                                End If
                            End If
                        End If
                    End Using
                End Using
            End If
        Next
        Console.ForegroundColor = ConsoleColor.Green
        Console.WriteLine()
        Console.Write("Done! (dumped ")
        If zip_sem Then
            If zip_sem_s Then
                Console.Write("zip file)")
            Else
                Console.Write("nothing)")
            End If
        ElseIf loc_doc Then
            Console.Write("project.xml & LockedDocument.rtf)")
        Else
            If prj_dmp Then
                Console.Write("project.xml, ")
            End If
            If audio_dmp Then
                Console.Write("background audio & ")
            End If
            Console.Write(img_int & " image(s))")
        End If
        Console.ReadKey()
        End
    End Sub
End Class