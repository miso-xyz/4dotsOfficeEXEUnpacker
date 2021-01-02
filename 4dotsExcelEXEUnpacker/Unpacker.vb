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
    Sub GetPacker()
        Dim manifestResourceNames As String() = asm.GetManifestResourceNames()
        Dim aa_ As New List(Of String)
        For Each Type In asm.GetTypes()
            aa_.Add(Type.Namespace)
        Next
        Dim ab_ As String
        For x = 0 To aa_.Count - 1
            ab_ += aa_(x) & "|"
        Next
        Console.ForegroundColor = ConsoleColor.Yellow
        Select Case aa_(1)
            Case "ConvertExcelToEXE4dots"
                Console.WriteLine("Packer: Excel To EXE Converter")
            Case "ConvertWordToEXE4dots"
                Console.WriteLine("Packer: Word To EXE Converter")
            Case "ConvertPowerpointToEXE4dots"
                Console.WriteLine("Packer: Powerpoint To EXE Converter")
            Case "PDFToEXEConverter"
                Console.WriteLine("Packer: PDF To EXE Converter")
            Case Else
                Console.WriteLine("Packer: ???")
        End Select
    End Sub

    Private Function MoveWithinArray(ByVal array As Array, ByVal source As Integer, ByVal dest As Integer) As Array
        Dim temp As Object = array.GetValue(source)
        System.Array.Copy(array, dest, array, dest + 1, source - dest)
        array.SetValue(temp, dest)
        Return array
    End Function

    Public Sub Extract()
        GetPacker()
        Dim prj_dmp As Boolean = False
        Dim audio_dmp As Boolean = False
        ext_path = My.Application.Info.DirectoryPath & "/4dotsOfficeEXEConverterUnpacker/" & asm_name & "/"
        IO.Directory.CreateDirectory(ext_path)
        Dim img_int As Integer = 0
        Dim manifestResourceNames As String() = asm.GetManifestResourceNames()
        For x = 0 To manifestResourceNames.Count - 1
            If manifestResourceNames(x).Contains("project.xml") Then
                manifestResourceNames = MoveWithinArray(manifestResourceNames, x, 0)
                Exit For
            End If
        Next
        For i As Integer = 0 To manifestResourceNames.Length - 1
            If manifestResourceNames(i).IndexOf("project.xml") >= 0 Then
                Using temp_str As New IO.MemoryStream()
                    Using FileStream As FileStream = File.Create(ext_path & "project.xml")
                        asm.GetManifestResourceStream(manifestResourceNames(i)).CopyTo(FileStream)
                    End Using
                    prj_dmp = True
                    prj.Load(ext_path & "project.xml")
                    Console.ForegroundColor = ConsoleColor.Magenta
                    Try
                        Console.WriteLine("Password: " & DecryptString(prj.SelectSingleNode("//Misc").Attributes.GetNamedItem("AskForPasswordValue").Value.ToString(), "493589549485043859430889230823"))
                    Catch ex As Exception
                        If IO.File.ReadAllText(ext_path & "project.xml").Contains("AskForPasswordValue") AndAlso prj.SelectSingleNode("//Misc").Attributes.GetNamedItem("AskForPasswordValue").Value <> "" Then
                            Console.ForegroundColor = ConsoleColor.Red
                            Console.WriteLine("Failed to retrieve password!")
                        Else
                            Console.ForegroundColor = ConsoleColor.Green
                            Console.WriteLine("Password: Nothing set")
                        End If
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
                            Try
                                If prj.SelectSingleNode("//Misc").Attributes.GetNamedItem("EncryptImages").Value.ToString = "True" Then
                                    Console.ForegroundColor = ConsoleColor.Green
                                    Console.WriteLine("Found Image N°" & img_int + 1 & "! (Crypted), decrypting...")
                                    Dim buffer2 As Byte() = DecryptBytes(memoryStream.ToArray(), "433424234234-93435849839453")
                                    'File.WriteAllBytes(manifestResourceNames(i), buffer2)
                                    Dim stream As MemoryStream = New MemoryStream(buffer2)
                                    Image.FromStream(stream).Save(manifestResourceNames(i), System.Drawing.Imaging.ImageFormat.Png)
                                    Console.ForegroundColor = ConsoleColor.Yellow
                                    Console.WriteLine("Dumping image n°" & img_int + 1 & "...")
                                    'Using fs As New FileStream(ext_path & "\image_" & i & ".png", FileMode.Create)
                                    '    memoryStream.CopyTo(fs)
                                    'End Using
                                    'item.Save(ext_path & manifestResourceNames(i), System.Drawing.Imaging.ImageFormat.Png)
                                    Console.WriteLine("Image N°" & img_int + 1 & " successfully dumped!")
                                    Console.ResetColor()
                                    img_int += 1
                                Else
                                    Console.ForegroundColor = ConsoleColor.Green
                                    Console.WriteLine("Found Image N°" & img_int + 1 & "! (Not Crypted), decrypting...")
                                    Image.FromStream(memoryStream).Save(manifestResourceNames(i), System.Drawing.Imaging.ImageFormat.Png)
                                    Console.ForegroundColor = ConsoleColor.Yellow
                                    Console.WriteLine("Dumping image n°" & img_int + 1 & "...")
                                    'Using fs As New FileStream(ext_path & "\image_" & i & ".png", FileMode.Create)
                                    '    memoryStream.CopyTo(fs)
                                    'End Using
                                    'item.Save(ext_path & manifestResourceNames(i), System.Drawing.Imaging.ImageFormat.Png)
                                    Console.ForegroundColor = ConsoleColor.Green
                                    Console.WriteLine("Image N°" & img_int + 1 & " successfully dumped!")
                                    Console.ResetColor()
                                    img_int += 1
                                End If
                            Catch
                            End Try
                        End If
                    End Using
                End Using
            End If
        Next
        Console.ForegroundColor = ConsoleColor.Green
        Console.WriteLine()
        Console.Write("Done! (dumped ")
        If prj_dmp Then
            Console.Write("project.xml, ")
        End If
        If audio_dmp Then
            Console.Write("background audio & ")
        End If
        Console.Write(img_int & " image(s))")
        Console.ReadKey()
        End
    End Sub
End Class