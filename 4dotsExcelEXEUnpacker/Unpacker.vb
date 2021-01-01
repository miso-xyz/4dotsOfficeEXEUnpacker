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

    Public Sub LoadXMLProjectInMemory()
        Using binaryReader As BinaryReader = New BinaryReader(asm.GetManifestResourceStream("project.xml"))
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
                prj.LoadXml(Encoding.[Default].GetString(memoryStream.ToArray()))
            End Using
        End Using
    End Sub

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

    Public Sub Extract()
        Dim audio_dmp As Boolean = False
        ext_path = My.Application.Info.DirectoryPath & "/4dotsExcelEXEUnpacker/" & asm_name & "/"
        IO.Directory.CreateDirectory(ext_path)
        Dim img_int As Integer = 0
        Dim manifestResourceNames As String() = asm.GetManifestResourceNames()
        For i As Integer = 2 To manifestResourceNames.Length - 1
            If manifestResourceNames(i).IndexOf("project.xml") >= 0 Then
                Using temp_str As New IO.MemoryStream()
                    Using FileStream As FileStream = File.Create(ext_path & "project.xml")
                        asm.GetManifestResourceStream(manifestResourceNames(i)).CopyTo(FileStream)
                    End Using
                    prj.Load(ext_path & "project.xml")
                    Console.ForegroundColor = ConsoleColor.Magenta
                    Console.WriteLine("Password: " & DecryptString(prj.SelectSingleNode("//Misc").Attributes.GetNamedItem("AskForPasswordValue").Value.ToString(), "493589549485043859430889230823"))
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
                                If CryptedImages() Then
                                    Console.ForegroundColor = ConsoleColor.Green
                                    Console.WriteLine("Found Image N°" & img_int & "! (Crypted), decrypting...")
                                    Dim buffer2 As Byte() = DecryptBytes(memoryStream.ToArray(), "433424234234-93435849839453")
                                    'File.WriteAllBytes(manifestResourceNames(i), buffer2)
                                    Dim stream As MemoryStream = New MemoryStream(buffer2)
                                    Dim item As Image = Image.FromStream(stream)
                                    Console.ForegroundColor = ConsoleColor.Yellow
                                    Console.WriteLine("Dumping image n°" & img_int & "...")
                                    item.Save(ext_path & manifestResourceNames(i), System.Drawing.Imaging.ImageFormat.Png)
                                    Console.WriteLine("Image N°" & img_int & " successfully dumped!")
                                    Console.ResetColor()
                                    img_int += 1
                                Else
                                    Console.ForegroundColor = ConsoleColor.Green
                                    Console.WriteLine("Found Image N°" & img_int & "! (Not Crypted), decrypting...")
                                    Dim item As Image = Image.FromStream(memoryStream)
                                    Console.ForegroundColor = ConsoleColor.Yellow
                                    Console.WriteLine("Dumping image n°" & img_int & "...")
                                    item.Save(ext_path & manifestResourceNames(i), System.Drawing.Imaging.ImageFormat.Png)
                                    Console.ForegroundColor = ConsoleColor.Green
                                    Console.WriteLine("Image N°" & img_int & " successfully dumped!")
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
        Console.Write("Done! (dumped ")
        If audio_dmp Then
            Console.Write("background audio & ")
        End If
        Console.Write(img_int & " image(s))")
        Console.ReadKey()
        End
    End Sub
End Class