Imports System.IO
Imports System.Xml
Imports System.Security.Principal
Imports Microsoft.VisualBasic.FileIO
Imports System.ComponentModel
Imports System.Threading

Module Module1

    Dim ConfigFileName As String = "ActivDeploy.Config.Xml" 'файл конфига по умолчанию
    Dim TmpFolderName As String = "DeployTmpFolder" 'папка где сохраняем файлы которые в данный момент заняты
    Dim confdoc As XmlDocument = New XmlDocument()
    Dim MySourceDir As String 'MySourceDir
    Dim DistDirectories() As String
    Dim SourceExtFiles() As String
    Dim SourceDirs() As String
    Dim hm As Boolean = False
    Dim np As Boolean = False
    Dim cc As Boolean = False
    Dim contentcompare As Boolean = False
    Dim subdir As Boolean = False
    Dim WithEvents worker As BackgroundWorker = New BackgroundWorker()
    Sub Main()
        'If IsUserAdministrator() = False Then
        '    ConsoleMsgPrint("Допустимо только с правами администратора...")
        '    Exit Sub
        'End If
        MySourceDir = AppDomain.CurrentDomain.BaseDirectory
        Dim clArgs() As String = Environment.GetCommandLineArgs()

        For i = 0 To clArgs.Length - 1
            If i > 0 Then
                Dim item As String = clArgs(i)
                Select Case item
                    Case "-h", "-help", "?", "-?"
                        Dim helpmsg As String
                        helpmsg = "параметры:" + vbCrLf +
                    "  имя xml-файла шаблонов и расположений (по умолчанию-ActivDeploy.Config.Xml)" + vbCrLf +
                    "  -s : включая подпапки" + vbCrLf +
                    "  -hm : (hide message) не выводить сообщения на экран" + vbCrLf +
                    "  -np : (not pause) не останавливаться после вывода сообщения на экран" + vbCrLf +
                    "  -cc : (create config) программа создаст config-файл, но делать ничего не будет" + vbCrLf +
                    "  -contentcompare : при всех прочих сравнивает и содержимое на идентичность"
                        ConsoleMsgPrint(helpmsg)
                        Exit Sub
                    Case "-hm" : hm = True
                    Case "-np" : np = True
                    Case "-cc" : cc = True
                    Case "-s" : subdir = True
                    Case "-contentcompare" : contentcompare = True
                    Case Else : ConfigFileName = item
                End Select
            End If
        Next
        If File.Exists(ConfigFileName) Then
            Dim fi As FileInfo = New FileInfo(ConfigFileName)
            ConfigDocRead(fi.FullName)
        Else
            'ReDim DistDirectories(1)
            'DistDirectories(0) = "D:\tmp1"
            'DistDirectories(1) = "D:\tmp2"

            ReDim DistDirectories(4)
            DistDirectories(0) = "\\10.0.0.4\netlogon\startbase"
            DistDirectories(1) = "\\s-rds04\c$\activ"
            DistDirectories(2) = "\\s-rds05\c$\activ"
            DistDirectories(3) = "\\s-rds06\c$\activ"
            DistDirectories(4) = "\\termvip\c$\activ"

            ReDim SourceExtFiles(2)
            SourceExtFiles(0) = "*.dll"
            SourceExtFiles(1) = "*.tlb"
            SourceExtFiles(2) = "*.xml"

            ReDim SourceDirs(0)
            SourceDirs(0) = "*"

            If ConfigDocCreate() = 1 OrElse cc = True Then Exit Sub
        End If

        If Directory.Exists(MySourceDir) = False Then
            ConsoleMsgPrint("Не найдена исходная папка...")
            Exit Sub
        End If
        worker.WorkerSupportsCancellation = True
        Console.Write("Идет обработка ")
        worker.RunWorkerAsync()

        For Each DistDir In DistDirectories
            'разбираемся с временной папкой
            If Directory.Exists(DistDir + "\" + TmpFolderName) Then
                DeleteFromTmpFolder(DistDir + "\" + TmpFolderName) 'очистили временную папку 
            Else
                Dim di As DirectoryInfo = New DirectoryInfo(DistDir)
                di.CreateSubdirectory(TmpFolderName)
            End If
            '----------------------
            Dim tmpfolderwasinuse As Boolean = False
            If MySourceDir <> String.Empty AndAlso DistDir <> MySourceDir Then
                If Directory.Exists(DistDir) Then
                    FilesProcessing(DistDir, tmpfolderwasinuse)
                    If subdir = True Then FoldersProcessing(DistDir, tmpfolderwasinuse)
                End If
            End If
            'удаляем временную папку если она не использовалась
            If tmpfolderwasinuse = False Then
                Try
                    Directory.Delete(DistDir + "\" + TmpFolderName, True)
                Catch
                End Try
            End If
        Next
        worker.CancelAsync()
        worker.Dispose()
        If np = False Then
            Console.WriteLine(vbCrLf + "Нажмите клавишу для окончания...")
            Console.ReadKey()
        End If
    End Sub

    Sub worker_DoWork(sender As Object, e As DoWorkEventArgs) Handles worker.DoWork
        Dim i As Integer = 0
        While (True)
            If worker.CancellationPending Then Exit While
            i += 1
            Console.CursorVisible = False
            Select Case (i Mod 4)
                Case 0 : Console.Write("|")
                Case 1 : Console.Write("/")
                Case 2 : Console.Write("-")
                Case 3 : Console.Write("\")
            End Select
            Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop)
            Thread.Sleep(150)
        End While
    End Sub
    Private Sub ConsoleMsgPrint(Msg As String)
        If Msg <> "" AndAlso hm = False Then Console.WriteLine(Msg)
        If np = False Then
            Console.WriteLine("Нажмите любую клавишу для продолжения...")
            Console.ReadKey()
        End If
    End Sub

    Private Sub FilesProcessing(DistDir As String, ByRef tmpfolderwasinuse As Boolean)
        Dim di As DirectoryInfo = New DirectoryInfo(MySourceDir)
        Dim ErrFileName As String
        Dim tmp = 0
        For Each SourceExt In SourceExtFiles
            Dim MyFiles() As FileInfo = di.GetFiles(SourceExt)
            For Each Fl In MyFiles
                Dim erroccured As Boolean = False
                Dim filesthesame As Boolean
                Dim ex As Exception = Nothing
                Dim tmpname = DistDir + "\" + Fl.Name
                ErrFileName = Fl.Name
                If File.Exists(tmpname) Then
                    'Debug.WriteLine(tmpname)
                    'Console.WriteLine(tmpname)
                    'If Fl.Name = "NomenRef.tlb" Then Stop
                    If FileCompare(Fl.FullName, tmpname, contentcompare) = False Then
                        filesthesame = False
                        Try
                            File.Delete(tmpname)
                        Catch
                            erroccured = True
                        End Try
                        If erroccured = True Then
                            Try
                                erroccured = False
                                Dim SourceFn As String = DistDir + "\" + Fl.Name
                                Dim ResultFn As String = DistDir + "\" + TmpFolderName + "\" + Fl.Name
                                If File.Exists(ResultFn) Then
                                    Dim i As Integer = 0
                                    Dim FnWithouExt As String
                                    Dim FnExt As String
                                    FnWithouExt = Path.GetFileNameWithoutExtension(Fl.Name)
                                    FnExt = Path.GetExtension(Fl.Name)
                                    ResultFn = DistDir + "\" + TmpFolderName + "\" + FnWithouExt + i.ToString + FnExt
                                    Do While File.Exists(ResultFn) = True
                                        i += 1
                                        ResultFn = DistDir + "\" + TmpFolderName + "\" + FnWithouExt + i.ToString + FnExt
                                    Loop
                                End If
                                File.Move(SourceFn, ResultFn)
                                tmpfolderwasinuse = True
                            Catch ex
                                erroccured = True
                            End Try
                        End If
                    Else
                        filesthesame = True
                    End If
                Else
                    filesthesame = False
                End If
                If erroccured = False AndAlso filesthesame = False Then
                    Try
                        File.Copy(Fl.FullName, tmpname)
                        File.SetCreationTime(tmpname, File.GetCreationTime(Fl.FullName))
                        File.SetLastWriteTime(tmpname, File.GetLastWriteTime(Fl.FullName))
                    Catch ex
                        erroccured = True
                    End Try
                End If
                If erroccured = True Then
                    If ex IsNot Nothing Then
                        ConsoleMsgPrint("Ошибка: " + ErrFileName + vbCrLf + ex.Message)
                    Else
                        ConsoleMsgPrint("Ошибка: " + ErrFileName + vbCrLf + "???")
                    End If
                End If
            Next
        Next
    End Sub

    Private Sub FoldersProcessing(DistDir As String, ByRef tmpfolderwasinuse As Boolean)
        Dim di As DirectoryInfo = New DirectoryInfo(MySourceDir)
        For Each SourceDir In SourceDirs
            Dim MyDirs() As DirectoryInfo = di.GetDirectories(SourceDir)
            For Each Dr In MyDirs
                Dim erroccured As Boolean = False
                Dim ex As Exception = Nothing
                If Directory.Exists(DistDir + "\" + Dr.Name) Then
                    Try
                        erroccured = False
                        Dim SourceDr As String = DistDir + "\" + Dr.Name
                        Dim ResultDr As String = DistDir + "\" + TmpFolderName + "\" + Dr.Name
                        If Directory.Exists(ResultDr) Then
                            Dim i As Integer = 0
                            ResultDr = DistDir + "\" + TmpFolderName + "\" + Dr.Name + i.ToString
                            Do While File.Exists(ResultDr) = True
                                i += 1
                                ResultDr = DistDir + "\" + TmpFolderName + "\" + Dr.Name + i.ToString
                            Loop
                        End If
                        Directory.Move(SourceDr, ResultDr)
                        tmpfolderwasinuse = True
                    Catch ex
                        erroccured = True
                    End Try
                End If
                If erroccured = False Then
                    Try
                        FileSystem.CopyDirectory(Dr.FullName, DistDir + "\" + Dr.Name)
                    Catch ex
                        erroccured = True
                    End Try
                End If
                If erroccured = True Then
                    If ex IsNot Nothing Then ConsoleMsgPrint("Ошибка: " + ex.Message) Else ConsoleMsgPrint("Ошибка: ???")
                End If
            Next
        Next
    End Sub

    Private Function ConfigDocCreate() As Integer
        '//(1) the xml declaration Is recommended, but Not mandatory
        Dim XmlDeclaration As XmlDeclaration = confdoc.CreateXmlDeclaration("1.0", "UTF-8", String.Empty)
        confdoc.AppendChild(XmlDeclaration)
        Dim root As XmlElement = confdoc.CreateElement("body")
        confdoc.AppendChild(root)

        Dim SourceStore As XmlElement = confdoc.CreateElement(String.Empty, "SourceStorage", String.Empty)
        SourceStore.SetAttribute("SourceDirectory", MySourceDir)
        root.AppendChild(SourceStore)

        Dim TmpFolder As XmlElement = confdoc.CreateElement(String.Empty, "TmpFolder", String.Empty)
        TmpFolder.SetAttribute("TmpFolderName", TmpFolderName)
        root.AppendChild(TmpFolder)

        Dim i As Int16
        If DistDirectories.Length > 0 Then
            i = 0
            Dim DistanationStoreXML As XmlElement = confdoc.CreateElement(String.Empty, "DistanationStorage", String.Empty)
            For Each item In DistDirectories
                DistanationStoreXML.SetAttribute("DistDirectory" + i.ToString, item)
                i += 1
            Next
            root.AppendChild(DistanationStoreXML)
        End If

        If SourceExtFiles.Length > 0 Then
            i = 0
            Dim SourceExtsXML As XmlElement = confdoc.CreateElement(String.Empty, "SourceExt", String.Empty)
            For Each item In SourceExtFiles
                SourceExtsXML.SetAttribute("SourceExtItem" + i.ToString, item)
                i += 1
            Next
            root.AppendChild(SourceExtsXML)
        End If

        If SourceDirs.Length > 0 Then
            i = 0
            Dim SourceDirsXML As XmlElement = confdoc.CreateElement(String.Empty, "SourceDir", String.Empty)
            For Each item In SourceDirs
                SourceDirsXML.SetAttribute("SourceDirItem" + i.ToString, item)
                i += 1
            Next
            root.AppendChild(SourceDirsXML)
        End If
        Try
            confdoc.Save(MySourceDir + ConfigFileName)
        Catch
            ConsoleMsgPrint("Ошибка при создании файла")
            Return 1
        End Try
        Return 0
    End Function

    Private Sub ConfigDocRead(PathToDoc As String)
        Dim document As XmlReader = New XmlTextReader(PathToDoc)
        confdoc.Load(document)
        document.Close() : document.Dispose()
        Dim node0 As XmlNode = confdoc.SelectSingleNode("/body/SourceStorage")
        MySourceDir = node0.Attributes("SourceDirectory").Value
        Dim node1 As XmlNode = confdoc.SelectSingleNode("/body/TmpFolder")
        TmpFolderName = node1.Attributes("TmpFolderName").Value
        Dim node2 As XmlNode = confdoc.SelectSingleNode("/body/DistanationStorage")
        For Each Attr As XmlAttribute In node2.Attributes
            If DistDirectories Is Nothing Then
                ReDim DistDirectories(0)
            Else
                ReDim Preserve DistDirectories(DistDirectories.Length)
            End If
            DistDirectories(DistDirectories.Length - 1) = Attr.Value
        Next
        Dim node3 As XmlNode = confdoc.SelectSingleNode("/body/SourceExt")
        If node3 IsNot Nothing Then
            For Each Attr As XmlAttribute In node3.Attributes
                If SourceExtFiles Is Nothing Then
                    ReDim SourceExtFiles(0)
                Else
                    ReDim Preserve SourceExtFiles(SourceExtFiles.Length)
                End If
                SourceExtFiles(SourceExtFiles.Length - 1) = Attr.Value
            Next
        End If
        Dim node4 As XmlNode = confdoc.SelectSingleNode("/body/SourceDir")
        If node4 IsNot Nothing Then
            For Each Attr As XmlAttribute In node4.Attributes
                If SourceDirs Is Nothing Then
                    ReDim SourceDirs(0)
                Else
                    ReDim Preserve SourceDirs(SourceDirs.Length)
                End If
                SourceDirs(SourceDirs.Length - 1) = Attr.Value
            Next
        End If
    End Sub
    Private Function FileCompare(file1 As String, file2 As String, contentcompare As Boolean) As Boolean
        Dim file1byte As Integer
        Dim file2byte As Integer
        Dim fs1 As FileStream
        Dim fs2 As FileStream

        '// Determine if the same file was referenced two times.
        If file1 = file2 Then Return True
        If File.GetCreationTime(file1) <> File.GetCreationTime(file2) Then Return False
        If File.GetLastWriteTime(file1) <> File.GetLastWriteTime(file2) Then Return False
        Dim fi1 As FileInfo = New FileInfo(file1)
        Dim fi2 As FileInfo = New FileInfo(file2)
        If fi1.Length <> fi2.Length Then Return False
        If contentcompare = False Then Return True

        '// Open the two files.
        fs1 = New FileStream(file1, FileMode.Open, FileAccess.Read)
        fs2 = New FileStream(file2, FileMode.Open, FileAccess.Read)

        '// Check the file sizes. If they are Not the same, the files 
        '// are Not the same.
        If fs1.Length <> fs2.Length Then
            '// Close the file
            fs1.Close()
            fs2.Close()
            '// Return false to indicate files are different
            Return False
        End If

        '// Read And compare a byte from each file until either a
        '// non-matching set of bytes Is found Or until the end of
        '// file1 Is reached.
        file1byte = fs1.ReadByte()
        file2byte = fs2.ReadByte()
        Do While (file1byte = file2byte) AndAlso (file1byte <> -1)
            '// Read one byte from each file.
            file1byte = fs1.ReadByte()
            file2byte = fs2.ReadByte()
        Loop

        '// Close the files.
        fs1.Close()
        fs2.Close()
        '// Return the success of the comparison. "file1byte" Is 
        '// equal to "file2byte" at this point only if the files are 
        '// the same.
        Return (file1byte - file2byte) = 0
    End Function
    Public Function IsUserAdministrator() As Boolean
        Dim isAdmin As Boolean
        Try
            Dim user As WindowsIdentity = WindowsIdentity.GetCurrent()
            Dim principal As WindowsPrincipal = New WindowsPrincipal(user)
            isAdmin = principal.IsInRole(WindowsBuiltInRole.Administrator)
        Catch ex As UnauthorizedAccessException
            isAdmin = False
        Catch ex As Exception
            isAdmin = False
        End Try
        Return isAdmin
    End Function

    Private Sub DeleteFromTmpFolder(FolderPath As String)
        Dim filePaths As String() = Directory.GetFiles(FolderPath)
        For Each filePath In filePaths
            Try
                File.Delete(filePath)
            Catch
            End Try
        Next
    End Sub
End Module
