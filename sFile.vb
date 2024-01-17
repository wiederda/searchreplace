Imports System.Security.Cryptography
Imports System.IO
Imports System.Security
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports System.Security.AccessControl
Imports System.Reflection
Imports System.Security.Principal
Imports System.Drawing

'<Assembly: AssemblyDelaySign(False)>
'<Assembly: AssemblyKeyFile("sFile.snk")>

Public Class s_File

    Public Enum ImageFormat
        Png = 0
        Jpg = 1
        Bmp = 2
        Gif = 3
    End Enum

    Public Enum ShortCut
        StartMenu = 0
        Desktop = 1
        TaskBar = 2
    End Enum

    Public Enum FileInfos
        Dateiversion = 0
        Productversion = 1
        FileDescription = 2
        FileBuildPart = 3
    End Enum

    Public Enum NTFSInherit
        SubFoldersAndFiles
        ThisFolderSubFoldersAndFiles
        ThisFolderAndSubFolders
        SubFoldersOnly
        ThisFolderAndFiles
        FilesOnly
        ThisFolderOnly
    End Enum

    ''' <summary>
    ''' Enumeration der Datei-Zeiten
    ''' </summary>
    Public Enum EnumFileAttributes As Integer
        LastAccessTime
        LastWriteTime
        CreationTime
    End Enum

    Public Enum StartFolder
        MyComputer = Environment.SpecialFolder.MyComputer
        MyDocuments = Environment.SpecialFolder.MyDocuments
        DesktopDirectory = Environment.SpecialFolder.DesktopDirectory
        ProgramFiles = Environment.SpecialFolder.CommonProgramFiles
        ProgramFilesX86 = Environment.SpecialFolder.CommonProgramFilesX86
    End Enum

    Public Enum Einheit
        KB = 0
        MB = 1
        GB = 2
    End Enum

    Public Enum FileTime
        LastWriteTime = 0
        LastWriteDate = 1
        LastWriteDateTime = 2
        LastAccessTime = 3
        LastAccessDate = 4
        LastAccessDateTime = 5
        CreationTime = 6
        CreationDate = 7
        CreationDateTime = 8
    End Enum

    Public Enum FileCompareStatus
        Identical = 0         ' identisch
        SizeDifferent = 1     ' Dateigröße verschieden
        ContentDifferent = 2  ' Inhalt verschieden
        [Error] = 3           ' Fehler
    End Enum

    Public Enum Special_Folder
        Startup = Environment.SpecialFolder.Startup
        Desktop = Environment.SpecialFolder.Desktop
    End Enum

    Private Declare Function SendMessage Lib "user32.dll" _
  Alias "SendMessageW" (
  ByVal hWnd As IntPtr,
  ByVal wMsg As Integer,
  ByVal wParam As Integer,
  ByVal lParam As Integer) As Integer

    Private Const EM_LINEFROMCHAR = &HC9
    Private Const EM_LINEINDEX = &HBB

    Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" _
  (ByVal lpExistingFileName As String, ByVal lpNewFileName As String,
   ByVal dwFlags As Long) As Boolean
    Private Const MOVEFILE_REPLACE_EXISTING = &H1 ' Überschreiben, falls vorhanden 
    Private Const MOVEFILE_COPY_ALLOWED = &H2 ' Kopieren auf anderes Laufwerk ermöglichen 
    Private Const MOVEFILE_DELAY_UNTIL_REBOOT = &H4 ' Erst bei Neustart ausführen 
    Private Const MOVEFILE_WRITE_THROUGH = &H8 ' Erst nach Erfolg zurückkehren

    ''' <summary>
    ''' Liefert Informationen zu einer Datei
    ''' </summary>
    ''' <param name="Pfad">Pfad zur Datei</param>
    ''' <param name="Wert">Wert, der zurückgegeben werden soll (LastWriteTime = 0, LastWriteDate = 1, LastWriteDateTime = 2, LastAccessTime = 3, LastAccessDate = 4, LastAccessDateTime = 5, CreationTime = 6, CreationDate = 7, CreationDateTime = 8)</param>
    ''' <param name="I">muss nicht geändert werden</param>
    ''' <returns></returns>
    Public Overloads Function FileInfo(ByVal Pfad As String, Optional ByVal Wert As FileTime = FileTime.LastAccessTime, Optional ByVal I As Integer = 0) As Date
        If System.IO.File.Exists(Pfad) = True Then
            Select Case Wert
                Case Wert.LastAccessDate
                    Return System.IO.File.GetLastAccessTime(Pfad).Date
                Case Wert.LastAccessTime
                    Return Format(System.IO.File.GetLastAccessTime(Pfad), "HH:mm")
                Case Wert.LastAccessDateTime
                    Return System.IO.File.GetLastAccessTime(Pfad)
                Case Wert.LastWriteDate
                    Return System.IO.File.GetLastWriteTime(Pfad).Date
                Case Wert.LastWriteTime
                    Return Format(System.IO.File.GetLastWriteTime(Pfad), "HH:mm")
                Case Wert.LastWriteDateTime
                    Return System.IO.File.GetLastWriteTime(Pfad)
                Case Wert.CreationDate
                    Return System.IO.File.GetCreationTime(Pfad).Date
                Case Wert.CreationTime
                    Return Format(System.IO.File.GetCreationTime(Pfad), "HH:mm")
                Case Wert.CreationDateTime
                    Return System.IO.File.GetCreationTime(Pfad)
            End Select
        End If
    End Function

    ''' <summary>
    ''' Liefert Informationen zu einer Datei
    ''' </summary>
    ''' <param name="Pfad">...</param>
    ''' <returns></returns>
    Public Overloads Function FileInfo(ByVal Pfad As String) As Double
        If System.IO.File.Exists(Pfad) = True Then
            Dim Info As New FileInfo(Pfad)
            Return Math.Round(Info.Length / 1024, 2)
        End If
    End Function

    ''' <summary>
    ''' Liefert Informationen zu einer Datei
    ''' </summary>
    ''' <param name="Pfad">...</param>
    ''' <param name="Wert">Wert, der zurückgegeben werden soll (Dateiversion = 0, Productversion = 1, FileDescription = 2, FileBuildPart = 3)</param>
    ''' <returns></returns>
    Public Overloads Function FileInfo(ByVal Pfad As String, Optional ByVal Wert As FileInfos = FileInfos.Productversion) As String
        If System.IO.File.Exists(Pfad) = True Then
            Dim Info As FileVersionInfo
            Info = FileVersionInfo.GetVersionInfo(Pfad)

            Select Case Wert
                Case Wert.Dateiversion
                    Return Info.FileVersion
                Case Wert.Productversion
                    Return Info.ProductVersion
                Case Wert.FileDescription
                    Return Info.FileDescription
                Case Wert.FileBuildPart
                    Return Info.FileBuildPart
            End Select
        End If
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="sFile1">...</param>
    ''' <param name="sFile2">....</param>
    ''' <returns></returns>
    Public Function FileReplace(ByVal sFile1 As String, ByVal sFile2 As String) As Boolean
        System.IO.File.Copy(sFile1, sFile2 & ".new", True)
        Return MoveFileEx(sFile2 & ".new", sFile2, MOVEFILE_REPLACE_EXISTING Or MOVEFILE_DELAY_UNTIL_REBOOT)
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="Message">Eintrag, der geschrieben werden soll</param>
    ''' <param name="LogFileName">Pfad zur Logdatei</param>
    Public Sub LogMessage(ByVal Message As String, ByVal LogFileName As String)

        Dim F As IO.FileInfo
        Dim X() As String, sFile As String

        X = LogFileName.Split("\")

        If X.Length >= 2 Then
            sFile = LogFileName
        Else
            sFile = Application.StartupPath & "\" & LogFileName
        End If

        ' F = New IO.FileInfo(sFile)
        Dim oStream As IO.StreamWriter = System.IO.File.AppendText(sFile)
        oStream.WriteLine(Message)

        oStream.Flush()
        oStream.Close()
        ' oStream = Nothing

        ' F = Nothing
    End Sub

    ''' <summary>
    ''' Vergleicht zwei Dateien
    ''' </summary>
    ''' <param name="sFile1">Pfad zur Datei</param>
    ''' <param name="sFile2">Pfad zur Datei</param>
    ''' <returns>Identical = 0 identisch, SizeDifferent = 1 Dateigröße verschieden, ContentDifferent = 2 Inhalt verschieden, [Error] = 3 Fehler</returns>
    Public Function FileCompare(ByVal sFile1 As String, ByVal sFile2 As String) As FileCompareStatus

        ' Prüfen, ob beide Dateien auch existieren
        If File.Exists(sFile1) AndAlso File.Exists(sFile2) Then
            Try
                ' zunächst Dateigröße vergleichen
                If New FileInfo(sFile1).Length <> New FileInfo(sFile2).Length Then
                    Return FileCompareStatus.SizeDifferent
                Else
                    ' jetzt Hash-Wert berechnen und vergleichen
                    Dim oHash As HashAlgorithm = HashAlgorithm.Create
                    Dim oStream As FileStream

                    ' Hash-Wert für 1. Datei ermitteln
                    oStream = New FileStream(sFile1, FileMode.Open)
                    Dim bHash1() As Byte = oHash.ComputeHash(oStream)
                    oStream.Close()

                    ' Hash-Wert für 2. Datei ermitteln
                    oStream = New FileStream(sFile2, FileMode.Open)
                    Dim bHash2() As Byte = oHash.ComputeHash(oStream)
                    oStream.Close()

                    ' jetzt die beiden Hash Byte-Arrays miteinander vergleichen
                    If BitConverter.ToString(bHash1) = BitConverter.ToString(bHash2) Then
                        ' Datei-Inhalt ist identisch
                        Return FileCompareStatus.Identical
                    Else
                        ' Datei-Inhalt unterscheidet sich
                        Return FileCompareStatus.ContentDifferent
                    End If

                End If
            Catch ex As Exception
                ' Fehler
                Return FileCompareStatus.Error
            End Try
        Else
            Return FileCompareStatus.Error
        End If
    End Function

    '''' <summary>
    '''' ....
    '''' </summary>
    '''' <param name="Pfad">Pfad zur Datei</param>
    'Public Function KillFile(ByVal Pfad As String) As Boolean
    '    Try
    '        ' Existiert die Datei?
    '        If IO.File.Exists(Pfad) Then
    '            ' Größe der Datei in Bytes ermitteln
    '            With New IO.FileInfo(Pfad)
    '                Dim fileLen As Long = .Length
    '                ' jetzt Datei zum Schreiben öffnen und blockweise 
    '                ' den Inhalt überschreiben
    '                Dim wStream As IO.FileStream = .OpenWrite()
    '                Dim curPos As Long = 0
    '                Dim emptyBytes(2047) As Byte

    '                While curPos < fileLen
    '                    Dim wLen As Long = fileLen - curPos
    '                    If wLen > emptyBytes.Length Then wLen = emptyBytes.Length
    '                    wStream.Write(emptyBytes, 0, wLen)
    '                    curPos += wLen
    '                End While

    '                ' Datei schließen
    '                wStream.Flush()
    '                wStream.Close()
    '            End With

    '            ' jetzt löschen
    '            IO.File.Delete(Pfad)
    '            Return True
    '        End If
    '    Catch ex As Exception
    '        Return False
    '    End Try
    'End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="sDir">...</param>
    ''' <param name="List">....</param>
    ''' <param name="Pattern">....</param>
    ''' <param name="SearchOption">....</param>
    Public Overloads Function GetFiles(ByVal sDir As String, ByVal List As ListBox, Optional ByVal Pattern As String = "*.*", Optional SearchOption As FileIO.SearchOption = FileIO.SearchOption.SearchAllSubDirectories) As String
        Dim sFile As String

        If Not sDir.EndsWith("\") Then sDir += "\"

        ' alle TXT-Dateien im Startverzeichnis einschl. Unterordner 
        ' in einer ListBox anzeigen
        Try
            For Each sFile In My.Computer.FileSystem.GetFiles(
                          sDir, SearchOption, Pattern)

                ' Dateiname mit relativer Pfadangabe zum Startverzeichnis ausgeben

                List.Items.Add(sFile)
            Next
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="directory">...</param>
    ''' <param name="Pattern">....</param>
    ''' <param name="SearchOption">....</param>
    Public Overloads Function GetFiles(ByVal directory As String, Optional SearchOption As SearchOption = SearchOption.AllDirectories, Optional Pattern As String = "*.*") As List(Of String)
        Return IO.Directory.GetFiles(directory, Pattern, SearchOption).ToList()
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="directory">...</param>
    ''' <param name="Pattern">....</param>
    ''' <param name="SearchOption">....</param>
    Public Overloads Function GetFilesCount(ByVal directory As String, Optional SearchOption As SearchOption = SearchOption.AllDirectories, Optional Pattern As String = "*.*") As Int32
        Return IO.Directory.GetFiles(directory, Pattern, SearchOption).Length
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="sDir">...</param>
    ''' <param name="List">....</param>
    ''' <param name="Pattern">....</param>
    ''' <param name="SearchOption">....</param>
    Public Overloads Function GetFiles(ByVal sDir As String, ByVal List As List(Of String), Optional ByVal Pattern As String = "*.*", Optional SearchOption As FileIO.SearchOption = FileIO.SearchOption.SearchAllSubDirectories) As String
        Dim sFile As String

        If Not sDir.EndsWith("\") Then sDir += "\"

        ' alle TXT-Dateien im Startverzeichnis einschl. Unterordner 
        ' in einer ListBox anzeigen
        Try
            For Each sFile In My.Computer.FileSystem.GetFiles(
                          sDir, SearchOption, Pattern)

                ' Dateiname mit relativer Pfadangabe zum Startverzeichnis ausgeben

                List.Add(sFile)
            Next
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="directory">...</param>
    ''' <param name="Pattern">....</param>
    ''' <param name="SearchOption">....</param>
    Public Overloads Function GetFiles(ByVal directory As String, List As ListBox, Optional SearchOption As SearchOption = SearchOption.AllDirectories, Optional Pattern As String = "*.*")
        Dim Dir As New List(Of String)
        Dir = IO.Directory.GetFiles(directory, Pattern, SearchOption).ToList()
        For i = 0 To Dir.Count - 1
            List.Items.Add(Dir(i))
        Next
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="sDir">...</param>
    ''' <param name="List">....</param>
    ''' <param name="SearchOption">....</param>
    Public Function GetFolder(ByVal sDir As String, ByVal List As ListBox, Optional SearchOption As FileIO.SearchOption = FileIO.SearchOption.SearchAllSubDirectories) As String
        Dim sFolder As String

        If Not sDir.EndsWith("\") Then sDir += "\"

        Try
            For Each sFolder In My.Computer.FileSystem.GetDirectories(sDir, SearchOption)
                List.Items.Add(sFolder)
            Next
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="sDir">...</param>
    ''' <param name="List">....</param>
    ''' <param name="SearchOption">....</param>
    Public Function GetFolder(ByVal sDir As String, ByVal List As List(Of String), Optional SearchOption As FileIO.SearchOption = FileIO.SearchOption.SearchAllSubDirectories) As String
        Dim sFolder As String

        If Not sDir.EndsWith("\") Then sDir += "\"

        Try
            For Each sFolder In My.Computer.FileSystem.GetDirectories(sDir, SearchOption)
                List.Add(sFolder)
            Next
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="sDir">...</param>
    Public Function GetFolderSize(ByVal sDir As String) As DataTable
        Dim table As DataTable = New DataTable
        table.Columns.Add("Ordner", GetType(String))
        table.Columns.Add("Größe_KB", GetType(String))
        table.Columns.Add("Größe_MB", GetType(String))

        Dim sFolder As String
        If Not sDir.EndsWith("\") Then sDir += "\"
        Try
            For Each sFolder In My.Computer.FileSystem.GetDirectories(sDir, FileIO.SearchOption.SearchTopLevelOnly)
                table.Rows.Add(sFolder, Math.Round(FolderSize_Help(sFolder) / 1024, 2), Math.Round(FolderSize_Help(sFolder) / 1024 / 1024, 2))
            Next
        Catch ex As Exception

        End Try
        Return table
    End Function

    Private Function FolderSize_Help(ByVal DirPath As String, Optional IncludeSubFolders As Boolean = True) As Long

        Dim lngDirSize As Long
        Dim objFileInfo As FileInfo
        Dim objDir As DirectoryInfo = New DirectoryInfo(DirPath)
        Dim objSubFolder As DirectoryInfo

        Try

            'add length of each file
            For Each objFileInfo In objDir.GetFiles()
                lngDirSize += objFileInfo.Length
            Next

            'call recursively to get sub folders
            'if you don't want this set optional
            'parameter to false 
            If IncludeSubFolders Then
                For Each objSubFolder In objDir.GetDirectories()
                    lngDirSize += FolderSize_Help(objSubFolder.FullName)
                Next
            End If

        Catch Ex As Exception

        End Try
        Return lngDirSize
    End Function

    ''' <summary>
    ''' Startet eine bestimmte Anwendung mit dem angegebenen Dokument
    ''' </summary>
    ''' <param name="ProgramFile">Dateiname der Anwendung</param>
    ''' <param name="DocumentFile">Dokument-Dateiname</param>
    ''' <returns>True, wenn die Anwendung gestartet werden konnte, andernfalls False.</returns>
    Public Overloads Function OpenDocument(ByVal ProgramFile As String, ByVal DocumentFile As String) As Boolean

        Try
            Dim pInfo As New Diagnostics.ProcessStartInfo
            With pInfo
                ' Anwendung, die gestartet werden soll
                .FileName = ProgramFile

                ' Parameter (Dokument)
                .Arguments = Chr(34) & DocumentFile & Chr(34)

                ' Anwendung starten
                .Verb = "open"
            End With
            Process.Start(pInfo)
            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Öffnet das Dokument mit der im System festgelegten Standard-Anwendung
    ''' </summary>
    ''' <param name="DocumentFile">Dokument-Dateiname</param>
    ''' <returns>True, wenn das Dokument geöffnet werden konnte, andernfalls False.</returns>
    Public Overloads Function OpenDocument(ByVal DocumentFile As String) As Boolean
        Try
            Dim pInfo As New Diagnostics.ProcessStartInfo
            With pInfo
                ' Dokument
                .FileName = DocumentFile

                ' verknüpfte Anwendung starten
                .Verb = "open"
            End With
            Process.Start(pInfo)
            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="Pfad">...</param>
    ''' <param name="arr">....</param>
    ''' <returns></returns>
    Public Function ArraySave(ByVal Pfad As String, ByVal arr As Object) As Boolean

        Dim fs As FileStream = Nothing
        Dim Success As Boolean = False

        Try
            ' Datei zum Schreiben öffnen
            fs = New FileStream(Pfad, FileMode.Create, FileAccess.Write)

            ' Array serialisieren und speichern
            Dim formatter As New BinaryFormatter()
            formatter.Serialize(fs, arr)
            Success = True

        Catch ex As Exception
        Finally
            ' Datei schließen
            If Not IsNothing(fs) Then fs.Close()
        End Try

        Return (Success)
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="Pfad">...</param>
    ''' <param name="arr">....</param>
    ''' <returns></returns>
    Public Function ArrayRead(ByVal Pfad As String, ByRef arr As Object) As Boolean

        Dim Success As Boolean = False

        ' Prüfen, ob Datei existiert
        If IO.File.Exists(Pfad) Then
            Dim fs As FileStream = Nothing
            Try
                ' Datei zum Lesen öffnen
                fs = New FileStream(Pfad, FileMode.Open, FileAccess.Read)

                ' Daten deserialiseren und dem Array zuweisen
                Dim formatter As New BinaryFormatter()
                arr = formatter.Deserialize(fs)
                Success = True

            Catch ex As Exception
            Finally
                ' Datei schließen
                If Not IsNothing(fs) Then fs.Close()
            End Try
        End If

        Return (Success)
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="Box">...</param>
    ''' <param name="oldText">....</param>
    ''' <param name="newText">....</param>
    Public Overloads Sub FindReplaceString(Box As TextBox, ByVal oldText As String, newText As String)
        Box.Text.Replace(FindString(Box, oldText), newText)
    End Sub

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="Pfad">...</param>
    ''' <param name="oldText">....</param>
    ''' <param name="newText">....</param>
    ''' <param name="append">....</param>
    Public Overloads Sub FindReplaceString(ByVal Pfad As String, ByVal oldText As String, newText As String, Optional append As Boolean = False)
        If append = False Then
            My.Computer.FileSystem.WriteAllText(Pfad, My.Computer.FileSystem.ReadAllText(Pfad).Replace(oldText, newText), False)
        ElseIf append = True Then
            My.Computer.FileSystem.WriteAllText(Pfad, My.Computer.FileSystem.ReadAllText(Pfad).Replace(oldText, newText), True)
        End If
    End Sub

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="Pfad">...</param>
    ''' <param name="box">....</param>
    ''' <param name="StrSuche">....</param>
    Public Overloads Function FindString(ByVal Pfad As String, ByVal box As ListBox, ByVal StrSuche As String, ByVal Splitter As Char)
        Dim SucheString As Array = StrSuche.Split(Splitter)

        For Each zeile As String In IO.File.ReadAllLines(Pfad)
            For Each suche As String In SucheString
                If zeile.Contains(suche) Then
                    box.Items.Add(zeile)
                End If
            Next
        Next
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="Pfad">...</param>
    ''' <param name="StrSuche">....</param>
    Public Overloads Function FindString(ByVal Pfad As String, ByVal StrSuche As String) As String
        For Each zeile As String In IO.File.ReadAllLines(Pfad)
            For Each suche As String In StrSuche
                If zeile.Contains(suche) Then
                    Return zeile
                End If
            Next
        Next
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="Box">...</param>
    ''' <param name="StrSuche">....</param>
    Public Overloads Function FindString(Box As TextBox, ByVal StrSuche As String) As String
        For Each zeile As String In IO.File.ReadAllLines(Box.Text)
            If zeile.Contains(StrSuche) Then
                Return zeile
            End If
        Next
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="Pfad">...</param>
    ''' <param name="StrSuche">....</param>
    ''' <param name="str">....</param>
    Public Overloads Function FindString(ByVal Pfad As String, ByVal StrSuche As String, Optional str As String = "") As String
        For Each zeile As String In IO.File.ReadAllLines(Pfad)
            If zeile.Contains(StrSuche) Then
                Return zeile
            End If
        Next
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="Pfad">...</param>
    ''' <param name="StrSuche">....</param>
    ''' <param name="List">....</param>
    ''' <param name="Encoding">....</param>
    ''' <param name="Split">....</param>
    ''' <param name="replaceold">....</param>
    ''' <param name="replacenew">....</param>
    Public Overloads Function FindString(ByVal Pfad As String, ByVal StrSuche As String, List As ListBox, Encoding As System.Text.Encoding, Optional Split As Char = "", Optional replaceold As String = "", Optional replacenew As String = "") As Int32
        For Each zeile As String In IO.File.ReadAllLines(Pfad, Encoding)
            If zeile.Contains(StrSuche) Then
                Dim sString() As String = zeile.Split(Split)
                List.Items.Add(sString.Last.Replace(replaceold, replacenew))
            End If
        Next
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="Pfad">...</param>
    ''' <param name="StartString">....</param>
    ''' <param name="EndString">...</param>
    ''' <param name="box">....</param>
    Public Overloads Function FindString(ByVal Pfad As String, ByVal StartString As String, ByVal EndString As String, ByVal box As ListBox)
        Dim lines = System.IO.File.ReadAllLines(Pfad), Nummer As String = ""
        Dim SucheString(1) As String
        SucheString(0) = StartString
        SucheString(1) = EndString

        For Each sFind In SucheString
            For i = 0 To lines.Length - 1
                If lines(i).StartsWith(sFind) Then
                    Nummer = Nummer & i & ","
                End If
            Next
        Next

        Help_String(Pfad, Nummer, box)
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="Pfad">...</param>
    ''' <param name="StrSuche">....</param>
    ''' <param name="I">....</param>
    Public Overloads Function FindString(ByVal Pfad As String, ByVal StrSuche As String, Optional ByVal I As Int32 = 0) As List(Of String)
        Dim s() As String = IO.File.ReadAllLines(Pfad), inputs() As String = {"", 0}
        For I = 0 To s.Length - 1
            If s(I).Contains(StrSuche) Then
                inputs = {s(I).ToString, I + 1}
            End If
        Next

        Dim lstWriteBits As List(Of String) = New List(Of String)(inputs)
        Return lstWriteBits
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="List">...</param>
    ''' <param name="StrSuche">....</param>
    Public Overloads Function FindString(ByVal List As List(Of String), ByVal StrSuche As String) As String

        For i = 0 To List.Count - 1
            If List.Item(i).Contains(StrSuche) Then
                Return List.Item(i)
            End If
        Next
    End Function

    Private Function Help_String(ByVal Pfad As String, ByVal Nummer As String, box As ListBox)
        Dim I() As String = Nummer.Split(",")

        For a = Int(I(0)) To Int(I(1))
            Dim Counter As String = System.IO.File.ReadAllLines(Pfad)(a)
            box.Items.Add(Counter)
        Next
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="sText">...</param>
    ''' <param name="nLen">....</param>
    ''' <returns></returns>
    Public Function Left(ByVal sText As String, ByVal nLen As Integer) As String
        If nLen > sText.Length Then nLen = sText.Length
        Return (sText.Substring(0, nLen))
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="sText">...</param>
    ''' <param name="nLen">....</param>
    ''' <returns></returns>
    Public Function Right(ByVal sText As String, ByVal nLen As Integer) As String
        If nLen > sText.Length Then nLen = sText.Length
        Return (sText.Substring(sText.Length - nLen))
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="cPfad1">...</param>
    ''' <param name="cPfad2">....</param>
    ''' <param name="overwrite">....</param>
    ''' <param name="I">....</param>
    ''' <returns></returns>
    Public Overloads Function SetFile(ByVal cPfad1 As String, ByVal cPfad2 As String, ByVal overwrite As Boolean, ByVal I As Integer) As Boolean
        If I = 1 Then
            Dim Pfad() As String = cPfad1.Split("\")
            If Not cPfad2.EndsWith("\") Then
                cPfad2 = cPfad2 & "\" & Pfad(Pfad.Length - 1)
            Else
                cPfad2 = cPfad2 & Pfad(Pfad.Length - 1)
            End If

        End If

        Try
            My.Computer.FileSystem.CopyFile(cPfad1, cPfad2, overwrite)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="dPfad">...</param>
    ''' <returns></returns>
    Public Overloads Function SetFile(ByVal dPfad As String) As String
        Try
            If My.Computer.FileSystem.FileExists(dPfad) Then
                My.Computer.FileSystem.DeleteFile(dPfad)
                Return "Datei wurde gelöscht"
            Else
                Return "Datei existiert nicht"
            End If
        Catch ex As Exception
            Return "Datei konnte nicht gelöscht werden"
        End Try
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="mPfad1">...</param>
    ''' <param name="mPfad2">....</param>
    ''' <param name="overwrite">....</param>
    ''' <returns></returns>
    Public Overloads Function SetFile(ByVal mPfad1 As String, ByVal mPfad2 As String, ByVal overwrite As Boolean) As Boolean
        Try
            My.Computer.FileSystem.MoveFile(mPfad1, mPfad2, overwrite)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="rPfad1">...</param>
    ''' <param name="rPfad2">....</param>
    ''' <returns></returns>
    Public Overloads Function SetFile(ByVal rPfad1 As String, ByVal rPfad2 As String) As Boolean
        Try
            My.Computer.FileSystem.RenameFile(rPfad1, rPfad2)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Kopiert einen Ordner
    ''' </summary>
    ''' <param name="cPfad1">Quellordner</param>
    ''' <param name="cPfad2">Zielordner</param>
    ''' <param name="overwrite">Gibt an, ob der Zielorner überschrieben werden soll</param>
    ''' <param name="löschen">Gibt an, ob der Zielorner gelöscht werden soll</param>
    ''' <param name="QuelltoZiel">Gibt an, ob Quell- und Zielordner den gleichen Namen haben (c:\Test\Backup) Backup wird übernommen</param>
    ''' <returns></returns>
    Public Overloads Function SetFolder(ByVal cPfad1 As String, ByVal cPfad2 As String, ByVal QuelltoZiel As Boolean, Optional ByVal overwrite As Boolean = True, Optional ByVal löschen As Boolean = True) As Boolean
        Try
            If QuelltoZiel = True Then
                Dim Folder() As String = cPfad1.Split("\")
                cPfad2 = cPfad2 & "\" & Folder(Folder.ToString.Length)
            End If

            If löschen = True And IO.Directory.Exists(cPfad2) Then
                IO.Directory.Delete(cPfad2, True)
            End If
            My.Computer.FileSystem.CopyDirectory(cPfad1, cPfad2, overwrite)

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' String zu SecureString
    ''' </summary>
    ''' <param name="source">String der umgewandert werden soll</param>
    ''' <returns></returns>
    Public Function ToSecureString(ByVal source As String) As SecureString
        If String.IsNullOrWhiteSpace(source) Then
            Return Nothing
        End If
        Dim result = New SecureString()
        For Each c As Char In source
            result.AppendChar(c)
        Next
        Return result
    End Function

    ''' <summary>
    ''' SecureString zu SecureString
    ''' </summary>
    ''' <param name="source">String der umgewandert werden soll</param>
    ''' <returns></returns>
    Public Function ToUnsecureString(ByVal source As SecureString) As String
        Dim returnValue = IntPtr.Zero
        Try
            returnValue = Marshal.SecureStringToGlobalAllocUnicode(source)
            Return Marshal.PtrToStringUni(returnValue)
        Finally
            Marshal.ZeroFreeGlobalAllocUnicode(returnValue)
        End Try
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="File">...</param>
    Public Overloads Sub CopyToClipboard(ByVal File As String)
        Dim oDataObject As New DataObject, tempFileArray(0) As String
        tempFileArray(0) = File
        oDataObject.SetData(DataFormats.FileDrop, True, tempFileArray)
        Clipboard.SetDataObject(oDataObject, True)
    End Sub

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="Text">...</param>
    ''' <param name="I">....</param>
    Public Overloads Sub CopyToClipboard(ByVal Text As String, ByVal I As Integer)
        Clipboard.SetText(Text)
    End Sub

    ''' <summary>
    ''' ........
    ''' </summary>
    ''' <param name="FileName">....</param>
    ''' <param name="ShortCutName">....</param>
    ''' <param name="ShortCutPath">....</param>
    ''' <param name="ShortCut">....</param>
    ''' <returns></returns>
    Public Function CreateShortCut(ByVal FileName As String, ByVal ShortCutName As String, Optional ShortCutPath As String = "", Optional ShortCut As ShortCut = Nothing) As Boolean

        Dim oShell As Object
        Dim oLink As Object

        If ShortCutPath = "" Then
            If ShortCut = 0 Then
                ShortCutPath = "C:\Users\" & Environment.UserName & "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"
            ElseIf ShortCut = 1 Then
                ShortCutPath = "C:\Users\" & Environment.UserName & "\Desktop"
            ElseIf ShortCut = 2 Then
                ShortCutPath = "C:\Users\" & Environment.UserName & "\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
            End If
        End If

        Try
            oShell = CreateObject("WScript.Shell")
            If IO.File.Exists(ShortCutPath & "\" & ShortCutName & ".lnk") = False Then
                oLink = oShell.CreateShortcut(ShortCutPath & "\" & ShortCutName & ".lnk")
                oLink.TargetPath = FileName
                oLink.WindowStyle = 1
                oLink.Save()
            End If
        Catch ex As Exception

        End Try
    End Function

    ''' <summary>
    ''' ........
    ''' </summary>
    ''' <param name="path">....</param>
    ''' <remarks></remarks>
    Public Shared Sub deleteEmptyLines(ByVal path As String)
        Dim tempFile As String = System.IO.Path.GetTempFileName()

        Using sr As StreamReader = New StreamReader(path)

            Using sw As StreamWriter = New StreamWriter(tempFile)

                While Not sr.EndOfStream
                    Dim line As String = sr.ReadLine()

                    If Not String.IsNullOrEmpty(line) AndAlso Not String.IsNullOrEmpty(line.Trim()) Then
                        sw.WriteLine(line)
                    End If
                End While
            End Using
        End Using

        System.IO.File.Copy(tempFile, path, True)
        System.IO.File.Delete(tempFile)
    End Sub

    ''' <summary>
    ''' Löscht die angegebene Zeile aus der Datei
    ''' </summary>
    ''' <param name="filename">Dateiname</param>
    ''' <param name="line">Zeilennummer der Zeile, die gelöscht werden soll.</param>
    ''' <remarks></remarks>
    Public Sub DelLineFromFile(ByVal filename As String, ByVal line As Integer)
        Try
            ' Datei in String-Attayeinlesen
            Dim lines As String() = My.Computer.FileSystem.ReadAllText(
              filename, System.Text.Encoding.Default).Split(vbCr)

            ' angegebenes Element (Zeile) aus der Liste entfernen
            If line > 0 AndAlso line <= lines.Length Then
                Dim oStream As IO.StreamWriter = Nothing
                Try
                    ' Stream-Objekt zum Speichern erstellen
                    oStream = New IO.StreamWriter(filename, False, System.Text.Encoding.Default)

                    ' alle Elemente der Liste zeilenweise speichern
                    Dim bNext As Boolean = False
                    For i As Integer = 0 To lines.Length - 1
                        If i + 1 <> line Then
                            If bNext Then oStream.Write(vbCr)
                            oStream.Write(lines(i))
                            bNext = True
                        End If
                    Next

                Catch ex As Exception
                Finally
                    ' Stream-Objekt schließen
                    If Not IsNothing(oStream) Then oStream.Close()
                End Try

            End If
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Ruft einen Eintrag im Windows-EventLog ab
    ''' </summary>
    ''' <param name="List">Beinhaltet die gefunden Einträge. 0=TimeGenerated,1=InstanceId,2=Source</param>
    ''' <param name="LastEntry">optional Bestimmt ob der Letzte oder erste Eintrag gesucht wird</param>
    ''' <param name="LogName">optional Name des EventLogs</param>
    Private Sub GetEventLogEntry(List As List(Of String), Optional LastEntry As Boolean = False, Optional LogName As String = "Application")

        Dim objLogs() As EventLog
        Dim objEntry As EventLogEntry
        Dim objLog As EventLog
        objLogs = EventLog.GetEventLogs()

        For Each objLog In objLogs
            If objLog.LogDisplayName = LogName Then Exit For
        Next

        If LastEntry = False Then
            List.Add(objLog.Entries(0).TimeGenerated)
            List.Add(objLog.Entries(0).InstanceId)
            List.Add(objLog.Entries(0).Source)
        ElseIf LastEntry = True Then
            List.Add(objLog.Entries(objLog.Entries.Count - 1).TimeGenerated)
            List.Add(objLog.Entries(objLog.Entries.Count - 1).InstanceId)
            List.Add(objLog.Entries(objLog.Entries.Count - 1).Source)
        End If
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="List"></param>
    ''' <param name="EventID"></param>
    ''' <param name="LogName">optional Name des EventLogs</param>
    ''' <param name="DatumZeit">optional Datum und Uhrzeit</param>
    ''' <param name="Datum">optional alle Einträge von dem Datum</param>
    Public Sub GetEventLogEntry(List As ListBox, EventID As Integer, Optional LogName As String = "Application", Optional DatumZeit As DateTime = Nothing, Optional Datum As String = "")
        Dim myEventLog As EventLog = New EventLog(LogName)

        For Each entry As EventLogEntry In myEventLog.Entries

            If DatumZeit <> Nothing Then
                If entry.InstanceId = EventID AndAlso entry.TimeGenerated >= DatumZeit Then
                    List.Items.Add(entry.Source)
                    List.Items.Add(entry.Message)
                    List.Items.Add(entry.TimeGenerated)
                    List.Items.Add(entry.MachineName)
                End If
            End If

            If Datum <> "" Then
                Dim d() As String = entry.TimeGenerated.ToString.Split(" ")
                If entry.InstanceId = EventID AndAlso d(0) = Datum Then
                    List.Items.Add(entry.Source)
                    List.Items.Add(entry.Message)
                    List.Items.Add(entry.TimeGenerated)
                    List.Items.Add(entry.MachineName)
                    List.Items.Add("")
                End If
            End If

            If Datum = "" And DatumZeit = Nothing Then
                If entry.InstanceId = EventID Then
                    List.Items.Add(entry.Source)
                    List.Items.Add(entry.Message)
                    List.Items.Add(entry.TimeGenerated)
                    List.Items.Add(entry.MachineName)
                    List.Items.Add("")
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="FileName">...</param>
    ''' <param name="WriteLine">....</param>
    ''' <param name="Append">....</param>
    Public Overloads Sub WriteFile(FileName As String, WriteLine As String, Append As Boolean)
        Dim writer As New System.IO.StreamWriter(FileName, Append)
        writer.WriteLine(WriteLine)
        writer.Flush()
        writer.Close()
    End Sub

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="FileName">...</param>
    ''' <param name="Line">....</param>
    ''' <param name="Text">....</param>
    Public Overloads Sub WriteFile(FileName As String, Line As Int32, Text As String)
        Dim lines() As String = IO.File.ReadAllLines(FileName)
        lines(Line) = Text
        IO.File.WriteAllLines(FileName, lines)
    End Sub

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="sSrcPath">...</param>
    ''' <param name="sDestPath">....</param>
    ''' <param name="bSubFolder">....</param>
    ''' <param name="bOverWrite">....</param>
    Public Sub CopyFolder(ByVal sSrcPath As String, ByVal sDestPath As String, Optional ByVal bSubFolder As Boolean = True, Optional ByVal bOverWrite As Boolean = True)

        ' Falls Zielordner nicht existiert, jetzt erstellen
        If Not System.IO.Directory.Exists(sDestPath) Then
            System.IO.Directory.CreateDirectory(sDestPath)
        End If

        ' zunächst alle Dateien des Quell-Ordners ermitteln
        ' und kopieren
        Dim sFiles() As String = System.IO.Directory.GetFiles(sSrcPath)
        Dim sFile As String
        For i As Integer = 0 To sFiles.Length - 1
            ' Falls Datei im Zielordner bereits existiert, nur
            ' kopieren, wenn Parameter "bOverWrite" auf True
            ' festgelegt ist
            sFile = sFiles(i).Substring(sFiles(i).LastIndexOf("\") + 1)
            If bOverWrite Or Not System.IO.File.Exists(sDestPath & "\" & sFile) Then
                System.IO.File.Copy(sFiles(i), sDestPath & "\" & sFile, True)
            End If
        Next i

        If bSubFolder Then
            ' jetzt alle Unterordner ermitteln und die CopyFolder-Funktion
            ' rekursiv aufrufen
            Dim sDirs() As String = System.IO.Directory.GetDirectories(sSrcPath)
            Dim sDir As String
            For i As Integer = 0 To sDirs.Length - 1
                If sDirs(i) <> sDestPath Then
                    sDir = sDirs(i).Substring(sDirs(i).LastIndexOf("\") + 1)
                    CopyFolder(sDirs(i), sDestPath & "\" & sDir, True, bOverWrite)
                End If
            Next i
        End If
    End Sub

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="sPfad">...</param>
    ''' <param name="LastLine">....</param>
    ''' <param name="FirstLine">....</param>
    Public Function GetLineFromFile(sPfad As String, Optional LastLine As Boolean = True, Optional FirstLine As Boolean = False) As String
        Dim sLine() As String = IO.File.ReadAllLines(sPfad)

        If LastLine = True And FirstLine = False Then
            Return sLine(sLine.Length - 1)
        ElseIf LastLine = False And FirstLine = True Then
            Return sLine(0)
        Else
            Return "Fehler"
        End If
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="IniPfad">...</param>
    ''' <param name="Filter">....</param>
    Public Function OpenFileDialog(IniPfad As String, Filter As String) As String
        Dim FileDialog As New OpenFileDialog
        FileDialog.InitialDirectory = IniPfad
        FileDialog.Filter = Filter
        FileDialog.ShowDialog()

        If DialogResult.OK Then
            Return FileDialog.FileName
        Else
            Return "Fehler"
        End If
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="IniPfad">...</param>
    ''' <param name="Filter">....</param>
    Public Function SaveFileDialog(IniPfad As String, Filter As String) As String
        Dim FileDialog As New SaveFileDialog
        FileDialog.InitialDirectory = IniPfad
        FileDialog.Filter = Filter
        FileDialog.ShowDialog()

        If DialogResult.OK Then
            Return FileDialog.FileName
        Else
            Return "Fehler"
        End If
    End Function

    ''' <summary>
    ''' ...
    ''' </summary>
    ''' <param name="RootFolder"></param>
    ''' <param name="NewFolder"></param>
    Public Function FolderBrowserDialog(RootFolder As StartFolder, Optional NewFolder As Boolean = False) As String
        Dim Folder As New FolderBrowserDialog
        Folder.RootFolder = RootFolder
        If NewFolder = False Then
            Folder.ShowNewFolderButton = False
        Else
            Folder.ShowNewFolderButton = True
        End If
        Folder.ShowDialog()

        If DialogResult.OK Then
            Return Folder.SelectedPath
        Else
            Return "Fehler"
        End If
    End Function

    'https://www.vbarchiv.net/tipps/tipp_1885-eintrag-ins-kontextmen-des-windows-explorers-net.html
    ''' <summary>
    ''' Fügt dem Kontextmenü des Windows Explorers einen Eintrag für einen Dateityp hinzu.
    ''' Bei Erfolg wird True zurückgegeben, sonst False.
    ''' Beispiel der Kommentare: extension=.js ,text=In JSEdit öffnen, command= C:\jsedit.exe "%1"
    ''' </summary>
    ''' <param name="extension">Der Dateityp. Beispiel: .txt</param>
    ''' <param name="text">Der Text des Eintrags. Beispiel: In JSEdit öffnen</param>
    ''' <param name="command">Der aufzurufende Befehl. Beispiel: C:\jsedit.exe "%1"</param>
    Public Function AddToExplorerContextMenu(ByVal extension As String,
      ByVal text As String, ByVal command As String) As Boolean
        ' Beispiel der Kommentare: 
        '   extension=.js 
        '   text=In JSEdit öffnen 
        '   command= C:\jsedit.exe "%1"
        Try
            ' Öffnen: HKEY_CLASSES_ROOT\.js
            Dim Extensionkey As RegistryKey = Registry.ClassesRoot.CreateSubKey(extension)
            ' Öffnen: HKEY_CLASSES_ROOT\.js\Shell
            Dim Shellkey As RegistryKey = Extensionkey.CreateSubKey("Shell")
            ' Öffnen: HKEY_CLASSES_ROOT\.js\Shell\In JSEdit bearbeiten
            Dim Entrykey As RegistryKey = Shellkey.CreateSubKey(text)
            ' Öffnen: HKEY_CLASSES_ROOT\.js\Shell\In JSEdit bearbeiten\command
            Dim Commandkey As RegistryKey = Entrykey.CreateSubKey("command")
            Commandkey.SetValue("", command)
            Commandkey.Close()
            Entrykey.Close()
            Shellkey.Close()
            Extensionkey.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    'https://www.vbarchiv.net/tipps/tipp_1885-eintrag-ins-kontextmen-des-windows-explorers-net.html

    ''' <summary>
    ''' Entfernt einen Eintrag eines Dateityüs aus dem Kontextmenü des Windows Explorers.
    ''' </summary>
    ''' <param name="extension">Siehe AddToExplorerContextMenu()</param>
    ''' <param name="text">Siehe AddToExplorerContextMenu()</param>
    Public Function RemoveFromExplorerContextMenu(ByVal extension As String, ByVal text As String) As Boolean
        Try
            ' Öffnen: HKEY_CLASSES_ROOT\.js
            Dim Extensionkey As RegistryKey = Registry.ClassesRoot.OpenSubKey(extension, True)
            ' Öffnen: HKEY_CLASSES_ROOT\.js\Shell
            Dim Shellkey As RegistryKey = Extensionkey.OpenSubKey("Shell", True)
            ' Entfernen: HKEY_CLASSES_ROOT\.js\Shell\In JSEdit bearbeiten
            Shellkey.DeleteSubKeyTree(text)
            Shellkey.Close()
            Extensionkey.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' ...
    ''' </summary>
    ''' <param name="fullPath"></param>
    Public Function GrantAccess(ByVal fullPath As String) As Boolean
        Dim dInfo As DirectoryInfo = New DirectoryInfo(fullPath)
        Dim dSecurity As DirectorySecurity = dInfo.GetAccessControl()
        dSecurity.AddAccessRule(New FileSystemAccessRule(New SecurityIdentifier(WellKnownSidType.WorldSid, Nothing), FileSystemRights.FullControl, InheritanceFlags.ObjectInherit Or InheritanceFlags.ContainerInherit, PropagationFlags.NoPropagateInherit, AccessControlType.Allow))
        dInfo.SetAccessControl(dSecurity)
        Return True
    End Function

    ''' <summary>
    ''' ...
    ''' </summary>
    ''' <param name="sDirectory"></param>
    ''' <param name="Pattern"></param>
    ''' <param name="SearchOption"></param>
    Public Function GetDirectoriesCount(ByVal sDirectory As String, Optional Pattern As String = "*.*", Optional SearchOption As SearchOption = SearchOption.AllDirectories) As Int32
        Dim Dirs() As String = Directory.GetDirectories(sDirectory, Pattern, SearchOption)
        Return Dirs.Count
    End Function

    ''' <summary>
    ''' ...
    ''' </summary>
    Public Function GetGuid() As String
        Return System.Guid.NewGuid.ToString()
    End Function

    ''' <summary>
    ''' ...
    ''' </summary>
    ''' <param name="list1"></param>
    ''' <param name="list2"></param>
    Public Function GetContains(ByVal list1 As List(Of String), ByVal list2 As List(Of String)) As List(Of String)
        Dim result As List(Of String) = New List(Of String)()
        result.AddRange(list1.Except(list2, StringComparer.OrdinalIgnoreCase))
        result.AddRange(list2.Except(list1, StringComparer.OrdinalIgnoreCase))
        Return result
    End Function

    ''' <summary>
    ''' ...
    ''' </summary>
    ''' <param name="Dir"></param>
    ''' <param name="list"></param>
    Public Overloads Function GetSubDirectories(Dir As String, List As ListBox)
        Dim subdirectoryEntries As String() = Directory.GetDirectories(Dir)

        For Each subdirectory As String In subdirectoryEntries
            List.Items.Add(LoadSubDirs(subdirectory))
        Next
    End Function

    ''' <summary>
    ''' ...
    ''' </summary>
    ''' <param name="Dir"></param>
    Public Overloads Function GetSubDirectories(Dir As String) As List(Of String)
        Dim subdirectoryEntries As String() = Directory.GetDirectories(Dir)

        For Each subdirectory As String In subdirectoryEntries
            LoadSubDir(subdirectory)
        Next
    End Function

    Private Overloads Function LoadSubDirs(ByVal dir As String) As ListBox
        Dim subdirectoryEntries As String() = Directory.GetDirectories(dir)

        For Each subdirectory As String In subdirectoryEntries
            LoadSubDirs(subdirectory)
        Next
    End Function

    Private Overloads Function LoadSubDir(ByVal dir As String) As List(Of String)
        Dim subdirectoryEntries As String() = Directory.GetDirectories(dir)

        For Each subdirectory As String In subdirectoryEntries
            LoadSubDir(subdirectory)
        Next
    End Function

    ''' <summary>
    ''' ...
    ''' </summary>
    ''' <param name="Pfad"></param>
    ''' <param name="List"></param>
    Public Overloads Function ReadFile(ByVal Pfad As String, ByVal List As List(Of String))
        Dim fileStream As FileStream = New FileStream(Pfad, FileMode.OpenOrCreate, FileAccess.ReadWrite)
        Dim streamReader As StreamReader = New StreamReader(CType(fileStream, Stream))
        streamReader.BaseStream.Seek(0L, SeekOrigin.Begin)

        While streamReader.Peek() > -1
            List.Add(streamReader.ReadLine())
        End While

        streamReader.Close()
        fileStream.Close()
    End Function

    ''' <summary>
    ''' ...
    ''' </summary>
    ''' <param name="Pfad"></param>
    ''' <param name="List"></param>
    Public Overloads Function ReadFile(ByVal Pfad As String, ByVal List As ListBox)
        Dim fileStream As FileStream = New FileStream(Pfad, FileMode.OpenOrCreate, FileAccess.ReadWrite)
        Dim streamReader As StreamReader = New StreamReader(CType(fileStream, Stream))
        streamReader.BaseStream.Seek(0L, SeekOrigin.Begin)

        While streamReader.Peek() > -1
            List.Items.Add(streamReader.ReadLine())
        End While

        streamReader.Close()
        fileStream.Close()
    End Function

    ''' <summary>
    ''' ...
    ''' </summary>
    ''' <param name="Pfad"></param>
    Public Overloads Function ReadFile(ByVal Pfad As String) As String()
        Dim fileStream As FileStream = New FileStream(Pfad, FileMode.OpenOrCreate, FileAccess.ReadWrite)
        Dim streamReader As StreamReader = New StreamReader(CType(fileStream, Stream))
        streamReader.BaseStream.Seek(0L, SeekOrigin.Begin)

        While streamReader.Peek() > -1
            Return {streamReader.ReadLine()}
        End While

        streamReader.Close()
        fileStream.Close()
    End Function

    ''' <summary>
    ''' Zeilennummer der aktuellen Eingabeposition ermitteln
    ''' </summary>
    ''' <param name="oTextBox"></param>
    Public Function GetCurrentLine(ByVal oTextBox As RichTextBox) As Integer
        GetCurrentLine = SendMessage(oTextBox.Handle, EM_LINEFROMCHAR, -1, 0) + 1
    End Function

    ''' <summary>
    ''' Zeilennummer der aktuellen Eingabeposition ermitteln
    ''' </summary>
    ''' <param name="oTextBox"></param>
    Public Function GetCurrentLine(ByVal oTextBox As TextBox) As Integer
        GetCurrentLine = SendMessage(oTextBox.Handle, EM_LINEFROMCHAR, -1, 0) + 1
    End Function

    ''' <summary>
    ''' Spaltennummer der aktuellen Zeile ermitteln
    ''' </summary>
    ''' <param name="oTextBox"></param>
    Public Function GetCurrentCol(ByVal oTextBox As RichTextBox) As Integer
        With oTextBox
            GetCurrentCol = .SelectionStart - SendMessage(.Handle, EM_LINEINDEX,
      GetCurrentLine(oTextBox) - 1, 0) + 1
        End With
    End Function

    ''' <summary>
    ''' Spaltennummer der aktuellen Zeile ermitteln
    ''' </summary>
    ''' <param name="oTextBox"></param>
    Public Function GetCurrentCol(ByVal oTextBox As TextBox) As Integer
        With oTextBox
            GetCurrentCol = .SelectionStart - SendMessage(.Handle, EM_LINEINDEX,
      GetCurrentLine(oTextBox) - 1, 0) + 1
        End With
    End Function

    ''' <summary>
    '''  Prüft, ob die angegeben Datei aktuell durch eine andere Anwendung in Benutzung ist
    ''' </summary>
    ''' <param name="sFile"></param>
    Public Function FileInUse(ByVal sFile As String) As Boolean
        '
        Dim bInUse As Boolean = False

        If IO.File.Exists(sFile) Then
            Try
                ' Versuch, Datei EXKLUSIV zu öffnen
                Dim F As Short = FreeFile()
                FileOpen(F, sFile, OpenMode.Binary, OpenAccess.ReadWrite, OpenShare.LockReadWrite)
                FileClose(F)
            Catch
                ' Bei Fehler ist die Datei in Benutzung
                bInUse = True
            End Try
        End If
        ' Rückgabewert
        Return (bInUse)
    End Function

    ''' <summary>
    ''' Gibt den Besitzer einer Datei zurück
    ''' </summary>
    ''' <param name="Filename">Dateiname</param>
    Public Function FileOwner(ByVal Filename As String) As String
        Dim oFile As New IO.FileInfo(Filename)
        Dim oFS As Security.AccessControl.FileSecurity = oFile.GetAccessControl
        Return oFS.GetOwner(GetType(Security.Principal.NTAccount)).Value
    End Function

    ''' <summary>
    ''' Datei sicher löschen
    ''' </summary>
    ''' <param name="Filename">Dateiname mit vollständiger Pfadangabe</param>
    ''' <returns>True, wenn die Datei gelöscht werden konnte, andernfalls False.</returns>
    Public Function KillFile(ByVal Filename As String) As Boolean
        Try
            ' Existiert die Datei?
            If IO.File.Exists(Filename) Then
                ' Größe der Datei in Bytes ermitteln
                With New IO.FileInfo(Filename)
                    Dim fileLen As Long = .Length
                    ' jetzt Datei zum Schreiben öffnen und blockweise 
                    ' den Inhalt überschreiben
                    Dim wStream As IO.FileStream = .OpenWrite()
                    Dim curPos As Long = 0
                    Dim emptyBytes(2047) As Byte

                    While curPos < fileLen
                        Dim wLen As Long = fileLen - curPos
                        If wLen > emptyBytes.Length Then wLen = emptyBytes.Length
                        wStream.Write(emptyBytes, 0, wLen)
                        curPos += wLen
                    End While

                    ' Datei schließen
                    wStream.Flush()
                    wStream.Close()
                End With

                ' jetzt löschen
                IO.File.Delete(Filename)
                Return True
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    '''''' <summary>
    '''''' Setzt NTFS Berechtigung auf den Übergeben Orderpfad
    '''''' </summary>
    '''''' <param name="sFolderPath">Pfad zum Ordner (Bsp.: C:\Mein\Ordner)</param>
    '''''' <param name="sNameAccount">Der Benutzername der in die Berechtigung geschrieben werden soll (Bsp.: Administrator oder UPNName von einer Active Directory Domäme MaxMustert@Test.com)</param>
    '''''' <param name="niNTFSInherit">Ob und wie Vererbt werden soll (Enum Werte)</param>
    '''''' <param name="fsrPermissions">Die Berechtigungen die gesetzt werden sollen. Mit OR trennen für mehrere</param>
    '''''' <param name="actAccess">Allow oder Deny. Zulassen oder Verweigern</param>
    '''''' <remarks></remarks>
    Public Sub SetFolderNTFSPermissions(ByVal sFolderPath As String, ByVal sNameAccount As String, Optional ByVal niNTFSInherit As NTFSInherit = NTFSInherit.ThisFolderSubFoldersAndFiles, Optional ByVal fsrPermissions As FileSystemRights = FileSystemRights.FullControl, Optional ByVal actAccess As AccessControlType = AccessControlType.Allow)
        Try
            Dim dInfo As New DirectoryInfo(sFolderPath)
            Dim dSecurity As DirectorySecurity = dInfo.GetAccessControl()
            Dim sid As SecurityIdentifier = WindowsIdentity.GetCurrent().User
            Dim myrules As Object
            myrules = dSecurity.GetAccessRules(True, True, GetType(Security.Principal.NTAccount))
            Dim iFlag As New InheritanceFlags
            Dim iProg As New PropagationFlags

            If niNTFSInherit = NTFSInherit.SubFoldersAndFiles Then
                iFlag = InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit
                iProg = PropagationFlags.InheritOnly
            ElseIf niNTFSInherit = NTFSInherit.SubFoldersOnly Then
                iFlag = InheritanceFlags.ContainerInherit
                iProg = PropagationFlags.InheritOnly
            ElseIf niNTFSInherit = NTFSInherit.ThisFolderAndFiles Then
                iFlag = InheritanceFlags.ObjectInherit
                iProg = PropagationFlags.None
            ElseIf niNTFSInherit = NTFSInherit.ThisFolderAndSubFolders Then
                iFlag = InheritanceFlags.ContainerInherit
                iProg = PropagationFlags.None
            ElseIf niNTFSInherit = NTFSInherit.ThisFolderSubFoldersAndFiles Then
                iFlag = InheritanceFlags.ContainerInherit Or InheritanceFlags.ObjectInherit
                iProg = PropagationFlags.None
            ElseIf niNTFSInherit = NTFSInherit.FilesOnly Then
                iFlag = InheritanceFlags.ObjectInherit
                iProg = PropagationFlags.InheritOnly
            ElseIf niNTFSInherit = NTFSInherit.ThisFolderOnly Then
                iFlag = InheritanceFlags.None
                iProg = PropagationFlags.None
            End If
            Dim AccessRule As New FileSystemAccessRule(sNameAccount, fsrPermissions, iFlag, iProg, actAccess)
            dSecurity.ModifyAccessRule(AccessControlModification.Add, AccessRule, True)
            dInfo.SetAccessControl(dSecurity)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Function GetZertDate(text As TextBox, Suche As String) As String
        For i = 0 To Val(text.Lines.Count) - 1
            If text.Lines(i).ToString.StartsWith(Suche) Then
                Return Trim(text.Lines(i + 1))
                Exit For
            End If
        Next
    End Function

    Private Function GetProcessText(ByVal process As String, ByVal param As String, Optional workingDir As String = "") As String
        Dim p As Process = New Process, output As String = "", out As String, err As String
        ' this is the name of the process we want to execute
        p.StartInfo.FileName = process
        If Not (workingDir = "") Then
            p.StartInfo.WorkingDirectory = workingDir
        End If
        p.StartInfo.Arguments = param
        ' need to set this to false to redirect output
        p.StartInfo.UseShellExecute = False
        p.StartInfo.RedirectStandardError = True
        p.StartInfo.RedirectStandardOutput = True
        p.StartInfo.CreateNoWindow = True
        ' start the process
        p.Start()

        out = p.StandardOutput.ReadToEnd
        err = p.StandardError.ReadToEnd

        If err <> "" Then
            output = "Error#:" & err
        ElseIf out <> "" Then
            output = "Output#:" & out
        End If

        'p.WaitForExit()
        Return output
    End Function

    Public Sub ConvertIMAGEtoICO(fileName As String, newfileName As String)
        Dim bmp As System.Drawing.Bitmap = System.Drawing.Image.FromFile(fileName, True)
        Dim ico As System.Drawing.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())
        Dim icofs As Stream = File.Create(newfileName)
        ico.Save(icofs)
        icofs.Close()
    End Sub

    Public Function GetAllFiles(ByVal directory As String, Pattern As String, SearchOption As SearchOption) As List(Of String)
        Return IO.Directory.GetFiles(directory, Pattern, SearchOption).ToList()
    End Function

    Public Function GetFolderSize(Folder As String, ByVal isSubFolder As Boolean) As String
        Dim di As DirectoryInfo = New DirectoryInfo(Folder)
        Return GetSize(di, True)
    End Function

    Private Function GetSize(ByVal dInfo As DirectoryInfo, ByVal isSubFolder As Boolean, Optional sEinheit As Einheit = Einheit.KB) As String
        Dim sizeInBytes As Long = 0
        sizeInBytes = dInfo.EnumerateFiles().Sum(Function(fi) fi.Length)
        If isSubFolder = True Then
            sizeInBytes += dInfo.EnumerateDirectories().Sum(Function(Dir) GetSize(Dir, True))
        End If
        If sEinheit = Einheit.KB Then
            Return Math.Round(sizeInBytes / 1024, 2)
        ElseIf sEinheit = Einheit.MB Then
            Return Math.Round(sizeInBytes / 1024 / 1024, 2)
        ElseIf sEinheit = Einheit.GB Then
            Return Math.Round(sizeInBytes / 1024 / 1024 / 1024, 2)
        End If
    End Function

    Public Function ImageConverter(Input As String, Format As ImageFormat) As String
        Dim picbox As New PictureBox
        Return ""
        If IO.File.Exists(Input) = True Then
            picbox.Image = Image.FromFile(Input)
            If Format = "Png" Then
                picbox.Image.Save(Right(Input, Input.Length - 3) & "png", System.Drawing.Imaging.ImageFormat.Png)
                Return Right(Input, Input.Length - 3) & "png"
            ElseIf Format = "Jpg" Then
                picbox.Image.Save(Right(Input, Input.Length - 3) & "jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
                Return Right(Input, Input.Length - 3) & "jpg"
            ElseIf Format = "Bmp" Then
                picbox.Image.Save(Right(Input, Input.Length - 3) & "bmp", System.Drawing.Imaging.ImageFormat.Bmp)
                Return Right(Input, Input.Length - 3) & "bmp"
            ElseIf Format = "Gif" Then
                picbox.Image.Save(Right(Input, Input.Length - 3) & "gif", System.Drawing.Imaging.ImageFormat.Gif)
                Return Right(Input, Input.Length - 3) & "gif"
            End If
        End If
    End Function

    Public Function FileToByteArray(ByVal FileName As String) As Byte()
        Dim _Buffer() As Byte = Nothing
        Try

            ' Open file for reading
            Dim _FileStream As New System.IO.FileStream(FileName, System.IO.FileMode.Open, System.IO.FileAccess.Read)

            ' attach filestream to binary reader

            Dim _BinaryReader As New System.IO.BinaryReader(_FileStream)

            ' get total byte length of the file

            Dim _TotalBytes As Long = New System.IO.FileInfo(FileName).Length

            ' read entire file into buffer

            _Buffer = _BinaryReader.ReadBytes(CInt(Fix(_TotalBytes)))

            ' close file reader

            _FileStream.Close()
            _FileStream.Dispose()
            _BinaryReader.Close()
        Catch _Exception As Exception
            ' Error

            Console.WriteLine("Exception caught in process: {0}", _Exception.ToString())
        End Try
        Return _Buffer
    End Function

    Public Overloads Function EncodeBase64(input As String) As String
        Return System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(input))
    End Function

    Public Overloads Function EncodeBase64(input As String, Optional x As Integer = 0) As String
        Return System.Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(input))
    End Function

    Public Overloads Function DecodeBase64(input As String) As String
        Return System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(input))
    End Function

    Public Overloads Function DecodeBase64(input As String, Optional x As Integer = 0) As String
        Return System.Text.ASCIIEncoding.ASCII.GetString(Convert.FromBase64String(input))
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="sDir">...</param>
    ''' <param name="List">....</param>
    ''' <param name="Pattern">....</param>
    ''' <param name="SearchOption">....</param>
    Public Overloads Function GetFiles(ByVal sDir As String, ByVal List As ListBox, Optional NoDirectory As String = "", Optional ByVal Pattern As String = "*.*", Optional SearchOption As FileIO.SearchOption = FileIO.SearchOption.SearchAllSubDirectories) As String
        Dim sFile As String

        If Not sDir.EndsWith("\") Then sDir += "\"

        ' alle TXT-Dateien im Startverzeichnis einschl. Unterordner 
        ' in einer ListBox anzeigen
        Try
            For Each sFile In My.Computer.FileSystem.GetFiles(sDir, SearchOption, Pattern)

                ' Dateiname mit relativer Pfadangabe zum Startverzeichnis ausgeben
                If NoDirectory = "" Then
                    List.Items.Add(sFile)
                ElseIf NoDirectory <> "" Then
                    If sFile.Contains(NoDirectory) = False Then
                        List.Items.Add(sFile)
                    End If
                End If
            Next
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    ''' <summary>
    ''' ....
    ''' </summary>
    ''' <param name="sDir">...</param>
    ''' <param name="List">....</param>
    ''' <param name="Pattern">....</param>
    ''' <param name="SearchOption">....</param>
    Public Overloads Function GetFiles(ByVal sDir As String, ByVal List As List(Of String), Optional NoDirectory As String = "", Optional ByVal Pattern As String = "*.*", Optional SearchOption As FileIO.SearchOption = FileIO.SearchOption.SearchAllSubDirectories) As String
        Dim sFile As String

        If Not sDir.EndsWith("\") Then sDir += "\"

        ' alle TXT-Dateien im Startverzeichnis einschl. Unterordner 
        ' in einer ListBox anzeigen
        Try
            For Each sFile In My.Computer.FileSystem.GetFiles(
                          sDir, SearchOption, Pattern)

                ' Dateiname mit relativer Pfadangabe zum Startverzeichnis ausgeben
                If NoDirectory = "" Then
                    List.Add(sFile)
                ElseIf NoDirectory <> "" Then
                    If sFile.Contains(NoDirectory) = False Then
                        List.Add(sFile)
                    End If
                End If
            Next
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    ''' <summary>
    ''' Berechnen Daten des Advent
    ''' </summary>
    ''' <param name="Jahr">das betreffende Jahr</param>
    ''' <returns></returns>
    Public Function Adventstage(Jahr As Short) As Date()
        ' Der erste Advent ist der erste Sonntag nach dem 26.11.
        Dim Advent(3), hAdv, Adv As Date, i As Short
        hAdv = DateSerial(Jahr, 11, 26)
        Adv = hAdv
        For i = 1 To 7
            hAdv = DateAdd(DateInterval.Day, i, DateSerial(Jahr, Adv.Month, Adv.Day))
            If hAdv.DayOfWeek = 0 Then
                Advent(0) = hAdv
                Exit For
            End If
        Next
        For i = 1 To 3
            Adv = DateAdd(DateInterval.Day, i * 7, DateSerial(Jahr, hAdv.Month, hAdv.Day))
            Advent(i) = Adv
        Next i
        Return Advent
    End Function

    ''' <summary>
    ''' Suchen der Pfade eines Files
    ''' </summary>
    ''' <param name="rootFolderPath">der Ausgangspfad</param>
    ''' <param name="fName">der Filename inkl. Extension</param>
    ''' <param name="fExt">die File-Extension</param>
    ''' <returns></returns>
    Public Function FindFilePathsByFileName(rootFolderPath As String, fName As String, fExt As String) As String()

        Dim i, j As Integer
        Dim fileLocation As DirectoryInfo
        Dim fList() As String
        Dim subDirs() As String = Directory.GetDirectories(rootFolderPath)

        j = 0
        For i = 0 To subDirs.Length - 1
            Try
                fileLocation = New DirectoryInfo(subDirs(i))
                For Each Direct In fileLocation.GetFiles()
                    If (Direct IsNot Nothing) Then
                        If (Path.GetExtension(Direct.ToString.ToLower) = fExt) Then
                            ' Debug.Print(Direct.FullName + vbTab + Direct.Name)
                            If Direct.Name = fName Then
                                ReDim Preserve fList(j)
                                fList(j) = Path.GetDirectoryName(Direct.FullName)
                                j += 1
                            End If
                        End If
                    End If
                Next
            Catch ex As Exception
            End Try
        Next
        Return fList
    End Function

    'Die zu ersetzenden Satzzeichen befinden sich hierbei allesamt nach [\ und vor dem anschließenden ] Zeichen.
    'In diesem Fall werden demnach die Zeichen ".,!?;:" aus dem String entfernt.
    Public Function ReplaceFromString(sText As String, sReplace As String, Optional sReplaceWith As String = "") As String
        'Return System.Text.RegularExpressions.Regex.Replace(sText, "[\.,!?;:]", "")
        Return System.Text.RegularExpressions.Regex.Replace(sText, "[\" & sReplace & "]", sReplaceWith)
    End Function

    ''' <summary>
    ''' 	finden von Dateien rekursiv
    ''' </summary>
    ''' <param name="path">der zu durchsuchende Pfad</param>
    ''' <param name="searchPattern">der File-Filter nach Extension</param>
    ''' <param name="inclPath">suchen auch in Unterordnern</param>
    ''' <param name="uCase">die Filenamen nur als Großbuchstaben</param>
    ''' <returns>die String-Liste der Filenamen</returns>
    Public Function MyGetFiles(path As String,
      searchPattern As String,
      Optional inclPath As Boolean = True,
      Optional uCase As Boolean = False) As List(Of String)

        ' Aufruf
        ' Dim files As List(Of String)
        ' files = MyGetFiles("C:", "*.txt")
        Dim a As New List(Of String)
        Try
            ' finde Files
            Dim files() As String
            files = Directory.GetFiles(path, searchPattern)
            For Each file As String In files
                If uCase Then file = file.ToUpper()
                a.Add(file)
            Next
        Catch ex As Exception
        End Try

        Try
            ' durchsuche Unterordner
            Dim subfolders() As String, b As New List(Of String)
            subfolders = Directory.GetDirectories(path)
            ' durchsuchen von Unterordnern mit rekursivem Aufruf
            For Each subfolder As String In subfolders
                a.AddRange(MyGetFiles(subfolder, searchPattern))
            Next
        Catch ex As Exception
        End Try

        If Not inclPath Then
            Dim i, idx As Short, t As String
            For i = 0 To a.Count - 1
                t = a(i)
                idx = t.LastIndexOf("") + 1
                t = t.Substring(idx)
                a(i) = t
                ' i += 1
            Next
        End If
        ' return die gefundenen Files
        Return a
    End Function

    ''' <summary>
    ''' Sortierung einer Datei-Namenliste, 
    ''' wo die Namen mit komplettem Pfad in einem Array angegeben werden,
    ''' nach einer von 3 verschiedenen Zeitangaben.
    ''' Die Zeiten LastAccessTime, LastWriteTime und CreationTime stehen zur Verfügung.
    ''' Das Array wird nach einer von diesen 3 Zeiten sortiert.
    ''' </summary>
    ''' <param name="theFileArray">das zu sortierende Array mit den kompletten Pfaden der zu sortierenden Dateien</param>
    ''' <param name="theTime">die Zeit, nach der sortiert werden soll</param>
    ''' <param name="doReverse">soll das Array nach der Sortierung umgekehrt werden, true|false</param>
    ''' <param name="delPath">sollen im Array die Pfade beseitigt werden, true|false</param>
    Public Sub SortFilesByDate(theFileArray() As String, theTime As EnumFileAttributes,
      Optional doReverse As Boolean = False,
      Optional delPath As Boolean = False)

        Dim ArrLen As Short = theFileArray.Length - 1
        Dim tmpValue, sp() As String, i As Short, flag As Boolean

        Select Case theTime
            Case EnumFileAttributes.CreationTime
                Do
                    flag = True
                    For i = 0 To ArrLen - 1
                        If File.GetCreationTime(theFileArray(i)) _
                          > File.GetCreationTime(theFileArray(i + 1)) Then
                            tmpValue = theFileArray(i)
                            theFileArray(i) = theFileArray(i + 1)
                            theFileArray(i + 1) = tmpValue
                            flag = False
                        End If
                    Next i
                Loop Until flag
            Case EnumFileAttributes.LastAccessTime
                Do
                    flag = True
                    For i = 0 To ArrLen - 1
                        If File.GetLastAccessTime(theFileArray(i)) _
                          > File.GetLastAccessTime(theFileArray(i + 1)) Then
                            tmpValue = theFileArray(i)
                            theFileArray(i) = theFileArray(i + 1)
                            theFileArray(i + 1) = tmpValue
                            flag = False
                        End If
                    Next i
                Loop Until flag
            Case EnumFileAttributes.LastWriteTime
                Do
                    flag = True
                    For i = 0 To ArrLen - 1
                        If File.GetLastWriteTime(theFileArray(i)) _
                          > File.GetLastWriteTime(theFileArray(i + 1)) Then
                            tmpValue = theFileArray(i)
                            theFileArray(i) = theFileArray(i + 1)
                            theFileArray(i + 1) = tmpValue
                            flag = False
                        End If
                    Next i
                Loop Until flag
        End Select

        If delPath Then    ' Pfade löschen
            For i = 0 To ArrLen
                sp = Split(theFileArray(i), "")
                theFileArray(i) = sp(sp.Length - 1)
            Next
        End If
        If doReverse Then Array.Reverse(theFileArray)   ' Feld umkehren
    End Sub
End Class


