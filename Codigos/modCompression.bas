Attribute VB_Name = "modCompression"
Option Explicit

Public PkContra() As Byte

'This structure will describe our binary file's
'size and number of contained files
Public Type FILEHEADER
    lngFileSize As Long                 'How big is this file? (Used to check integrity)
    lngNumFiles As Long                 'How many files are inside?
End Type

'This structure will describe each file contained
'in our binary file
Public Type INFOHEADER
    
    lngFileStart As Long            'Where does the chunk start?
    lngFileSize As Long             'How big is this chunk of stored data?
    strFileName As String * 32      'What's the name of the file this data came from?
    lngFileSizeUncompressed As Long 'How big is the file compressed
End Type

Public Enum srcFileType
    Graphics
    Ambient
    Music
    Wav
    Scripts
    Map
    Interface
    Fuentes
    Skin
    Minimap
    Patch
End Enum

Public Windows_Temp_Dir As String

Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, FreeBytesToCaller As Currency, bytesTotal As Currency, FreeBytesTotal As Currency) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function Compress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function UnCompress Lib "zlib.dll" Alias "uncompress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Public Sub GenerateContra()
'***************************************************
'Author: ^[GS]^
'Last Modification: 17/06/2012 - ^[GS]^
'
'***************************************************

'on error resume next
    Dim Contra As String
    Dim loopc As Byte
    
    Contra = PasswordResources
    
    Erase PkContra
    
    If LenB(Contra) <> 0 Then
        ReDim PkContra(Len(Contra) - 1)
        For loopc = 0 To UBound(PkContra)
            PkContra(loopc) = Asc(mid(Contra, loopc + 1, 1))
        Next loopc
    End If
    
End Sub

Public Sub Compress_Data(ByRef Data() As Byte)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Compresses binary data avoiding data loses
'*****************************************************************

    Dim Dimensions As Long
    Dim DimBuffer As Long
    Dim BufTemp() As Byte
    Dim BufTemp2() As Byte
    Dim loopc As Long
    
    'Get size of the uncompressed data
    Dimensions = UBound(Data)
    
    'Prepare a buffer 1.06 times as big as the original size
    DimBuffer = Dimensions * 1.06
    ReDim BufTemp(DimBuffer)
    
    'Compress data using zlib
    Compress BufTemp(0), DimBuffer, Data(0), Dimensions
    
    'Deallocate memory used by uncompressed data
    Erase Data
    
    'Get rid of unused bytes in the compressed data buffer
    ReDim Preserve BufTemp(DimBuffer - 1)
    
    'Copy the compressed data buffer to the original data array so it will return to caller
    Data = BufTemp
    
    'Deallocate memory used by the temp buffer
    Erase BufTemp
    
    'Encrypt the first byte of the compressed data for extra security
    If UBound(PkContra) <= UBound(Data) And UBound(PkContra) <> 0 Then
        For loopc = 0 To UBound(PkContra)
            Data(loopc) = Data(loopc) Xor PkContra(loopc)
        Next loopc
    End If
End Sub

Public Sub Decompress_Data(ByRef Data() As Byte, ByVal OrigSize As Long)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Decompresses binary data
'*****************************************************************

    Dim BufTemp() As Byte
    Dim loopc As Integer
    
    ReDim BufTemp(OrigSize - 1)
    
    'Des-encrypt the first byte of the compressed data
    If UBound(PkContra) <= UBound(Data) And UBound(PkContra) <> 0 Then
        For loopc = 0 To UBound(PkContra)
            Data(loopc) = Data(loopc) Xor PkContra(loopc)
        Next loopc
    End If
    
    UnCompress BufTemp(0), OrigSize, Data(0), UBound(Data) + 1
    
    ReDim Data(OrigSize - 1)
    
    Data = BufTemp
    
    Erase BufTemp
End Sub

Private Sub encryptHeaderFile(ByRef FileHead As FILEHEADER)
    'Each different variable is encrypted with a different key for extra security
    With FileHead
        .lngNumFiles = .lngNumFiles Xor 37816
        .lngFileSize = .lngFileSize Xor 245378169
    End With
End Sub

Private Sub encryptHeaderInfo(ByRef InfoHead As INFOHEADER)
    Dim EncryptedFileName As String
    Dim loopc As Long
    
    For loopc = 1 To Len(InfoHead.strFileName)
        If loopc Mod 2 = 0 Then
            EncryptedFileName = EncryptedFileName & Chr(Asc(mid(InfoHead.strFileName, loopc, 1)) Xor 12)
        Else
            EncryptedFileName = EncryptedFileName & Chr(Asc(mid(InfoHead.strFileName, loopc, 1)) Xor 23)
        End If
    Next loopc
    
    'Each different variable is encrypted with a different key for extra security
    With InfoHead
        .lngFileSize = .lngFileSize Xor 221872469
        .lngFileSizeUncompressed = .lngFileSizeUncompressed Xor 447915732
        .lngFileStart = .lngFileStart Xor 172379447
        .strFileName = EncryptedFileName
    End With
End Sub

Public Function General_Drive_Get_Free_Bytes(ByVal DriveName As String) As Currency
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 6/07/2004
'
'**************************************************************

    Dim RetVal As Long
    Dim FB As Currency
    Dim BT As Currency
    Dim FBT As Currency
    
    RetVal = GetDiskFreeSpace(Left(DriveName, 2), FB, BT, FBT)
    
    General_Drive_Get_Free_Bytes = FB * 10000 'convert result to actual size in bytes
End Function

Public Function General_Get_Temp_Dir() As String
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2005
'Gets windows temporary directory
'**************************************************************

 Const MAX_LENGTH = 512
   Dim s As String
   Dim c As Long
   s = Space$(MAX_LENGTH)
   c = GetTempPath(MAX_LENGTH, s)
   If c > 0 Then
       If c > Len(s) Then
           s = Space$(c + 1)
           c = GetTempPath(MAX_LENGTH, s)
       End If
   End If
   General_Get_Temp_Dir = IIf(c > 0, Left$(s, c), "")
End Function

'--------------------------------------------------------------------------------
' Project    :       Compresor
' Procedure  :       extractFiles
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LORWIK
' Date-Time  :       9/25/2023-23:44:57
'
' Parameters :       File_Type (srcFileType)
'--------------------------------------------------------------------------------
Public Function extractFiles(ByVal File_Type As srcFileType) As Boolean
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Extracts all files from a resource file
'*****************************************************************

    Dim loopc As Long
    Dim SourceFilePath As String
    Dim OutputFilePath As String
    Dim SourceFile As Integer
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim handle As Integer

On Local Error GoTo errhandler

    Select Case File_Type
        Case Graphics
            SourceFilePath = SrcPath & "\Graficos" & Extension
            OutputFilePath = OutPath & "\Graficos\"
            
        Case Ambient
            SourceFilePath = SrcPath & "\Ambient" & Extension
            OutputFilePath = OutPath & "\Ambient\"
        
        Case Music
            SourceFilePath = SrcPath & "\Musica" & Extension
            OutputFilePath = OutPath & "\Musica\"
            
        Case Wav
            SourceFilePath = SrcPath & "\Sounds" & Extension
            OutputFilePath = OutPath & "\Wav\"
            
        Case Scripts
            SourceFilePath = SrcPath & "\Scripts" & Extension
            OutputFilePath = OutPath & "\Init\"
            
        Case Map
            SourceFilePath = SrcPath & "\Mapas" & Extension
            OutputFilePath = OutPath & "\Mapas\"
            
        Case Interface
            SourceFilePath = SrcPath & "\Interface" & Extension
            OutputFilePath = OutPath & "\Interface\"
            
        Case Fuentes
            SourceFilePath = SrcPath & "\Fuentes" & Extension
            OutputFilePath = OutPath & "\Fuentes\"
            
        Case Skin
            SourceFilePath = SrcPath & FrmMain.txtSkinName.Text & Extension
            OutputFilePath = OutPath & "\Skins\" & FrmMain.txtSkinName.Text & "\"
            
        Case Minimap
            SourceFilePath = SrcPath & "\Minimap" & Extension
            OutputFilePath = OutPath & "\Minimapa\"
            
        Case Patch
            SourceFilePath = SrcPath & "\Patch" & Extension
            OutputFilePath = OutPath & "\Patch\"
        
        Case Else
            Exit Function
    End Select
        
    
    'Open the binary file
    SourceFile = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As SourceFile
    
    'Extract the FILEHEADER
    Get SourceFile, 1, FileHead
    
    'Desencrypt FILEHEADER
    encryptHeaderFile FileHead
    
    'Check the file for validity
    If LOF(SourceFile) <> FileHead.lngFileSize Then ' - Pass1.lngFileSize - 1 Then
         MsgBox "El archivo de recursos " & SourceFilePath & " esta corrupto.", , "Error"
        Close SourceFile
        Erase InfoHead
        Exit Function
    End If
    
    'Size the InfoHead array
    ReDim InfoHead(FileHead.lngNumFiles - 1)

    'Extract the INFOHEADER
    Get SourceFile, , InfoHead

    'Extract all of the files from the binary file
    For loopc = 0 To UBound(InfoHead)
        'Desencrypt each INFOHEADER befWO accessing the data
        encryptHeaderInfo InfoHead(loopc)
        
        'Check if there is enough memory
        If InfoHead(loopc).lngFileSizeUncompressed > General_Drive_Get_Free_Bytes(Left(App.Path, 3)) Then
            MsgBox "No tienes suficiente espacio en el disco para seguir descomprimiendo archivos."
            Exit Function
        End If

        'Resize the byte data array
        ReDim SourceData(InfoHead(loopc).lngFileSize - 1)
        
        'Get the data
        Get SourceFile, InfoHead(loopc).lngFileStart, SourceData
        
        'Decompress all data
        Decompress_Data SourceData, InfoHead(loopc).lngFileSizeUncompressed
        
        'Get a free handler
        handle = FreeFile

        'Create a new file and put in the data
        Open OutputFilePath & InfoHead(loopc).strFileName For Binary As handle
        
        Put handle, , SourceData
        
        Close handle
        
        Erase SourceData
        
        DoEvents
    Next loopc
    
    'Close the binary file
    Close SourceFile
    
    Erase InfoHead
    
    extractFiles = True
Exit Function

errhandler:
    Close SourceFile
    Erase SourceData
    Erase InfoHead
    'Display an error message if it didn't work
    MsgBox "Unable to decode binary file. Reason: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error"
End Function

Public Function compressFiles(ByVal File_Type As srcFileType) As Boolean
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Comrpesses all files to a resource file
'*****************************************************************

    Dim SourceFilePath As String
    Dim SourceFileExtension As String
    Dim OutputFilePath As String
    Dim SourceFile As Long
    Dim OutputFile As Long
    Dim SourceFileName As String
    Dim SourceData() As Byte
    Dim FileHead As FILEHEADER
    Dim InfoHead() As INFOHEADER
    Dim FileNames() As String
    Dim lngFileStart As Long
    Dim loopc As Long
    Dim tmplng As Long
    
'Set up the error handler
On Local Error GoTo errhandler

    Dim SourcePath As String
    
    SourcePath = OutPath
    
    Select Case File_Type
        Case Graphics
            SourceFileExtension = ".png"
            SourceFilePath = SourcePath & "\Graficos\"
            OutputFilePath = SrcPath & "\Graficos" & Extension
            
        Case Ambient
            SourceFileExtension = ".wav"
            SourceFilePath = SourcePath & "\Ambient\"
            OutputFilePath = SrcPath & "\Ambient" & Extension
        
        Case Music
            SourceFileExtension = ".mp3"
            SourceFilePath = SourcePath & "\Musica\"
            OutputFilePath = SrcPath & "\Musica" & Extension
            
        Case Wav
            SourceFileExtension = ".wav"
            SourceFilePath = SourcePath & "\Wav\"
            OutputFilePath = SrcPath & "\Sounds" & Extension
            
        Case Scripts
            SourceFileExtension = ".*"
            SourceFilePath = SourcePath & "\Init\"
            OutputFilePath = SrcPath & "\Scripts" & Extension
            
        Case Map
            SourceFileExtension = ".csm"
            SourceFilePath = SourcePath & "\Mapas\"
            OutputFilePath = SrcPath & "\Mapas" & Extension
            
        Case Interface
            SourceFileExtension = ".bmp"
            SourceFilePath = SourcePath & "\Interface\"
            OutputFilePath = SrcPath & "\Interface" & Extension
            
        Case Fuentes
            SourceFileExtension = ".*"
            SourceFilePath = SourcePath & "\Fuentes\"
            OutputFilePath = SrcPath & "\Fuentes" & Extension
            
        Case Skin
            SourceFileExtension = ".bmp"
            SourceFilePath = SourcePath & "\Skins\" & FrmMain.txtSkinName.Text & "\"
            OutputFilePath = SrcPath & "\" & FrmMain.txtSkinName.Text & Extension
                    
        Case Minimap
            SourceFileExtension = ".*"
            SourceFilePath = SourcePath & "\Minimapa\"
            OutputFilePath = SrcPath & "\Minimap" & Extension
            
        Case Patch
            SourceFileExtension = ".*"
            SourceFilePath = SourcePath & "\Parches\"
            OutputFilePath = SrcPath & "\Patch" & Extension
        
        Case Else
            Exit Function
            
    End Select
    
    'Get first file in the directoy
    SourceFileName = Dir$(SourceFilePath & "*" & SourceFileExtension, vbNormal)
    
    SourceFile = FreeFile
    
    'Get all other files i nthe directory
    While SourceFileName <> ""
        FileHead.lngNumFiles = FileHead.lngNumFiles + 1
        
        ReDim Preserve FileNames(FileHead.lngNumFiles - 1)
        FileNames(FileHead.lngNumFiles - 1) = LCase(SourceFileName)
        
        'Search new file
        SourceFileName = Dir$()
    Wend
    
    'If we found none, be can't compress a thing, so we exit
    If FileHead.lngNumFiles = 0 Then
        MsgBox "There are no files of extension " & SourceFileExtension & " in " & SourceFilePath & ".", , "Error"
        Exit Function
    End If
    
    'Sort file names alphabetically (this will make patching much easier).
    General_Quick_Sort FileNames(), 0, UBound(FileNames)
    
    'Resize InfoHead array
    ReDim InfoHead(FileHead.lngNumFiles - 1)
        
    'Destroy file if it previuosly existed
    If Dir(OutputFilePath, vbNormal) <> "" Then
        Kill OutputFilePath
    End If
    
    'Open a new file
    OutputFile = FreeFile
    Open OutputFilePath For Binary Access Read Write As OutputFile
    
    For loopc = 0 To FileHead.lngNumFiles - 1
        'Find a free file number to use and open the file
        
        SourceFile = FreeFile
        Open SourceFilePath & FileNames(loopc) For Binary Access Read Lock Write As SourceFile
        
        'Store file name
        InfoHead(loopc).strFileName = FileNames(loopc)
        
        'Find out how large the file is and resize the data array appropriately
        ReDim SourceData(LOF(SourceFile) - 1)
        
        'Store the value so we can decompress it later on
        InfoHead(loopc).lngFileSizeUncompressed = LOF(SourceFile)
        
        'Get the data from the file
        Get SourceFile, , SourceData
        'If loopc = 0 Then SourceData = "115792!"
        'Compress it
        Compress_Data SourceData
        
        'Save it to a temp file
        Put OutputFile, , SourceData
        
        'Set up the file header
        FileHead.lngFileSize = FileHead.lngFileSize + UBound(SourceData) + 1
        
        'Set up the info headers
        InfoHead(loopc).lngFileSize = UBound(SourceData) + 1
        
        Erase SourceData
        
        'Close temp file
        Close SourceFile
        
        DoEvents
    Next loopc
    
    'Finish setting the FileHeader data
    FileHead.lngFileSize = FileHead.lngFileSize + CLng(FileHead.lngNumFiles) * Len(InfoHead(0)) + Len(FileHead)
    
    'Set InfoHead data
    lngFileStart = Len(FileHead) + CLng(FileHead.lngNumFiles) * Len(InfoHead(0)) + 1
    For loopc = 0 To FileHead.lngNumFiles - 1
        InfoHead(loopc).lngFileStart = lngFileStart
        lngFileStart = lngFileStart + InfoHead(loopc).lngFileSize
        'Once an InfoHead index is ready, we encrypt it
        encryptHeaderInfo InfoHead(loopc)
    Next loopc
    
    'Desencrypt File Header
    encryptHeaderFile FileHead
    
    'InfoHead.lngFileSize = InfoHead.lngFileSize + 15
    'Get all data stored so far
    ReDim SourceData(LOF(OutputFile) - 1)
    Seek OutputFile, 1
    Get OutputFile, , SourceData
    
    Seek OutputFile, 1
    
    'Store the data in the file
    Put OutputFile, , FileHead
    Put OutputFile, , InfoHead
    Put OutputFile, , SourceData
    
    'Close the file
    Close OutputFile
    
    Erase InfoHead
    Erase SourceData
    
    compressFiles = True
Exit Function

errhandler:
    Erase SourceData
    Erase InfoHead
    'Display an error message if it didn't work
    MsgBox "Unable to create binary file. Reason: " & Err.Number & " : " & Err.Description, vbOKOnly, "Error"
End Function

Public Function extractFile(ByVal File_Type As srcFileType, ByVal file_name As String, Optional ByVal TempDir As Boolean = False) As Boolean
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 10/13/2004
'Extracts all files from a resource file
'*****************************************************************

    Dim loopc As Long
    
    Dim SourceFilePath As String
    Dim OutputFilePath As String
    
    Dim SourceData() As Byte
    Dim InfoHead As INFOHEADER
    Dim handle As Integer
    
'Set up the error handler
On Local Error GoTo errhandler
    
    Select Case File_Type
        Case Graphics
            SourceFilePath = SrcPath & "\Graficos" & Extension
            OutputFilePath = OutPath & "\Graficos\"
            
        Case Ambient
            SourceFilePath = SrcPath & "\Ambient" & Extension
            OutputFilePath = OutPath & "\Ambient\"
        
        Case Music
            SourceFilePath = SrcPath & "\Musica" & Extension
            OutputFilePath = OutPath & "\Musica\"
            
        Case Wav
            SourceFilePath = SrcPath & "\Sounds" & Extension
            OutputFilePath = OutPath & "\Wav\"
            
        Case Scripts
            SourceFilePath = SrcPath & "\Scripts" & Extension
            OutputFilePath = OutPath & "\Init\"
            
        Case Map
            SourceFilePath = SrcPath & "\Mapas" & Extension
            OutputFilePath = OutPath & "\Mapas\"
            
        Case Interface
            SourceFilePath = SrcPath & "\Interface" & Extension
            OutputFilePath = OutPath & "\Interface\"
            
        Case Fuentes
            SourceFilePath = SrcPath & "\Fuentes" & Extension
            OutputFilePath = OutPath & "\Fuentes\"
            
        Case Skin
            SourceFilePath = SrcPath & FrmMain.txtSkinName.Text & Extension
            OutputFilePath = OutPath & "\Skins\" & FrmMain.txtSkinName.Text & "\"
            
        Case Minimap
            SourceFilePath = SrcPath & "\Minimap" & Extension
            OutputFilePath = OutPath & "\Minimapa\"
            
        Case Patch
            SourceFilePath = SrcPath & "\Patch" & Extension
            OutputFilePath = OutPath & "\Patch\"
        
        Case Else
            Exit Function
    End Select
    
    '¿Queremos descomprimir en la carpeta temporal?
    If TempDir Then OutputFilePath = Windows_Temp_Dir
    
    'Find the Info Head of the desired file
    InfoHead = File_Find(SourceFilePath, file_name)
    
    If InfoHead.strFileName = "" Or InfoHead.lngFileSize = 0 Then Exit Function

    'Open the binary file
    handle = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As handle
    
    'Check the file for validity
    'If LOF(handle) <> InfoHead.lngFileSize Then
    '    Close handle
    '    MsgBox "Resource file " & SourceFilePath & " seems to be corrupted.", , "Error"
    '    Exit Function
    'End If
    
    'Make sure there is enough space in the HD
    If InfoHead.lngFileSizeUncompressed > General_Drive_Get_Free_Bytes(Left$(App.Path, 3)) Then
        Close handle
        MsgBox "There is not enough drive space to extract the compressed file.", , "Error"
        Exit Function
    End If
    
    'Extract file from the binary file
    
    'Resize the byte data array
    ReDim SourceData(InfoHead.lngFileSize - 1)
    
    'Get the data
    Get handle, InfoHead.lngFileStart, SourceData
    
    'Decompress all data
    Decompress_Data SourceData, InfoHead.lngFileSizeUncompressed
    
    'Close the binary file
    Close handle
    
    'Get a free handler
    handle = FreeFile
    
    Open OutputFilePath & InfoHead.strFileName For Binary As handle
    
    Put handle, 1, SourceData
    
    Close handle
    
    Erase SourceData
        
    extractFile = True
Exit Function

errhandler:
    Close handle
    Erase SourceData
    'Display an error message if it didn't work
    'MsgBox "Unable to decode binary file. Reason: " & Err.number & " : " & Err.Description, vbOKOnly, "Error"
End Function

Public Function Extract_File_Memory(ByVal File_Type As srcFileType, ByVal file_name As String, ByRef SourceData() As Byte) As Boolean
 '********************************************
'Author: ???
'Last Modify Date: ???
'Extra archivos en memoria
'*********************************************

    Dim loopc As Long
    Dim SourceFilePath As String
    Dim InfoHead As INFOHEADER
    Dim handle As Integer
   
On Local Error GoTo errhandler
   
    Select Case File_Type
        Case Graphics
                SourceFilePath = SrcPath & "\Graficos" & Extension
            
        Case Music
                SourceFilePath = SrcPath & "\Musica" & Extension
        
        Case Wav
                SourceFilePath = SrcPath & "\Sounds" & Extension

        Case Scripts
                SourceFilePath = SrcPath & "\Scripts" & Extension

        Case Interface
                SourceFilePath = SrcPath & "\Interface" & Extension

        Case Map
                SourceFilePath = SrcPath & "\Mapas" & Extension

        Case Ambient
                SourceFilePath = SrcPath & "\Ambient" & Extension
                
        Case Fuentes
                SourceFilePath = SrcPath & "\Fuentes" & Extension
                
        Case Minimap
                SourceFilePath = SrcPath & "\Minimap" & Extension
                
        Case Else
            Exit Function
    End Select
   
    InfoHead = File_Find(SourceFilePath, file_name)
   
    If InfoHead.strFileName = "" Or InfoHead.lngFileSize = 0 Then Exit Function
 
    handle = FreeFile
    Open SourceFilePath For Binary Access Read Lock Write As handle
   
    If InfoHead.lngFileSizeUncompressed > General_Drive_Get_Free_Bytes(Left$(App.Path, 3)) Then
        Close handle
        MsgBox "There is not enough drive space to extract the compressed file.", , "Error"
        Exit Function
    End If
   
   
    ReDim SourceData(InfoHead.lngFileSize - 1)
   
    Get handle, InfoHead.lngFileStart, SourceData
        Decompress_Data SourceData, InfoHead.lngFileSizeUncompressed
    Close handle
       
    Extract_File_Memory = True
Exit Function
 
errhandler:
    Close handle
    Erase SourceData
End Function

Public Sub DeleteFile(ByVal file_path As String)
'*****************************************************************
'Author: Juan Martín Dotuyo Dodero
'Last Modify Date: 3/03/2005
'Deletes a resource files
'*****************************************************************

    Dim handle As Integer
    Dim Data() As Byte
    
    On Error GoTo ERROR_HANDLER
    
    'We open the file to delete
    handle = FreeFile
    Open file_path For Binary Access Write Lock Read As handle
    
    'We replace all the bytes in it with 0s
    ReDim Data(LOF(handle) - 1)
    Put handle, 1, Data
    
    'We close the file
    Close handle
    
    'Now we delete it, knowing that if they retrieve it (some antivirus may create backup copies of deleted files), it will be useless
    Kill file_path
    
    Exit Sub
    
ERROR_HANDLER:
    Kill file_path
        
End Sub
Public Function General_File_Exists(ByVal file_path As String, ByVal File_Type As VbFileAttribute) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Checks to see if a file exists
'*****************************************************************

    If Dir(file_path, File_Type) = "" Then
        General_File_Exists = False
    Else
        General_File_Exists = True
    End If
End Function

Public Sub General_Quick_Sort(ByRef SortArray As Variant, ByVal First As Long, ByVal Last As Long)
 '********************************************
'Author: ???
'Last Modify Date: ???
'Extra archivos en memoria
'*********************************************

    Dim Low As Long, High As Long
    Dim temp As Variant
    Dim List_Separator As Variant
   
    Low = First
    High = Last
    List_Separator = SortArray((First + Last) / 2)
    Do While (Low <= High)
        Do While SortArray(Low) < List_Separator
            Low = Low + 1
        Loop
        Do While SortArray(High) > List_Separator
            High = High - 1
        Loop
        If Low <= High Then
            temp = SortArray(Low)
            SortArray(Low) = SortArray(High)
            SortArray(High) = temp
            Low = Low + 1
            High = High - 1
        End If
    Loop
    If First < High Then General_Quick_Sort SortArray, First, High
    If Low < Last Then General_Quick_Sort SortArray, Low, Last
End Sub

Public Function Get_Extract(ByVal File_Type As srcFileType, ByVal file_name As String) As String
 '********************************************
'Author: ???
'Last Modify Date: ???
'Extra archivos en memoria
'*********************************************

    extractFile File_Type, LCase$(file_name), True
    Get_Extract = Windows_Temp_Dir & LCase$(file_name)
End Function

Public Function File_Find(ByVal resource_file_path As String, ByVal file_name As String) As INFOHEADER
 '********************************************
'Author: ???
'Last Modify Date: ???
'Extra archivos en memoria
'*********************************************
 
On Error GoTo errhandler
 
    Dim Max As Integer
    Dim Min As Integer
    Dim mid As Integer
    Dim file_handler As Integer
   
    Dim file_head As FILEHEADER
    Dim info_head As INFOHEADER
   
    If Len(file_name) < Len(info_head.strFileName) Then _
        file_name = file_name & Space$(Len(info_head.strFileName) - Len(file_name))
   
    file_handler = FreeFile
    Open resource_file_path For Binary Access Read Lock Write As file_handler
   
    Get file_handler, 1, file_head
    
    'Desencrypt File Header
    encryptHeaderFile file_head
   
    Min = 1
    Max = file_head.lngNumFiles
   
    Do While Min <= Max
        mid = (Min + Max) / 2
       
        Get file_handler, CLng(Len(file_head) + CLng(Len(info_head)) * CLng((mid - 1)) + 1), info_head
        
        'Once an InfoHead index is ready, we encrypt it
        encryptHeaderInfo info_head
               
        If file_name < info_head.strFileName Then
            If Max = mid Then
                Max = Max - 1
            Else
                Max = mid
            End If
        ElseIf file_name > info_head.strFileName Then
            If Min = mid Then
                Min = Min + 1
            Else
                Min = mid
            End If
        Else
            File_Find = info_head
           
            Close file_handler
            Exit Function
        End If
    Loop
   
errhandler:
    Close file_handler
    File_Find.strFileName = ""
    File_Find.lngFileSize = 0
End Function
