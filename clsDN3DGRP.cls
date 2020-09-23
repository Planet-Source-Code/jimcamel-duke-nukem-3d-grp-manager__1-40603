VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDN3DGRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Duke3D GRP file format found at http://www.advsys.net/ken/build.htm

'An API which can be used to copy data formats to byte arrays
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'These functions are used to create a temporary file
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Const MAX_PATH As Integer = 260

'Header is 12 Bytes, that is "KenSilverman" (modest, wasnt he?)
Private strHeader As String * 12

'Each File is comprised of 16 bytes, 12 for the name, 4 for the size
'Since the LONG file type is 4 bytes, it will do quite nicely
Private Type typpackedFile
    filename As String * 12
    Filelength As Long
    FilePosition As Long
End Type

'Hold the number of files
Private numFile As Long

'Temporary files go in here before they get saved.
Private tmpAdded() As String

'An array holding all the files in the grp
Private PackedFiles() As typpackedFile

'Just so we know where the grp File is for extraction
Private grpFilePath As String

'So we know whether the file has changed or not
Private hasChanged As Boolean

'The first function called, used to load the data out of the grp file
Public Function LoadFile(FilePath As String) As Boolean
Dim filenum As Integer, i As Integer

'Get a free file number
filenum = FreeFile

'Store the filepath in case we wish to extract anything
grpFilePath = FilePath

'Open the file for reading
Open FilePath For Binary Access Read As #filenum
'Since we declared strHeader as string * 12, we'll get the 12 bytes
Get #filenum, , strHeader

'Check to see that this is a duke3d grp file
If strHeader <> "KenSilverman" Then LoadFile = False: Exit Function

'The next 4 bytes are the number of files
Get #filenum, , numFile

'Redim the number of files
ReDim PackedFiles(1 To numFile)

'The first file position is 16 bytes (for the header)
'and 16*number of files into the file
PackedFiles(1).FilePosition = 17 + (16 * numFile)

'Read in the details for each of the files
For i = 1 To numFile
    With PackedFiles(i)
        Get #filenum, , .filename
        Get #filenum, , .Filelength
        'If I > 1 then update the file position
        If i > 1 Then
            .FilePosition = PackedFiles(i - 1).FilePosition + PackedFiles(i - 1).Filelength
        End If
    End With
Next i

Close #filenum

LoadFile = True
hasChanged = False

End Function

'This sub deletes a file
Public Sub deleteFile(index As Integer)
    Dim i As Integer
    'Are we deleting a saved file or a tempfile
    If index > numFile Then
        'Temp File
        index = index - numFile
        For i = index + 1 To UBound(tmpAdded)
            tmpAdded(i - 1) = tmpAdded(i)
        Next i
        ReDim Preserve tmpAdded(UBound(tmpAdded) - 1)
    Else
        'Saved File
        For i = index + 1 To numFile
            PackedFiles(i - 1) = PackedFiles(i)
        Next i
        numFile = numFile - 1
        ReDim Preserve PackedFiles(1 To numFile)
    End If
    hasChanged = True
End Sub

'This function adds a temporary file
Public Function addFile(FilePath As String) As String
    ReDim Preserve tmpAdded(UBound(tmpAdded) + 1)
    tmpAdded(UBound(tmpAdded)) = FilePath
    addFile = TrimFileName(FilePath)
    hasChanged = True
End Function

Public Sub saveFile(Optional FilePath As String)
    Dim inFile As Integer, grpFile As Integer, outFile As Integer
    Dim Data() As Byte, newFile As String * 12, i As Integer, tmpNumFile As Long
    Dim numFileArray(3) As Byte, tempFile As String
    
    If Len(FilePath) = 0 Then FilePath = grpFilePath
    
    If InStr(1, FilePath, ".") = 0 Then FilePath = FilePath & ".grp"
    
    tempFile = GetUniqueFilename
    
    'Open the grp File to read from
    If Len(grpFilePath) > 0 Then
        grpFile = FreeFile: Open grpFilePath For Binary Access Read As #grpFile
    End If
    
    'Open the output file to write to
    outFile = FreeFile: Open tempFile For Binary Access Write As #outFile
    
    If Len(grpFilePath) > 0 Then
        Get #grpFile, , newFile 'Get the header
    Else
        newFile = "KenSilverman"
    End If
    
    Put #outFile, , newFile 'Print the header
    tmpNumFile = numFile + UBound(tmpAdded) 'Add one to the number of files
    
    'Copy the long into a byte array
    CopyMemory numFileArray(0), ByVal VarPtr(tmpNumFile), Len(tmpNumFile)
    
    Put #outFile, , numFileArray 'Copy in the new number of files
    
    ReDim Data(3)
    If Len(grpFilePath) > 0 Then
        Get #grpFile, , Data 'Move the filepoint 4 right
        
        'Enough memory to hold all the file descriptions
        ReDim Data((16 * numFile) - 1)
        
        Get #grpFile, , Data 'Get the file descriptions
        Put #outFile, , Data 'Add them
    End If
    
    'For each new file to add
    For i = 1 To UBound(tmpAdded)
        'The filename has to be exactly 12 characters
        newFile = TrimFileName(tmpAdded(i))
        'Put our new file in there
        Put #outFile, , newFile
        'Load the file to find out how many bytes it is
        inFile = FreeFile: Open tmpAdded(i) For Binary Access Read As #inFile
        'Get the length of the file
        tmpNumFile = LOF(inFile)
        'Copy it into a byte array
        CopyMemory numFileArray(0), ByVal VarPtr(tmpNumFile), Len(tmpNumFile)
        'And save it in the file
        Put #outFile, , numFileArray
        Close #inFile 'Close the file for the moment
    Next i
    
    If Len(grpFilePath) > 0 Then
        'Copy all the files from one GRP file to the new one
        For i = 1 To numFile
            ReDim Data(FileSize(i) - 1)
            Get #grpFile, FilePosition(i), Data
            Put #outFile, , Data
        Next i
    End If
    
    'Copy all the temporary files into the GRP file
    For i = 1 To UBound(tmpAdded)
        inFile = FreeFile: Open tmpAdded(i) For Binary Access Read As #inFile
        ReDim Data(LOF(inFile) - 1)
        Get #inFile, , Data
        Put #outFile, , Data
        Close #inFile
    Next i
    
    'Remove all the temporary files
    Erase tmpAdded
    ReDim tmpAdded(0)
        
    Close #outFile: If Len(grpFilePath) > 0 Then Close #grpFile
    
    'Copy the tempfile to the new name
    FileCopy tempFile, FilePath
    
    'Delete the tempfile
    Kill tempFile
    
    hasChanged = False
    
    'And load in the new descriptor
    LoadFile FilePath
End Sub

Public Sub extractFile(index As Integer, FilePath As String)
    Dim inFile As Integer, outFile As Integer
    Dim Data() As Byte
    
    'Redim the Data to hold the File
    ReDim Data(FileSize(index) - 1)
    
    'Open the grp File to read from
    inFile = FreeFile: Open grpFilePath For Binary Access Read As #inFile
    
    'Open the output file to write to
    outFile = FreeFile: Open FilePath For Binary Access Write As #outFile
    
    'Get all the data
    Get #inFile, FilePosition(index), Data
    'And write it back out
    Put #outFile, , Data
    
    'And close both the files
    Close #inFile
    Close #outFile
    
End Sub


'Get the number of files
Public Property Get NumberOfFiles() As Integer
    NumberOfFiles = numFile
End Property

'Get the file name
Public Property Get filename(index As Integer) As String
    filename = PackedFiles(index).filename
End Property

'Get the file size
Public Property Get FileSize(index As Integer) As Long
    FileSize = PackedFiles(index).Filelength
End Property

'Get the file position
Public Property Get FilePosition(index As Integer) As Long
    FilePosition = PackedFiles(index).FilePosition
End Property

'Seeing as we're only allowed file names of 12 characters length
'This function is used to shrink files to that long
Private Function TrimFileName(ByVal filename As String) As String
    filename = Replace(filename, "/", "\")
    If InStr(1, filename, "\") > 0 Then
        filename = Mid(filename, InStrRev(filename, "\") + 1)
    End If
    If Len(filename) > 12 Then
        If InStr(1, filename, ".") > 0 Then
            filename = Left(filename, 8) + Mid(filename, InStr(1, filename, "."))
        End If
    End If
    TrimFileName = filename
End Function

'This function gets a Unique name for the temporary file
Private Function GetUniqueFilename(Optional Path As String = "", Optional Prefix As String = "", Optional UseExtension As String = "") As String

Dim wUnique As Long, lngRet As Long
Dim lpTempFileName As String

    wUnique = 0
    If Path = "" Then Path = CurDir
    lpTempFileName = String$(MAX_PATH, 0)
    lngRet = GetTempFileName(Path, Prefix, wUnique, lpTempFileName)

    lpTempFileName = Left(lpTempFileName, InStr(lpTempFileName, Chr$(0)) - 1)
    Call Kill(lpTempFileName)
    If Len(UseExtension) > 0 Then lpTempFileName = Left(lpTempFileName, Len(lpTempFileName) - 3) & UseExtension
    GetUniqueFilename = lpTempFileName

End Function

Private Sub Class_Initialize()
ReDim tmpAdded(0)
End Sub

Public Property Get Changed() As Boolean
Changed = hasChanged
End Property
