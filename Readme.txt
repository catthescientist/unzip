ActiveX dll for zip/unzip files using Shell complex.
Can be used as reference dll in vb6 and vb.net.

"unzip" class have two functions and two subroutines.

Public Function ArchiveRead(ByVal ZipName As String) As String()
Try to open *.zip file "ZipName" and returns array of files in archive.

Public Function ExtractFromZip(ByVal ZipName As String, ByVal DestFolder As String,_
	ByVal nofitem As Integer, ByVal movetotrash As Boolean) As String
Try to open *.zip file "ZipName" and extract "nofitem" file from filelist to "DestFolder".
If "movetotrash" is true, extracted file write in specian array and it could be deleted automatically using "DeleteTemp" procedure.
Returns name of extracted file.

Public Sub AppendToZip(ByVal ZipName As String, ByVal FileName As String, ByVal deletefile As Boolean)
Try to open *.zip file "ZipName" (create if not exists) as folder. Move (if "deletefile" is true) or copy choosed file "FileName" to this folder.

Public Sub DeleteTemp()
After using some extracted or writed files we can delete it.
