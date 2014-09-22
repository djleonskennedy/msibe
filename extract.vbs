Dim BinaryName : BinaryName = "<Binary Name>"
Dim ExtractedObject : ExtractedObject = <File Name>

Function ExtractBinary(BinaryName, OutputFile) 
    Const msiReadStreamAnsi = 2 
    Dim oDatabase, View, Record, FSO, Stream, BinaryData 
    Set oDatabase = Session.Database 
    Set View = oDatabase.OpenView("SELECT * FROM Binary WHERE Name = '" & BinaryName & "'") 
    View.Execute 
    Set Record = View.Fetch 
    BinaryData = Record.ReadStream(2, Record.DataSize(2), msiReadStreamAnsi) 
    Set FSO = CreateObject("Scripting.FileSystemObject") 
    Set Stream = FSO.CreateTextFile(OutputFile, True) 
    Stream.Write BinaryData 
    Stream.Close 
    Set FSO = Nothing 
End Function
Dim TempFolder : TempFolder = Session.Property("TempFolder")
ExtractBinary BinaryName,TempFolder + ExtractedObject
