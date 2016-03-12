Dim MyExcel
Set MyExcel = GetObject("\\C:\VBSTest.xls")

strComputer = "."
'// Note 4 forward slashes!
strDirToMonitor = "::{031E4825-7B94-4DC3-B131-E946B44C8DD5}"
'// Monitor Above every 10 secs...
strTime = "10"

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colMonitoredEvents = objWMIService.ExecNotificationQuery _
    ("SELECT * FROM __InstanceOperationEvent WITHIN " & strTime & " WHERE " _
        & "Targetinstance ISA 'CIM_DirectoryContainsFile' and " _
            & "TargetInstance.GroupComponent= " _
                & "'Win32_Directory.Name=" & Chr(34) & strDirToMonitor & Chr(34) & "'")
 

Do While True
    Set objEventObject = colMonitoredEvents.NextEvent()

    Select Case objEventObject.Path_.Class
        Case "__InstanceCreationEvent"
            MsgBox "A new file was just created: " & _
                objEventObject.TargetInstance.PartComponent
            
            MyFile = StrReverse(objEventObject.TargetInstance.PartComponent)
            '// Get the string to the left of the first \ and reverse it
            MyFile = (StrReverse(Left(MyFile, InStr(MyFile, "\") - 1)))
            MyFile = Mid(MyFile, 1, Len(MyFile) - 1)
            With MyExcel.Worksheets(1)
                 .Range("A65536").End(-4162).Offset(1, 0).Value = MyFile
            End With
            Exit Do
        Case "__InstanceDeletionEvent"
            MsgBox "A file was just deleted: " & _
                objEventObject.TargetInstance.PartComponent
            MyFile = StrReverse(objEventObject.TargetInstance.PartComponent)
            '// Get the string to the left of the first \ and reverse it
            MyFile = (StrReverse(Left(MyFile, InStr(MyFile, "\") - 1)))
            MyFile = Mid(MyFile, 1, Len(MyFile) - 1)
            With MyExcel.Worksheets(1)
                 .Range("A65536").End(-4162).Offset(1, 0).Value = MyFile
            End With
            Exit Do
        Case "__InstanceModificationEvent"
            MsgBox "A file was just modified: " & _
                objEventObject.TargetInstance.PartComponent
            Exit Do
    End Select
Loop