' ExcelMerge 실행 런처 (터미널 창 없이 조용히 실행)
' - GUI 앱이므로 콘솔이 필요 없다 → pythonw(창 없는 파이썬)로 실행.
' - .bat 은 구조상 콘솔 창이 잠깐이라도 뜨지만, 이 .vbs 는 창을 전혀 띄우지 않는다.
' - 사용법: 이 파일(run.vbs)을 더블클릭. (오류 로그가 필요하면 run.bat 로 실행)
Option Explicit
Dim shell, fso, here, script, launched
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

here = fso.GetParentFolderName(WScript.ScriptFullName)
shell.CurrentDirectory = here
script = """" & here & "\excel_diff_merge.py"""

launched = False
On Error Resume Next
' 1순위: pythonw (PATH 의 파이썬), 2순위: pyw (py 런처의 창 없는 변형)
Err.Clear
shell.Run "pythonw " & script, 0, False   ' 0 = 창 숨김, False = 종료 대기 안 함
If Err.Number = 0 Then launched = True

If Not launched Then
    Err.Clear
    shell.Run "pyw " & script, 0, False
    If Err.Number = 0 Then launched = True
End If
On Error GoTo 0

If Not launched Then
    MsgBox "Python(pythonw)을 찾지 못했습니다. Python 설치 여부와 PATH 를 확인해 주세요.", _
           vbExclamation, "ExcelMerge"
End If
