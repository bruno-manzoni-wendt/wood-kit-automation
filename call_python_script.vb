Sub call_python()

    Dim objShell As Object
    Dim PythonExe, PythonScript As String

    Set objShell = VBA.CreateObject("Wscript.Shell")
    
    PythonExe = """User\AppData\Local\Programs\Python\Python312\python.exe"""
    PythonScript = """\path\to\python\script.py"""
    
    objShell.Run PythonExe & PythonScript