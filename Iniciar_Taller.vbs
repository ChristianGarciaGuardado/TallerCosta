' Script para iniciar Costa Taller Mecánico sin mostrar ventana CMD
' Creado para: Lucas Costa y Leonardo Costa

Set oShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")

' Ruta de la aplicación
strPath = "C:\TallerCosta"

' Verificar que la carpeta existe
If Not oFSO.FolderExists(strPath) Then
    MsgBox "No se encontró la carpeta C:\TallerCosta." & vbCrLf & _
           "Verificá que la aplicación esté instalada correctamente.", _
           vbCritical, "Costa Taller Mecánico"
    WScript.Quit
End If

' Esperar 1 segundo y abrir el navegador
WScript.Sleep 1500
oShell.Run "http://192.168.0.12:8080", 1, False

' Ejecutar la app Flask en segundo plano (sin ventana CMD visible)
' El 0 al final = ventana oculta
oShell.Run "cmd /c cd C:\TallerCosta && C:\TallerCosta\venv\Scripts\activate.bat && python app.py", 0, False

