 WinWaitActive("Authentication Required","","10")
 If WinExists("Authentication Required") Then
 Send("nidhina_nanu{TAB}")
 Send("Password123")
 ;Send("password{Enter}")
 EndIf

