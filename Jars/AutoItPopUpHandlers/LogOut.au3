WinWaitActivate("The page at http://ratingsgateway-qa.mhf.mhc says:","","10")

If WinExists("The page at http://ratingsgateway-qa.mhf.mhc says:") Then
 Send("{Enter}")
EndIf