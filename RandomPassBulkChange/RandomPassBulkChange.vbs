' Letters is a string that we will be pulling characters out of when we generate our random password.
Const LETTERS = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()"



'************************************************************
'
' getRandomPass function will generate a random password and return it.
'
'************************************************************
Function getRandomPass
	Dim password, passwordLen, min, max, randomNumber
	passwordLen = 8								  ' Password length is currently set at 8. Adjust to your needs.
	max=len(LETTERS)
	min=1

	strLen=len(LETTERS)
    	For i = 1 to passwordLen
		Randomize
		randomNumber = Int((max-min+1)*Rnd+min)
        	password = password & Mid(LETTERS, RandomNumber, 1)
   	Next
	getRandomPass = password
End Function



'************************************************************
'
' Block below changes passwords and creates our .csv file.
'
'************************************************************
Dim goodUser
goodUser = "Administrator"							' goodUser will be the account you don't want to change or who you currently are. Adjust to your needs.
set fs=CreateObject("Scripting.FileSystemObject")
set objFile = fs.CreateTextFile("c:\changedAccounts.csv", True)			' File path and creation to .csv file. Adjust to your needs.
set objComputer = GetObject("WinNT://.")
objComputer.Filter = ARRAY("user")

For each objUser In objComputer
		If NOT (goodUser = objUser.Name) Then
			newPassword = getRandomPass				' Changing password here.	
			objUser.SetPassword newPassword 
			objUser.setInfo
			objFile.Write objUser.Name & "," & newPassword & vbCrLf	' Prints user and new password in csv file.
		End If
Next

objFile.close
