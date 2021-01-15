Dim sInput
height = InputBox("Aka je vasa vyska? (v centimetroch) (zadajte iba cisla)", "Najviac Smart Program v Historii")
If IsEmpty(height) Then
	MsgBox "Pridite aj nabuduce :3", 0, "Dovidenia!"
Else
	If IsNumeric(height) Then
		MsgBox "Ste " & height & " cm vysoky!", 0, "Vypocet"
	Else
		MsgBox "Neviem sice tvoju vysku ale viem ze si mentalne retardovany.", 0, "Vypocet"
	End If
End If