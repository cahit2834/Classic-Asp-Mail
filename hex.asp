<%Private Function hexValue(ByVal intHexLength)

	Dim intLoopCounter
	Dim strHexValue

	'Randomise the system timer
	Randomize Timer()

	'Generate a hex value
	For intLoopCounter = 1 to intHexLength

		'Genreate a radom decimal value form 0 to 15
		intHexLength = CInt(Rnd * 1000) Mod 16

		'Turn the number into a hex value
		Select Case intHexLength
			Case 1
				strHexValue = "1"
			Case 2
				strHexValue = "2"
			Case 3
				strHexValue = "3"
			Case 4
				strHexValue = "4"
			Case 5
				strHexValue = "5"
			Case 6
				strHexValue = "6"
			Case 7
				strHexValue = "7"
			Case 8
				strHexValue = "8"
			Case 9
				strHexValue = "9"
			Case 10
				strHexValue = "A"
			Case 11
				strHexValue = "B"
			Case 12
				strHexValue = "C"
			Case 13
				strHexValue = "D"
			Case 14
				strHexValue = "E"
			Case 15
				strHexValue = "H"
			Case Else
				strHexValue = "Z"
		End Select

		'Place the hex value into the return string
		hexValue = hexValue & strHexValue
	Next
End Function


%>
