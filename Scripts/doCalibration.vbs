Option Explicit

Sub doCalibration(ByRef Attarr1, ByRef Attarr2 , ByRef prevCalibFile, objCalibFile as String, objToneVoltage as String, strCalibFilePath as String)

'	Dim objCalibFile As String
'	objCalibFile = ">CalibrationFile.Value"

'	Dim objToneVoltage As String
'	objToneVoltage = "AStim.ToneVoltage"

	If CInt(Read(objCalibFile)) <> CInt(prevCalibFile) Then
		'%%%%%% Enter paths for calibration files for Speaker 1 & 2 %%%%%%%%%
		Dim CalibFile As Integer
		Dim CalibFileName As String
		CalibFile = CInt(Read(objCalibFile))

		Select Case CalibFile
			Case 0
			   	CalibFileName = strCalibFilePath & "20100301 Speaker 3585 straight - 1-88kHz out of frame - 9v fft8192.txt"
				Write_(objToneVoltage,9)
			Case 1
			   	CalibFileName = strCalibFilePath & "20100301 EarBar 2 - speaker 3585 - 1-88kHz out of frame - 9v fft8192.txt"
				Write_(objToneVoltage,9)
			Case 2
			   	CalibFileName = strCalibFilePath & "20100301 EarBar 1 - speaker 3585 - 1-88kHz out of frame - 9v fft8192.txt"
				Write_(objToneVoltage,9)


		End Select
		prevCalibFile = CalibFile

		Dim oFFS1 As Object, oFFile1 As Object, ts1 As Object
		Set oFFS1 = CreateObject("Scripting.FileSystemObject")
		Set oFFile1 = oFFS1.GetFile(CalibFileName)
		Set ts1 = oFFile1.OpenAsTextStream

		Dim oFFS2 As Object, oFFile2 As Object, ts2 As Object
		Set oFFS2 = CreateObject("Scripting.FileSystemObject")
		Set oFFile2 = oFFS2.GetFile(CalibFileName)
		Set ts2 = oFFile2.OpenAsTextStream

		'%%%%%% Initialize and load calibration arrays %%%%%%%%%
		Dim count As Long
		count = 0
		While Not ts1.AtEndOfStream
			Attarr1(count) = ts1.ReadLine
			count = count + 1
		Wend
		ts1.Close

		count = 0
		While Not ts2.AtEndOfStream
			Attarr2(count) = ts2.ReadLine
			count = count + 1
		Wend
		ts2.Close
		Println("New calibration loaded")
	End If
End Sub

Function calcAttenuation(currFrequency as Long, currAmplitude as Integer, byref arrAtt as Variant) as Double
	Dim attVal as Double
	Dim attLookup as Double

	attLookup = (currFrequency - 1000)/100
	If attLookup > UBound(arrAtt) Then
		attVal = 120
	Else
		attVal = arrAtt(attLookup) - Abs(currAmplitude)
	End If
	If attVal > 120 Then
		attVal = 120
	ElseIf attVal < 0 Then
		attVal = 0
	End If
	
	calcAttenuation = attVal
End Function