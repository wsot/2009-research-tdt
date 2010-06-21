Option Explicit
Const ORDER_RANDOM = 1
Const ORDER_SEQUENCED = 2

Function mapGenerateArrayFromInterface(outputPath as String, freqOutputFilename as String, ampOutputFilename as String, generateDatedFiles as Boolean) as Variant

	Dim seqOrder As Integer
	seqOrder = CInt(Read(">seqOrder"))

	Dim freqMin As Long, freqMax As Long, freqStep As Long
	Dim ampMin As Integer, ampMax As Integer, ampStep As Integer
	Dim reps As Integer
	freqMin = CLng(Read(">FreqMin"))
	freqMax = CLng(Read(">FreqMax"))
	freqStep = CLng(Read(">FreqStep"))

	If freqMax < freqMin Then
		freqMin = freqMax
		freqMax = CLng(Read(">FreqMin"))
		Write_(">FreqMin", freqMin)
		Write_(">FreqMax", freqMax)
	End If

	ampMin = CInt(Read(">AmpMin"))
	ampMax = CInt(Read(">AmpMax"))
	ampStep = CInt(Read(">AmpStep"))

	If ampMax < ampMin Then
		ampMin = ampMax
		ampMax = CLng(Read(">AmpMin"))
		Write_(">AmpMin", ampMin)
		Write_(">AmpMax", ampMax)
	End If

	reps = CInt(Read(">Reps"))
	
	mapGenerateArrayFromInterface = mapGenerateArrays(seqOrder, freqMin, freqMax, freqStep, ampMin, ampMax, ampStep, reps)
	mapWriteToFile(mapGenerateArrayFromInterface, outputPath, freqOutputFilename, ampOutputFilename, generateDatedFiles)

End Function

Function mapGenerateArrays(seqOrder as Integer, freqMin as Long, freqMax as Long, freqStep as Long, ampMin as Integer, ampMax as Integer, ampStep as Integer, reps as Integer) as Variant
	'if use previous, then load the previous file... not yet implemented
	If seqOrder = 3 Then
		Exit Function
	End If

	Dim pairArr() As Variant
    Dim freq As Long 'frequency currently being processed
    Dim amp As Integer 'amp currently being processed
    Dim rep As Integer 'rep currently being processed
    
    'variables used for 'randomisation' - used to swap around variables
    Dim swapWith As Long
    Dim swapVar As Variant

	'stores the number of frequency and amplitude steps; used for iteration in building the arrays
	Dim freqSteps As Integer
	Dim ampSteps As Integer

	'calculate number of frequencies
	freqSteps = Int((freqMax - freqMin) / freqStep)
	If (freqMax - freqMin) Mod freqStep = 0 Then
		freqSteps = freqSteps + 1
	End If

	'calculate number of amplitudes
	ampSteps = Int((ampMax - ampMin) / ampStep)
	If (ampMax - ampMin) Mod ampStep = 0 Then
		ampSteps = ampSteps + 1
	End If

	'calculate the total number of frequency/amplitude pairs
	Dim entryCount As Long, currentEntry As Long
	entryCount = freqSteps * ampSteps * reps

	'pairArr holds a list of all frequency/amplitude pairs with comma separation
    ReDim pairArr(0 To entryCount - 1)
    currentEntry = 0

	'build list of all frequency/amplitude pairs, comma separated
    For freq = freqMin To freqMax Step freqStep
    	For amp = ampMin To ampMax Step ampStep
			For rep = 1 To reps
				pairArr(currentEntry) = freq & "," & amp
				currentEntry = currentEntry + 1
			Next
		Next
    Next

	If seqOrder = ORDER_RANDOM Then 'if the order is 'randomised', then randomise the order of the frequency/intensity pairs - ideally 'pseudorandomisation' should be employed to ensure that the same frequency is not repeated twice in a row - this is not currently done
	    For currentEntry = 0 To entryCount - 1
	        swapWith = Int(Rnd() * entryCount)
	        While swapWith = entryCount
	        	swapWith = Int(Rnd() * entryCount)
	        Wend
	        swapVar = pairArr(currentEntry)
	        pairArr(currentEntry) = pairArr(swapWith)
	        pairArr(swapWith) = swapVar
	    Next
	End If

	mapGenerateArrays = pairArr

End Function

Function mapWriteToFile(pairArr as Variant, outputPath as String, freqOutputFilename as String, ampOutputFilename as String, generateDatedFiles as Boolean)

	Dim objFSO As Object
	Dim objFreqTS As Object, objAmpTS As Object

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFreqTS = objFSO.CreateTextFile(outputPath & freqOutputFilename & ".txt", True)
	Set objAmpTS = objFSO.CreateTextFile(outputPath & ampOutputFilename & ".txt", True)

	Dim currentEntry As Long

	Dim swapWith As Long
    Dim pairVal As String 'holds string value of frequency/amplitude pair
    Dim pairSplit As Variant 'holds split array of frequency/amplitude pair during processing
    Dim freq As Long 'frequency currently being processed
    Dim amp As Integer 'amp currently being processed
    Dim rep As Integer 'rep currently being processed

	'iterate through each string freq/amp pair, split and write to file
    For currentEntry = 0 To UBound(pairArr)
    	pairVal = pairArr(currentEntry)
    	pairSplit = Split(pairVal, ",") 'split freq/amp string pair
		freq = CLng(pairSplit(0)) 'get freq
		objFreqTS.WriteLine(freq) 'write to frequency file
    	amp = CInt(pairSplit(1)) 'get amp from split
    	objAmpTS.WriteLine(amp) 'write to amplitude file
    Next

	'close files
	objFreqTS.close
	objAmpTS.close

	If generateDatedFiles Then
		Dim strTimestamp As String
		strTimestamp = Year(Now()) & Month(Now()) & Day(Now()) & "_" & Hour(Now()) & Minute(Now()) & Second(Now())
		objFSO.CopyFile(outputPath & freqOutputFilename & ".txt",  outputPath & strTimestamp & "_" & freqOutputFilename & ".txt")
		objFSO.CopyFile(outputPath & ampOutputFilename & ".txt", outputPath & strTimestamp & "_" & ampOutputFilename & ".txt")
	End If


	Set objFSO = Nothing
End Function

Function mapReadFromFile(freqInputFilename as String, ampInputFilename as String)

		'initially dim pair array as 616 items; enough for 10-70dB at 10dB steps, and 1-88kHz in 1kHz step
		Dim upperBound as Integer
		upperBound = 615
		Dim pairArr() as String
		ReDim pairArr(upperBound)
		
		'intCount is used to track the actual number of items (so list can be resized down if required at the end)
		Dim intCount as Integer
		intCount = 0

		'create file system object to open files
		Dim objFSO As Object
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		
		'open frequency list and get text stream for reading
		Dim objFreqFile As Object, objFreqTS As Object
		Set objFreqFile = objFSO.GetFile(freqInputFilename)
		Set objFreqTS = objFreqFile.OpenAsTextStream
		
		'open amplitude list and get text stream for reading
		Dim objAmpFile As Object, objAmpTS As Object
		Set objAmpFile = objFSO.GetFile(ampInputFilename)
		Set objAmpTS = objAmpFile.OpenAsTextStream
				
		While Not objFreqTS.AtEndOfStream
			pairArr(intCount) = objFreqTS.ReadLine & "," & objAmpTS.ReadLine
			intCount = intCount + 1
			If intCount = (upperBound + 1) Then 'check if the pair array is now full, and needs expanding
				upperBound = upperBound + 100
				ReDim Preserve pairArr(upperBound)
			End If
		Wend
		
		objFreqTS.Close
		objAmpTS.Close
		
		ReDim Preserve pairArr(intCount - 1)
		
		mapReadFromFile = pairArr
		
End Function