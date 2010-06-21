Option Explicit

Function calcTimeRemaining(stimRemaining, stimPeriod) as Variant
	Dim timeRem(2) as Integer
	
	Dim dMsecRemain as Double
	dMsecRemain = stimRemaining * stimPeriod
	timeRem(0) = Int(dMsecRemain / 3600000)
	timeRem(1) = Int((dMsecRemain Mod 3600000)/ 60000)
	timeRem(2) = Int((dMsecRemain Mod 60000)/ 1000)
	
	calcTimeRemaining = timeRem
End Function