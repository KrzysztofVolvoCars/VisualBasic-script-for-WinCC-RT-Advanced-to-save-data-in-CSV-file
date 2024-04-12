Sub VBFunction_1()

'Tip:
' 1. Use the <CTRL+SPACE> or <CTRL+I> shortcut to open a list of all objects and functions
' 2. Write the code using the HMI Runtime object.
'  Example: HmiRuntime.Screens("Screen_1").
' 3. Use the <CTRL+J> shortcut to create an object reference.
'Write the code as of this position:



'===========================================================================================
'init - Rev 1.0 from 2024-04-10
'===========================================================================================
Dim f, fso
Const ForAppending = 8, ForWritting = 2
Dim MyVar
Dim DataToFile
Dim Path
Dim FileName

Dim PCTIME_Year
Dim PCTIME_Month
Dim PCTIME_DayofYear
Dim PCTIME_Day
Dim PCTIME_Week
Dim PCTIME_Hour
Dim PCTIME_Minute

Dim MotorName
Dim Current_DateTime
Dim Current_DateTimeNanoSEC
Dim Current_DateTimeSec
Dim Current_Value
Dim MotorOn_DateTime
Dim MotorOn_DateTimeNanoSEC
Dim MotorOn_DateTimeSec
Dim MotorOn_Value

Dim RealTrigTemp
Dim BoolTrigTemp



'===========================================================================================
'Configuratie
'===========================================================================================
MyVar = Now
DataToFile = ""

PCTIME_Year = DatePart("yyyy",Now)
PCTIME_Month = DatePart("m",Now)
PCTIME_DayofYear = DatePart("y",Now)
PCTIME_Day = DatePart("d",Now)
PCTIME_Week = DatePart("ww",Now)
PCTIME_Hour = DatePart("h",Now)
PCTIME_Minute = DatePart("n",Now)
 
Path = "C:\log\"       'locatie logfile'
FileName = "PLC12122_AllMotors_" & PCTIME_Year & PCTIME_DayofYear & PCTIME_Hour &  ".csv"    'Naam logfile '


'===========================================================================================
'prog
'===========================================================================================
Set fso = CreateObject("scripting.fileSystemObject")

'request data'
MotorName=SmartTags("MotorLoggingDB_Simulation_O_DataLog_MotorName")
Current_DateTime = SmartTags("MotorLoggingDB_Simulation_O_DataLog_DataCurrentLog_DateTime")
Current_Value = SmartTags("MotorLoggingDB_Simulation_O_DataLog_DataCurrentLog_Current")
MotorOn_DateTime = SmartTags("MotorLoggingDB_Simulation_O_DataLog_DataBoolLog_DateTime")
MotorOn_Value = SmartTags("MotorLoggingDB_Simulation_O_DataLog_DataBoolLog_Bool")

RealTrigTemp = SmartTags("MotorLoggingDB_Simulation_O_DataLog_DataCurrentLog_CurrentTriger")
BoolTrigTemp = SmartTags("MotorLoggingDB_Simulation_O_DataLog_DataBoolLog_BoolTriger")

Current_DateTimeNanoSEC= CStr(SmartTags("MotorLoggingDB_Simulation_O_DataLog_DataCurrentLog_DateTime_NANOSECOND")/1000000)
MotorOn_DateTimeNanoSEC= CStr(SmartTags("MotorLoggingDB_Simulation_O_DataLog_DataBoolLog_DateTime_NANOSECOND")/1000000)

' Convert from 1 or 2 digits to 3 digits'
If Len(Current_DateTimeNanoSEC) < 3 Then
   Current_DateTimeNanoSEC = Right("000" & Current_DateTimeNanoSEC, 3)
End If 

If Len(MotorOn_DateTimeNanoSEC) < 3 Then
   MotorOn_DateTimeNanoSEC = Right("000" & MotorOn_DateTimeNanoSEC, 3)
End If  

'Curent'

If RealTrigTemp = True Then
	   		 
   If (fso.FileExists(Path & FileName)) Then
		'file ready for appending'
		Set f = fso.OpenTextFile (Path & FileName, ForAppending, True)
	Else
		'No file > create new'
		Set f = fso.OpenTextFile (Path & FileName, ForWritting, True)
		
		DataToFile = "Motor Name; Date and Time; Value; Descryption;"
		DataToFile = DataToFile & vbCrLf
	f.Write DataToFile
	End If
	' RealTrigTemp = SmartTags("Data_block_1_TriggerReal")
	 
	 DataToFile = "MotorName: " & MotorName &";"& Current_DateTime & "." &Current_DateTimeNanoSEC &";" & Current_Value & ";" & "Current" & ";" 
	 DataToFile = DataToFile & vbCrLf
	 f.Write DataToFile
	  
     f.Close
 
End If

'Motor ON'
If BoolTrigTemp = True Then
	 
	  		 
   If (fso.FileExists(Path & FileName)) Then
		'file ready for appending'
		Set f = fso.OpenTextFile (Path & FileName, ForAppending, True)
	Else
		'No file > create new'
		Set f = fso.OpenTextFile (Path & FileName, ForWritting, True)
		
		DataToFile = "Motor Name; Date and Time; Value; Descryption;"
		DataToFile = DataToFile & vbCrLf
	f.Write DataToFile
	End If
	' RealTrigTemp = SmartTags("Data_block_1_TriggerReal")
	 
	 DataToFile = "MotorName: " & MotorName &";"& MotorOn_DateTime  & "." &MotorOn_DateTimeNanoSEC &";" & MotorOn_Value & ";" & "Motor ON/OFF" & ";" 
	 DataToFile = DataToFile & vbCrLf
	 f.Write DataToFile
	 
     f.Close 
 
End If

' 1. Use the <CTRL+SPACE> or <CTRL+I> shortcut to open a list of all objects and functions
' 2. Write the code using the HMI Runtime object.
'  Example: HmiRuntime.Screens("Screen_1").
' 3. Use the <CTRL+J> shortcut to create an object reference.
'Write the code as of this position:

End Sub
