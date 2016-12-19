Option Explicit
'************************** CommonCode Library START
' Instructions
'To include this script,
'Execute CreateObject("Scripting.FileSystemObject").openTextFile(sFileFullPath).readAll()|Where sFileFullPath is likely C:\LOGS\
' Open
Dim oCommonCode:set oCommonCode = New zCCzCommonCodeClass
Dim CC:set CC = oCommonCode 'Alias
' CommonCodeClass (intended to be a static class)
Class zCCzCommonCodeClass
	'This class serves as a host to other classes in this library.
	Private LoDebugger,LoEnum
	' Class_Initialize
	Private Sub Class_Initialize()
		set LoDebugger = New zCCzDebugClass
		set LoEnum = New zCCzEnumClass
	end sub
	Public Property Get oDebugger:set oDebugger = LoDebugger:End Property
	Public Property Get D:set D = LoDebugger:End Property 'Alias
	Public Property Get oEnum:oEnum = LoEnum:End Property
	Public Property Get E:set E = LoEnum:End Property 'Alias
	'	*** MISC
	Public Function IsStringable(vVariable)
		On Error Resume Next
		CStr vVariable
		IsStringable = not CBool(Err.Number)
	end Function
	Public Function IsEqualV2(ByVal vVar1,ByVal vVar2)
		'This function's definition of Equal is broad.
		Dim i
		IsEqualV2 = False
		if CC.VarTypeV2(vVar1)=CC.E.eString or CC.VarTypeV2(vVar2)=CC.E.eString or CC.VarTypeV2(vVar1)=CC.E.eRegExpMatch or CC.VarTypeV2(vVar2)=CC.E.eRegExpMatch then
			if not (CC.IsStringable(vVar1) and CC.IsStringable(vVar2)) then Exit Function
			IsEqualV2 = not CBool(StrComp(UCase(CStr(vVar1)),UCase(CStr(vVar2))))
		elseif CC.VarTypeV2(vVar1)=CC.E.eArray or CC.VarTypeV2(vVar2)=CC.E.eArray then
			if not (CC.VarTypeV2(vVar1)=CC.E.eArray and CC.VarTypeV2(vVar2)=CC.E.eArray) then Exit Function
			if CC.Size(vVar1) <> CC.Size(vVar2) then Exit Function
			if CC.Size(vVar1)=0 then IsEqualV2=True:Exit Function
			For i=0 to UBound(vVar1)
				if not CC.IsEqualV2(vVar1(i),vVar2(i)) then Exit Function
			Next
			IsEqualV2 = True
		else
			On Error Resume Next
			IsEqualV2 = CBool(vVar1=vVar2)
		end if
	end Function
	Public Function IsBlankV2(vVariable)
		'This function serves as an exception-handled alternative to IsBlank.
		Select Case VarTypeV2(vVariable)
		Case CC.E.eDictionary
			IsBlankV2 = not CBool(vVariable.Count)
		Case CC.E.eArray
			On Error Resume Next
			UBound vVariable
			IsBlankV2 = CBool(Err.Number)
		Case CC.E.eString
			IsBlankV2 = CBool(vVariable="")
		Case CC.E.eByte,CC.E.eBoolean,CC.E.eInteger,CC.E.eLong,CC.E.eSingle,CC.E.eDouble,CC.E.eCurrency,CC.E.eDecimal,CC.E.eError
			IsBlankV2 = not CBool(vVariable)
		Case CC.E.eTextStream,CC.E.eDate,CC.E.eObject
			IsBlankV2 = False
		Case CC.E.eEmpty,CC.E.eNull,CC.E.eUnknown,CC.E.eNothing
			IsBlankV2 = True
		Case Else
			IsBlankV2 = True
			CC.D.Debug 2,"EXCEPTION\IsBlankV2\Received unhandled type of vVariable"
		end Select
	end Function
	Public Function VarTypeV2(vVariable)
		'This function serves as an enumerated alternative to TypeName, and a more specific alternative to VarType.
		Select Case TypeName(vVariable)
		Case "Dictionary"
			VarTypeV2 = CC.E.eDictionary
		Case "TextStream"
			VarTypeV2 = CC.E.eTextStream
		Case "Variant()"
			VarTypeV2 = CC.E.eArray
		Case "Byte"
			VarTypeV2 = CC.E.eByte
		Case "Boolean"
			VarTypeV2 = CC.E.eBoolean
		Case "Integer"
			VarTypeV2 = CC.E.eInteger
		Case "Long"
			VarTypeV2 = CC.E.eLong
		Case "Single"
			VarTypeV2 = CC.E.eSingle
		Case "Double"
			VarTypeV2 = CC.E.eDouble
		Case "Currency"
			VarTypeV2 = CC.E.eCurrency
		Case "Decimal"
			VarTypeV2 = CC.E.eDecimal
		Case "Date"
			VarTypeV2 = CC.E.eDate
		Case "String"
			VarTypeV2 = CC.E.eString
		Case "Empty"
			VarTypeV2 = CC.E.eEmpty
		Case "Null"
			VarTypeV2 = CC.E.eNull
		Case "Unknown"
			VarTypeV2 = CC.E.eUnknown
		Case "Nothing"
			VarTypeV2 = CC.E.eNothing
		Case "Error"
			VarTypeV2 = CC.E.eError
		Case "Object"
			VarTypeV2 = CC.E.eObject
		Case "IMatchCollection2"
			VarTypeV2 = CC.E.eRegExpMatches
		Case "IMatch2"
			VarTypeV2 = CC.E.eRegExpMatch
		Case Else
			VarTypeV2 = CC.E.eUnknown
			CC.D.Debug 2,"EXCEPTION\TypeV2\Received unhandled TypeName: "&TypeName(vVariable)
		end Select
	end Function
	Public Function BuildPath(ByVal sPath)
		Dim oFSO:set oFSO = CreateObject("Scripting.FileSystemObject")
		BuildPath = False
		' Filter
		if not IsStringable(sPath) or sPath="" then CC.D.Debug 2,"ERROR\BuildPath\Received invalid sPath.":Exit Function
		'
		If oFSO.FolderExists(sPath) then
			BuildPath = True
		elseif BuildPath(oFSO.GetParentFolderName(sPath)) then
			BuildPath = True
			oFSO.CreateFolder sPath
		end If
	End Function
	'	*** MISC
	'	*** CONTAINER
	Public Function HasKey(vKey,cContainer)
		HasKey = false
		Select Case CC.VarTypeV2(cContainer)
		Case CC.E.eArray
			if IsNumeric(vKey) then
				HasKey = CBool(0<=vKey and vKey<=UBound(cContainer))
			end if
		Case CC.E.eDictionary
			HasKey = cContainer.Exists(vKey)
		end Select
	end Function
	Public Function Size(vVariable)
		Size=0
		Select Case CC.VarTypeV2(vVariable)
		Case CC.E.eArray
			On Error Resume Next
			Size = UBound(vVariable)+1
		Case CC.E.eDictionary,CC.E.eRegExpMatches
			Size = vVariable.Count
		Case CC.E.eString
			Size = Len(vVariable)
		end Select
	end Function
	Public Function IsContainer(vVariable)
		Select Case CC.VarTypeV2(vVariable)
		Case CC.E.eDictionary,CC.E.eRegExpMatches,CC.E.eArray
			IsContainer = True
		Case Else
			IsContainer = False
		end Select
	end Function
	Public Function ArAppend(vVariable,ByRef aArray)
		' Filter
		if CC.VarTypeV2(aArray)=CC.E.eEmpty then aArray = Array()
		if not IsArray(aArray) then CC.D.Debug 2,"ERROR\ArAppend\Received invalid aArray.":Exit Function
		' 
		if CC.Size(aArray)=0 then
			ReDim aArray(0)
		else
			ReDim Preserve aArray(UBound(aArray)+1)
		end if
		aArray(UBound(aArray)) = vVariable
		ArAppend = aArray
	end Function
	Public Function ArErase(iIndex,aArray)
		Dim i
		ArErase = aArray
		' Filter
		if not CBool(CC.Size(aArray)) then Exit Function
		if not IsArray(aArray) then CC.D.Debug 2,"ERROR\ArErase\Received invalid aArray.":Exit Function
		if not IsNumeric(iIndex) then CC.D.Debug 2,"ERROR\ArErase\Received invalid iIndex.":Exit Function
		'
		For i=Cint(iIndex) to UBound(aArray)-1
			aArray(i) = aArray(i+1)
		Next
		ReDim Preserve aArray(UBound(aArray)-1)
		ArErase = aArray
	end Function
	Public Function GetFlattenedContainer(ByVal cContainer)
		Dim aReturning()
		' Filter
		if not CC.IsContainer(cContainer) then CC.D.Debug 2,"ERROR\GetFlattenedContainer\Received invalid cContainer.":Exit Function
		'
		GetFlattenedContainerHelper cContainer,aReturning
		GetFlattenedContainer = aReturning
	end Function
	Private Function GetFlattenedContainerHelper(cContainer,aReturning)
		Dim vKey,vValue
		'
		For Each vKey in CC.GetKeys(cContainer):vValue=CC.GetValue(vKey,cContainer)
			if CC.IsContainer(vValue) then 
				GetFlattenedContainerHelper vValue,aReturning
			else
				ArAppend vValue,aReturning
			end if
		Next
	end Function
	Public Function GetRegExpPatternFromContainer(cContainer)
		' Filter
		if not CC.IsContainer(cContainer) then CC.D.Debug 2,"ERROR\GetRegExpPatternFromContainer\Received invalid cContainer.":Exit Function
		'
		GetRegExpPatternFromContainer = "\b("&Join(CC.GetFlattenedContainer(cContainer),"|")&")\b"
	end Function
	Public Function GetKeys(cContainer)
		Dim aKeys,i
		' Filter
		if not CC.IsContainer(cContainer) then CC.D.Debug 2,"ERROR\ToDict\Received invalid cContainer.":Exit Function
		'
		Select Case CC.VarTypeV2(cContainer)
		Case CC.E.eArray,CC.E.eRegExpMatches
			aKeys = Array()
			For i=0 to UBound(cContainer)
				ArAppend i,aKeys
			Next
		Case CC.E.eDictionary
			aKeys = cContainer.Keys
		end Select
		GetKeys = aKeys
	end Function
	Public Function GetValue(vKey,cContainer)
		Dim vValue
		' Filter
		if not CC.IsContainer(cContainer) then CC.D.Debug 2,"ERROR\GetValue\Received invalid cContainer.":Exit Function
		'
		On Error Resume Next
		Select Case CC.VarTypeV2(cContainer)
		Case CC.E.eArray,CC.E.eDictionary
			GetValue = cContainer(vKey)
		Case CC.E.eRegExpMatches
			GetValue = cContainer(vKey).Value
		end Select
	end Function
	Public Function NarrateContainer(cContainer)
		' Filter
		if not CC.IsContainer(cContainer) then CC.D.Debug 2,"ERROR\NarrateContainer\Received invalid cContainer":Exit Function
		'
		NarrateContainer = NarrateContainerHelper(cContainer,0)
	end Function
	Private Function NarrateContainerHelper(cContainer,ByVal iRecursionLevel)
		Dim vKey,vValue,i,oMatch,bDoOnce
		'
		For Each vKey in CC.GetKeys(cContainer) : vValue = CC.GetValue(vKey,cContainer)
			if not CBool(bDoOnce) and not CBool(iRecursionLevel) then
				bDoOnce = vbTrue
			else
				NarrateContainerHelper = NarrateContainerHelper&vbNewLine
			end if
			if CC.IsContainer(vValue) then
				NarrateContainerHelper = NarrateContainerHelper&Indent(iRecursionLevel+1)&vKey&":Container. Size:"&CC.Size(vValue)&" iRecursionLevel:"&iRecursionLevel+1
				NarrateContainerHelper = NarrateContainerHelper&NarrateContainerHelper(vValue,iRecursionLevel+1)
			elseif CC.IsStringable(vValue) then
				NarrateContainerHelper = NarrateContainerHelper&Indent(iRecursionLevel+1)&vKey&":"&vValue
			else
				NarrateContainerHelper = NarrateContainerHelper&Indent(iRecursionLevel+1)&vKey&":EXCEPTION\Unprintable value."
			end if
		Next
	end Function
	Public Function Indent(iIndentLevel)
		Indent = ""
		On Error Resume Next
		Indent = Space(iIndentLevel)
	end Function
	Public Function IsInContainer(vVariable,cContainer)
		Dim vKey,vValue
		IsInContainer = False
		' Filter
		if not CC.IsContainer(cContainer) then Exit Function
		'
		Select Case CC.VarTypeV2(cContainer)
		Case CC.E.eArray,CC.E.eDictionary,CC.E.eRegExpMatches
			For Each vKey in CC.GetKeys(cContainer):vValue=CC.GetValue(vKey,cContainer)
				if CC.IsEqualV2(vValue,vVariable) then
					IsInContainer = True
					Exit For
				end if
			Next
		Case Else
			CC.D.Debug 2,"EXCEPTION\IsInContainer\Received unhandled type of cContainer"
		end Select
	end Function
	Public Function IsInContainerV2(vVariable,cContainer)
		'This version searches every level of the container.
		Dim vKey,vValue
		IsInContainerV2 = False
		' Filter
		if not CC.IsContainer(cContainer) then Exit Function
		'
		Select Case CC.VarTypeV2(cContainer)
		Case CC.E.eArray,CC.E.eDictionary
			For Each vKey in CC.GetKeys(cContainer):vValue=CC.GetValue(vKey,cContainer)
				if CC.IsEqualV2(vValue,vVariable) then
					IsInContainerV2 = True
					Exit For
				elseif CC.IsContainer(vValue) then
					if CC.IsInContainerV2(vVariable,vValue) then
						IsInContainerV2 = True
						Exit For
					end if
				end if
			Next
		Case CC.E.eRegExpMatches
			IsInContainerV2 = IsInContainer(vVariable,cContainer)
		Case Else
			CC.D.Debug 2,"EXCEPTION\IsInContainerV2\Received unhandled type of cContainer"
		end Select
	end Function
	Public Function Find(vVariable,cContainer)
		Dim vKey,vValue
		' Filter
		if not CC.IsContainer(cContainer) then Exit Function
		'
		Select Case CC.VarTypeV2(cContainer)
		Case CC.E.eArray,CC.E.eDictionary,CC.E.eRegExpMatches
			For Each vKey in CC.GetKeys(cContainer):vValue=CC.GetValue(vKey,cContainer)
				if CC.IsEqualV2(vValue,vVariable) then
					Find = vKey
					Exit For
				end if
			Next
		Case Else
			CC.D.Debug 2,"EXCEPTION\Find\Received unhandled type of cContainer"
		end Select
	end Function
	Public Function FindV2(vVariable,cContainer)
		'This version searches every level of the container.
		Dim vKey,vValue
		' Filter
		if not CC.IsContainer(cContainer) then Exit Function
		'
		Select Case CC.VarTypeV2(cContainer)
		Case CC.E.eArray,CC.E.eDictionary
			For Each vKey in CC.GetKeys(cContainer):vValue=CC.GetValue(vKey,cContainer)
				if CC.IsEqualV2(vValue,vVariable) then
					FindV2 = vKey
					Exit For
				elseif CC.IsContainer(vValue) then
					FindV2 = FindV2(vVariable,vValue)
					if not IsEmpty(FindV2) then
						Exit For
					end if
				end if
			Next
		Case CC.E.eRegExpMatches
			FindV2 = CC.Find(vVariable,cContainer)
		Case Else
			CC.D.Debug 2,"EXCEPTION\Find\Received unhandled type of cContainer"
		end Select
	end Function
	Public Function FindBaseKey(vVariable,cContainer)
		'This function searches every level of a container, and returns the key of the first(base) level.
		Dim vKey,vValue
		' Filter
		if not CC.IsContainer(cContainer) then Exit Function
		'
		Select Case CC.VarTypeV2(cContainer)
		Case CC.E.eArray,CC.E.eDictionary,CC.E.eRegExpMatches
			For Each vKey in CC.GetKeys(cContainer):vValue=CC.GetValue(vKey,cContainer)
				if CC.IsEqualV2(vValue,vVariable) then
					FindBaseKey = vKey
					Exit For
				elseif CC.IsContainer(vValue) then
					if CC.IsInContainerV2(vVariable,vValue) then
						FindBaseKey = vKey
						Exit For
					end if
				end if
			Next
		Case Else
			CC.D.Debug 2,"EXCEPTION\FindBaseKey\Received unhandled type of cContainer"
		end Select
	end Function
	'	*** CONTAINER
end Class
' DebugClass
Class zCCzDebugClass
	'Declare
	Public iThreshold '-1:BypassFilter 0:IgnoreEverything 1:Important 2:Limited 3:Success 4:SemiVerbose 5:Verbose 6:Flooding
	Public iIndentLevel
	Private oLogFileHandle,LsLogFullPath
	' Class_Initialize
	Private Sub Class_Initialize()
		' Option Defaults
		iThreshold = 0
		iIndentLevel = 0
		LsLogFullPath = "C:\LOGS\CCDefaultLog.log"
	end sub
	' sLogFullPath
	Public Property Let sLogFullPath(sParam)
		Dim oFSO:set oFSO = CreateObject("Scripting.FileSystemObject")
		' Establish oLogFileHandle,LsLogFullPath
		LsLogFullPath = ""
		set oLogFileHandle = Nothing
		CC.BuildPath oFSO.GetParentFolderName(sParam)
		if not oFSO.FolderExists(sParam) then
			LsLogFullPath=sParam
			set oLogFileHandle = oFSO.OpenTextFile(LsLogFullPath,2,vbTrue)
		end if
	end Property
	Public Property Get sLogFullPath:sLogFullPath=LsLogFullPath:end Property
	' Debug
	Public Sub Debug(iDebugLevel,vVariable)
		Dim sPrintString
		'  Filter
		if not IsNumeric(iDebugLevel) then CC.D.Debug -1,"ERROR\Debug\Received invalid iDebugLevel.":Exit Sub
		if not IsDebugLevelMet(iDebugLevel) then Exit Sub
		'  Interpret vVariable
		if CC.IsContainer(vVariable) then
			sPrintString = "Debug\Narrating Container.."&vbNewLine&CC.NarrateContainer(vVariable)
		elseif not CC.IsStringable(vVariable) then
			sPrintString = "EXCEPTION\Debug\Unable to stringize vVariable"
		else
			sPrintString = CStr(vVariable)
		end if
		'  Indent
		sPrintString = CC.Indent(iIndentLevel)&sPrintString
		sPrintString = Replace(sPrintString,vbNewLine,vbNewLine&CC.Indent(iIndentLevel))
		'  MsgBox
		if iDebugLevel=1 then MsgBox sPrintString
		'  Log
		' TryInit oLogFileHandle via Property sLogFullPath
		if CC.IsBlankV2(oLogFileHandle) then me.sLogFullPath=LsLogFullPath
		if not CC.IsBlankV2(oLogFileHandle) then oLogFileHandle.WriteLine(sPrintString)
	end Sub
	' IsDebugLevelMet
	Public Function IsDebugLevelMet(iDebugLevel)
		' Filter
		if not IsNumeric(iDebugLevel) then CC.D.Debug -1,"ERROR\IsDebugLevelMet\Received invalid iDebugLevel.":Exit Function
		'
		IsDebugLevelMet=CBool(iDebugLevel=-1 or (iThreshold<>0 and iDebugLevel<=iThreshold))
	end Function
	Public Sub ViewLog
		CreateObject("WScript.Shell").Run LsLogFullPath
	end Sub
end Class
' EnumClass (intended to be a static class)
Class zCCzEnumClass
	Private lNextID
	Private LeDefault,LeDictionary,LeTextStream,LeArray,LeByte,LeBoolean,LeInteger,LeLong,LeSingle,LeDouble,LeCurrency,LeDecimal
	Private LeDate,LeString,LeEmpty,LeNull,LeUnknown,LeNothing,LeError,LeObject
	Private LeRegExpMatches,LeRegExpMatch
	' Available Enumerations
	'Misc
	Public Property Get eDefault : eDefault = Enumerate(LeDefault) : End Property
	'Types
	Public Property Get eDictionary : eDictionary = Enumerate(LeDictionary) : End Property
	Public Property Get eTextStream : eTextStream = Enumerate(LeTextStream) : End Property
	Public Property Get eArray : eArray = Enumerate(LeArray) : End Property
	Public Property Get eByte : eByte = Enumerate(LeByte) : End Property
	Public Property Get eBoolean : eBoolean = Enumerate(LeBoolean) : End Property
	Public Property Get eInteger : eInteger = Enumerate(LeInteger) : End Property
	Public Property Get eLong : eLong = Enumerate(LeLong) : End Property
	Public Property Get eSingle : eSingle = Enumerate(LeSingle) : End Property
	Public Property Get eDouble : eDouble = Enumerate(LeDouble) : End Property
	Public Property Get eCurrency : eCurrency = Enumerate(LeCurrency) : End Property
	Public Property Get eDecimal : eDecimal = Enumerate(LeDecimal) : End Property
	Public Property Get eDate : eDate = Enumerate(LeDate) : End Property
	Public Property Get eString : eString = Enumerate(LeString) : End Property
	Public Property Get eEmpty : eEmpty = Enumerate(LeEmpty) : End Property
	Public Property Get eNull : eNull = Enumerate(LeNull) : End Property
	Public Property Get eUnknown : eUnknown = Enumerate(LeUnknown) : End Property
	Public Property Get eNothing : eNothing = Enumerate(LeNothing) : End Property
	Public Property Get eError : eError = Enumerate(LeError) : End Property
	Public Property Get eObject : eObject = Enumerate(LeObject) : End Property
	Public Property Get eRegExpMatches : eRegExpMatches = Enumerate(LeRegExpMatches) : End Property
	Public Property Get eRegExpMatch : eRegExpMatch = Enumerate(LeRegExpMatch) : End Property
	'Template.. (don't forget to declare L~)
	'Public Property Get ~ : ~ = Enumerate(L~) : End Property
	' 
	Private Function Enumerate(ByRef lLocalEnumeration)
		if not IsEmpty(lLocalEnumeration) then
			Enumerate = lLocalEnumeration
			Exit Function
		else
			if IsEmpty(lNextID) then
				lNextID = CLng(1)
			elseif lNextID=2147483647 then
				CC.D.Debug 1,"ERROR\Enumerate\Reached ID Capacity."
			else
				lNextID = lNextID + 1
			end if
			'
			lLocalEnumeration = lNextID
			Enumerate = lNextID
		end if
	end Function
end Class
'************************** CommonCode Library END

CC.D.iThreshold = 5
Test1
Test2
Test4
CC.D.ViewLog

Function Test4
	Dim aArray,vKey,vValue
	aArray = Array("Egg1","Egg2","Egg3","Egg4")
	For Each vKey in CC.GetKeys(aArray) :vValue = CC.GetValue(vKey,aArray)
		CC.D.Debug 5,"vKey:"&vKey&" vValue:"&vValue
	Next
end Function

Function Test2
	Dim aArray1,aArray2,sString
	'
	aArray1 = Array("Egg1","Egg2","Egg3")
	aArray2 = Array("Egg4","Egg5","Egg6")
	CC.ArAppend aArray2,aArray1
	sString = CC.GetRegExpPatternFromContainer(aArray1)
	CC.D.Debug 5,sString
end Function

Function Test1
	Dim d2dDict:set d2dDict = CreateObject("Scripting.Dictionary")
	d2dDict.Add "Key0",Array("egg1","Egg2")
	d2dDict.Add "Key1",Array("egg3","Egg4")

	CC.D.Debug 5,d2dDict
	CC.D.Debug 5,CC.GetRegExpPatternFromContainer(d2dDict)

	Dim aArray
	aArray = Array("z1","z2","z3")
	CC.D.Debug 5,CC.GetRegExpPatternFromContainer(aArray)

	CC.D.Debug 5,CC.IsInContainerV2("EG2",d2dDict)
	CC.D.Debug 5,CC.IsInContainerV2("EGG2",d2dDict)

	CC.D.Debug 5,CC.IsInContainer(Array("egg1","Egg3"),d2dDict)
	CC.D.Debug 5,CC.IsInContainer(Array("egg1","Egg2"),d2dDict)

	CC.D.Debug 5,CC.FindV2("z3",aArray)
	CC.D.Debug 5,CC.FindV2("Egg2",d2dDict)
	CC.D.Debug 5,CC.FindBaseKey("Egg2",d2dDict)
	
	CC.D.Debug 4,1
end Function













































