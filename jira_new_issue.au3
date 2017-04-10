; JIRA new issue from TAB delimited txt file

#include <IE.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <Array.au3>
#include <File.au3>
#include <TrayConstants.au3>


Opt("WinTitleMatchMode", 2)


;;;;;;;;;;;;;;  Read data from file (tab delimited)

Global $aInFile
_FileReadToArray(@ScriptDir & "\new_issues.txt", $aInFile)

;_ArrayDisplay($aInFile, "Array:")

$aDimension = StringSplit($aInFile[1], @TAB)
Global $aOutArray[$aInFile[0] + 1][$aDimension[0]]
For $iIndex1 = 1 To $aInFile[0]
	$aRecord = StringSplit($aInFile[$iIndex1], @TAB)
	For $iIndex2 = 0 To $aRecord[0] - 1
		$aOutArray[$iIndex1][$iIndex2] = $aRecord[$iIndex2 + 1]
	Next
Next
$aOutArray[0][0] = $aInFile[0]
$aOutArray[0][1] = $aDimension[0]


$aResult = $aOutArray
;_ArrayDisplay($aResult, "Array:")



;;;;;;;;;;  attach to existing IE window

;ConsoleWrite("hello..." & @CRLF)

;Local $oIE = _IECreate("https://intranet.ersteopen.net/Portal.Node/portal", 0, 1, 1, 1)
;_IELoadWait($oIE)
Sleep(2000)
$oIE = _IEAttach("Open Network")

;;;;;;;;;;;;;;  Submit data to JIRA


Local $iRows = UBound($aResult, $UBOUND_ROWS)
For $i = 0 + 2 To $iRows - 1

	TrayTip("JIRA", "creating issue (" & $i - 1 & "/" & $iRows - 2 & ")", 0, $TIP_ICONASTERISK)

	$pid = 18541
	$issuetype = 11
	$summary = _URIEncode($aResult[$i][3])
	$description = _URIEncode($aResult[$i][4])
	$components = 17912
	$duedate = "14.04.2017"
	$priority = 3
	$reporter = "H5011OB"
	$lab_requirement = _URIEncode($aResult[$i][15])
	$lab_release = _URIEncode($aResult[$i][17])
	$lab_tasktype = _URIEncode($aResult[$i][16])
	$lab_component = _URIEncode($aResult[$i][14])
	$lab_project = _URIEncode($aResult[$i][8])
	$projectname = _URIEncode($aResult[$i][7])
	$projectnum = _URIEncode($aResult[$i][8])
	$projectmanager = _URIEncode($aResult[$i][12])
	$requirement = _URIEncode($aResult[$i][11])
	$release = _URIEncode($aResult[$i][17])
	$versions = 29132 ; id of 2017-S07

	_IENavigate($oIE, "http://jira.s-mxs.net/secure/CreateIssueDetails!init.jspa?pid=" & $pid & "&issuetype=" & $issuetype & "&summary=" & $summary & "&description=" & $description & "&components=" & $components & "&duedate=" & $duedate & "&priority=" & $priority & "&reporter=" & $reporter & "&labels=" & $lab_requirement & "&labels=" & $lab_release & "&labels=" & $lab_tasktype & "&labels=" & $lab_component & "&labels=" & $lab_project & "&customfield_14834=" & $projectname & "&customfield_14833=" & $projectmanager & "&customfield_14835=" & $projectnum & "&customfield_13291=" & $requirement & "&versions=" & $versions & "")

	Sleep(3000)
	_IELoadWait($oIE)

	Local $oForm = _IEFormGetCollection($oIE, 1) ; 2nd form on the page
	_IEFormSubmit($oForm)
	Sleep(3000)
	_IELoadWait($oIE)


Next



Exit




Func _URIEncode($sData)
	; Prog@ndy
	Local $aData = StringSplit(BinaryToString(StringToBinary($sData, 4), 1), "")
	Local $nChar
	$sData = ""
	For $i = 1 To $aData[0]
		; ConsoleWrite($aData[$i] & @CRLF)
		$nChar = Asc($aData[$i])
		Switch $nChar
			Case 45, 46, 48 To 57, 65 To 90, 95, 97 To 122, 126
				$sData &= $aData[$i]
			Case 32
				$sData &= "+"
			Case Else
				$sData &= "%" & Hex($nChar, 2)
		EndSwitch
	Next
	Return $sData
EndFunc   ;==>_URIEncode

Func _URIDecode($sData)
	; Prog@ndy
	Local $aData = StringSplit(StringReplace($sData, "+", " ", 0, 1), "%")
	$sData = ""
	For $i = 2 To $aData[0]
		$aData[1] &= Chr(Dec(StringLeft($aData[$i], 2))) & StringTrimLeft($aData[$i], 2)
	Next
	Return BinaryToString(StringToBinary($aData[1], 1), 4)
EndFunc   ;==>_URIDecode

;MsgBox(0, '', _URIDecode(_URIEncode("testäöü fv")))



