; JIRA new issue link from TAB delimited txt file

#include <IE.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <Array.au3>
#include <File.au3>
#include <TrayConstants.au3>


Opt("WinTitleMatchMode", 2)


;;;;;;;;;;;;;;  Read data from file (tab delimited)

Global $aInFile
_FileReadToArray(@ScriptDir & "\new_issue_links.txt", $aInFile)

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

;Exit

;;;;;;;;;;  attach to existing IE window

;ConsoleWrite("hello..." & @CRLF)

;Local $oIE = _IECreate("https://intranet.ersteopen.net/Portal.Node/portal", 0, 1, 1, 1)
;_IELoadWait($oIE)
Sleep(2000)
$oIE = _IEAttach("Open Network")

;;;;;;;;;;;;;;  Submit data to JIRA


Local $iRows = UBound($aResult, $UBOUND_ROWS)
For $i = 0 + 2 To $iRows - 1

	TrayTip("JIRA", "creating issue link (" & $i - 1 & "/" & $iRows - 2 & ")", 0, $TIP_ICONASTERISK)

	$issue1 = _URIEncode($aResult[$i][0])
	$issue2 = _URIEncode($aResult[$i][1])
	$linkdesc = _URIEncode($aResult[$i][2])   ; relates to, blocks, is blocked by, depends on, contains, is child of, tests, is tested by
	$comment = _URIEncode($aResult[$i][3])


	_IENavigate($oIE, "https://jira.s-mxs.net/secure/LinkJiraIssue!default.jspa?key=" & $issue1 & "&issueKeys=" & $issue2 & "&linkDesc=" & $linkdesc & "&comment=" & $comment & "")

	Sleep(2000)
	_IELoadWait($oIE)

	;Exit

	Local $oForm = _IEFormGetCollection($oIE, 1) ; 2nd form on the page
	_IEFormSubmit($oForm)
	Sleep(2000)
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



