on run argv
        set defaultWorkbookPath to "/Users/Rakesh.Patel/Documents/__code/git/rakeshpatelfa/fa_workday_etl/implementations/elt_transform_excel/dist/Runner.xlsm"
	set defaultMacroName to "modDispatcher.RunSelectedTransform"
	set workbookPath to defaultWorkbookPath
	set macroName to defaultMacroName
	set workbookName to my fileNameFromPosixPath(workbookPath)
	set macroCandidates to {}
	set macroRan to false
	
	if (count of argv) >= 1 then
		set workbookPath to item 1 of argv
		set workbookName to my fileNameFromPosixPath(workbookPath)
	end if
	if (count of argv) >= 2 then
		set macroName to item 2 of argv
	end if
	
	set end of macroCandidates to macroName
	set end of macroCandidates to workbookName & "!" & macroName
	
	tell application "Microsoft Excel"
		activate
		set display alerts to false
		set visible to true
		
		set targetWorkbook to missing value
		repeat with wb in workbooks
			if (name of wb) is workbookName then
				set targetWorkbook to wb
				exit repeat
			end if
		end repeat
		
		if targetWorkbook is missing value then
			set targetWorkbook to open (POSIX file workbookPath)
		end if
		
		delay 1
		
		repeat with candidateMacro in macroCandidates
			try
				run VB macro (contents of candidateMacro)
				set macroRan to true
				exit repeat
			end try
		end repeat
		
		if macroRan is false then
			error "Cannot run VB macro. Tried: " & macroCandidates as string
		end if
		
		save targetWorkbook
	end tell
	
	return "Ran " & macroName & " against " & workbookPath
end run

on fileNameFromPosixPath(filePath)
	set AppleScript's text item delimiters to "/"
	set pathParts to text items of filePath
	set AppleScript's text item delimiters to ""
	return item -1 of pathParts
end fileNameFromPosixPath
