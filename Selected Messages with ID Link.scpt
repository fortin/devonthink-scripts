-- Import selected Mail messages to DEVONthink with message link and user-specified parameters.
-- Modified to include a user dialog for selecting the database, folder, and tags.

property pNoSubjectString : "(no subject)"

tell application "Mail"
	try
		-- Get the selection
		set theSelection to the selection
		if the length of theSelection is less than 1 then error "One or more messages must be selected."

		-- Get DEVONthink databases and groups
		tell application id "DNtp"
			set dbList to name of every database
			if (count of dbList) is 0 then error "No DEVONthink databases available."
			
			-- Ask user to select a database
			set dbChoice to choose from list dbList with prompt "Select a DEVONthink Database:" default items {item 1 of dbList}
			if dbChoice is false then return
			set theDatabase to database named (item 1 of dbChoice)

			-- Get groups from the selected database
			set groupList to name of every record in theDatabase whose type is group
			if (count of groupList) is 0 then error "No folders available in the selected database."
			
			-- Ask user to select a folder
			set groupChoice to choose from list groupList with prompt "Select a Folder in the Database:" default items {item 1 of groupList}
			if groupChoice is false then return
			set theGroup to (get record at (item 1 of groupChoice) in theDatabase)

			-- Ask user for tags
			set tagInput to text returned of (display dialog "Enter tags (comma-separated):" default answer "")

			-- Convert tag input to a list
			set AppleScript's text item delimiters to ","
			set tagList to text items of tagInput
			set AppleScript's text item delimiters to ""

			-- Process each selected message
			repeat with theMessage in theSelection
				my importMessage(theMessage, theGroup, tagList)
			end repeat
		end tell
	on error error_message number error_number
		if error_number is not -128 then display alert "Mail" message error_message as warning
	end try
end tell

on importMessage(theMessage, theGroup, tagList)
	tell application "Mail"
		try
			tell theMessage
				set {theDateReceived, theDateSent, theSender, theSubject, theSource, theReadFlag, theID} to {the date received, the date sent, the sender, subject, the source, the read status, message id}
			end tell
			set msgID to "message://%3c" & theID & "%3e"
			if theSubject is equal to "" then set theSubject to pNoSubjectString

			-- Import message into DEVONthink
			tell application id "DNtp"
				set newRecord to (create record with {name:theSubject & ".eml", type:unknown, creation date:theDateSent, modification date:theDateReceived, source:(theSource as string), unread:(not theReadFlag)} in theGroup)
				set URL of newRecord to msgID
				set tags of newRecord to tagList
			end tell
		on error error_message number error_number
			if error_number is not -128 then display alert "Mail" message error_message as warning
		end try
	end tell
end importMessage