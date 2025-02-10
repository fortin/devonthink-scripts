-- Import selected Mail messages to DEVONthink with message link.
-- Created by Christian Grunenberg on Mon Apr 19 2004.
-- Modified by BLUEFROG/Jim Neumann Tue Jan 24 2023
-- Copyright (c) 2004-2023. All rights reserved.

(* Uses the email's message ID as the imported file's URL instead of mailto: *)

-- this string is used when the message subject is empty
property pNoSubjectString : "(no subject)"

tell application "Mail"
	try
		tell application id "DNtp"
			if not (exists current database) then error "No database is in use."
			set theGroup to preferred import destination
		end tell
		set theSelection to the selection
		if the length of theSelection is less than 1 then error "One or more messages must be selected."
		repeat with theMessage in theSelection
			my importMessage(theMessage, theGroup)
		end repeat
	on error error_message number error_number
		if error_number is not -128 then display alert "Mail" message error_message as warning
	end try
end tell

on importMessage(theMessage, theGroup)
	tell application "Mail"
		try
			tell theMessage
				set {theDateReceived, theDateSent, theSender, theSubject, theSource, theReadFlag, theID} to {the date received, the date sent, the sender, subject, the source, the read status, message id}
			end tell
			set msgID to "message://%3c" & theID & "%3e"
			if theSubject is equal to "" then set theSubject to pNoSubjectString
			tell application id "DNtp"
				set newRecord to (create record with {name:theSubject & ".eml", type:unknown, creation date:theDateSent, modification date:theDateReceived, source:(theSource as string), unread:(not theReadFlag)} in theGroup)
				set URL of newRecord to msgID -- Set the URL here
			end tell
		on error error_message number error_number
			if error_number is not -128 then display alert "Mail" message error_message as warning
		end try
	end tell
end importMessage


