on open location this_URL
	set the messageId to text 11 thru -1 of this_URL
	tell application "Microsoft Outlook"
		activate
		open message id messageId
	end tell
end open location
