-- Open a .docx in Microsoft Word for Mac, refresh its tables of contents and save a PDF.
-- The input .docx is closed without saving; other open documents are left alone.
--
-- Usage: osascript word_pdf.applescript /abs/path/in.docx /abs/path/out.pdf
on run argv
	set src to item 1 of argv
	set dst to item 2 of argv
	set fname to do shell script "basename " & quoted form of src
	with timeout of 600 seconds
		tell application "Microsoft Word"
			open file name src
			-- a large file can still be loading when open returns
			repeat 120 times
				if exists document fname then exit repeat
				delay 1
			end repeat
			set d to document fname
			set n to count of tables of contents of d
			repeat with i from 1 to n
				update (table of contents i of d)
			end repeat
			save as d file name dst file format format PDF
			close d saving no
		end tell
	end timeout
end run
