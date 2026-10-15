#!/bin/sh
# Make sandboxed Microsoft Word for Mac ask once for access to a folder (default: your home).
# Word's Grant File Access dialog offers exactly the folder of the file it opens, so a throwaway
# document is placed in that folder, opened and closed, then deleted. Click "Grant Access".
#
# Usage: sh word_grant_access.sh [folder]
set -eu
dir=$(cd "${1:-$HOME}" && pwd)
doc="$dir/docx_fix_grant_access.docx"
trap 'rm -f "$doc"' EXIT
python3 - "$doc" <<'EOF'
import sys, zipfile
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
with zipfile.ZipFile(sys.argv[1], 'w') as z:
    z.writestr('[Content_Types].xml',
               '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
               '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
               '<Default Extension="xml" ContentType="application/xml"/>'
               '<Override PartName="/word/document.xml" '
               'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    z.writestr('_rels/.rels',
               '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
               'Target="word/document.xml"/></Relationships>')
    z.writestr('word/document.xml',
               f'<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="{W}"><w:body><w:p><w:r>'
               '<w:t>docx_fix: grant access</w:t></w:r></w:p></w:body></w:document>')
EOF
osascript - "$doc" <<'EOF'
on run argv
	set src to item 1 of argv
	set fname to do shell script "basename " & quoted form of src
	with timeout of 300 seconds
		tell application "Microsoft Word"
			activate
			open file name src
			close document fname saving no
		end tell
	end timeout
end run
EOF
echo "Word can now open files under $dir"
