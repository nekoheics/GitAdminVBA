Attribute VB_Name = "ModuleGitFilesContents"
Rem .gitattributes
'# Auto detect text files and perform LF normalization
'* text=auto
'
'*.bas       text    eol=crlf
'*.cls       text    eol=crlf
'*.frm       text    eol=crlf
'*.frx       binary  eol=crlf
'*.dcm       text    eol=crlf
'*.vbaproj   text    eol=crlf
'
'*.wsf       text    eol=crlf
'*.bat       text    eol=crlf
'
'*.cls linguist-language=VBA
'*.dcm linguist-language=VBA
'*.vbaproj linguist-language=INI
'
'# file encording
'*.bas working-tree-encoding=sjis
'*.cls working-tree-encoding=sjis
'*.dcm working-tree-encoding=sjis
'*.frm working-tree-encoding=sjis
'
'*.bas encoding=sjis
'*.cls encoding=sjis
'*.dcm encoding=sjis
'*.frm encoding=sjis
'
'*.bas diff=sjis
'*.cls diff=sjis
'*.dcm diff=sjis
'*.frm diff=sjis

Rem .gitignore
'*.tmp
'*.xl*
'‾$*.xl*
'bin/old
'!bin/*.xl*
'!src/*

Rem settings.json
'{
'  "[markdown]": {
'    "editor.wordWrap": "on",
'    "editor.quickSuggestions": {
'      "comments": "off",
'      "strings": "off",
'      "other": "off"
'    },
'    "files.encoding": "utf8",
'  },
'  "files.encoding": "shiftjis",
'  "files.associations": {
'    "*.bas": "vb",
'    "*.cls": "vb",
'    "*.dcm": "vb",
'    "*.frm": "vb"
'  }
'}
