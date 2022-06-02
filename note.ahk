#NoEnv
#SingleInstance force
#Include JSON.ahk

; Gui
<#2::
  clipboard := ""
  Send ^c
  ClipWait  ; Wait for the clipboard to contain text.
  StringReplace clipboard, clipboard, %A_Space% `r`n, %A_Space%, All
  StringReplace clipboard, clipboard, `r`n, %A_Space%, All

  ParamsFile := A_Desktop . "/autohotkey/parameters.json"
  FileRead, params, %ParamsFile%
  obj := JSON.load(params)

  Gui, Destroy
  Gui, Add, Radio, vEmmay Checked, Emmay
  Gui, Add, Radio, x+10 vKat, Kat Media

  Gui, Add, Text, x10 y+20, Table Name:
  Gui, Add, Edit, x+10 y+-15 vTableName, % obj.tablename

  Gui, Add, Edit, x10 y+20 vContentField, % obj.content
  Gui, Add, Edit, r4 vContent w360, % clipboard

  Gui, Add, Edit, y+20 vNoteField, % obj.note
  Gui, Add, Edit, x+10 w330 r1 vNote

  Gui, Add, Button, Default w80 gOK, OK
  Gui, Add, Button, x+5 w80 gCancel, Cancel
  Gui, Show
  GuiControl, Focus, Note
  Return

  OK: 
  {
    Gui, Submit
    
    StringReplace Content, Content, \, /, All
    Data = {"fields":{"%ContentField%": "%Content%", "%NoteField%": "%Note%"}}

    Sleep 1000
    WinHttp := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    WinHttp.Open("POST", "https://api.airtable.com/v0/" obj.baseapp "/" TableName, false)
    WinHttp.SetRequestHeader("Content-Type", "application/json")
    WinHttp.SetRequestHeader("Authorization", "Bearer " obj.key)
    WinHttp.Send(Data)
    ; MsgBox, % WinHttp.ResponseText
    Return
  }
  Cancel:
  {
    GuiClose:
    GuiEscape:
      Gui, Destroy
    Return
  }