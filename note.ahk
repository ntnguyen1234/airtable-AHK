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
  clipboard := Trim(clipboard, " `t")

  ParamsFile := A_Desktop . "/autohotkey/parameters.json"
  FileRead, params, %ParamsFile%
  obj := JSON.load(params)

  Gui, Destroy
  
  Gui, Font, s14 q5, Comic Sans MS
  Gui, Add, Text, x10, Project
  Gui, Add, Text, x+160, Table

  ProjectDropList := ""
  For Project, BaseAPI in obj.projects
  {
    If (obj.defaultProject = Project)
    {
      ProjectDropList := ProjectDropList Project "||"
    }
    Else
    {
      ProjectDropList := ProjectDropList Project "|"
    }
  }
  Gui, Add, DropDownList, x10 vProject, %ProjectDropList%
  
  TableDropList := ""
  For i, Table in obj.tables
  {
    If (obj.defaultTable = Table)
    {
      TableDropList := TableDropList Table "||"
    }
    Else
    {
      TableDropList := TableDropList Table "|"
    }
  }
  Gui, Add, DropDownList, x+10 vTable, %TableDropList%

  Gui, Add, Edit, x10 y+20 vContentField, % obj.content
  Gui, Add, Edit, r10 vContent w480, % clipboard

  Gui, Add, Edit, y+20 vNoteField, % obj.note
  Gui, Add, Edit, x+10 w400 r1 vNote

  Gui, Add, Button, Default w80 gOK, OK
  Gui, Add, Button, x+5 w80 gCancel, Cancel
  Gui, Show
  GuiControl, Focus, Note
  Return

  OK: 
  {
    Gui, Submit
    For P, BaseAPI in obj.projects
    {
      if (Project = P)
      {
        API_Base := BaseAPI
      }
    }
    StringReplace Content, Content, \, /, All
    Data = {"fields":{"%ContentField%": "%Content%", "%NoteField%": "%Note%"}}

    Sleep 1000
    WinHttp := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    WinHttp.Open("POST", "https://api.airtable.com/v0/" API_Base "/" Table, false)
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