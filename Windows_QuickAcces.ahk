; #Requires AutoHotkey v2.0
; #SingleInstance force   
; Persistent()
; #Warn All
; Windows_QuickAcces([A_ScriptDir])
Windows_QuickAcces(L_oPath){
  ; From: #Lib
  ; Created By Emmanuel d
  Static S_AddedItems := []
	L_Shell :=  ComObject("Shell.Application")
  For L_i, L_Path in L_oPath {
    If !DirExist(L_Path)
      Throw Error(A_ThisFunc "(1) You need to pass a existing directory: '" L_Path "'")
    Loop Files, L_Path, "D"
      L_Path := A_LoopFileFullPath ; Convert relative path to full
		for e in L_Shell.Namespace("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}").Items() 
			if (e.Path = L_path)                        ; The Directory is already added
				Break                                     ; Goto next Directory
		; pin
    S_AddedItems.Push(L_path)
		L_Shell.Namespace(L_Path).Self.InvokeVerb("pintohome") ; OK
		; Run('*pintohome "' L_path '"')   ; OK
		; L_F := Func("Explorer_UnPinFromQuickAcces").Bind(L_path)
		}
  ; ObjRelease(L_Shell)
  OnExit(UnPin_From_QuickAcces)
  Return
	UnPin_From_QuickAcces(*){
    L_Shell :=  ComObject("Shell.Application")
    For L_i, L_Path in S_AddedItems 
      for e in L_Shell.Namespace("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}").Items()
        if (e.Path = L_Path)
          e.InvokeVerb("unpinfromhome")
    }
	}