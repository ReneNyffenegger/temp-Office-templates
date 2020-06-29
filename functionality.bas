option explicit

sub addReferenceToPersonalXlsb() ' {
 '
 '  run "personal.xlsb!addReferenceToPersonalXlsb"
 '
 '  application.VBE.activeVBProject.references.AddFromFile "C:\Users\r.nyffenegger\AppData\Roaming\Microsoft\Excel\XLSTART\Personal.xlsb"
    application.VBE.activeVBProject.references.AddFromFile environ$("appdata") & "\Microsoft\Excel\XLSTART\Personal.xlsb"
end sub ' }

sub removeReferenceToPersonalXlsb() ' {
    application.VBE.activeVBProject.references.remove application.VBE.ActiveVBProject.References("Personal_xlsb")
end sub ' }

sub copyCellWithoutNewLine() ' {
 '
 '  This sub copies the value of the currenlty selected cell (activeCell) into
 '  the clipboard WITHOUT also adding the (imho unnecessary) new line.
 '
 '  This sub is triggered by ctrl-q  ( See thisWorkbook.bas )
 '

 '
 '  https://renenyffenegger.ch/notes/development/languages/VBA/Win-API/examples/clipboard/index#vba-winapi-put-text-into-clipboard
 '
 '  https://stackoverflow.com/a/14696083/180275
 '

'   dim dataObj As New MSForms.DataObject

'   DataObj.SetText ActiveCell.Value 'depending what you want, you could also use .Formula here
'   DataObj.PutInClipboard

   dim memory          as long
   dim lockedMemory    as long
   dim text4clipboard  as string

   text4clipboard = activeCell.value

   memory = GlobalAlloc(GHND, len(text4clipboard) + 1)
   if memory = 0 then
      msgBox "GlobalAlloc failed"
      exit sub
   end if

   lockedMemory = GlobalLock(memory)
   if lockedMemory = 0 then
      msgBox "GlobalLock failed"
      exit sub
   end if

   lockedMemory = lstrcpy(lockedMemory, text4clipboard)

   call GlobalUnlock(memory)

   if openClipboard(0) = 0 Then
      msgBox "openClipboard failed"
      exit sub
   end if

   call EmptyClipboard()

   call SetClipboardData(CF_TEXT, memory)

   if CloseClipboard() = 0 then
      msgBox "CloseClipboard failed"
   end if

end sub ' }

sub addModule() ' {
 '
 '  Add a VBA module to the current project
 '
    application.VBE.activeVBProject.vbComponents.add(vbext_ct_StdModule)
end sub ' }

sub removeModule(nameOrNum as variant) ' {
    application.VBE.ActiveVBProject.VBComponents.Remove application.VBE.ActiveVBProject.VBComponents(nameOrNum)
end sub ' }
