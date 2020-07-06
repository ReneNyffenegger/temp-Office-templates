option explicit

sub addReferenceToPersonalXlsb() ' {
 '
 '  Keyboard shortcut that is assigned to this function is ctrl+F11 (see workbook_open in thisWorkbook.bas)
 '
 '  run "personal.xlsb!addReferenceToPersonalXlsb"
 '

  '
  ' Determine if Personal_xlsb was already added:
    on error resume next
    dim n as string
    n = application.VBE.activeVBProject.references("Personal_xlsb").name
    on error goto 0

    if n = "" then
       application.VBE.activeVBProject.references.addFromFile environ$("appdata") & "\Microsoft\Excel\XLSTART\Personal.xlsb"
    else
       debug.print "reference to Personal.xlsb was already added"
    end if

    application.VBE.activevbProject.references.addFromGuid ("{0002E157-0000-0000-C000-000000000046}", 5, 3)
 
'err_:
'   msgBox err.number & ": " & err.description
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

'
'     2020-07-06: Functionality also found in 00_ModuleLoader
'
' sub removeModule(nameOrNum as variant) ' {
'     application.VBE.ActiveVBProject.VBComponents.Remove application.VBE.ActiveVBProject.VBComponents(nameOrNum)
' end sub ' }

sub add_00ModuleLoader() ' {

    dim mdl as vbide.vbComponent
    set mdl = application.VBE.activeVBProject.vbComponents.import("C:\Users\r.nyffenegger\github\lib\VBAModules\Common\00_ModuleLoader.bas")
    mdl.name = "ModuleLoader"

end sub ' }
