option explicit

sub copyCellWithoutNewLine() ' {
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

'  text4clipboard = "This text was placed into the clipboard via VBA"
   text4clipboard = activeCell.value

   memory       = GlobalAlloc(GHND, len(text4clipboard) + 1)
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

