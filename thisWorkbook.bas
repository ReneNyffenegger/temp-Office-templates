option explicit

sub workbook_open() ' {

  ' msgBox "Workbook was opened"

    application.onKey "^q", "copyCellWithoutNewLine"

end sub ' }
