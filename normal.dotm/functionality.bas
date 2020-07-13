option explicit

sub hlp() ' {
    debug.print("help (functionality.bas for Word)")
    debug.print("")
    debug.print("   page()")
end sub ' }

sub page() ' {

    with thisDocument.pageSetup

         debug.print("leftMargin:   " & pointsToCentimeters(.leftMargin  ) & " cm")
         debug.print("rightMargin:  " & pointsToCentimeters(.rightMargin ) & " cm")
         debug.print("topMargin:    " & pointsToCentimeters(.topMargin   ) & " cm")
         debug.print("bottomMargin: " & pointsToCentimeters(.bottomMargin) & " cm")
'        debug.print("footerMargin: " & pointsToCentimeters(.footerMargin) & " cm")  ' Excel

    end with

end sub ' }
