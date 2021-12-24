option explicit

dim withEvents app as application

private sub app_newWorkbook(ByVal wb As workbook) ' {
 '  msgBox "New Workbook: " & wb.name
end sub ' }

private sub app_SheetSelectionChange(byVal sh As Object, byVal target as range) ' {
  ' msgBox Sh.Name & " " & Target.Address(0, 0)
end sub ' }

sub workbook_open() ' {

'   msgBox "Workbook was opened: " & me.name

    set app = application

    application.onKey "^q"    , "copyCellWithoutNewLine"
    application.onKey "^s"    , "save_"
    application.onKey "^{F11}", "addReferenceToPersonalXlsb" ' Default for ctrl+F11 is to open a «macro workbook», which apparently almost nobody uses anymore

end sub ' }
