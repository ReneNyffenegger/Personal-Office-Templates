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

  '
  ' Add reference to «Microsoft Visual Basic for Applications Extensibility 5.3»
  '
    on error goto errReference
    application.VBE.activevbProject.references.addFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 3
    exit sub
 
 errReference:

    if err.number = 32813 then
    '
    '  Name conflicts with existing module, project, or object library
    '
       debug.print "reference was already added"
    else
       msgBox err.number & ": " & err.description
    end if

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

function addModule(optional subMain as boolean = false) as vbide.vbComponent ' {
 '
 '  Add a VBA module to the current project
 '
    if application.VBE.activeVBProject.name = "Personal_xlsb" then ' {
       debug.print "activeVBProject = Personal.xlsb"
       exit function
    end if ' }


    set addModule = application.VBE.activeVBProject.vbComponents.add(vbext_ct_StdModule)

    if subMain then ' {

       dim cm as vbide.codeModule
       set cm = addModule.codeModule

       cm.insertLines 1, "option explicit"
       cm.insertLines 2, ""
       cm.insertLines 3, "sub main()"
       cm.insertLines 4, ""
       cm.insertLines 5, "end sub"

    end if ' }

end function ' }

' {
'     2020-07-06: Functionality also found in 00_ModuleLoader
'
' sub removeModule(nameOrNum as variant)
'     application.VBE.ActiveVBProject.VBComponents.Remove application.VBE.ActiveVBProject.VBComponents(nameOrNum)
' end sub ' }

sub r1c1() ' {
    application.referenceStyle = xlR1C1
end sub ' }

sub a1() ' {
    application.referenceStyle = xlA1
end sub ' }

sub add_00ModuleLoader() ' {

    if application.VBE.activeVBProject.name = "Personal_xlsb" then ' {
       debug.print "activeVBProject = Personal.xlsb"
       exit sub
    end if ' }

    if isModuleNamePresent("ModuleLoader") then ' {
        debug.print "ModuleLoader was already added"
        exit sub
    end if ' }

    dim mdl as vbide.vbComponent

    set mdl = application.VBE.activeVBProject.vbComponents.import("C:\Users\r.nyffenegger\github\lib\VBAModules\Common\00_ModuleLoader.bas")
    mdl.name = "ModuleLoader"

end sub ' }

function isModuleNamePresent(moduleName as string) ' {

    dim mdl as vbide.vbComponent

    for each mdl in application.VBE.activeVBProject.vbComponents ' {

        if mdl.name = moduleName then
           isModuleNamePresent = true
           exit function
        end if 

    next mdl ' }

    isModuleNamePresent = false
    
end function ' }

sub createTestConstellation(testFileName as string) ' {

    if application.VBE.activeVBProject.name = "Personal_xlsb" then ' {
       debug.print "activeVBProject = Personal.xlsb"
       exit sub
    end if ' }

    if isModuleNamePresent("loadTestModule") then ' {
        debug.print "loadTestModule was already added"
        exit sub
    end if ' }

    add_00ModuleLoader 

    dim mdlLoadTestFile as vbide.vbComponent
    set mdlLoadTestFile = addModule
    mdlLoadTestFile.name = "loadTestModule"

    replace testFileName, "/", "\" ' make vim syntax highlightin happy: \\"

    dim cm as vbide.codeModule
    set cm = mdlLoadTestFile.codeModule
    cm.insertLines 1, "option explicit"
    cm.insertLines 2, ""
    cm.insertLines 3, "sub loadTestFile()"
    cm.insertLines 4, "  loadOrReplaceModuleWithFile ""testFile"", """ & testFileName & """"
    cm.insertLines 5, "  application.run ""main"""
    cm.insertLines 6, "  appActivate application.caption"
    cm.insertLines 7, "end sub"

    application.onKey "^m", activeWorkbook.name & "!loadTestFile"

    debug.print("end")
    debug.print("loadTestFile")
    debug.print("main")

end sub ' }

sub saveAsXlsm(fileName as string) ' {

    activeWorkbook.saveAs fileName, xlOpenXMLWorkbookMacroEnabled

end sub ' }

sub clearUsedRange() ' {

     with activeSheet.usedRange
         .select
         .clearFormats
         .clearContents
     end with

end sub ' }

' sub showSheets(optional wb as workbook) ' {
' 
'     if wb is nothing then
'        set wb = activeWorkbook
'     end if
' 
'     dim sh as worksheet
' 
'     for each sh in wb.worksheets
' 
' 
'     next xh
' 
' end sub ' }

sub hlp() ' {
    debug.print("help (functionality.vb for Excel)")
    debug.print "  addModule true|false          ' true: with sub main"
    debug.print "  createTestConstellation  """"   ' ""p:\ath\to\file"""
    debug.print "  add_00ModuleLoader"
    debug.print "  saveAsXlsm ""filename""                     ' save activeWorkbook in xlsm format"
    debug.print "  clearUsedRange                            ' select usedRange (for verification), then call clearContents and clearFormats on activeSheet.usedRange"
'   debug.print "  showSheets [wb]                             | show name of sheets in workbook indicated by wb. If wb is missing, use current workbook"
end sub ' }

sub save_() ' {
    
    if activeWorkbook.path = "" then
       dim ret as variant
       ret = application.getSaveAsFilename(initialFilename := environ("userprofile"), fileFilter := "xlsm, *.xlsm, xlsx, *.xlsx")

       if ret <> false then

          if     right(ret, 5) = ".xlsm" then
                 activeWorkbook.saveAs ret, xlOpenXMLWorkbookMacroEnabled
          elseif right(ret, 5) = ".xlsx" then
                 activeWorkbook.saveAs ret, xlOpenXMLWorkbook
          else
                 msgBox ("Unknown extension")
          end if
       end if
    else
       activeWorkbook.save
    end if

end sub ' }
