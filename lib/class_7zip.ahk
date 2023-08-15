
/*  7zip Wrapper Klasse in spezieller Version für Addendum für Albis on Windows


  EXAMPLE:


  #NoEnv

  fileName := "zipTest.zip"

  if !(zipFile := new 7Zip(fileName)) {
    msgbox % "Failed loading 7Zip library"
    ExitApp
  }

  debug( "hide           = " zipFile.option._hide() )
  debug( "yes            = " zipFile.option._yes() )
  debug( "password       = " zipFile.option._password() )
  debug( "sfx            = " zipFile.option._sfx() )
  debug( "volumeSize     = " zipFile.option._volumeSize() )
  debug( "workingDir     = " zipFile.option._workingDir() )
  debug( "compressLevel  = " zipFile.option._compressLevel() )
  debug( "compressType   = " zipFile.option._compressType() )
  debug( "recurse        = " zipFile.option._recurse() )
  debug( "includeFile    = " zipFile.option._includeFile() )
  debug( "excludeFile    = " zipFile.option._excludeFile() )
  debug( "includeArchive = " zipFile.option._includeArchive() )
  debug( "excludeArchive = " zipFile.option._excludeArchive() )
  debug( "overwrite      = " zipFile.option._overwrite() )
  debug( "extractPaths   = " zipFile.option._extractPaths() )
  debug( "output         = " zipFile.option._output() )

  debug( "fileCount : " zipFile.getFileCount() )
  debug( "version   : " zipFile.getVersion() )
  debug( "list      : " zipFile.list() )

  debug( A_ScriptDir "\ziptemp" )


  zipFile.extract( A_ScriptDir "\ziptemp" )
  zipFile.close()

  debug( "end")

  ExitApp

  debug( message ) {
   if( A_IsCompiled == 1 )
     return
    message .= "`n"
    FileAppend %message%, * ; send message to stdout

*/

class 7Zip {

  dllName     := ""
  archiveFile := ""
  module      := 0
  option      := new 7Zip.OptionSpec()

  __New( archiveFile, ZiplipPath:="", encoding := "UTF-8", dbg:=false ) {

    this.ziplibPath:= ZipLibPath
    this.encoding := encoding
    this.dbg := dbg

    currDir := ZiplipPath ? ZiplipPath : FileUtil.getDir( A_LineFile ) "\dll\64bit"
    ;~ MsgBox, % currDir "\7-zip64.dll"
    if (r := DllCall("LoadLibrary", "Str", currDir "\7-zip64.dll")) {
      this.dllName := "7-zip64"
    } else if (r := DllCall("LoadLibrary", "Str", currDir "\7-zip32.dll")) {
      this.dllName := "7-zip32"
    } else {
      FileAppend, % "Can't load dll: " ErrorLevel ", r: " r, *   ; schreibt in *stdout
      return ErrorLevel += 1
    }

    this.module      := r
    this.archiveFile := archiveFile

  }

  /**
  * list files in archive
  *
  * @param {String} archiveFile
  * @return {String} response on success, 0 on failure.
  *
  */
  list() {
    commandline := "l """ this.archiveFile """"
    commandline .= this.option._hide()
    commandline .= this.option._password()
    return this._run( commandline )
  }

  /**
  * add file to archive
  *
  * @param {String} file     file to add
  * @return {String} response on success, 0 on failure.
  */
  add( file ) {
    /*    Adds files to archive.

    Examples
    7z a archive1.zip subdir\
    adds all files and subfolders from folder subdir to archive archive1.zip. The filenames in archive will
    contain subdir\ prefix.

    7z a archive2.zip .\subdir\*
    adds all files and subfolders from folder subdir to archive archive2.zip. The filenames in archive will
    not contain subdir\ prefix.

    cd /D c:\dir1\
    7z a c:\archive3.zip dir2\dir3\
    The filenames in archive c:\archive3.zip will contain dir2\dir3\ prefix, but they
    will not contain c:\dir1\ prefix.

    7z a Files.7z *.txt -r
    adds all *.txt files from current folder and its subfolders to archive Files.7z.

    Notes
    7-Zip doesn't use the system wildcard parser. 7-Zip doesn't follow the archaic rule by which *.*
    means any file. 7-Zip treats *.* as matching the name of any file that has an extension.
    To process all files, you must use a * wildcard.

    */

    IF InStr(file, "@")      ; to add a files from a list
      commandline := "a """ this.archiveFile """ "  file   " "
    else
      commandline := "a """ this.archiveFile """ """ file """"

    commandline .= this.option._hide()
    commandline .= this.option._volumeSize()
    commandline .= this.option._stream()
    commandline .= this.option._password()
    commandline .= this.option._compressLevel()
    commandline .= this.option._compressType()
    commandline .= this.option._compressOpts()
    commandline .= this.option._recurse()
    commandline .= this.option._sfx()
    commandline .= this.option._workingDir()
    commandline .= this.option._includeFile()
    commandline .= this.option._excludeFile()

    SciTEOutput("commandline: " commandLine)

    return this._run( commandline )
  }

  /**
  * delete file to archive
  *
  * @param {String} file  fileName to delete
  * @return {String} response on success, 0 on failure.
  */
  delete( file ) {
    commandline  = "d """ this.archiveFile """ """ file """"
    commandline .= this.option._hide()
    commandline .= this.option._password()
    commandline .= this.option._compressLevel()
    commandline .= this.option._compressType()
    commandline .= this.option._recurse()
    commandline .= this.option._sfx()
    commandline .= this.option._workingDir()
    commandline .= this.option._includeFile()
    commandline .= this.option._excludeFile()
    return this._run( commandline )
  }

  /**
  * extract files from archive
  *
  * @param {String} path   directory to extract files
  * @return {String} response on success, 0 on failure.
  */
  extract( path="" ) {
    commandline := ( this.option.extractPaths ? "x" : "e" ) " """ this.archiveFile """"
    commandline .= this.option._hide()
    commandline .= this.option._stream()
    commandline .= this.option._recurse()
    commandline .= this.option._overwrite()
    commandline .= this.option._password()
    commandline .= this.option._workingDir()
    commandline .= this.option._yes()
    commandline .= this.option._includeArchive()
    commandline .= this.option._excludeArchive()
    commandline .= this.option._includeFile()
    commandline .= this.option._excludeFile()

    if ( path != "" ) {
      FileCreateDir, % path
      commandline .= " -o""" path """"
    }

    return this._run( commandline )
  }

  /**
  * check archive integrity
  *
  * @return {Boolean} true on success, false otherwise
  */
  checkArchive() {
    Return DllCall( this.dllName "\SevenZipCheckArchive", "AStr", this.archiveFile, "int", 0 )
  }

  /**
  * get type of archive
  *
  * @return {Number}
  *   0 : Unknown type
  *   1 : ZIP type
  *   2 : 7Z type
  *  -1 : Failure
  */
  getArchiveType() {
    Return DllCall( this.dllName "\SevenZipGetArchiveType", "AStr", this.archiveFile )
  }

  /**
  * get number of files in archive
  *
  * @return {Number} number of files
  */
  getFileCount() {
    Return DllCall( this.dllName "\SevenZipGetFileCount", "AStr", this.archiveFile )
  }

  configDialog() {
    ;
    ; Function: 7Zip_ConfigDialog
    ; Description:
    ;      Shows the about dialog for 7-zip32.dll
    ; Syntax: 7Zip_ConfigDialog(hWnd)
    ; Parameters:
    ;      hWnd - handle of owner window
    ;
    Return DllCall( this.dllName "\SevenZipConfigDialog", "Ptr", 0, "ptr",0, "int",0 )
  }

  queryFunctionList( iFunction = 0 ) {
    Return DllCall( this.dllName "\SevenZipQueryFunctionList", "int", iFunction )
  }

  /*
  * get version of 7-zip32.dll
  *
  * @return {String} version
  */
  getVersion() {
    aRet := DllCall( this.dllName "\SevenZipGetVersion", "Short" )
    Return SubStr(aRet,1,1) . "." . SubStr(aRet,2)
  }

  /*
  * get sub version of 7-zip32.dll
  *
  * @return {String} subversion
  */
  getSubVersion() {
    return DllCall( this.dllName "\SevenZipGetSubVersion", "Short" )
  }

  /**
  * free 7-zip32.dll library in memory.
  */
  close() {
    DllCall("FreeLibrary", "Ptr", this.option.hwnd)
    this.module      := 0
    this.archiveFile := ""
  }

  setOwnerWindowEx( sProcFunc ) {
    ;
    ; Function: 7Zip_SetOwnerWindowEx
    ; Description:
    ;      Appoints the call-back function in order to receive the information of the compressing/unpacking
    ; Syntax: 7Zip_SetOwnerWindowEx(hWnd, sProcFunc)
    ; Parameters:
    ;      sProcFunc - Callback function name
    ;      hWnd - handle of window (calling application), can be 0
    ; Return Value:
    ;      True on success, false otherwise
    ; Related: 7Zip_KillOwnerWindowEx
    ; Example:
    ;      file:example_callback.ahk
    ;
    Address := RegisterCallback(sProcFunc, "F", 4)
    Return DllCall( this.dllName "\SevenZipSetOwnerWindowEx","Ptr", 0 , "ptr", Address )
  }

  ;
  killOwnerWindowEx() {
    ;
    ; Function: 7Zip_KillOwnerWindowEx
    ; Description:
    ;      Removes the callback
    ; Syntax: 7Zip_KillOwnerWindowEx(hWnd)
    ; Parameters:
    ;      hWnd - Handle to parent or owner window
    ; Return Value:
    ;      True on success, false otherwise
    ; Related: 7Zip_SetOwnerWindowEx
    Return DllCall( this.dllName "\SevenZipKillOwnerWindowEx" , "Ptr", 0 )
  }

  SetUnicodeMode(state) {
  ; SevenZipSetUnicodeMode
  Return DllCall( this.dllName "\SevenZipSetUnicodeMode", "Int", state)
  }

  ; FUNCTIONS BELOW - CREDIT TO LEXIKOS -------------------------------------------------------

  openArchive( archiveFile ) {
    ;
    ; Function: 7Zip_OpenArchive
    ; Description:
    ;      Open archive and return handle for use with 7Zip_FindFirst
    ; Syntax: 7Zip_OpenArchive(archiveFile, [hWnd])
    ; Parameters:
    ;      archiveFile - Path of archive
    ;      hWnd - Handle of calling window
    ; Return Value:
    ;      Handle for use with 7Zip_FindFirst function, 0 on error.
    ; Remarks:
    ;      Nil
    ; Related: 7Zip_CloseArchive, 7Zip_FindFirst , File Info Functions
    ; Example:
    ;      hArc := 7Zip_OpenArchive("C:\Path\To\Archive.7z")
    ;
    Return DllCall( this.dllName "\SevenZipOpenArchive", "Ptr", 0, "AStr", archiveFile, "int", 0 )
  }

  closeArchive( archiveFile ) {
    ;
    ; Function: 7Zip_CloseArchive
    ; Description:
    ;      Closes the archive handle
    ; Syntax: 7Zip_CloseArchive(hArc)
    ; Parameters:
    ;      hArc - Handle retrived from 7Zip_OpenArchive
    ; Return Value:
    ;      -1 on error
    ; Remarks:
    ;      Nil
    ; Related: 7Zip_OpenArchive
    ; Example:
    ;      7Zip_CloseArchive(hArc)
    ;
    Return DllCall( this.dllName "\SevenZipCloseArchive", "Ptr", 0, "AStr", archiveFile, "int", 0 )
  }

  findFirst( hArc, sSearch, o7zip__info="" ) {
    ;
    ; Function: 7Zip_FindFirst
    ; Description:
    ;      Find first file for search criteria in archive
    ; Syntax: 7Zip_FindFirst(hArc, sSearch, [o7zip__info])
    ; Parameters:
    ;      hArc - handle of archive (returned from 7Zip_OpenArchive)
    ;      sSearch - Search string (wildcards allowed)
    ;      o7zip__info - (Optional) Name of object to recieve details of file.
    ; Return Value:
    ;      Object with file details on success. If 3rd param was 0, returns true on success. False on failure.
    ; Remarks:
    ;      If third param is omitted, details are returned in a new object.
    ;      If it is set to 0, details are not retrieved. (You can use the other functions to get details.)
    ; Related: 7Zip_FindNext , 7Zip_OpenArchive , File Info Functions
    ; Example:
    ;      file:example_archive_info.ahk
    ;
    if (o7zip__info == 0) {
      r := DllCall(this.dllName "\SevenZipFindFirst", "Ptr", hArc, "AStr", sSearch, "ptr", 0)
      return ( r ? 0 : 1 ), ErrorLevel := (r ? r : ErrorLevel)
    }
    if ! IsObject(o7zip__info)
      o7zip__info := {}
    VarSetCapacity(tINDIVIDUALINFO , 558, 0)
    If DllCall(this.dllName "\SevenZipFindFirst", "Ptr", hArc, "AStr", sSearch, "ptr", &tINDIVIDUALINFO)
      Return 0
    o7zip__info.OriginalSize   := NumGet(tINDIVIDUALINFO , 0, "UInt")
    o7zip__info.CompressedSize := NumGet(tINDIVIDUALINFO , 4, "UInt")
    o7zip__info.CRC            := NumGet(tINDIVIDUALINFO , 8, "UInt")
  ; uFlag                      := NumGet(tINDIVIDUALINFO , 12, "UInt") ;always 0
  ; uOSType                    := NumGet(tINDIVIDUALINFO , 16, "UInt") ;always 0
    o7zip__info.Ratio          := NumGet(tINDIVIDUALINFO , 20, "UShort")
    o7zip__info.Date           := this._dosDateTimeToStr(NumGet(tINDIVIDUALINFO , 22, "UShort"),NumGet(tINDIVIDUALINFO , 24, "UShort"))
    o7zip__info.FileName       := StrGet(&tINDIVIDUALINFO+26 ,513,"CP0")
    o7zip__info.Attribute      := StrGet(&tINDIVIDUALINFO+542,8  ,"CP0")
    o7zip__info.Mode           := StrGet(&tINDIVIDUALINFO+550,8  ,"CP0")

    return o7zip__info
  }

  findNext( hArc, o7zip__info="" ) {
    ;
    ; Function: 7Zip_FindNext
    ; Description:
    ;      Find next file for search criteria in archive
    ; Syntax: 7Zip_FindNext(hArc, [o7zip__info])
    ; Parameters:
    ;      hArc - handle of archive (returned from 7Zip_OpenArchive)
    ;      o7zip__info - (Optional) Name of object to recieve details of file.
    ; Return Value:
    ;      Object with file details on success. If 2nd param was 0, returns true on success. False on failure.
    ; Remarks:
    ;      If second param is omitted, details are returned in a new object.
    ;      If it is set to 0, details are not retrieved. (You can use the other functions to get details.)
    ; Related: 7Zip_FindFirst , 7Zip_OpenArchive, File Info Functions
    ; Example:
    ;      file:example_archive_info.ahk
    ;
    if (o7zip__info = 0)
    {
      r := DllCall(this.dllName "\SevenZipFindFirst", "Ptr", hArc, "AStr", sSearch, "ptr", 0)
      return ( r ? 0 : 1 ), ErrorLevel := (r ? r : ErrorLevel)
    }
    if !IsObject(o7zip__info)
      o7zip__info := {}
    VarSetCapacity(tINDIVIDUALINFO , 558, 0)
    if DllCall(this.dllName "\SevenZipFindNext","Ptr", hArc, "ptr", &tINDIVIDUALINFO)
      Return 0

    o7zip__info.OriginalSize   := NumGet(tINDIVIDUALINFO , 0, "UInt")
    o7zip__info.CompressedSize := NumGet(tINDIVIDUALINFO , 4, "UInt")
    o7zip__info.CRC            := NumGet(tINDIVIDUALINFO , 8, "UInt")
    o7zip__info.Ratio          := NumGet(tINDIVIDUALINFO , 20, "UShort")
    o7zip__info.Date           := this._dosDateTimeToStr(NumGet(tINDIVIDUALINFO , 22, "UShort"),NumGet(tINDIVIDUALINFO , 24, "UShort"))
    o7zip__info.FileName       := StrGet(&tINDIVIDUALINFO+26 ,513,"CP0")
    o7zip__info.Attribute      := StrGet(&tINDIVIDUALINFO+542,8  ,"CP0")
    o7zip__info.Mode           := StrGet(&tINDIVIDUALINFO+550,8  ,"CP0")

    return o7zip__info
  }

  getFileName(hArc) {
    ;
    ; Function: File Info Functions
    ; Description:
    ;      Using handle hArc, get info of file(s) in archive.
    ; Syntax: 7Zip_<InfoFunction>(hArc)
    ; Parameters:
    ;      7Zip_GetFileName - Get file name
    ;      7Zip_GetArcOriginalSize - Original size of file
    ;      7Zip_GetArcCompressedSize - Compressed size
    ;      7Zip_GetArcRatio - Compression ratio
    ;      7Zip_GetDate - Date
    ;      7Zip_GetTime - Time
    ;      7Zip_GetCRC - File CRC
    ;      7Zip_GetAttribute - File Attribute
    ;      7Zip_GetMethod - Compression method (LZMA or PPMD)
    ; Return Value:
    ;      -1 on error
    ; Remarks:
    ;      See included example for details
    ; Related: 7Zip_OpenArchive , 7Zip_FindFirst
    ; Example:
    ;      file:example_archive_info.ahk
    ;
    VarSetCapacity( tNameBuffer,513 )
    If !DllCall(this.dllName "\SevenZipGetFileName", "Ptr", hArc, "ptr", &tNameBuffer, "int", 513)
      Return StrGet(&tNameBuffer,513,"CP0")
  }

  getArcOriginalSize(hArc) {
    Return DllCall(this.dllName "\SevenZipGetArcOriginalSize", "Ptr", hArc)
  }

  getArcCompressedSize(hArc) {
    Return DllCall(this.dllName "\SevenZipGetArcCompressedSize", "Ptr", hArc)
  }

  getArcRatio(hArc) {
    Return DllCall(this.dllName "\SevenZipGetArcRatio", "Ptr", hArc, "short")
  }

  getDate(hArc) {
    Return this._dosDate(DllCall(this.dllName "\SevenZipGetDate", "Ptr", hArc, "Short"))
  }

  getTime(hArc) {
    Return this._dosTime(DllCall(this.dllName "\SevenZipGetTime", "Ptr", hArc, "Short"))
  }

  getCRC(hArc) {
    Return DllCall(this.dllName "\SevenZipGetCRC", "Ptr", hArc, "UInt")
  }

  getAttribute(hArc) {
    return DllCall(this.dllName "\SevenZipGetAttribute", "Ptr", hArc)
  }

  getMethod(hArc) {
    VarSetCapacity(sBUFFER,8)
    if !DllCall(this.dllName "\SevenZipGetMethod" , "Ptr", hArc , "ptr", &sBuffer,"int", 8)
      Return StrGet(&sBUFFER, 8, "CP0")
  }

  ; FUNCTIONS FOR INTERNAL USE --------------------------------------------------------------------------------------------------
  _run( commandLine ) {

    /*
    -----------------------------------------------------------------------
    int WINAPI SevenZip(const HWND _hwnd, LPCSTR _szCmdLine,
                        LPSTR _szOutput, const DWORD _dwSize)
    -----------------------------------------------------------------------
    Number  1
    Function
        Compression thawing and the like is done.

    Argument
        _hwnd       The window handle of the application which calls 7-zip32.dll.
                    7-zip32.dll EnableWindow () executes the execution time vis-a-vis this window and controls the operation of the window.
                    It is the console application where the window does not exist when and, when it is not necessary to appoint, NULL is transferred.
        _szCmdLine  The command character string which transfers to 7-zip32.dll.
        _szOutput   The buffer because 7-zip32.dll returns the result.
                    When it is global memory or like that, it is necessary to be locked.
        _dwSize     Größe des Puffers.
                    When the result exceeds designated size, it is economized in this size.
                    If size is 1 or more, always NULL letter is added lastly.

    Return value
      At the time of normal termination 0.
      The time of error a quantity other than 0.
                    As for error value not yet verification.

    */

	If IsFunc("debug")
		dbgfunc := Func("debug")
	else if IsFunc("SciteOutput")
		dbgfunc := Func("SciteOutput")
	If IsFunc(dbgFunc) && this.dbg {
        msg := StrReplace(commandLine, q " " q, q "`n  " q)
        msg := StrReplace(msg, q " -" , q "`n-")
		%dbgFunc%( msg)
      }


    nSize := 32768
    VarSetCapacity(tOutBuffer,nSize)
    retVal := DllCall(this.dllName "\SevenZip"
            , "Ptr", ""
            , "AStr", GetANSI(commandLine)
            , "Ptr", &tOutBuffer
            , "Int", nSize)
    If ! ErrorLevel
      return StrGet(&tOutBuffer,nSize, "CP0"), ErrorLevel := retVal ; "CP0"
    else
      return 0
  }


  _dosDate( ByRef DosDate ) {
    day   := DosDate & 0x1F
    month := (DosDate<<4) & 0x0F
    year  := ((DosDate<<8) & 0x3F) + 1980
    return "" . year . "/" . month . "/" . day
  }

  _dosTime( ByRef DosTime ) {
    sec   := (DosTime & 0x1F) * 2
    min   := (DosTime<<4) & 0x3F
    hour  := (DosTime<<10) & 0x1F
    return "" . hour . ":" . min . ":" . sec
  }

  _dosDateTimeToStr( ByRef DosDate, ByRef DosTime ) {
    VarSetCapacity(FileTime,8)
    DllCall("DosDateTimeToFileTime", "UShort", DosDate, "UShort", DosTime, "UInt", &FileTime)
    VarSetCapacity(SystemTime, 16, 0)
    If (!NumGet(FileTime,"UInt") && !NumGet(FileTime,4,"UInt"))
     Return 0
    DllCall("FileTimeToSystemTime", "PTR", &FileTime, "PTR", &SystemTime)
    Return NumGet(SystemTime,6,"short") ;date
      . "/" . NumGet(SystemTime,2,"short") ;month
      . "/" . NumGet(SystemTime,0,"short") ;year
      . " " . NumGet(SystemTime,8,"short") ;hours
      . ":" . ((StrLen(tvar := NumGet(SystemTime,10,"short")) = 1) ? "0" . tvar : tvar) ;minutes
      . ":" . ((StrLen(tvar := NumGet(SystemTime,12,"short")) = 1) ? "0" . tvar : tvar) ;seconds
    ;      . "." . NumGet(SystemTime,14,"short") ;milliseconds
  }

  class OptionSpec {

    hide           := true    ;Callback is called (bool);a,d,e,x,u
    stream         := "p1"    ;-bs (Set output stream for output/error/progress line) switch
    yes            := 0       ;assume Yes on all queries;e,x
    password       := ""      ;Password (string);a,d,e,x,u
    SFX            := ""      ;Self extracting archive module name (string);a,u
    volumeSize     := 0       ;Create volumes of specified sizes (integer);a
    workingDir     := ""      ;Sets working directory for temporary base archive (string);a,d,e,x,u
    compressLevel  := 5       ;0-9 (level);a,d,u
    compressType   := "7z"    ;7z,gzip,zip,bzip2,tar,iso,udf (string);a
    compressOpts   := "mo=LZMA2"      ;
    recurse        := 0       ;0:Disable, 1:Enable, 2:Enable only for wildcard names;a,d,e,x,u
    includeFile    := ""      ;Specifies filenames and wildcards or list file that specify processed files (string);a,d,e,x,u
    excludeFile    := ""      ;Specifies what filenames or (and) wildcards must be excluded from operation (string);a,d,e,x,u
    includeArchive := ""      ;Include archive filenames (string);e,x
    excludeArchive := ""      ;Exclude archive filenames (string);e,x
    overwrite      := 0       ;0:Overwrite All, 1:Skip extracting of existing, 2:Auto rename extracting file, 3:auto rename existing file;e,x
    extractPaths   := 1       ;Extract full paths (default 1);e,x
    output         := ""      ;Output directory (string);e,x

    toOption( key, val ) {
      return val ? " -" key val : ""
    }

    toBooleanOption( key, val ) {
      return val ? " -" key : ""
    }

    toIncludeOption( key, val ) {
      if ! (val)
        return
        val := """" val """"
      return (SubStr(val,1,1) == "@") ? " -" key val : " -" key "!" val
    }

    _hide() {
      ;return this.toBooleanOption( "hide", this.hide )
    }

    _stream() {
    /*
    -bs (Set output stream for output/error/progress line) switch
      Syntax
      -bs{o|e|p}{0|1|2}
          {id}	Stream Type
          o	standard output messages
          e	error messages
          p	progress information
      {N}	Stream Destination
          0	disable stream
          1	redirect to stdout stream
          2	redirect to stderr stream
          Default values: o1, e2, p1.

      Examples
      7z t a.7z -bse1 > messages.txt
      tests a.7z archive and sends error messages to stdout that is redirected to messages.txt

      7z a -si -so -bsp2 -txz -an < file.tar > file.tar.xz
      compresses file.tar (from stdin) to file.tar.xz (stdout stream) and shows progress information in stderr stream that can be seen at console window.
    */
       return this.toOption( "bs", this.stream )
    }

    _yes() {
      return this.toBooleanOption( "y", this.yes )
    }

    _password() {
      return this.toOption( "p", this.password )
    }

    _sfx() {
      return FileExist(this.SFX) ? " -sfx" . this.SFX : ""
    }

    _volumeSize() {
      return this.toOption( "v", this.volumeSize )
    }

    _workingDir() {
      return this.toIncludeOption( "w", this.workingDir )
    }

    _compressLevel() {
      return this.toOption( "mx", this.compressLevel )
    }

    _compressType() {
      return this.toOption( "t", this.compressType )
    }

    _compressOpts() {
      return " " this.compressOpts
    }

    _recurse() {
      if this.recurse = 1
        return " -r"
      if this.recurse = 2
        return " -r0"
      Else
        return " -r-"
    }

    _includeFile() {
      return this.toIncludeOption( "i", this.includeFile )
    }

    _excludeFile() {
      return this.toIncludeOption( "x", this.excludeFile )
    }

    _includeArchive() {
      return this.toIncludeOption( "ai", this.includeArchive )
    }

    _excludeArchive() {
      return this.toIncludeOption( "ax", this.excludeArchive )
    }

    _overwrite() {
      if (this.overwrite = 0)
        return " -aoa"
      else if (this.overwrite = 1)
        return " -aos"
      else if (this.overwrite = 2)
        return " -aou"
      else if (this.overwrite = 3)
        return " -aot"
      Else
        return " -aoa"
    }

    _extractPaths() {
      return this.toOption( "o", this.extractPaths )
    }

    _output() {
      return this.toOption( "o", this.output )
    }

  }

}

GetANSI(ByRef Str) {
   VarSetCapacity(Tmp, StrPut(Str, "UTF-8"))
   StrPut(Str, &Tmp, "UTF-8")
   Return StrGet(&Tmp, "CP0")
}



class FileUtil {

	static void := FileUtil._init()

	_init() {
	}

	__New() {
		throw Exception( "FileUtil is a static class, dont instante it!", -1 )
	}

	getDir( path ) {
		path := RegExReplace( path, "^(.*?)\\$", "$1" )
		if( this.isDir(path) )
			return path
		return this.getParentDir( path )
	}

	getParentDir( path ) {
		path := RegExReplace( path, "^(.*?)\\$", "$1" )
		path := RegExReplace( path, "^(.*)\\.+?$", "$1" )
		return path
	}

	getExt( filePath ) {
		SplitPath, % filePath, fileName, fileDir, fileExtention, fileNameWithoutExtension, DriveName
		StringLower, fileExtention, fileExtention
		return fileExtention
	}

  /**
  * File extention is matched width extentionPattern
  *
  * @param {string} filePath
  * @param {string} extentionPattern
  * @exmaple
  *   FileUtil.isExt("cue|mdx")
  */
	isExt( filePath, extentionPattern ) {

		IfNotExist %filePath%
			return false

		if ( RegExMatch( filePath, "i).*\.(" extentionPattern ")$" ) ) {
			return true
		} else {
			return false
		}

	}

	getFileName( filePath, withExt:=true ) {
		filePath := RegExReplace( filePath, "^(.*?)\\$", "$1" )
		SplitPath, filePath, fileName, fileDir, fileExtention, fileNameWithoutExtension, DriveName
		if( withExt == true )
			return fileName
		return fileNameWithoutExtension
	}

	getFiles( path, pattern=".*", includeDir=false, recursive=false ) {

		files := []

		if ( this.isFile(path) ) {
			if RegExMatch( path, pattern )
				files.Insert( path )

		} else {

		currDir := this.getDir( path )
			Loop, %currDir%\*, % includeDir, % recursive
			{
					if not RegExMatch( A_LoopFileFullPath, pattern )
						continue
					files.Insert( A_LoopFileFullPath )
			}

			this._sortArray( files )

		}

      return files

	}

	getFile( pathDirOrFile, pattern=".*" ) {

		if ( pathDirOrFile == "" or this.isFile(pathDirOrFile) )  {
			return pathDirOrFile
		}

        files := this.getFiles( pathDirOrFile, pattern )

        if ( files.MaxIndex() > 0 ) {
        	return files[ 1 ]
        }

        return ""

	}

	isDir( path ) {
		if( ! this.exist(path) )
			return false
		FileGetAttrib, attr, %path%
		Return InStr( attr, "D" )
	}

	isFile( path ) {
		if( ! this.exist(path) )
			return false
		FileGetAttrib, attr, %path%
		Return ! InStr( attr, "D" )
	}

	readProperties( path ) {

		prop := []

		Loop, Read, %path%
		{

			If RegExMatch(A_LoopReadLine, "^#.*" )
				continue

			splitPosition := InStr(A_LoopReadLine, "=" )

			If ( splitPosition = 0 ) {
				key := A_LoopReadLine
				val := ""
			} else {
				key := SubStr( A_LoopReadLine, 1, splitPosition - 1 )
				val := SubStr( A_LoopReadLine, splitPosition + 1 )
			}

			prop[ Trim(key) ] := Trim(val)

		}

		return prop

	}

	makeDir( path ) {
		FileCreateDir, %path%
	}

	makeParentDir( path, forDirectory=true ) {
		if ( forDirectory == true ) {
			parentDir := this.getParentDir( path )
		} else {
			parentDir := this.getDir( path )
		}
		FileCreateDir, % parentDir
	}

	exist( path ) {
		return FileExist( path )
	}

	delete( path, recursive=1 ) {
		if ( this.isFile(path) ) {
			FileDelete, % path
		} else if( this.isDir(path) ) {
			FileRemoveDir, % path, % recursive
		}
	}

	move( src, trg, overwrite=1 ) {
		if ( ! this.exist(src) )
			return
		this.makeParentDir( trg, this.isDir(src) )
		FileMove, % src, % trg, % overwrite
	}

	copy( src, trg, overwrite=1 ) {
		if ( ! this.exist(src) )
			return
		this.makeParentDir( trg, this.isDir(src) )
		if ( this.isDir(src) ) {
			FileCopyDir, % src, % trg, % overwrite
		} else {
			FileCopy, % src, % trg, % overwrite
		}
	}

  /**
  * get file size
  *
  * @param {path} filePath
  * @return size (byte)
  */
	getSize( path ) {
		FileGetSize, size, % path
		return size
	}

  /**
  * get time
  *
  * @param {path} file path
  * @param {witchTime} M: modification time (default), C: creation time, A: last access time
  * @return YYYYMMDDHH24MISS
  */
	getTime( path, whichTime="M" ) {
		FileGetTime, var, % path, % whichTime
		return var
	}

    /**
    * Get Symbolic Link Information
    *
    * @param  filePath   path to check if it is symbolic link
    * @param  srcPath    path to linked by filePath
    * @param  linkType   link type ( file or directory )
    * @return true if filepath is symbolic link
    */
	isSymlink( filePath, ByRef srcPath="", ByRef linkType="" ) {

		IfNotExist, % filePath
			return false

		if RegExMatch(filePath,"^\w:\\?$") ; false if it is a root directory
			return false

		SplitPath, filePath, fn, parentDir

		result := this.cli( "/c dir /al """ (InStr(FileExist(filePath),"D") ? parentDir "\" : filePath) """" )

		if RegExMatch(result,"<(.+?)>.*?\b" fn "\b.*?\[(.+?)\]",m) {
			linkType:= m1, srcPath := m2
			if ( linkType == "SYMLINK" )
  			linkType := "file"
			else if ( linkType == "SYMLINKD" )
  			linkType := "directory"
			return true
		} else {
			return false
		}

	}

  /**
  * make symbolic link
  *
  * @param src  source path (real file)
  * @param trg  target path (path to used as link)
  */
    makeLink( src, trg ) {

  	if this.isSymlink( trg ) {
  		this.delete( trg )
  	}

		this.makeParentDir( trg, this.isDir(src) )
		if ( this.isDir(src) ) {
			cmd := "/c mklink /d """ trg """ """ src """"
		} else {
			cmd := "/c mklink /f """ trg """ """ src """"
		}
		this.cli( cmd )

  }

  /**
  * run command and return result
  *
  * @param  command	 command
  * @return command execution result
  */
	cli( command ) {

		dhw := A_DetectHiddenWindows
		DetectHiddenWindows,On
		Run, %ComSpec% /k,,Hide UseErrorLevel, pid
		if not ErrorLevel
		{
			while ! WinExist("ahk_pid" pid)
				Sleep,100
			DllCall( "AttachConsole","UInt",pid )
		}
		DetectHiddenWindows, % dhw

		; debug( "command :`n`t" command )
		shell := ComObjCreate("WScript.Shell")
		try {
			exec := shell.Exec( comspec " " command )
			While ! exec.Status
				sleep, 100
			result := exec.StdOut.readAll()
		}
		catch e
          return "error: " e.what " - " e.message

		; debug( "result :`n`t" result )
		DllCall("FreeConsole")
		Process Close, %pid%

		return result

	}

	_sortArray( Array ) {
	  t := Object()
	  for k, v in Array
	    t[RegExReplace(v,"\s")]:=v
	  for k, v in t
	    Array[A_Index] := v
	  return Array
	}

}