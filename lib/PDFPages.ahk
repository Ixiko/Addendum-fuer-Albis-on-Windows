; original written by malcev 2022
; https://www.autohotkey.com/boards/viewtopic.php?f=76&t=103992&sid=c4836422e75965e4487395e6860c659f

PDFPages(file) {    ;-- function to count pdf pages in file


	 VarSetCapacity(GUID, 16)
	 DllCall("ole32\CLSIDFromString", "wstr", IID_RandomAccessStream := "{905A0FE1-BC53-11DF-8C49-001E4FC686DA}", "ptr", &GUID)
	 DllCall("ShCore\CreateRandomAccessStreamOnFile", "wstr", file, "uint", Read := 0, "ptr", &GUID, "ptr*", IRandomAccessStream)
	 CreateClass("Windows.Data.Pdf.PdfDocument", IPdfDocumentStatics := "{433A0B5F-C007-4788-90F2-08143D922599}", PdfDocumentStatics)
	 DllCall(NumGet(NumGet(PdfDocumentStatics+0)+8*A_PtrSize), "ptr", PdfDocumentStatics, "ptr", IRandomAccessStream, "ptr*", PdfDocument)   ; LoadFromStreamAsync
	 WaitForAsync(PdfDocument)
	 DllCall(NumGet(NumGet(PdfDocument+0)+7*A_PtrSize), "ptr", PdfDocument, "uint*", PageCount)
	 ObjReleaseClose(IRandomAccessStream)
	 ObjReleaseClose(PdfDocumentStatics)
	 ObjReleaseClose(PdfDocument)

 Return PageCount
}

; from AHK-Forum to use with some cool Filestream functions on PDF files
; https://www.autohotkey.com/boards/viewtopic.php?f=76&t=103992&sid=c4836422e75965e4487395e6860c659f
CreateClass(string, interface := "", ByRef Class := "") {

   CreateHString(string, hString)

   if (interface = "")
      result := DllCall("Combase.dll\RoActivateInstance", "ptr", hString, "ptr*", Class, "uint")
   else    {
      VarSetCapacity(GUID, 16)
      DllCall("ole32\CLSIDFromString", "wstr", interface, "ptr", &GUID)
      result := DllCall("Combase.dll\RoGetActivationFactory", "ptr", hString, "ptr", &GUID, "ptr*", Class, "uint")
   }

   if (result != 0)    {
      if (result = 0x80004002)
         msgbox No such interface supported
      else if (result = 0x80040154)
         msgbox Class not registered
      else
         msgbox error: %result%
      ExitApp
   }

   DeleteHString(hString)
}
CreateHString(string, ByRef hString){
    DllCall("Combase.dll\WindowsCreateString", "wstr", string, "uint", StrLen(string), "ptr*", hString)
}
DeleteHString(hString){
   DllCall("Combase.dll\WindowsDeleteString", "ptr", hString)
}
WaitForAsync(ByRef Object){

   AsyncInfo := ComObjQuery(Object, IAsyncInfo := "{00000036-0000-0000-C000-000000000046}")
   loop   {
      DllCall(NumGet(NumGet(AsyncInfo+0)+7*A_PtrSize), "ptr", AsyncInfo, "uint*", status)   ; IAsyncInfo.Status
      if (status != 0)      {
         if (status != 1)         {
            DllCall(NumGet(NumGet(AsyncInfo+0)+8*A_PtrSize), "ptr", AsyncInfo, "uint*", ErrorCode)   ; IAsyncInfo.ErrorCode
            msgbox AsyncInfo status error: %ErrorCode%
            ExitApp
         }
         ObjRelease(AsyncInfo)
         break
      }
      sleep 10
   }
   DllCall(NumGet(NumGet(Object+0)+8*A_PtrSize), "ptr", Object, "ptr*", ObjectResult)   ; GetResults
   ObjReleaseClose(Object)
   Object := ObjectResult

}
ObjReleaseClose(ByRef Object) {

   if Object   {
      if (Close := ComObjQuery(Object, IClosable := "{30D5A829-7FA4-4026-83BB-D75BAE4EA99E}"))      {
         DllCall(NumGet(NumGet(Close+0)+6*A_PtrSize), "ptr", Close)   ; Close
         ObjRelease(Close)
      }
      ObjRelease(Object)
   }
   Object := ""

}
