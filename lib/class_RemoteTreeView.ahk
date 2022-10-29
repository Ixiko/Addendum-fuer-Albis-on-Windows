;--------------------------------------------------------------------------------------------------
; Title:  Remote TreeView class
;         This class allows a script to work with TreeViews controlled by a third party process.
;
;         8/30/2012  Released for beta testing
;
;         8/31/2012  Added a wait parameter to SetSelection
;                    Changed name of ExpandCollapse to Expand
;                    Changed default WaitTime of Expand to 0
;
;         9/2/2012	 Removed GetState method and replaced it with the IsBold, IsChecked, IsExpanded
; 					 and IsSelected methods.
;
;         9/7/2012   Added Check method.
;                    For ease of use, changed parameters of SetSelection, Expand and IsChecked methods.
;
;         9/17/2012  Extended the "Full" option of GetNext() to allow it to transverse sub-trees.
;                    Given a starting node, all decendents of that node will be  transversed depth
;                    first. Those nodes which are not descendants of the starting node will NOT be
;                    visited. To traverse the entire tree, including the real root, pass zero as the
;                    starting node.
;
;         11/02/2014 Fix for GetText and ddditional function from just me
;                    See http://ahkscript.org/boards/viewtopic.php?f=5&t=4998#p29339
;
class RemoteTreeView{

	__New(TVHnd)	{
		;----------------------------------------------------------------------------------------------
		; Method: __New
		;         Stores the TreeView's Control HWnd in the object for later use
		;
		; Parameters:
		;         TVHnd			- HWND of the TreeView control
		;
		; Returns:
		;         Nothing
		;
		this.TVHnd := TVHnd
	}

	SetSelection(pItem)	{
		;----------------------------------------------------------------------------------------------
		; Method: SetSelection
		;         Makes the given item selected and expanded. Optionally scrolls the
		;         TreeView so the item is visible.
		;
		; Parameters:
		;         pItem			- Handle to the item you wish selected
		;
		; Returns:
		;         TRUE if successful, or FALSE otherwise
		;
		; ORI SendMessage TVM_SELECTITEM, TVGN_CARET|TVSI_NOSINGLEEXPAND, pItem, , % "ahk_id " this.TVHnd
		; sle118 :
		SendMessage TVM_SELECTITEM, TVGN_FIRSTVISIBLE, pItem, , % "ahk_id " this.TVHnd
		return ErrorLevel
	}

	GetSelection()	{
		;----------------------------------------------------------------------------------------------
		; Method: GetSelection
		;         Retrieves the currently selected item
		;
		; Parameters:
		;         None
		;
		; Returns:
		;         Handle to the selected item if successful, Null otherwise.
		;
		SendMessage TVM_GETNEXTITEM, TVGN_CARET, 0, , % "ahk_id " this.TVHnd
		return ErrorLevel
	}

	GetRoot()	{
		;----------------------------------------------------------------------------------------------
		; Method: GetRoot
		;         Retrieves the root item of the treeview
		;
		; Parameters:
		;         None
		;
		; Returns:
		;         Handle to the topmost or very first item of the tree-view control
		;         if successful, NULL otherwise.
		;
		SendMessage TVM_GETNEXTITEM, TVGN_ROOT, 0, , % "ahk_id " this.TVHnd
		return ErrorLevel
	}

	GetParent(pItem)	{
		;----------------------------------------------------------------------------------------------
		; Method: GetParent
		;         Retrieves an item's parent
		;
		; Parameters:
		;         pItem			- Handle to the item
		;
		; Returns:
		;         Handle to the parent of the specified item. Returns
		;         NULL if the item being retrieved is the root node of the tree.
		;
		SendMessage TVM_GETNEXTITEM, TVGN_PARENT, pItem, , % "ahk_id " this.TVHnd
		return ErrorLevel
	}

	GetChild(pItem)	{
		;----------------------------------------------------------------------------------------------
		; Method: GetChild
		;         Retrieves an item's first child
		;
		; Parameters:
		;         pItem			- Handle to the item
		;
		; Returns:
		;         Handle to the first Child of the specified item, NULL otherwise.
		;
		SendMessage TVM_GETNEXTITEM, TVGN_CHILD, pItem, , % "ahk_id " this.TVHnd
		return ErrorLevel
	}

	GetNext(pItem = 0, flag = "")	{

		;----------------------------------------------------------------------------------------------
		; Method: GetNext
		;         Returns the handle of the sibling below the specified item (or 0 if none).
		;
		; Parameters:
		;         pItem			- (Optional) Handle to the item
		;
		;         flag          - (Optional) "FULL" or "F"
		;
		; Returns:
		;         This method has the following modes:
		;              1. When all parameters are omitted, it returns the handle
		;                 of the first/top item in the TreeView (or 0 if none).
		;
		;              2. When the only first parameter (pItem) is present, it returns the
		;                 handle of the sibling below the specified item (or 0 if none).
		;                 If the first parameter is 0, it returns the handle of the first/top
		;                 item in the TreeView (or 0 if none).
		;
		;              3. When the second parameter is "Full" or "F", the first time GetNext()
		;                 is called the hItem passed is considered the "root" of a sub-tree that
		;                 will be transversed in a depth first manner. No nodes except the
		;                 decendents of that "root" will be visited. To traverse the entire tree,
		;                 including the real root, pass zero in the first call.
		;
		;                 When all descendants have benn visited, the method returns zero.
		;
		; Example:
		;				hItem = 0  ; Start the search at the top of the tree.
		;				Loop
		;				{
		;					hItem := MyTV.GetNext(hItem, "Full")
		;					if not hItem  ; No more items in tree.
		;						break
		;					ItemText := MyTV.GetText(hItem)
		;					MsgBox The next Item is %hItem%, whose text is "%ItemText%".
		;				}
		;

		static Root := -1

		if (RegExMatch(flag, "i)^\s*(F|Full)\s*$")) {
			if (Root = -1) {
				Root := pItem
			}
			SendMessage TVM_GETNEXTITEM, TVGN_CHILD, pItem, , % "ahk_id " this.TVHnd
			if (ErrorLevel = 0) {
				SendMessage TVM_GETNEXTITEM, TVGN_NEXT, pItem, , % "ahk_id " this.TVHnd
				if (ErrorLevel = 0) {
					Loop {
						SendMessage TVM_GETNEXTITEM, TVGN_PARENT, pItem, , % "ahk_id " this.TVHnd
						if (ErrorLevel = Root) {
							Root := -1
							return 0
						}
						pItem := ErrorLevel
						SendMessage TVM_GETNEXTITEM, TVGN_NEXT, pItem, , % "ahk_id " this.TVHnd
					} until ErrorLevel
				}
			}
			return ErrorLevel
		}

		Root := -1
		if (!pItem)
			SendMessage TVM_GETNEXTITEM, TVGN_ROOT, 0, , % "ahk_id " this.TVHnd
		else
			SendMessage TVM_GETNEXTITEM, TVGN_NEXT, pItem, , % "ahk_id " this.TVHnd
		return ErrorLevel
	}

	GetPrev(pItem)	{
		;----------------------------------------------------------------------------------------------
		; Method: GetPrev
		;         Returns the handle of the sibling above the specified item (or 0 if none).
		;
		; Parameters:
		;         pItem			- Handle to the item
		;
		; Returns:
		;         Handle of the sibling above the specified item (or 0 if none).
		;
		SendMessage TVM_GETNEXTITEM, TVGN_PREVIOUS, pItem, , % "ahk_id " this.TVHnd
		return ErrorLevel
	}

	Expand(pItem, DoExpand = true)	{
		;----------------------------------------------------------------------------------------------
		; Method: Expand
		;         Expands or collapses the specified tree node
		;
		; Parameters:
		;         pItem			- Handle to the item
		;
		;         flag			- Determines whether the node is expanded or collapsed.
		;                         true : expand the node (default)
		;                         false : collapse the node
		;
		;
		; Returns:
		;         Nonzero if the operation was successful, or zero otherwise.
		;
		flag := DoExpand ? TVE_EXPAND : TVE_COLLAPSE
		SendMessage TVM_EXPAND, flag, pItem, , % "ahk_id " this.TVHnd
		return ErrorLevel
	}

	Check(pItem, fCheck, Force = true)	{

		;----------------------------------------------------------------------------------------------
		; Method: Check
		;         Changes the state of a treeview item's check box
		;
		; Parameters:
		;         pItem			- Handle to the item
		;
		;         fCheck        - If true, check the node
		;                         If false, uncheck the node
		;
		;         Force			- If true (default), prevents this method from failing due to
		;                         the node having an invalid initial state. See IsChecked
		;                         method for more info.
		;
		; Returns:
		;         Returns true if if successful, otherwise false
		;
		; Remarks:
		;         This method makes pItem the current selection.
		;
		SavedDelay := A_KeyDelay
		SetKeyDelay 30

		CurrentState := this.IsChecked(pItem, false)
		if (CurrentState = -1)
			if (Force) {
				ControlSend, , {Space}, % "ahk_id " this.TVHnd
				CurrentState := this.IsChecked(pItem, false)
			}
			else
				return false

		if (CurrentState and not fCheck) or (not CurrentState and fCheck )
			ControlSend, , {Space}, % "ahk_id " this.TVHnd

		SetKeyDelay %SavedDelay%
		return true
	}

    GetText(pItem)    {

		;----------------------------------------------------------------------------------------------
		; Method: GetText
		;         Retrieves the text/name of the specified node
		;
		; Parameters:
		;         pItem         - Handle to the item
		;
		; Returns:
		;         The text/name of the specified Item. If the text is longer than 127, only
		;         the first 127 characters are retrieved.
		;
		; Fix from just me (http://ahkscript.org/boards/viewtopic.php?f=5&t=4998#p29339)
		;

        TVM_GETITEM := A_IsUnicode ? TVM_GETITEMW : TVM_GETITEMA

        WinGet ProcessId, pid, % "ahk_id " this.TVHnd
        hProcess := OpenProcess(PROCESS_VM_OPERATION|PROCESS_VM_READ
                               |PROCESS_VM_WRITE|PROCESS_QUERY_INFORMATION
                               , false, ProcessId)

        ; Try to determine the bitness of the remote tree-view's process
        ProcessIs32Bit := A_PtrSize = 8 ? False : True
        If (A_Is64bitOS) && DllCall("Kernel32.dll\IsWow64Process", "Ptr", hProcess, "UIntP", WOW64)
            ProcessIs32Bit := WOW64

        Size := ProcessIs32Bit ?  60 : 80 ; Size of a TVITEMEX structure

        _tvi := VirtualAllocEx(hProcess, 0, Size, MEM_COMMIT, PAGE_READWRITE)
        _txt := VirtualAllocEx(hProcess, 0, 256,  MEM_COMMIT, PAGE_READWRITE)

        ; TVITEMEX Structure
        VarSetCapacity(tvi, Size, 0)
        NumPut(TVIF_TEXT|TVIF_HANDLE, tvi, 0, "UInt")
        If (ProcessIs32Bit)
        {
            NumPut(pItem, tvi,  4, "UInt")
            NumPut(_txt , tvi, 16, "UInt")
            NumPut(127  , tvi, 20, "UInt")
        }
        Else
        {
            NumPut(pItem, tvi,  8, "UInt64")
            NumPut(_txt , tvi, 24, "UInt64")
            NumPut(127  , tvi, 32, "UInt")
        }

        VarSetCapacity(txt, 256, 0)
        WriteProcessMemory(hProcess, _tvi, &tvi, Size)
        SendMessage TVM_GETITEM, 0, _tvi, ,  % "ahk_id " this.TVHnd
        ReadProcessMemory(hProcess, _txt, txt, 256)

        VirtualFreeEx(hProcess, _txt, 0, MEM_RELEASE)
        VirtualFreeEx(hProcess, _tvi, 0, MEM_RELEASE)
        CloseHandle(hProcess)

        return txt
    }

	EditLabel(pItem)	{
		;----------------------------------------------------------------------------------------------
		; Method: EditLabel
		;         Begins in-place editing of the specified item's text, replacing the text of the
		;         item with a single-line edit control containing the text. This method implicitly
		;         selects and focuses the specified item.
		;
		; Parameters:
		;         pItem			- Handle to the item
		;
		; Returns:
		;         Returns the handle to the edit control used to edit the item text if successful,
		;         or NULL otherwise. When the user completes or cancels editing, the edit control
		;         is destroyed and the handle is no longer valid.
		;
		TVM_EDITLABEL := A_IsUnicode ? TVM_EDITLABELW : TVM_EDITLABELA
		SendMessage TVM_EDITLABEL, 0, pItem, , % "ahk_id " this.TVHnd
		return ErrorLevel
	}

	GetCount()	{
		;----------------------------------------------------------------------------------------------
		; Method: GetCount
		;         Returns the total number of expanded items in the control
		;
		; Parameters:
		;         None
		;
		; Returns:
		;         Returns the total number of expanded items in the control
		;
		SendMessage TVM_GETCOUNT, 0, 0, , % "ahk_id " this.TVHnd
		return ErrorLevel
	}

	IsChecked(pItem, Force = true)	{
		;----------------------------------------------------------------------------------------------
		; Method: IsChecked
		;         Retrieves an item's checked status
		;
		; Parameters:
		;         pItem			- Handle to the item
		;
		;         Force			- If true (default), forces the node to return a valid state.
		;                         Since this involves toggling the state of the check box, it
		;                         can have undesired side effects. Make this false to disable
		;                         this feature.
		; Returns:
		;         Returns 1 if the item is checked, 0 if unchecked.
		;
		;         Returns -1 if the checkbox state cannot be determined because no checkbox
		;         image is currently associated with the node and Force is false.
		;
		; Remarks:
		;         Due to a "feature" of Windows, a checkbox can be displayed even if no checkbox image
		;         is associated with the node. It is important to either check the actual value returned
		;         or make the Force parameter true.
		;
		;         This method makes pItem the current selection.
		;
		SavedDelay := A_KeyDelay
		SetKeyDelay 30

		this.SetSelection(pItem)
		SendMessage TVM_GETITEMSTATE, pItem, 0, , % "ahk_id " this.TVHnd
		State := ((ErrorLevel & TVIS_STATEIMAGEMASK) >> 12) - 1

		if (State = -1 and Force) {
			ControlSend, , {Space 2}, % "ahk_id " this.TVHnd
			SendMessage TVM_GETITEMSTATE, pItem, 0, , % "ahk_id " this.TVHnd
			State := ((ErrorLevel & TVIS_STATEIMAGEMASK) >> 12) - 1
		}

		SetKeyDelay %SavedDelay%
		return State
	}

	IsBold(pItem)	{
		;----------------------------------------------------------------------------------------------
		; Method: IsBold
		;         Check if a node is in bold font
		;
		; Parameters:
		;         pItem			- Handle to the item
		;
		; Returns:
		;         Returns true if the item is in bold, false otherwise.
		;
		SendMessage TVM_GETITEMSTATE, pItem, 0, , % "ahk_id " this.TVHnd
		return (ErrorLevel & TVIS_BOLD) >> 4
	}

	IsExpanded(pItem)	{
	;----------------------------------------------------------------------------------------------
	; Method: IsExpanded
	;         Check if a node has children and is expanded
	;
	; Parameters:
	;         pItem			- Handle to the item
	;
	; Returns:
	;         Returns true if the item has children and is expanded, false otherwise.
	;
		SendMessage TVM_GETITEMSTATE, pItem, 0, , % "ahk_id " this.TVHnd
		return (ErrorLevel & TVIS_EXPANDED) >> 5
	}

	IsSelected(pItem)	{
	;----------------------------------------------------------------------------------------------
	; Method: IsSelected
	;         Check if a node is Selected
	;
	; Parameters:
	;         pItem			- Handle to the item
	;
	; Returns:
	;         Returns true if the item is selected, false otherwise.
	;
		SendMessage TVM_GETITEMSTATE, pItem, 0, , % "ahk_id " this.TVHnd
		return (ErrorLevel & TVIS_SELECTED) >> 1
	}

}


;====================================================================================
;
;	Functions
;
;====================================================================================


OpenProcess(DesiredAccess, InheritHandle, ProcessId){
	;----------------------------------------------------------------------------------------------
	; Function: OpenProcess
	;         Opens an existing local process object.
	;
	; Parameters:
	;         DesiredAccess - The desired access to the process object.
	;
	;         InheritHandle - If this value is TRUE, processes created by this process will inherit
	;                         the handle. Otherwise, the processes do not inherit this handle.
	;
	;         ProcessId     - The Process ID of the local process to be opened.
	;
	; Returns:
	;         If the function succeeds, the return value is an open handle to the specified process.
	;         If the function fails, the return value is NULL.
	;
	return DllCall("OpenProcess"
	             , "Int", DesiredAccess
				 , "Int", InheritHandle
				 , "Int", ProcessId
				 , "Ptr")
}

CloseHandle(hObject){
;----------------------------------------------------------------------------------------------
; Function: CloseHandle
;         Closes an open object handle.
;
; Parameters:
;         hObject       - A valid handle to an open object
;
; Returns:
;         If the function succeeds, the return value is nonzero.
;         If the function fails, the return value is zero.
;
	return DllCall("CloseHandle"
	             , "Ptr", hObject
				 , "Int")
}

VirtualAllocEx(hProcess, Address, Size, AllocationType, ProtectType){
;----------------------------------------------------------------------------------------------
; Function: VirtualAllocEx
;         Reserves or commits a region of memory within the virtual address space of the
;         specified process, and specifies the NUMA node for the physical memory.
;
; Parameters:
;         hProcess      - A valid handle to an open object. The handle must have the
;                         PROCESS_VM_OPERATION access right.
;
;         Address       - The pointer that specifies a desired starting address for the region
;                         of pages that you want to allocate.
;
;                         If you are reserving memory, the function rounds this address down to
;                         the nearest multiple of the allocation granularity.
;
;                         If you are committing memory that is already reserved, the function rounds
;                         this address down to the nearest page boundary. To determine the size of a
;                         page and the allocation granularity on the host computer, use the GetSystemInfo
;                         function.
;
;                         If Address is NULL, the function determines where to allocate the region.
;
;         Size          - The size of the region of memory to be allocated, in bytes.
;
;         AllocationType - The type of memory allocation. This parameter must contain ONE of the
;                          following values:
;								MEM_COMMIT
;								MEM_RESERVE
;								MEM_RESET
;
;         ProtectType   - The memory protection for the region of pages to be allocated. If the
;                         pages are being committed, you can specify any one of the memory protection
;                         constants:
;								 PAGE_NOACCESS
;								 PAGE_READONLY
;								 PAGE_READWRITE
;								 PAGE_WRITECOPY
;								 PAGE_EXECUTE
;								 PAGE_EXECUTE_READ
;								 PAGE_EXECUTE_READWRITE
;								 PAGE_EXECUTE_WRITECOPY
;
; Returns:
;         If the function succeeds, the return value is the base address of the allocated region of pages.
;         If the function fails, the return value is NULL.
;
	return DllCall("VirtualAllocEx"
				 , "Ptr", hProcess
				 , "Ptr", Address
				 , "UInt", Size
				 , "UInt", AllocationType
				 , "UInt", ProtectType
				 , "Ptr")
}

VirtualFreeEx(hProcess, Address, Size, FType){
	;----------------------------------------------------------------------------------------------
	; Function: VirtualFreeEx
	;         Releases, decommits, or releases and decommits a region of memory within the
	;         virtual address space of a specified process
	;
	; Parameters:
	;         hProcess      - A valid handle to an open object. The handle must have the
	;                         PROCESS_VM_OPERATION access right.
	;
	;         Address       - The pointer that specifies a desired starting address for the region
	;                         to be freed. If the dwFreeType parameter is MEM_RELEASE, lpAddress
	;                         must be the base address returned by the VirtualAllocEx function when
	;                         the region is reserved.
	;
	;         Size          - The size of the region of memory to be allocated, in bytes.
	;
	;                         If the FreeType parameter is MEM_RELEASE, dwSize must be 0 (zero). The function
	;                         frees the entire region that is reserved in the initial allocation call to
	;                         VirtualAllocEx.
	;
	;                         If FreeType is MEM_DECOMMIT, the function decommits all memory pages that
	;                         contain one or more bytes in the range from the Address parameter to
	;                         (lpAddress+dwSize). This means, for example, that a 2-byte region of memory
	;                         that straddles a page boundary causes both pages to be decommitted. If Address
	;                         is the base address returned by VirtualAllocEx and dwSize is 0 (zero), the
	;                         function decommits the entire region that is allocated by VirtualAllocEx. After
	;                         that, the entire region is in the reserved state.
	;
	;         FreeType      - The type of free operation. This parameter can be one of the following values:
	;								MEM_DECOMMIT
	;								MEM_RELEASE
	;
	; Returns:
	;         If the function succeeds, the return value is a nonzero value.
	;         If the function fails, the return value is 0 (zero).
	;
	return DllCall("VirtualFreeEx"
				 , "Ptr", hProcess
				 , "Ptr", Address
				 , "UINT", Size
				 , "UInt", FType
				 , "Int")
}

WriteProcessMemory(hProcess, BaseAddress, Buffer, Size, ByRef NumberOfBytesWritten = 0){
;----------------------------------------------------------------------------------------------
; Function: WriteProcessMemory
;         Writes data to an area of memory in a specified process. The entire area to be written
;         to must be accessible or the operation fails
;
; Parameters:
;         hProcess      - A valid handle to an open object. The handle must have the
;                         PROCESS_VM_WRITE and PROCESS_VM_OPERATION access right.
;
;         BaseAddress   - A pointer to the base address in the specified process to which data
;                         is written. Before data transfer occurs, the system verifies that all
;                         data in the base address and memory of the specified size is accessible
;                         for write access, and if it is not accessible, the function fails.
;
;         Buffer        - A pointer to the buffer that contains data to be written in the address
;                         space of the specified process.
;
;         Size          - The number of bytes to be written to the specified process.
;
;         NumberOfBytesWritten
;                       - A pointer to a variable that receives the number of bytes transferred
;                         into the specified process. This parameter is optional. If NumberOfBytesWritten
;                         is NULL, the parameter is ignored.
;
; Returns:
;         If the function succeeds, the return value is a nonzero value.
;         If the function fails, the return value is 0 (zero).
;
	return DllCall("WriteProcessMemory"
				 , "Ptr", hProcess
				 , "Ptr", BaseAddress
				 , "Ptr", Buffer
				 , "Uint", Size
				 , "UInt*", NumberOfBytesWritten
				 , "Int")
}

ReadProcessMemory(hProcess, BaseAddress, ByRef Buffer, Size, ByRef NumberOfBytesRead = 0){
	;----------------------------------------------------------------------------------------------
	; Function: ReadProcessMemory
	;         Reads data from an area of memory in a specified process. The entire area to be read
	;         must be accessible or the operation fails
	;
	; Parameters:
	;         hProcess      - A valid handle to an open object. The handle must have the
	;                         PROCESS_VM_READ access right.
	;
	;         BaseAddress   - A pointer to the base address in the specified process from which to
	;                         read. Before any data transfer occurs, the system verifies that all data
	;                         in the base address and memory of the specified size is accessible for read
	;                         access, and if it is not accessible the function fails.
	;
	;         Buffer        - A pointer to a buffer that receives the contents from the address space
	;                         of the specified process.
	;
	;         Size          - The number of bytes to be read from the specified process.
	;
	;         NumberOfBytesWritten
	;                       - A pointer to a variable that receives the number of bytes transferred
	;                         into the specified buffer. If lpNumberOfBytesRead is NULL, the parameter
	;                         is ignored.
	;
	; Returns:
	;         If the function succeeds, the return value is a nonzero value.
	;         If the function fails, the return value is 0 (zero).
	;
	return DllCall("ReadProcessMemory"
	             , "Ptr", hProcess
				 , "Ptr", BaseAddress
				 , "Ptr", &Buffer
				 , "UInt", Size
				 , "UInt*", NumberOfBytesRead
				 , "Int")
}

