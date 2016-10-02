Attribute VB_Name = "wow64ext"
Option Explicit

'USER32
Public Declare Function CallWindowProcW Lib "USER32" (ByVal lpCode As Long, Optional ByVal lParam1 As Long, Optional ByVal lParam2 As Long, Optional ByVal lParam3 As Long, Optional ByVal lParam4 As Long) As Long

Private Const INIT                  As String = "<INIT>"
Private Const code                  As String = "<CODE>"
Private Const FINI                  As String = "<FINI>"
Private Const THUNK_CALLCODE        As String = INIT & "6A33E80000000083042405CB" & code & "E800000000C7442404230000008304240DCB" & FINI
Private ASM_CALLCODE(0 To 255)      As Byte

Private Const IMAGE_DOS_SIGNATURE As Integer = &H5A4D
Private Const IMAGE_NT_SIGNATURE As Long = &H4550&

Private Const IMAGE_NUMBEROF_DIRECTORY_ENTRIES As Long = 16&
Private Const IMAGE_DIRECTORY_ENTRY_EXPORT As Integer = 0

Private Type LIST_ENTRY64
    Flink                           As Currency
    Blink                           As Currency
End Type

Private Type UNICODE_STRING64
    Length                          As Integer
    MaximumLength                   As Integer
    dummy                           As Long
    Buffer                          As Currency
End Type

Private Type ANSI_STRING64
    Length                          As Integer
    MaximumLength                   As Integer
    dummy                           As Long
    Buffer                          As Currency
End Type

Private Type NT_TIB64
    ExceptionList                   As Currency
    StackBase                       As Currency
    StackLimit                      As Currency
    SubSystemTib                    As Currency
    FiberData                       As Currency
    ArbitraryUserPointer            As Currency
    Self                            As Currency
End Type

Private Type CLIENT_ID64
    UniqueProcess                   As Currency
    UniqueThread                    As Currency
End Type

Private Type TEB64
    NtTib                           As NT_TIB64
    EnvironmentPointer              As Currency
    ClientId                        As CLIENT_ID64
    ActiveRpcHandle                 As Currency
    ThreadLocalStoragePointer       As Currency
    ProcessEnvironmentBlock         As Currency
    LastErrorValue                  As Long
    CountOfOwnedCriticalSections    As Long
    CsrClientThread                 As Currency
    Win32ThreadInfo                 As Currency
    User32Reserved(0 To 25)         As Long
    'rest of the structure is not defined for now, as it is not needed
End Type

Private Type LDR_DATA_TABLE_ENTRY64
    InLoadOrderLinks                As LIST_ENTRY64
    InMemoryOrderLinks              As LIST_ENTRY64
    InInitializationOrderLinks      As LIST_ENTRY64
    DllBase                         As Currency
    EntryPoint                      As Currency
    SizeOfImage                     As Long
    dummy                           As Long
    FullDllName                     As UNICODE_STRING64
    BaseDllName                     As UNICODE_STRING64
    Flags                           As Long
    LoadCount                       As Integer
    TlsIndex                        As Integer
    HashLinks                       As LIST_ENTRY64
    LoadedImports                   As Currency
    EntryPointActivationContext     As Currency
    PatchInformation                As Currency
    ForwarderLinks                  As LIST_ENTRY64
    ServiceTagLinks                 As LIST_ENTRY64
    StaticLinks                     As LIST_ENTRY64
    ContextInformation              As Currency
    OriginalBase                    As Currency
    LoadTime                        As Currency
End Type

Private Type PEB_LDR_DATA64
    Length                          As Long
    Initialized                     As Long
    SsHandle                        As Currency
    InLoadOrderModuleList           As LIST_ENTRY64
    InMemoryOrderModuleList         As LIST_ENTRY64
    InInitializationOrderModuleList As LIST_ENTRY64
    EntryInProgress                 As Currency
    ShutdownInProgress              As Long
    dummy                           As Long
    ShutdownThreadId                As Currency
End Type

Private Type PEB64
    InheritedAddressSpace           As Byte
    ReadImageFileExecOptions        As Byte
    BeingDebugged                   As Byte
    BitField                        As Byte
    dummy01                         As Long
    Mutant                          As Currency
    ImageBaseAddress                As Currency
    ldr                             As Currency
    ProcessParameters               As Currency
    SubSystemData                   As Currency
    ProcessHeap                     As Currency
    FastPebLock                     As Currency
    AtlThunkSListPtr                As Currency
    IFEOKey                         As Currency
    CrossProcessFlags               As Currency
    UserSharedInfoPtr               As Currency
    SystemReserved                  As Long
    AtlThunkSListPtr32              As Long
    ApiSetMap                       As Currency
    TlsExpansionCounter             As Currency
    TlsBitmap                       As Currency
    TlsBitmapBits(0 To 1)           As Long
    ReadOnlySharedMemoryBase        As Currency
    HotpatchInformation             As Currency
    ReadOnlyStaticServerData        As Currency
    AnsiCodePageData                As Currency
    OemCodePageData                 As Currency
    UnicodeCaseTableData            As Currency
    NumberOfProcessors              As Long
    NtGlobalFlag                    As Long
    CriticalSectionTimeout          As Currency
    HeapSegmentReserve              As Currency
    HeapSegmentCommit               As Currency
    HeapDeCommitTotalFreeThreshold  As Currency
    HeapDeCommitFreeBlockThreshold  As Currency
    NumberOfHeaps                   As Long
    MaximumNumberOfHeaps            As Long
    ProcessHeaps                    As Currency
    GdiSharedHandleTable            As Currency
    ProcessStarterHelper            As Currency
    GdiDCAttributeList              As Currency
    LoaderLock                      As Currency
    OSMajorVersion                  As Long
    OSMinorVersion                  As Long
    OSBuildNumber                   As Integer
    OSCSDVersion                    As Integer
    OSPlatformId                    As Long
    ImageSubsystem                  As Long
    ImageSubsystemMajorVersion      As Long
    ImageSubsystemMinorVersion      As Currency
    ActiveProcessAffinityMask       As Currency
    GdiHandleBuffer(0 To 29)        As Currency
    PostProcessInitRoutine          As Currency
    TlsExpansionBitmap              As Currency
    TlsExpansionBitmapBits(0 To 31) As Long
    SessionId                       As Currency
    AppCompatFlags                  As Currency
    AppCompatFlagsUser              As Currency
    pShimData                       As Currency
    AppCompatInfo                   As Currency
    CSDVersion                      As UNICODE_STRING64
    ActivationContextData           As Currency
    ProcessAssemblyStorageMap       As Currency
    SystemDefaultActivationContextData As Currency
    SystemAssemblyStorageMap        As Currency
    MinimumStackCommit              As Currency
    FlsCallback                     As Currency
    FlsListHead                     As LIST_ENTRY64
    FlsBitmap                       As Currency
    FlsBitmapBits(0 To 3)           As Long
    FlsHighIndex                    As Currency
    WerRegistrationData             As Currency
    WerShipAssertPtr                As Currency
    pContextData                    As Currency
    pImageHeaderHash                As Currency
    TracingFlags                    As Currency
End Type

Private Type IMAGE_DOS_HEADER
    e_magic                         As Integer
    e_cblp                          As Integer
    e_cp                            As Integer
    e_crlc                          As Integer
    e_cparhdr                       As Integer
    e_minalloc                      As Integer
    e_maxalloc                      As Integer
    e_ss                            As Integer
    e_sp                            As Integer
    e_csum                          As Integer
    e_ip                            As Integer
    e_cs                            As Integer
    e_lfarlc                        As Integer
    e_ovno                          As Integer
    e_res(0 To 3)                   As Integer
    e_oemid                         As Integer
    e_oeminfo                       As Integer
    e_res2(0 To 9)                  As Integer
    e_lfanew                        As Long
End Type

Private Type IMAGE_FILE_HEADER
    Machine                         As Integer
    NumberOfSections                As Integer
    TimeDateStamp                   As Long
    PointerToSymbolTable            As Long
    NumberOfSymbols                 As Long
    SizeOfOptionalHeader            As Integer
    Characteristics                 As Integer
End Type

Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress                  As Long
    Size                            As Long
End Type

Private Type IMAGE_OPTIONAL_HEADER64
    Magic                           As Integer
    MajorLinkerVersion              As Byte
    MinorLinkerVersion              As Byte
    SizeOfCode                      As Long
    SizeOfInitializedData           As Long
    SizeOfUnitializedData           As Long
    AddressOfEntryPoint             As Long
    BaseOfCode                      As Long
    ImageBase                       As Currency
    SectionAlignment                As Long
    FileAlignment                   As Long
    MajorOperatingSystemVersion     As Integer
    MinorOperatingSystemVersion     As Integer
    MajorImageVersion               As Integer
    MinorImageVersion               As Integer
    MajorSubsystemVersion           As Integer
    MinorSubsystemVersion           As Integer
    Win32VersionValue               As Long
    SizeOfImage                     As Long
    SizeOfHeaders                   As Long
    CheckSum                        As Long
    SubSystem                       As Integer
    DllCharacteristics              As Integer
    SizeOfStackReserve              As Currency
    SizeOfStackCommit               As Currency
    SizeOfHeapReserve               As Currency
    SizeOfHeapCommit                As Currency
    LoaderFlags                     As Long
    NumberOfRvaAndSizes             As Long
    DataDirectory(0 To 15)          As IMAGE_DATA_DIRECTORY
End Type

Private Type IMAGE_NT_HEADERS64
    Signature                       As Long
    FileHeader                      As IMAGE_FILE_HEADER
    OptionalHeader                  As IMAGE_OPTIONAL_HEADER64
End Type

Private Type IMAGE_EXPORT_DIRECTORY
    Characteristics                 As Long
    TimeDateStamp                   As Long
    MajorVersion                    As Integer
    MinorVersion                    As Integer
    name                            As Long
    Base                            As Long
    NumberOfFunctions               As Long
    NumberOfNames                   As Long
    AddressOfFunctions              As Long
    AddressOfNames                  As Long
    AddressOfNameOrdinals           As Long
End Type

Private Function calltest(ByVal func As Long, ParamArray args()) As Long
    Dim sThunk                      As String
    Dim argc                        As Long
    Dim params()                    As Long
    Dim i                           As Long
    
    argc = UBound(args) + 1&
    ReDim params(IIf(argc > 0&, argc - 1&, 0&))
    
    For i = LBound(args) To UBound(args)
        params(i) = CLng(args(i))
    Next
    
    sThunk = Replace$("CC 55 89 E5 81 E4 F0 FF FF FF 8B 4D 0C 89 C8 83 F8 04 73 02 B0 04 C1 E0 03 29 C4 8B 55 10 31 C0 39 C8 74 09 8B 1C 82 89 1C C4 40 EB F3 8B 4C 24 04 8B 45 08 FF 10 8B 55 14 89 02 89 EC 5D C3", " ", vbNullString)
    Call PutThunk(sThunk, ASM_CALLCODE)
    Call CallWindowProcW(VarPtr(ASM_CALLCODE(0)), VarPtr(func), argc, VarPtr(params(0)), VarPtr(calltest))
End Function

Private Function X64Call(ByVal func As Currency, ParamArray args()) As Currency
    Dim sThunk                      As String
    Dim argc                        As Long
    Dim params()                    As Currency
    Dim i                           As Long
    
    argc = UBound(args) + 1&
    ReDim params(IIf(argc > 0&, argc - 1&, 0&))
    
    For i = LBound(args) To UBound(args)
        params(i) = CCur(args(i))
    Next
    
    sThunk = Replace$(THUNK_CALLCODE, INIT, "5589E581E4F0FFFFFF")
    sThunk = Replace$(sThunk, code, "678B4D0C89C8A8017502FFC083F8047302B004C1E00329C4678B551031C039C8740D674C8B04C24C8904C4FFC0EBEF488B0C24488B5424084C8B4424104C8B4C2418678B450867FF10678B551467488902")
    sThunk = Replace$(sThunk, FINI, "89EC5DC3")
    Call PutThunk(sThunk, ASM_CALLCODE)
    Call CallWindowProcW(VarPtr(ASM_CALLCODE(0)), VarPtr(func), argc, VarPtr(params(0)), VarPtr(X64Call))
End Function

Private Sub GetMem64(ByVal dstMem As Long, ByVal srcMem As Currency, ByVal sz As Long)
    Dim sThunk                      As String
    
    If (dstMem = 0&) Or (srcMem = 0@) Or (sz = 0&) Then
        Exit Sub
    End If
    
    sThunk = Replace$(THUNK_CALLCODE, INIT, "57568B7C240C8B7424108B4C2414")
    sThunk = Replace$(sThunk, code, "488B3689C883E003C1E902F3A585C0740D83F801740766A583F8027401A4")
    sThunk = Replace$(sThunk, FINI, "5E5FC3")
    Call PutThunk(sThunk, ASM_CALLCODE)
    Call CallWindowProcW(VarPtr(ASM_CALLCODE(0)), dstMem, VarPtr(srcMem), sz)
End Sub

Private Function CmpMem64(ByVal dstMem As Long, ByVal srcMem As Currency, ByVal sz As Long) As Boolean
    Dim sThunk                      As String
    
    If (dstMem = 0&) Or (srcMem = 0@) Or (sz = 0&) Then
        Exit Function
    End If
    
    sThunk = Replace$(THUNK_CALLCODE, INIT, "57568B7C240C8B7424108B4C241431D2")
    sThunk = Replace$(sThunk, code, "488B3689C883E003C1E902F3A7751785C0741183F801740966A7750A83F8027403A67502FFC2")
    sThunk = Replace$(sThunk, FINI, "89D05E5FC3")
    Call PutThunk(sThunk, ASM_CALLCODE)
    CmpMem64 = CallWindowProcW(VarPtr(ASM_CALLCODE(0)), dstMem, VarPtr(srcMem), sz)
End Function

Private Function GetTEB64() As Currency
    Dim sThunk                      As String
    Dim bErr                        As Boolean
    
    sThunk = Replace$(THUNK_CALLCODE, INIT, "8B442404")
    sThunk = Replace$(sThunk, code, "4C8920")
    sThunk = Replace$(sThunk, FINI, "31C0C3")
    Call PutThunk(sThunk, ASM_CALLCODE)
    bErr = CallWindowProcW(VarPtr(ASM_CALLCODE(0)), VarPtr(GetTEB64))
    
    If bErr Then
        Err.Raise -1, , "Unable to get x64 TEB"
    End If
End Function

Private Function GetModuleHandle64(ByRef lpModuleName As String) As Currency
    Dim TEB64                       As TEB64
    Dim PEB64                       As PEB64
    Dim ldr64                       As PEB_LDR_DATA64
    Dim head                        As LDR_DATA_TABLE_ENTRY64
    Dim lastEntry                   As Currency
    Dim dllName                     As String
    
    Call GetMem64(VarPtr(TEB64), GetTEB64(), Len(TEB64))
    Call GetMem64(VarPtr(PEB64), TEB64.ProcessEnvironmentBlock, Len(PEB64))
    Call GetMem64(VarPtr(ldr64), PEB64.ldr, Len(ldr64))
    
    lastEntry = PEB64.ldr + ToCurrency(&H10&) '0x010 = offsetof(PEB_LDR_DATA64, InLoadOrderModuleList);
    head.InLoadOrderLinks.Flink = ldr64.InLoadOrderModuleList.Flink
    
    Do
        Call GetMem64(VarPtr(head), head.InLoadOrderLinks.Flink, Len(head))
        
        dllName = Space$(head.BaseDllName.Length \ 2)
        Call GetMem64(StrPtr(dllName), head.BaseDllName.Buffer, head.BaseDllName.Length)
        
        If (StrComp(lpModuleName, dllName, vbTextCompare) = 0) Then
            GetModuleHandle64 = head.DllBase
            Exit Function
        End If
    Loop While (head.InLoadOrderLinks.Flink <> lastEntry)
    
    Err.Raise -1, , "Module not found"
End Function

Private Function GetProcedureAddress64(ByRef funcName As String) As Currency
    Dim modBase                     As Currency
    Dim idh                         As IMAGE_DOS_HEADER
    Dim inh                         As IMAGE_NT_HEADERS64
    Dim ied                         As IMAGE_EXPORT_DIRECTORY
    Dim rvaTable()                  As Long
    Dim ordTable()                  As Integer
    Dim nameTable()                 As Long
    Dim name()                      As Byte
    Dim i                           As Long
    
    modBase = GetNTDLL64()
    Call GetMem64(VarPtr(idh), modBase, Len(idh))
    Call GetMem64(VarPtr(inh), modBase + ToCurrency(idh.e_lfanew), Len(inh))
    
    With inh.OptionalHeader.DataDirectory(IMAGE_DIRECTORY_ENTRY_EXPORT)
        If (.VirtualAddress = 0&) Then
            Exit Function
        End If
        
        Call GetMem64(VarPtr(ied), modBase + ToCurrency(.VirtualAddress), Len(ied))
    End With
    
    ReDim rvaTable(ied.NumberOfFunctions - 1)
    Call GetMem64(VarPtr(rvaTable(0)), modBase + ToCurrency(ied.AddressOfFunctions), ied.NumberOfFunctions * &H4&) '0x004 = sizeof(DWORD)
    
    ReDim ordTable(ied.NumberOfNames - 1)
    Call GetMem64(VarPtr(ordTable(0)), modBase + ToCurrency(ied.AddressOfNameOrdinals), ied.NumberOfNames * &H2&) '0x002 = sizeof(WORD)
    
    ReDim nameTable(ied.NumberOfNames - 1)
    Call GetMem64(VarPtr(nameTable(0)), modBase + ToCurrency(ied.AddressOfNames), ied.NumberOfNames * &H4&) '0x004 = sizeof(DWORD)
    
    'name = StrConv("LdrGetProcedureAddress" + vbNullChar, vbFromUnicode)
    name = StrConv(funcName + vbNullChar, vbFromUnicode)
    
    ' lazy search, there is no need to use binsearch for just one function
    For i = 0 To ied.NumberOfNames - 1&
        If (CmpMem64(VarPtr(name(0)), modBase + ToCurrency(nameTable(i)), UBound(name) + 1&)) Then
            GetProcedureAddress64 = modBase + ToCurrency(rvaTable(ordTable(i)))
            Exit Function
        End If
    Next
    
    Err.Raise -1, , "Function not found"
End Function

Private Sub PutThunk(ByVal sThunk As String, ByRef bvRet() As Byte)
    Dim i                           As Long
 
    For i = 0 To Len(sThunk) - 1 Step 2
        bvRet((i / 2)) = CByte("&H" & Mid$(sThunk, i + 1&, 2))
    Next i
End Sub

Public Function GetNTDLL64() As Currency
    GetNTDLL64 = GetModuleHandle64("ntdll.dll")
End Function

Public Function GetLdrGetProcedureAddress64() As Currency
    GetLdrGetProcedureAddress64 = GetProcedureAddress64("LdrGetProcedureAddress")
End Function

Public Function GetProcAddress64(ByVal hModule As Currency, ByRef funcName As String) As Currency
    Dim LdrGetProcedureAddress      As Currency
    Dim fName                       As ANSI_STRING64
    Dim name()                      As Byte
    Dim funcRet                     As Currency
    
    LdrGetProcedureAddress = GetLdrGetProcedureAddress64
    
    name = StrConv(funcName + vbNullChar, vbFromUnicode)
    fName.Length = Len(funcName)
    fName.MaximumLength = fName.Length + 1
    fName.Buffer = ToCurrency(VarPtr(name(0)))
    
    Call X64Call(LdrGetProcedureAddress, hModule, ToCurrency(VarPtr(fName)), 0@, ToCurrency(VarPtr(GetProcAddress64)))
End Function

Public Function ToCurrency(ByVal lVal As Long) As Currency
    ToCurrency = lVal * 0.0001@
End Function
