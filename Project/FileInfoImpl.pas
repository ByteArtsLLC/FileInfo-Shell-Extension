unit FileInfoImpl;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  System.Win.ComObj, Winapi.ActiveX, FileInfo_TLB, System.Win.StdVCL,
  Winapi.ShlObj, Winapi.Windows, Winapi.ShellAPI;

type

  TFileInfoContextMenu = class(TAutoObject, IFileInfoContextMenu, IShellExtInit, IContextMenu)
    private
      FFileName: string;
      FMenuItemIndex: UINT;
      procedure ShowFileInfo(WinHandle: HWND; const Filename: string);

    protected
      { IShellExtInit Methods }
      { Initialize the context menu if a files was selected }
      function IShellExtInit.Initialize = ShellExtInitialize;
      function ShellExtInitialize(
            pidlFolder: PItemIDList;
            lpdobj: IDataObject;
            hKeyProgID: HKEY): HResult; stdcall;

      { IContextMenu Methods }
      { Initializes the context menu and it decides which items appear in it,
        based on the flags you pass }
      function QueryContextMenu(
            Menu: HMENU;
            IndexMenu, CmdFirst, CmdLast, Flags: UINT): HResult; stdcall;

      { Execute the command, which will be the upload to Amazon or Azure }
      function InvokeCommand(var lpici: TCMInvokeCommandInfo): HResult; stdcall;
      { Set help string on the Explorer status bar when the menu item is selected }
      function GetCommandString(
            idCmd: UINT_PTR;
            uFlags: UINT;
            pwReserved: PUINT;
            pszName: LPSTR;
            cchMax: UINT): HResult; stdcall;

  end;

  { the new class factory }
  TFileInfoObjectFactory = class(TAutoObjectFactory)
    protected
      procedure ApproveShellExtension(
            &Register: Boolean;
            const ClsID: string);
      function GetProgID: string; override;
    public
      procedure UpdateRegistry(Register: Boolean); override;
  end;

implementation

uses System.Win.ComServ, System.SysUtils, System.Win.Registry;

const
  MENU_ITEM_INDEX = 0;

type
  TFileTimeKind = (ftkCreated, ftkModified);

function FileTime(const Filename: string; const Kind: TFileTimeKind): string;
var
  fadValue: TWin32FileAttributeData;
  stTime: TSystemTime;
  ftLocal: WinApi.Windows._FILETIME;
  dtValue: TDateTime;
begin
  Result := 'error';

  if not GetFileAttributesEx(PChar(Filename), GetFileExInfoStandard, @fadValue) then
    exit;

  if (Kind = ftkCreated) then
  begin
    if not FileTimeToLocalFileTime(fadValue.ftCreationTime, ftLocal) then
      exit;
  end
  else
  begin
    if not FileTimeToLocalFileTime(fadValue.ftLastWriteTime, ftLocal) then
      exit;
  end;

  if not FileTimeToSystemTime(ftLocal, stTime) then
    exit;

  dtValue := SystemTimeToDateTime(stTime);

  Result := Format('%s %s', [DateToStr(dtValue), TimeToStr(dtValue)]);
end;


function GetFileSize(const Filename: string): Int64;
var
  srRecord: TSearchRec;
begin
  Result := 0;

  if FileExists(Filename) then
  begin
    if FindFirst(Filename, faNormal, srRecord) = 0 then
      Result := srRecord.Size;
  end;
end;


{ TFileInfoImpl }


function TFileInfoContextMenu.GetCommandString(
      idCmd: UINT_PTR;
      uFlags: UINT;
      pwReserved: PUINT;
      pszName: LPSTR;
      cchMax: UINT): HResult;
begin
  Result := E_INVALIDARG;

  if (idCmd = MENU_ITEM_INDEX) and (uFlags = GCS_HELPTEXT) then
  begin
    StrLCopy(PWideChar(pszName), PWideChar('gGet file info'), cchMax);
    Result := NOERROR;
  end;
end;


function TFileInfoContextMenu.InvokeCommand(var lpici: TCMInvokeCommandInfo): HResult;
var
  wIndex: Word;
begin
  Result := E_FAIL;

  wIndex := HiWord(Integer(lpici.lpVerb));
  if (wIndex <> 0) then
    Exit;

  { if the index matches the index for the menu, show the options }
  wIndex := LoWord(Integer(lpici.lpVerb));

  //OutputDebugString(PWideChar(Format('wIndex=%d FMenuItemIndex=%d', [wIndex, FMenuItemIndex])));

  if (wIndex = MENU_ITEM_INDEX) then
  begin
    try
      ShowFileInfo(lpici.hwnd, FFileName);

    except
      on E: Exception do
        MessageBox(lpici.hwnd, PWideChar(E.Message), 'FileInfo', MB_ICONERROR);

    end;
    Result := NOERROR;
  end;
end;


function TFileInfoContextMenu.QueryContextMenu(
      Menu: HMENU;
      IndexMenu, CmdFirst, CmdLast, Flags: UINT): HResult;
var
  mnuFileInfoItem: TMenuItemInfo;
  sMenuCaption: String;
begin
  { only adding one menu FileInfoItem, so generate the result code accordingly }
  Result := MakeResult(SEVERITY_SUCCESS, 0, 3);

  { store the menu FileInfoItem index }
  FMenuItemIndex := IndexMenu;

  { specify what the menu says, depending on where it was spawned }
  if (Flags = CMF_NORMAL) then // from the desktop
    sMenuCaption := 'Show File Info..'
  else if (Flags and CMF_VERBSONLY) = CMF_VERBSONLY then // from a shortcut
    sMenuCaption := 'Show File Info (from Shortcut)..'
  else if (Flags and CMF_EXPLORE) = CMF_EXPLORE then // from explorer
    sMenuCaption := 'Show File Info..'
  else
    { fail for any other value }
    Result := E_FAIL;

  if Result <> E_FAIL then
  begin
    if FileExists(FFilename) then
    begin
      // show file date info as menu caption
      sMenuCaption :=
          Format('[%s] ', [ExtractFileExt(FFilename)]) +
          'Created: ' + FileTime(FFilename, ftkCreated) +
          ' - Mod: ' + FileTime(FFilename, ftkModified);
    end;

    FillChar(mnuFileInfoItem, SizeOf(TMenuItemInfo), #0);
    mnuFileInfoItem.cbSize := SizeOf(TMenuItemInfo);
    mnuFileInfoItem.fMask := MIIM_STRING or MIIM_ID;
    mnuFileInfoItem.fType := MFT_STRING;
    mnuFileInfoItem.wID := CmdFirst;
    mnuFileInfoItem.dwTypeData := PWideChar(sMenuCaption);
    mnuFileInfoItem.cch := Length(sMenuCaption);

    InsertMenuItem(Menu, IndexMenu, True, mnuFileInfoItem);
  end;
end;


function TFileInfoContextMenu.ShellExtInitialize(
      pidlFolder: PItemIDList;
      lpdobj: IDataObject;
      hKeyProgID: HKEY): HResult;
var
  DataFormat: TFormatEtc;
  StrgMedium: TStgMedium;
  Buffer: array [0 .. MAX_PATH] of Char;
begin
  Result := E_FAIL;

  { Check if an object was defined }
  if lpdobj = nil then
    Exit;

  { Prepare to get information about the object }
  DataFormat.cfFormat := CF_HDROP;
  DataFormat.ptd := nil;
  DataFormat.dwAspect := DVASPECT_CONTENT;
  DataFormat.lindex := -1;
  DataFormat.tymed := TYMED_HGLOBAL;

  if lpdobj.GetData(DataFormat, StrgMedium) <> S_OK then
    Exit;

  { The implementation now support only one file }
  if DragQueryFile(StrgMedium.hGlobal, $FFFFFFFF, nil, 0) = 1 then
  begin
    SetLength(FFileName, MAX_PATH);
    DragQueryFile(StrgMedium.hGlobal, 0, @Buffer, SizeOf(Buffer));
    FFileName := Buffer;
    Result := NOERROR;
  end
  else
  begin
    // Don't show the Menu if more then one file was selected
    FFileName := EmptyStr;
    Result := E_FAIL;
  end;

  { http://msdn.microsoft.com/en-us/library/ms693491(v=vs.85).aspx }
  ReleaseStgMedium(StrgMedium);
end;

procedure TFileInfoContextMenu.ShowFileInfo(WinHandle: HWND; const Filename: string);
var
  sPath, sName, sUnits, sInfo, sCreated, sModified: string;
  nSize: int64;
  fSize: double;
const
  ONE_KB = 1024;
  ONE_MB = 1024 * 1024;
  ONE_GB = 1024 * 1024 * 1024;
begin
  sPath := ExtractFilePath(Filename);
  sName := ExtractFileName(Filename);
  sCreated := FileTime(Filename, ftkCreated);
  sModified := FileTime(Filename, ftkModified);

  nSize := GetFileSize(Filename);
  if (nSize >= ONE_GB) then
  begin
    fSize := nSize / ONE_GB;
    sUnits := 'GB';
  end
  else if (nSize >= ONE_MB) then
  begin
    fSize := nSize / ONE_MB;
    sUnits := 'MB';
  end
  else
  begin
    fSize := nSize / ONE_KB;
    sUnits := 'KB';
  end;

  sInfo :=
        'Filename: ' + sName + #10#13 +
        'Path: ' + sPath + #10#13 +
        Format('Size: %.3f %s', [fSize, sUnits]) + #10#13 +
        'Date Created:   ' + sCreated + #10#13 +
        'Date Modified: ' + sModified + #10#13;

  MessageBox(WinHandle, PWideChar(sInfo), 'File Info', 0); //MB_ICONINFORMATION);
end;

{ TFileInfoObjectFactory }


{ Required to registration for Windows NT/2000 }
procedure TFileInfoObjectFactory.ApproveShellExtension(
      &Register: Boolean;
      const ClsID: string);
Const
  WinNTRegKey = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved';
var
  Reg: TRegistry;
begin
  Reg := TRegistry.Create;

  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;

    if not Reg.OpenKey(WinNTRegKey, True) then
      Exit;

    { register the extension appropriately }
    if &Register then
      Reg.WriteString(ClsID, Description)
    else
      Reg.DeleteValue(ClsID);
  finally
    Reg.Free;
  end;

end;


function TFileInfoObjectFactory.GetProgID: string;
begin
  { ProgID not required for shell extensions }
  Result := '';
end;


procedure TFileInfoObjectFactory.UpdateRegistry(Register: Boolean);
Const
  ContextKey = '*\shellex\ContextMenuHandlers\%s';
begin
  { perform normal registration }
  inherited UpdateRegistry(Register);

  { Registration required for Windows NT/2000 }
  ApproveShellExtension(Register, GUIDToString(ClassID));

  { if this server is being registered, register the required key/values
    to expose it to Explorer }
  if Register then
    CreateRegKey(Format(ContextKey, [ClassName]), '', GUIDToString(ClassID), HKEY_CLASSES_ROOT)
  else
    DeleteRegKey(Format(ContextKey, [ClassName]));
end;

initialization

TFileInfoObjectFactory.Create(ComServer, TFileInfoContextMenu, CLASS_FileInfoContextMenu,
      ciMultiInstance, tmApartment);

end.
