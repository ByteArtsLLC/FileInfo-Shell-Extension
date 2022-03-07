unit FileInfo_TLB;

// ************************************************************************ //
// WARNING
// -------
// The types declared in this file were generated from data read from a
// Type Library. If this type library is explicitly or indirectly (via
// another type library referring to this type library) re-imported, or the
// 'Refresh' command of the Type Library Editor activated while editing the
// Type Library, the contents of this file will be regenerated and all
// manual modifications will be lost.
// ************************************************************************ //

// $Rev: 52393 $
// File generated on 3/1/17 3:20:38 PM from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\Users\Scott\Documents\RAD Studio\Projects\Windows shell extension\FileInfo\FileInfo (1)
// LIBID: {CBCADBFF-8D95-43B9-AA37-CC004CA0D9FF}
// LCID: 0
// Helpfile:
// HelpString:
// DepndLst:
//   (1) v2.0 stdole, (C:\Windows\SysWOW64\stdole2.tlb)
// SYS_KIND: SYS_WIN32
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers.
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
{$ALIGN 4}

interface

uses Winapi.Windows, System.Classes, System.Variants, System.Win.StdVCL, Vcl.Graphics, Vcl.OleServer, Winapi.ActiveX;


// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:
//   Type Libraries     : LIBID_xxxx
//   CoClasses          : CLASS_xxxx
//   DISPInterfaces     : DIID_xxxx
//   Non-DISP interfaces: IID_xxxx
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  FileInfoMajorVersion = 1;
  FileInfoMinorVersion = 0;

  LIBID_FileInfo: TGUID = '{CBCADBFF-8D95-43B9-AA37-CC004CA0D9FF}';

  IID_IFileInfoContextMenu: TGUID = '{7B09FCFC-9B2D-4BEB-92A6-BE327B888857}';
  CLASS_FileInfoContextMenu: TGUID = '{5F36EC37-DDB0-40AA-836F-1A492442C4B9}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary
// *********************************************************************//
  IFileInfoContextMenu = interface;
  IFileInfoContextMenuDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library
// (NOTE: Here we map each CoClass to its Default Interface)
// *********************************************************************//
  FileInfoContextMenu = IFileInfoContextMenu;


// *********************************************************************//
// Interface: IFileInfoContextMenu
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {7B09FCFC-9B2D-4BEB-92A6-BE327B888857}
// *********************************************************************//
  IFileInfoContextMenu = interface(IDispatch)
    ['{7B09FCFC-9B2D-4BEB-92A6-BE327B888857}']
  end;

// *********************************************************************//
// DispIntf:  IFileInfoContextMenuDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {7B09FCFC-9B2D-4BEB-92A6-BE327B888857}
// *********************************************************************//
  IFileInfoContextMenuDisp = dispinterface
    ['{7B09FCFC-9B2D-4BEB-92A6-BE327B888857}']
  end;

// *********************************************************************//
// The Class CoFileInfoContextMenu provides a Create and CreateRemote method to
// create instances of the default interface IFileInfoContextMenu exposed by
// the CoClass FileInfoContextMenu. The functions are intended to be used by
// clients wishing to automate the CoClass objects exposed by the
// server of this typelibrary.
// *********************************************************************//
  CoFileInfoContextMenu = class
    class function Create: IFileInfoContextMenu;
    class function CreateRemote(const MachineName: string): IFileInfoContextMenu;
  end;

implementation

uses System.Win.ComObj;

class function CoFileInfoContextMenu.Create: IFileInfoContextMenu;
begin
  Result := CreateComObject(CLASS_FileInfoContextMenu) as IFileInfoContextMenu;
end;

class function CoFileInfoContextMenu.CreateRemote(const MachineName: string): IFileInfoContextMenu;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_FileInfoContextMenu) as IFileInfoContextMenu;
end;

end.

