library FileInfo;

uses
  ComServ,
  FileInfo_TLB in 'FileInfo_TLB.pas',
  FileInfoImpl in 'FileInfoImpl.pas';

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer,
  DllInstall;

{$R *.TLB}

{$R *.RES}

begin
end.
