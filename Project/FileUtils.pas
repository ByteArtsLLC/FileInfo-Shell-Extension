unit FileUtils;

interface

{$WARN SYMBOL_PLATFORM OFF}

uses
  System.Classes,
  System.Types,
  System.IOUtils,
  System.Generics.Collections;


/// <summary>
/// Returns the size of a file.
/// </summary>
function GetFileSize(const Filename: string): Int64;

function GetFileTime(const Filename: string): TDateTime;


implementation

uses
  System.SysUtils,
  System.Generics.Defaults;



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


function GetFileTime(const Filename: string): TDateTime;
var
  srRecord: TSearchRec;
begin
  Result := 0;

  if FileExists(Filename) then
  begin
    if FindFirst(Filename, faNormal, srRecord) = 0 then
      Result := srRecord.TimeStamp;
  end;
end;


end.
