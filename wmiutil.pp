{
 https://forum.lazarus.freepascal.org/index.php?topic=51125.0#msg374984
 http://www.stack.nl/~marcov/files/wmiutil.zip
}

unit wmiutil;

interface
{$mode delphi}
uses
  Windows,
  SysUtils,
  StrUtils,
  ActiveX,
  ComObj,
  Variants;

Type
  TEatenType = {$ifdef fpc} {$ifdef ver3_0}pulong{$else}Ulong{$endif}{$else}Integer{$endif}; // type for eaten parameter MkParseDisplayName
  oEnumIterator = record
                      mainobj: OleVariant;
                      oEnum  : IEnumVariant;
                      IterItem : OleVariant;
                      IterVal  : LongWord;
                      function Enumerate(v:olevariant):oEnumIterator;
                      function GetEnumerator :oEnumIterator;
                      function MoveNext:Boolean;
                      property Current:OleVariant read iteritem;
                   end;

function OleVariantToText(aVar:OleVariant):string;
function GetWMIObject(const objectName: String): IDispatch;

Implementation

{ oEnumIterator}

function oEnumIterator.getenumerator :oEnumIterator;
begin
 result:=self;
end;

Function oEnumIterator.Enumerate(v :olevariant):oEnumIterator;
begin
  mainobj:=v;
  oEnum  := IUnknown(mainobj._NewEnum) as IEnumVariant;
  result:=self;
end;

Function  oEnumIterator.MoveNext:boolean;
begin
  result:=(oEnum.Next(1, iteritem, iterval) = s_ok);
end;

function OleVariantToText(aVar:OleVariant):string;
// mostly quickdump for WMI researchpurposes
var
    i : integer;
begin
  Result:='';
  if not VarIsNull(aVar) then
    if VarIsArray(aVar) then
      begin
        result:='{';
        for i :=VarArrayLowBound(aVar,1) to vararrayhighbound(aVar,1)  do
          begin
            if i<>0 then
              result:=result+',';
            result:=result+OleVariantToText(vararrayget(aVar,[i]));
          end;
        result:=result+'}';
      end
    else
      Result:=VarToStr(aVar);
end;

function GetWMIObject(const objectName: String): IDispatch;
var
  chEaten: TEatenType;
  BindCtx: IBindCtx;
  Moniker: IMoniker;
begin
    OleCheck(CreateBindCtx(0, bindCtx));
    OleCheck(MkParseDisplayName(BindCtx, StringToOleStr(objectName), pulong(@chEaten), Moniker));
    OleCheck(Moniker.BindToObject(BindCtx, nil, IDispatch, Result));
end;

end.
