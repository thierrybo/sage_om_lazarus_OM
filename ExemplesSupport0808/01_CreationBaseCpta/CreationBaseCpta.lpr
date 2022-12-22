{******************************************************************************}
{                                                                              }
{ Lazarus Component                                                            }
{ Copyright (c) 2013 MBG partenaires                                           }
{ Unit owner: Thierry Bothorel                                                 }
{ Version: 1                                                                   }
{ Subversion:                                                                  }
{   $Id:: CreationBaseCpta.lpr 12 2013-07-05 21:00:21Z TBOTHOREL             $ }
{                                                                              }
{******************************************************************************}

program CreationBaseCpta;

{$mode objfpc}{$H+}

uses
  Interfaces, // sinon Error: Undefined symbol: WSRegisterCustomImageList
  SysUtils,  // sinon Error: Identifier not found "Exception"
  ActiveX,
  objets100clib
  ;

var
  StreamCpta  : TAxcBSCPTAApplication100c;
  BaseCpta    : IBSCPTAApplication3;

function CreeBaseCpta(var ABaseCpta: IBSCPTAApplication3; ANomBase: string):
    Boolean;
begin
  try
    Writeln(UTF8ToAnsi('Création de la base comptable '), ANomBase, ' en cours');
    ABaseCpta.Name := ANomBase;
    ABaseCpta.Create();
    Result := true;
  except on E: Exception do
    begin
      Writeln(UTF8ToAnsi('Erreur en création de base comptable '),
        ABaseCpta.Name, ' : ', sLineBreak, E.ClassName, ': ', E.Message);
      Result := false;
    end;
  end;
end;

begin
  // Initialize COM. ------------------------------------------
  CoInitializeEx(nil, COINIT_MULTITHREADED);

  StreamCpta  := TAxcBSCPTAApplication100c.Create(nil);
  BaseCpta    := StreamCpta.OleServer;
  try
    if CreeBaseCpta(BaseCpta, 'E:\DATA\tmp\test1.mae') then
    begin
        Writeln(UTF8ToAnsi('Base comptable correctement créée !'));
    end;
  finally
    FreeAndNil(StreamCpta);
    CoUnInitialize;
    Readln;
  end;
end.
