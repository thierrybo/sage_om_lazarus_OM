{******************************************************************************}
{                                                                              }
{ Lazarus Component                                                            }
{ Copyright (c) 2013 MBG partenaires                                           }
{ Unit owner: Thierry Bothorel                                                 }
{ Version: 1                                                                   }
{ Subversion:                                                                  }
{   $Id:: CreationObjetClient.lpr 23 2013-07-06 11:54:08Z TBOTHOREL          $ }
{                                                                              }
{******************************************************************************}

program CreationObjetClient;

{$APPTYPE CONSOLE}{$mode objfpc}{$H+}

uses
  Interfaces, // sinon Error: Undefined symbol: WSRegisterCustomImageList
  SysUtils,  // sinon Error: Identifier not found "Exception"
  ActiveX,
  Objets100Lib_3_0_TLB,
  Commun
  ;

var
  StreamCpta  : TAxcBSCPTAApplication3;
  BaseCpta    : IBSCPTAApplication3;
  ObjetClient : IBOClient3;

begin
  // Initialize COM. ------------------------------------------
  CoInitializeEx(nil, COINIT_MULTITHREADED);

  StreamCpta  := TAxcBSCPTAApplication3.Create(nil);
  BaseCpta    := StreamCpta.OleServer;

  try
    if OuvreBaseCpta(BaseCpta,
        'C:\Temp\BIJOU1553.MAE',
        '<Administrateur>'
        ) then
    begin
      try
        ObjetClient := BaseCpta.FactoryClient.Create as IBOClient3;
        Writeln(UTF8ToAnsi('Objet client créé !'));
      except
        on E:Exception do
          Writeln(
            UTF8ToAnsi('Erreur en création d''un nouvel objet client : '),
            sLineBreak,
            E.Classname,
            ': ',
            E.Message);
      end;
    end;
  finally
    FermeBaseCpta(BaseCpta);
    FreeAndNil(StreamCpta);
    CoUnInitialize;
    Readln;
  end;
end.

