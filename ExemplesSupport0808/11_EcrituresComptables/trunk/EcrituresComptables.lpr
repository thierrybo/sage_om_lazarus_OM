{******************************************************************************}
{                                                                              }
{ Lazarus Component                                                            }
{ Copyright (c) 2013 MBG partenaires                                           }
{ Unit owner: Thierry Bothorel                                                 }
{ Version: 1                                                                   }
{ Subversion:                                                                  }
{   $Id::                                                                    $ }
{                                                                              }
{******************************************************************************}

program EcrituresComptables;

{$APPTYPE CONSOLE}{$mode objfpc}{$H+}

uses
  Interfaces, // sinon Error: Undefined symbol: WSRegisterCustomImageList
  SysUtils,  // sinon Error: Identifier not found "Exception"
  ActiveX,
  Objets100Lib_3_0_TLB,
  commun;

var
  StreamCpta  : TAxcBSCPTAApplication3;
  BaseCpta    : IBSCPTAApplication3;
  Ecriture    : IBOEcriture3;

begin
  // Initialize COM. ------------------------------------------
  CoInitializeEx(nil, COINIT_MULTITHREADED);

  StreamCpta  := TAxcBSCPTAApplication3.Create(nil);
  BaseCpta    := StreamCpta.OleServer;

  try
    try
      if OuvreBaseCpta(BaseCpta,
        'C:\Documents and Settings\All Users\Documents\Sage\Sage Entreprise\BIJOU.MAE',
        '<Administrateur>'
        ) then
      begin
        Ecriture := BaseCpta.FactoryEcriture.Create as IBOEcriture3;
        with Ecriture do
        begin
          Journal     := BaseCpta.FactoryJournal.ReadNumero('BEU');
          Date        := StrToDateTime('03/07/09');
          Tiers       := BaseCpta.FactoryClient.ReadNumero('CARAT');
          EC_Intitule := 'Acompte';
          EC_RefPiece := 'FA1234';
          EC_Piece    := Journal.NextEC_Piece[StrToDateTime('03/07/09')];
          EC_Montant  := 123.45;
          EC_Sens     := EcritureSensTypeCredit;
          SetDefault;
          WriteDefault;
        end;
      end;
    except
      on E:Exception do
      begin
        Writeln(E.Classname, ': ', E.Message);
        Readln;
      end;
    end;
  finally
    FermeBaseCpta(BaseCpta);
    FreeAndNil(StreamCpta);
    CoUnInitialize;
    Readln;
  end;
end.

