{******************************************************************************}
{                                                                              }
{ Lazarus Component                                                            }
{ Copyright (c) 2013 MBG partenaires                                           }
{ Unit owner: Thierry Bothorel                                                 }
{ Version: 1                                                                   }
{ Subversion:                                                                  }
{   $Id:: EcrituresComptables.lpr 37 2013-07-07 09:50:20Z TBOTHOREL          $ }
{                                                                              }
{******************************************************************************}

program EcrituresComptables;

{$APPTYPE CONSOLE}{$mode objfpc}{$H+}

uses
  Interfaces, // sinon Error: Undefined symbol: WSRegisterCustomImageList
  SysUtils,  // sinon Error: Identifier not found "Exception"
  ActiveX,
  objets100clib_tlb,
  commun;

var
  StreamCpta  : TAxcBSCPTAApplication100c;
  BaseCpta    : IBSCPTAApplication3;
  Ecriture    : IBOEcriture3;

begin
  // Initialize COM. ------------------------------------------
  CoInitializeEx(nil, COINIT_MULTITHREADED);

  StreamCpta  := TAxcBSCPTAApplication100c.Create(nil);
  BaseCpta    := StreamCpta.OleServer;

  try
    try
      // Si on utilise l'ouverture SQL
      if OuvreBaseCptaSql(BaseCpta,
        '(local)\SAGE2017',
        'BIJOU_V7',
        '<Administrateur>'
        ) then
      // Si on utilise l'ouverture .mae
      //if OuvreBaseCpta(BaseCpta,
      //  'E:\DATA\Gestion\BIJOU-SQL2017\V7\BIJOU_V7.MAE',
      //  '<Administrateur>') then
      begin
        Ecriture := BaseCpta.FactoryEcriture.Create as IBOEcriture3;
        with Ecriture do
        begin
          Journal     := BaseCpta.FactoryJournal.ReadNumero('BEU');
          Date        := StrToDateTime('03/07/19');
          Tiers       := BaseCpta.FactoryClient.ReadNumero('CARAT');
          EC_Intitule := 'Acompte';
          EC_RefPiece := 'FA1234';
          // ci -dessous il faut lui mettre le n° de pièce suivante (next)
          EC_Piece    := Journal.NextEC_Piece[StrToDateTime('03/07/19')];
          EC_Montant  := 123.45;
          EC_Sens     := EcritureSensTypeCredit; // VB = .EC_Sens = EcritureSensType.EcritureSensTypeCredit
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

