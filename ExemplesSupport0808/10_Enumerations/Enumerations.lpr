{******************************************************************************}
{                                                                              }
{ Lazarus Component                                                            }
{ Copyright (c) 2013 MBG partenaires                                           }
{ Unit owner: Thierry Bothorel                                                 }
{ Version: 1                                                                   }
{ Subversion:                                                                  }
{   $Id:: Enumerations.lpr 35 2013-07-07 09:19:25Z TBOTHOREL                 $ }
{                                                                              }
{******************************************************************************}

program Enumerations;

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
  Journal     : IBOJournal3;

begin
  // Initialize COM. ------------------------------------------
  CoInitializeEx(nil, COINIT_MULTITHREADED);

  StreamCpta  := TAxcBSCPTAApplication100c.Create(nil);
  BaseCpta    := StreamCpta.OleServer;

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
      Journal             := BaseCpta.FactoryJournal.Create as IBOJournal3;
      Journal.SetDefault;  // Pas dans VB. Pzsse type numérotation de manuelle à continue
      Journal.JO_Type     := JournalTypeVente;
      Journal.JO_Num      := 'VTE2';
      Journal.JO_Intitule := 'Vente 2';
      Journal.Write_;
    end;
  finally
    FermeBaseCpta(BaseCpta);
    FreeAndNil(StreamCpta);
    CoUnInitialize;
    Readln;
  end;
end.

