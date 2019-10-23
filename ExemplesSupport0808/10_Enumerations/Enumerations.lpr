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

program Enumerations;

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
  Journal     : IBOJournal3;

begin
  // Initialize COM. ------------------------------------------
  CoInitializeEx(nil, COINIT_MULTITHREADED);

  StreamCpta  := TAxcBSCPTAApplication3.Create(nil);
  BaseCpta    := StreamCpta.OleServer;

  try
    if OuvreBaseCpta(BaseCpta,
      'C:\Documents and Settings\All Users\Documents\Sage\Sage Entreprise\BIJOU.MAE',
      '<Administrateur>'
      ) then
    begin
      Journal             := BaseCpta.FactoryJournal.Create as IBOJournal3;
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

