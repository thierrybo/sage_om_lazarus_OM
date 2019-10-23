{******************************************************************************}
{                                                                              }
{ Lazarus Component                                                            }
{ Copyright (c) 2013 MBG partenaires                                           }
{ Unit owner: Thierry Bothorel                                                 }
{ Version: 1                                                                   }
{ Subversion:                                                                  }
{   $Id:: CreationCollection.lpr 27 2013-07-06 12:47:32Z TBOTHOREL           $ }
{                                                                              }
{******************************************************************************}

program CreationCollection;

{$APPTYPE CONSOLE}{$mode objfpc}{$H+}

uses
  Interfaces, // sinon Error: Undefined symbol: WSRegisterCustomImageList
  SysUtils,  // sinon Error: Identifier not found "Exception"
  ActiveX,
  Objets100Lib_3_0_TLB,
  commun;

var
  StreamCpta : TAxcBSCPTAApplication3;
  BaseCpta   : IBSCPTAApplication3;
  CollTiers  : IBICollection;
  Tiers      : IBOTiers3;
  I          : Integer;

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
        CollTiers := BaseCpta.FactoryTiers.List;
        Writeln(
          UTF8ToAnsi('La base de données contient '),
          CollTiers.Count,
          ' tiers.');
        Tiers := CollTiers.Item[CollTiers.Count] as IBOTiers3;
        Writeln('Le dernier tiers se nomme ', Tiers.CT_Intitule, '.');
        Writeln(sLineBreak, 'Liste des tiers :');
        for I := 1 to CollTiers.Count do
        begin
          Writeln((CollTiers.Item[I] as IBOTiers3).CT_Num);
        end;
      except
        on E:Exception do
          Writeln(
            UTF8ToAnsi('Erreur en création d''une collection de tiers : '),
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

