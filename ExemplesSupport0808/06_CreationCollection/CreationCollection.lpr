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
  objets100clib_tlb,
  commun;

var
  StreamCpta : TAxcBSCPTAApplication100c;
  BaseCpta   : IBSCPTAApplication3;
  CollTiers  : IBICollection;
  Tiers      : IBOTiers3;
  iTiers     : OleVariant; // Usage itération standard Delphi IenumVariant
  IEnum      : IEnumVARIANT; // Usage itération standard Delphi IenumVariant
  Nombre     : LongWord; // Usage itération standard Delphi IenumVariant

begin
  // Initialize COM. ------------------------------------------
  CoInitializeEx(nil, COINIT_MULTITHREADED);

  StreamCpta  := TAxcBSCPTAApplication100c.Create(nil);
  BaseCpta    := StreamCpta.OleServer;

  try
    if OuvreBaseCptaSql(BaseCpta,
      '(local)\SAGE2017',
      'BIJOU_V7',
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
        IEnum := CollTiers._NewEnum as IEnumVARIANT;
        while IEnum.Next(1, iTiers, Nombre) = S_OK do
        begin
          Writeln((IUnknown(iTiers) as IBOTiers3).CT_Num);
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

