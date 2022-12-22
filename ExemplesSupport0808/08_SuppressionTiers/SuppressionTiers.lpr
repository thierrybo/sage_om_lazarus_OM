{******************************************************************************}
{                                                                              }
{ Lazarus Component                                                            }
{ Copyright (c) 2013 MBG partenaires                                           }
{ Unit owner: Thierry Bothorel                                                 }
{ Version: 1                                                                   }
{ Subversion:                                                                  }
{   $Id:: SuppressionTiers.lpr 31 2013-07-06 20:40:57Z TBOTHOREL             $ }
{                                                                              }
{******************************************************************************}

program SuppressionTiers;

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

function SupprimeTiers(var ABaseCpta: IBSCPTAApplication3; ANumTiers: string):
    Boolean;
  var
    ObjetTiers : IBOTiers3;
  begin
    try
      ObjetTiers  := ABaseCpta.FactoryTiers.ReadNumero(ANumTiers);
      ObjetTiers.Remove;
      Result      := true;
    except on E: Exception do
      begin
      Writeln(
        'Erreur en suppression de l''enregistrement ',
        ANumTiers,
        ' dans la table F_COMPTET : ',
        sLineBreak,
        E.ClassName,
        ': ',
        E.Message);
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
    if OuvreBaseCptaSql(BaseCpta,
      '(local)\SAGE2017',
      'BIJOU_V7',
      '<Administrateur>'
      ) then
    begin
      if SupprimeTiers(BaseCpta, 'ZASUPPRIMER') then
        Writeln(UTF8ToAnsi('Tiers ZASUPPRIMER supprimé !'));
      //if SupprimeTiers(BaseCpta, 'CARAT') then
      //  Writeln(UTF8ToAnsi('Tiers CARAT supprimé !'));
    end;
  finally
    FermeBaseCpta(BaseCpta);
    FreeAndNil(StreamCpta);
    CoUnInitialize;
    Readln;
  end;
end.

