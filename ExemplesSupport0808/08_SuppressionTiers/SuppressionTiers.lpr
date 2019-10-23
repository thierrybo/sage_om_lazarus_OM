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

program SuppressionTiers;

{$APPTYPE CONSOLE}{$mode objfpc}{$H+}

uses
  Interfaces, // sinon Error: Undefined symbol: WSRegisterCustomImageList
  SysUtils,  // sinon Error: Identifier not found "Exception"
  ActiveX,
  Objets100Lib_3_0_TLB,
  commun
  ;

var
  StreamCpta  : TAxcBSCPTAApplication3;
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

  StreamCpta  := TAxcBSCPTAApplication3.Create(nil);
  BaseCpta    := StreamCpta.OleServer;

  try
    if OuvreBaseCpta(BaseCpta,
      'C:\Documents and Settings\All Users\Documents\Sage\Sage Entreprise\BIJOU.MAE',
      '<Administrateur>'
      ) then
    begin
      if SupprimeTiers(BaseCpta, 'ZAN') then
        Writeln(UTF8ToAnsi('Tiers ZAN supprimé !'));
      if SupprimeTiers(BaseCpta, 'CARAT') then
        Writeln(UTF8ToAnsi('Tiers CARAT supprimé !'));
    end;
  finally
    FermeBaseCpta(BaseCpta);
    FreeAndNil(StreamCpta);
    CoUnInitialize;
    Readln;
  end;
end.

