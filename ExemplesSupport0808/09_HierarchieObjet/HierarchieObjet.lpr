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

program HierarchieObjet;

{$APPTYPE CONSOLE}{$mode objfpc}{$H+}

uses
  Interfaces, // sinon Error: Undefined symbol: WSRegisterCustomImageList
  SysUtils,  // sinon Error: Identifier not found "Exception"
  ActiveX,
  Objets100Lib_3_0_TLB, commun;

var
  StreamCpta  : TAxcBSCPTAApplication3;
  BaseCpta    : IBSCPTAApplication3;

procedure AfficheHierarchie(ABanque: IBOTiersBanque3);
begin
  try
    Writeln('La banque est : ', ABanque.BT_Intitule);
    Writeln(
      'Le stream est la base comptable : ',
      (ABanque.Stream as IBSCPTAApplication3).Name);
    Writeln(UTF8ToAnsi('L''objet maître est le tiers : '), ABanque.Tiers.CT_Num);

    { Au préalable, créer une seconde ABanque pour le tiers CARAT }
    Writeln(
      'La seconde banque est : ',
      (ABanque.FactoryTiersBanque.List.Item[2] as IBOTiersBanque3)
        .BT_Intitule);
  except
    on E:Exception do
      Writeln(E.Classname, ': ', E.Message);
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
      AfficheHierarchie((BaseCpta.FactoryTiers.ReadNumero('CARAT').
        FactoryTiersBanque.List.Item[1]) as IBOTiersBanque3);
    end;
  finally
    FermeBaseCpta(BaseCpta);
    FreeAndNil(StreamCpta);
    CoUnInitialize;
    Readln;
  end;
end.

