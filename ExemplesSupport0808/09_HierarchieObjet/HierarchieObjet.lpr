{******************************************************************************}
{                                                                              }
{ Lazarus Component                                                            }
{ Copyright (c) 2013 MBG partenaires                                           }
{ Unit owner: Thierry Bothorel                                                 }
{ Version: 1                                                                   }
{ Subversion:                                                                  }
{   $Id:: HierarchieObjet.lpr 33 2013-07-06 21:02:49Z TBOTHOREL              $ }
{                                                                              }
{******************************************************************************}

program HierarchieObjet;

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

procedure AfficheHierarchie(ABanque: IBOTiersBanque3);
begin
  try
    Writeln('La banque est : ', ABanque.BT_Intitule);
    { #todo : La sortie de (ABanque.Stream as IBSCPTAApplication3).Name est vide}
    Writeln(
      'Le stream est la base comptable : ',
      //(ABanque.Stream as IBSCPTAApplication3).Name); // Si on utilise l'ouverture .mae
      (ABanque.Stream as IBSCPTAApplication3).CompanyDatabaseName); // Si on utilise l'ouverture sql
    Writeln(UTF8ToAnsi('L''objet maître est le tiers : '), ABanque.Tiers.CT_Num);

    { Au préalable, créer une seconde Banque pour le tiers CARAT }
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

