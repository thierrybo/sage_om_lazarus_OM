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

program LectureDevise;

{$APPTYPE CONSOLE}{$mode objfpc}{$H+}

uses
  Interfaces, // sinon Error: Undefined symbol: WSRegisterCustomImageList
  SysUtils,  // sinon Error: Identifier not found "Exception"
  ActiveX,
  Objets100Lib_3_0_TLB,
  Commun
  ;

var
  StreamCpta      : TAxcBSCPTAApplication3;
  BaseCpta        : IBSCPTAApplication3;
  IntituleDevise  : string;
  Devise          : IBPDevise2;

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
      try
        IntituleDevise := 'Dollar US';
        Devise         := BaseCpta.FactoryDevise.ReadIntitule(IntituleDevise);
        With Devise do
        begin
          Writeln(UTF8ToAnsi('Intitulé : '), D_Intitule);
          Writeln(UTF8ToAnsi('Unité : ')   , D_Monnaie);
          Writeln(UTF8ToAnsi('Cours ')     , FormatFloat('#0.##', D_Cours));
        end;
      except
        on E:Exception do
          Writeln(
            'Erreur en lecture de l''enregistrement ',
            IntituleDevise,
            ' de la table P_DEVISE : ',
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

