{******************************************************************************}
{                                                                              }
{ Lazarus Component                                                            }
{ Copyright (c) 2013 MBG partenaires                                           }
{ Unit owner: Thierry Bothorel                                                 }
{ Version: 1                                                                   }
{ Subversion:                                                                  }
{   $Id:: LectureDevise.lpr 20 2013-07-06 11:47:04Z TBOTHOREL                $ }
{                                                                              }
{******************************************************************************}

program LectureDevise;

{$APPTYPE CONSOLE}{$mode objfpc}{$H+}

uses
  Interfaces, // sinon Error: Undefined symbol: WSRegisterCustomImageList
  SysUtils,  // sinon Error: Identifier not found "Exception"
  ActiveX,
  objets100clib,
  Commun
  ;

var
  StreamCpta      : TAxcBSCPTAApplication100c;
  BaseCpta        : IBSCPTAApplication3;
  IntituleDevise  : string;
  Devise          : IBPDevise2;

begin
  // Initialize COM. ------------------------------------------
  CoInitializeEx(nil, COINIT_MULTITHREADED);

  StreamCpta  := TAxcBSCPTAApplication100c.Create(nil);
  BaseCpta    := StreamCpta.OleServer;

  try
    if OuvreBaseCpta(BaseCpta,
        'E:\DATA\Gestion\BIJOU-SQL2017\V7\BIJOU_V7.MAE',
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

