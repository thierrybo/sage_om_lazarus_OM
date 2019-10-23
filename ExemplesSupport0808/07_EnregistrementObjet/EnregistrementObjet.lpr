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

program EnregistrementObjet;

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
  Client     : IBOClient3;

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
        Client := BaseCpta.FactoryClient.Create as IBOClient3;
        With Client do
        begin
          CT_Num       := 'BOLLE';
          CT_Intitule  := 'Bolle Virginie';
          TiersPayeur  := Client;
          CompteGPrinc := BaseCpta.FactoryCompteG.ReadNumero('4110000');
          Write;
        end;
        Writeln(UTF8ToAnsi('Client correctement créé'));
      except
        on E:Exception do
          Writeln(
            UTF8ToAnsi('Erreur en création du client : '),
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

