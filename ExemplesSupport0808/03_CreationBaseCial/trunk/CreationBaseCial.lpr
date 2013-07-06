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

program CreationBaseCial;

{$APPTYPE CONSOLE}{$mode objfpc}{$H+}

uses
  Interfaces, // sinon Error: Undefined symbol: WSRegisterCustomImageList
  SysUtils,  // sinon Error: Identifier not found "Exception"
  ActiveX,
  Objets100Lib_3_0_TLB
  ;

var
  StreamCpta  : TAxcBSCPTAApplication3;
  BaseCpta    : IBSCPTAApplication3;
  StreamCial  : TAxcBSCIALApplication3;
  BaseCial    : IBSCIALApplication3;

function CreeBaseCpta(var ABaseCpta: IBSCPTAApplication3; ANomBase: string):
      Boolean;
begin
  try
    Writeln(UTF8ToAnsi('Création de la base comptable '), ANomBase, ' en cours');
    ABaseCpta.Name := ANomBase;
    ABaseCpta.Create();
    Result := true;
  except on E: Exception do
    begin
      Writeln(UTF8ToAnsi('Erreur en création de base comptable '),
        ABaseCpta.Name,
        ' : ',
        sLineBreak,
        E.ClassName,
        ': ',
        UTF8ToAnsi(E.Message));
      Result := false;
    end;
  end;
end;

function OuvreBaseCpta(var ABaseCpta: IBSCPTAApplication3; ANomBaseCpta: string;
    AUtilisateur: string = ''; AMotDePasse: string = ''): Boolean;
begin
  try
    ABaseCpta.Name := ANomBaseCpta;
    if (AUtilisateur <> '') then
    begin
      ABaseCpta.Loggable.UserName := AUtilisateur;
      ABaseCpta.Loggable.UserPwd  := AMotDePasse;
    end;
    ABaseCpta.Open;
    Result := true;
  except on E: Exception do
    begin
      Writeln(
        'Erreur en ouverture de base comptable ',
        ABaseCpta.Name,
        ' : ',
        sLineBreak,
        E.ClassName,
        ': ',
        UTF8ToAnsi(E.Message));
      Result := false;
    end;
  end;
end;

function FermeBaseCpta(var ABaseCpta: IBSCPTAApplication3): Boolean;
begin
  try
    if ((ABaseCpta <> nil) and ABaseCpta.IsOpen) then
    begin
      ABaseCpta.Close;
      Result := true;
    end;
  except on E: Exception do
    begin
      Writeln(
        'Erreur en fermeture de base comptable ',
        ABaseCpta.Name,
        ' : ',
        sLineBreak,
        E.ClassName,
        ': ',
        UTF8ToAnsi(E.Message));
      Result := false;
    end;
  end;
end;

function CreeBaseCial(var ABaseCial: IBSCIALApplication3; ANomBaseCial: string;
    var ABaseCpta: IBSCPTAApplication3; ANomBaseCpta: string):  Boolean;
begin
  try
    if not FileExists(ANomBaseCpta) then
    begin
      CreeBaseCpta(ABaseCpta, ANomBaseCpta);
    end;
    if OuvreBaseCpta(ABaseCpta, ANomBaseCpta) then
    begin
      Writeln(
        UTF8ToAnsi('Création de la base commerciale '),
        ANomBaseCial,
        ' en cours');
      ABaseCial.CptaApplication  := ABaseCpta;
      ABaseCial.Name             := ANomBaseCial;
      ABaseCial.Create();
      FermeBaseCpta(ABaseCpta);
      Result := true;
    end
    else
    begin
      Result := false;
    end;
  except on E: Exception do
    begin
      Writeln(
        UTF8ToAnsi('Erreur en création de base commerciale '),
        ABaseCial.Name,
        ' : ',
        sLineBreak,
        E.ClassName,
        ': ',
        UTF8ToAnsi(E.Message));
      Result := false;
    end;
  end;
end;

begin
  // Initialize COM. ------------------------------------------
  CoInitializeEx(nil, COINIT_MULTITHREADED);

  StreamCpta  := TAxcBSCPTAApplication3.Create(nil);
  BaseCpta    := StreamCpta.OleServer;
  StreamCial  := TAxcBSCIALApplication3.Create(nil);
  BaseCial    := StreamCial.OleServer;

  try
    if CreeBaseCial(BaseCial, 'C:\Temp\test2.gcm', BaseCpta, 'C:\Temp\test2.mae')
        then
    begin
        Writeln(UTF8ToAnsi('Base commerciale créée !'));
    end;
  finally
    FreeAndNil(StreamCial);
    FreeAndNil(StreamCpta);
    CoUnInitialize;
    Readln;
  end;
end.

