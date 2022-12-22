{******************************************************************************}
{                                                                              }
{ Lazarus Component                                                            }
{ Copyright (c) 2013 MBG partenaires                                           }
{ Unit owner: Thierry Bothorel                                                 }
{ Version: 1                                                                   }
{ Subversion:                                                                  }
{   $Id:: CreationBaseCial.lpr 22 2013-07-06 11:52:17Z TBOTHOREL             $ }
{                                                                              }
{******************************************************************************}

program CreationBaseCial;

{$APPTYPE CONSOLE}{$mode objfpc}{$H+}

uses
  Interfaces, // sinon Error: Undefined symbol: WSRegisterCustomImageList
  SysUtils,  // sinon Error: Identifier not found "Exception"
  ActiveX,
  objets100clib_tlb;

var
  StreamCpta  : TAxcBSCPTAApplication100c;
  BaseCpta    : IBSCPTAApplication3;
  StreamCial  : TAxcBSCIALApplication100c;
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
        E.Message);
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
        E.Message);
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
        E.Message);
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
  StreamCial  := TAxcBSCIALApplication100c.Create(nil);
  BaseCial    := StreamCial.OleServer;

  try
    if CreeBaseCial(BaseCial, 'E:\DATA\Gestion\SQL2017\100c_v7\TEST_OM_ExemplesSupport0808.gcm',
      BaseCpta, 'E:\DATA\Gestion\SQL2017\100c_v7\TEST_OM_ExemplesSupport0808.mae')
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

