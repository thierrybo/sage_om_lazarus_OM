{******************************************************************************}
{                                                                              }
{ Lazarus Component                                                            }
{ Copyright (c) 2009 MBG partenaires                                           }
{ Unit owner: Thierry Bothorel                                                 }
{ Version: 1                                                                   }
{ Subversion:                                                                  }
{   $Id::                                                                    $ }
{                                                                              }
{******************************************************************************}

program OuvertureFermetureBaseCpta;

uses
  Interfaces, // sinon Error: Undefined symbol: WSRegisterCustomImageList
  SysUtils,  // sinon Error: Identifier not found "Exception"
  ActiveX,
  Objets100Lib_3_0_TLB
  ;

var
  StreamCpta  : TAxcBSCPTAApplication3;
  BaseCpta    : IBSCPTAApplication3;

function OuvreBaseCpta(
  var ABaseCpta : IBSCPTAApplication3;
  ANomBaseCpta  : string;
  AUtilisateur  : string = '';
  AMotDePasse   : string = ''): Boolean;
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

begin
  // Initialize COM. ------------------------------------------
  CoInitializeEx(nil, COINIT_MULTITHREADED);

  StreamCpta  := TAxcBSCPTAApplication3.Create(nil);
  BaseCpta    := StreamCpta.OleServer;
  try

    { Au préalable, créer un utilisateur DURANT ayant pour mot de passe 1234 : }
    if (OuvreBaseCpta(BaseCpta, 'C:\Temp\BIJOU.MAE', 'DURANT', '1234')) then
    begin
      writeln('Base comptable ', BaseCpta.Name, ' ouverte !');
      if FermeBaseCpta(BaseCpta) then
      begin
        Writeln('Base comptable ', BaseCpta.Name, UTF8ToAnsi(' fermée !'));
      end;
    end;
  finally
    FreeAndNil(StreamCpta);
    CoUnInitialize;
    Readln;
  end;
end.

