{
 https://laurent-dardenne.developpez.com/articles/Delphi/2005/langage/win32/iterateurIEnumVariant/
}

unit IterateurEnumVariant;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, Windows, SysUtils, ActiveX, ComObj;

type
  // Conteneur générique pour une collection IEnumVariant
  TCustomEnumerateurEnumVariant = class(TObject)
  private
    // Contient la collection
    FObjetCollection: IEnumVariant;
    // Accés remote possible
    FItem: olevariant;
  public
    constructor Create(Enum: IDispatch);
    function GetNextItem: boolean;
    function GetCollection(ACollection: IDispatch): IEnumVariant;

    property Item: olevariant read FItem;
    property NextItem: boolean read GetNextItem;
  end;

  TEnumerateur = class;

  // Conteneur spécialisé pour une interface implémentant IEnumVariant
  TEnumVariant = class(TCustomEnumerateurEnumVariant)
    //Méthode nécessaire pour la prise en charge de l'itération
    function GetEnumerator: TEnumerateur;
  end;

  // Enumérateur pour la classe conteneur.
  // Il doit implémenter une suite de méthodes comme indiquée ci-dessous.
  TEnumerateur = class(TObject)
  strict private
    FListe: TEnumVariant; //Référence la collection concernée par l'itération
  public
    constructor Create(AList: TEnumVariant);
    //Membre de classse nécessaire pour la prise en charge de l'itération
    function GetCurrent: IUnknown;
    function MoveNext: boolean;    // Méthode publique
    property Current: IUnknown read GetCurrent; // Propriété en Read Only
  end;

implementation

{ TCustomEnumeratorEnumVariant --------------------------------------------------- }
constructor TCustomEnumerateurEnumVariant.Create(Enum: IDispatch);
  // On utilise une interface IDispatch afin de pouvoir passer différent type d'interface
  // Comme on ne connaît pas son 'type' on appellera dynamiquement la méthode '_NewEnum' commune aux interfaces
  // implémentant IEnumVariant (renvoi une collection de variant).
begin
  // E2003 : Identificateur non déclaré : '_NewEnum'
  //FObjetEnumerateur:=IUnKnown(Enum._NewEnum) as IEnumVariant
  FObjetCollection := GetCollection(Enum);
end;

function TCustomEnumerateurEnumVariant.GetNextItem: boolean;
  // Récupére dans la propriété FItem la prochaine valeur de la collection FObjetCollection.
  // Renvoi False si la fin de la collection est atteinte.
var
  NombreElement: longword;
begin
  Result := (FObjetCollection.Next(1, FItem, NombreElement) = S_OK);
end;

function TCustomEnumerateurEnumVariant.GetCollection(ACollection:
  IDispatch): IEnumVariant;
  // Renvoi la collection par l'appel à '_NewEnum' que chaque Interface IEnumVariant implémente.
  // Déclenche 'EOleSysError: Nom inconnu' si l'interface n'implémente pas '_NewEnum'

var
  VarResult: olevariant; // Résultat de la méthode Invoke
  Params: TDispParams;    // Tableau de paramètre, ici 0
  lProperty: WideString;  // Contient le nom de la propriété à appeler
  lDispID: integer;       // Numéro de DispId de la propriété
  Resultat: HResult;      // Résultat de l'appel COM
  ExcepInfo: TExcepInfo;  // En cas d'erreur d'appel

begin
  Result := nil;
  lProperty := '_NewEnum';
  FillChar(Params, SizeOf(DispParams), 0);

  // Recherche le numéro d'identificateurs de dispatch (dispID) de la propriété contenue dans lProperty
  OLECheck(ACollection.GetIDsOfNames(GUID_NULL, @lProperty,
    1, LOCALE_USER_DEFAULT, @ldispid));

  // Appel de la propriété de l'interface
  Resultat := ACollection.Invoke(ldispid, GUID_NULL, 0, DISPATCH_PROPERTYGET, Params, @VarResult,
    @ExcepInfo, nil);

  // En cas d'erreur on doit lever une exception
  if Resultat <> 0 then DispatchInvokeError(Resultat, ExcepInfo);

  //Transtype le résultat obtenu en Collection de variant
  Result := IUnknown(VarResult) as IEnumVariant;
end;

{ TEnumVariant --------------------------------------------------- }
// Les Interface WMI ISWbemXXXXSet sont des collections
// On utilise une interface IDispatch afin de pouvoir passer différent type d'interface WMI
// Comme on ne connait pas son 'type' on appellera dynamiquement la méthode commune aux interfaces WMI
// qui renvoi la collection de variant.

function TEnumVariant.GetEnumerator: TEnumerateur;
  // Appelé par le compilateur
begin
  Result := TEnumerateur.Create(Self);
end;

{ TEnumerateur  ---------------------------------------------------}

constructor TEnumerateur.Create(AList: TEnumVariant);
  // Créé par la classe 'itérable'
begin
  FListe := AList;
end;

{Destructor TEnumerateur.Destroy;
begin
 inherited;
end;
}

function TEnumerateur.GetCurrent: IUnknown;
  //Renvoi la dernière valeur lue à partir de la collection FObjetEnumerateur: IEnumVariant
begin
  Result := FListe.Item;
end;

function TEnumerateur.MoveNext: boolean;
  // Se positionne sur la prochaine valeur si elle existe.
begin
  Result := FListe.GetNextItem;
end;

end.
