{******************************************************************************}
{                                                                              }
{ Lazarus Component                                                            }
{ Copyright (c) 2013 MBG partenaires                                           }
{ Unit owner: Thierry Bothorel                                                 }
{ Version: 1                                                                   }
{ Subversion:                                                                  }
{   $Id:: EcrituresComptables.lpr 37 2013-07-07 09:50:20Z TBOTHOREL          $ }
{                                                                              }
{******************************************************************************}

program ExoDocuments;

{$APPTYPE CONSOLE}{$mode objfpc}{$H+}

uses
  Interfaces, // sinon Error: Undefined symbol: WSRegisterCustomImageList
  SysUtils,  // sinon Error: Identifier not found "Exception"
  ActiveX,
  Objets100cLib_TLB,
  commun,
  IterateurEnumVariant,
  wmiutil;

var
  StreamCial    : TAxcBSCIALApplication100c;
  BaseCial      : IBSCIALApplication3;
  DocVente      : IBODocumentVente3;
  DocStock      : IBODocumentStock3;

function CreeEnteteStockME(
  var ABaseCial : IBSCIALApplication3;
  ADepot        : string;
  ADateDoc      : TDateTime): IBODocumentStock3;
var
  DocStock      : IBODocumentStock3;
begin

  try
    DocStock := ABaseCial.FactoryDocumentStock.CreateType(DocumentTypeStockMouvIn);
    with DocStock do
    begin
      DO_Date       := ADateDoc;
      DepotStockage := ABaseCial.FactoryDepot.ReadIntitule(ADepot);
      SetDefault;
      WriteDefault;
    end;
    Result := DocStock;
  except on E: Exception do
    begin
      Writeln(E.ClassName, ': ', E.Message);
      Result := nil;
    end;
  end;
end;

function CreeEnteteVenteBC(
  var ABaseCial : IBSCIALApplication3;
  AClient       : string;
  ADateDoc      : TDateTime): IBODocumentVente3;
var
  DocVente      : IBODocumentVente3;
begin

  try
    DocVente := ABaseCial.FactoryDocumentVente.CreateType(
                DocumentTypeVenteCommande);
    with DocVente do
    begin
      DO_Date := ADateDoc;
      SetDefaultClient(
        ABaseCial.CptaApplication.FactoryTiers.ReadNumero(AClient)
        as IBOClient3);
      Write_;
    end;
    Result := DocVente;

  except on E: Exception do
    begin
      Writeln(E.ClassName, ': ', E.Message);
      Result := nil;
    end;
  end;
end;

function CreeEnteteVenteFA(
  var ABaseCial : IBSCIALApplication3;
  AClient       : string;
  ADateDoc      : TDateTime): IBODocumentVente3;
var
  DocVente      : IBODocumentVente3;
begin

  try
    DocVente := ABaseCial.FactoryDocumentVente.CreateType(
                DocumentTypeVenteFacture);
    with DocVente do
    begin
      DO_Date := ADateDoc;
      SetDefaultClient(
        ABaseCial.CptaApplication.FactoryTiers.ReadNumero(AClient)
        as IBOClient3);
      Write_;
    end;
    Result := DocVente;

  except on E: Exception do
    begin
      Writeln(E.ClassName, ': ', E.Message);
      Result := nil;
    end;
  end;
end;

function CreeEnteteVenteFR(
  var ABaseCial : IBSCIALApplication3;
  AClient       : string;
  ADateDoc      : TDateTime): IBODocumentVente3;
var
  DocVente      : IBODocumentVente3;
begin

  try
    DocVente := ABaseCial.FactoryDocumentVente.CreateFacture(DocProvenanceRetour);
    with DocVente do
    begin
      DO_Date := ADateDoc;
      SetDefaultClient(
        ABaseCial.CptaApplication.FactoryTiers.ReadNumero(AClient)
        as IBOClient3);
      Write_;
    end;
    Result := DocVente;

  except on E: Exception do
    begin
      Writeln(E.ClassName, ': ', E.Message);
      Result := nil;
    end;
  end;
end;

function CreeLigneStock(
  var ADocStock       : IBODocumentStock3;
  ARefArticle         : string;
  AQte                : Double;
  ACodeEmplacement    : string = ''): IBODocumentStockLigne3;
var
  Depot               : IBODepot3;
  LigneDocStock       : IBODocumentStockLigne3;
  DocLigneEmplacement : IBODocumentLigneEmplacement;
begin

  try
    Depot             := ADocStock.DepotStockage;
    LigneDocStock     := ADocStock.FactoryDocumentLigne.Create
                         as IBODocumentStockLigne3;
    LigneDocStock.SetDefaultArticleReference(ARefArticle, AQte);

    if ACodeEmplacement = '' then
    begin

      { Mouvemente l'emplacement principal }
      LigneDocStock.WriteDefault;
    end
    else
    begin

      { Mouvemente l'emplacement sélectionné }
      LigneDocStock.Write_;
      DocLigneEmplacement := LigneDocStock.FactoryDocumentLigneEmplacement
                             .Create as IBODocumentLigneEmplacement;
      with DocLigneEmplacement do
      begin
        DL_Qte      := AQte;
        Emplacement := Depot.FactoryDepotEmplacement.ReadCode(ACodeEmplacement);
        WriteDefault;
      end;
    end;
    Result := LigneDocStock;

  except on E: Exception do
    begin
      Writeln(E.ClassName, ': ', E.Message);
      Result := nil;
    end;
  end;
end;

function CreeLot(
  var AArticle       : IBOArticle3;
  var ADepot         : IBODepot3;
  ANumLot            : string;
  ADatePeremption    : TDateTime;
  AComplementSerieLot: string): IBOArticleDepotLot;
var
  Lot                : IBOArticleDepotLot;
  iArticleDepot      : OleVariant; // Usage itération standard Delphi IenumVariant
  IEnum              : IEnumVARIANT; // Usage itération standard Delphi IenumVariant
  Nombre             : LongWord; // Usage itération standard Delphi IenumVariant
begin

  try
    { Parcours des dépôts de l'article : }
	{ TODO: ATTENTION déplacer l'appel jusqu'à la collection (.List) dans une
		variable sinon à chaque appel il doit recréer la liste => peut être long }
    //IEnum := IUnknown(AArticle.FactoryArticleDepot.List._NewEnum) as IEnumVARIANT;
    IEnum := AArticle.FactoryArticleDepot.List._NewEnum as IEnumVARIANT;
    while IEnum.Next(1, iArticleDepot, Nombre) = S_OK do
    begin

      { Identification du dépôt de l'article correspondant à l'intitulé : }
      if (IUnknown(iArticleDepot) as IBOArticleDepot3).Depot.DE_Intitule =
          ADepot.DE_Intitule then
      begin

        { Création d'un nouveau lot : }
        Lot := (IUnknown(iArticleDepot) as IBOArticleDepot3).
               FactoryArticleDepotLot.Create as IBOArticleDepotLot;

        with Lot do
        begin
          NoSerie         := ANumLot;
          DatePeremption  := ADatePeremption;
          Complement      := AComplementSerieLot;
        end;

        Break;

      end;
    end;

    Result := Lot;

  except on E: Exception do
    begin
      Writeln(E.ClassName, ': ', E.Message);
      Result := nil;
    end;
  end;

end;

function CreeLigneStockLot(
  var ADocStock      : IBODocumentStock3;
  ARefArticle        : string;
  AQte               : Double;
  ANumLot            : string;
  ADatePeremption    : TDateTime;
  AComplementSerieLot: string;
  ACodeEmplacement   : string = ''): IBODocumentStockLigne3;
var
  BaseCial           : IBSCIALApplication3;
  Article            : IBOArticle3;
  Depot              : IBODepot3;
  Lot                : IBOArticleDepotLot;
  LigneDocStock      : IBODocumentStockLigne3;
begin

  try
    BaseCial  := ADocStock.Stream as IBSCIALApplication3;
    Article   := BaseCial.FactoryArticle.ReadReference(ARefArticle);
    Depot     := ADocStock.DepotStockage;

    { Création d'un lot en fonction de l'article et du dépôt :}
    Lot := CreeLot(
              Article,
              Depot,
              ANumLot,
              ADatePeremption,
              AComplementSerieLot);

    LigneDocStock := ADocStock.FactoryDocumentLigne.Create
                     as IBODocumentStockLigne3;
    LigneDocStock.SetDefaultLot(Lot, AQte);
    LigneDocStock.WriteDefault;

    Result := LigneDocStock;

  except on E: Exception do
    begin
      Writeln(E.ClassName, ': ', E.Message);
      Result := nil;
    end;
  end;

end;

function CreeLigneVente(
  var ADocVente : IBODocumentVente3;
  ARefArticle   : string;
  AQte          : Double): IBODocumentVenteLigne3;
var
  LigneDocVente : IBODocumentVenteLigne3;
begin

  try
      LigneDocVente := ADocVente.FactoryDocumentLigne.Create as
                       IBODocumentVenteLigne3;
      with LigneDocVente do
      begin
        SetDefaultArticleReference(ARefArticle, AQte);
        SetDefaultRemise;
        WriteDefault;
      end;

      Result := LigneDocVente;

  except on E: Exception do
    begin
      Writeln(E.ClassName, ': ', E.Message);
      Result := nil;
    end;
  end;

end;

function CreeLigneVenteArtNomenclature(
  var ADocVente : IBODocumentVente3;
  ARefArticle   : string;
  AQteCompose   : Double): IBODocumentVenteLigne3;
var
  LigneDocVente        : IBODocumentVenteLigne3;
  BaseCial             : IBSCIALApplication3;
  LArticleCompose      : IBOArticle3;
  QteComposant         : Double;
  iArticleNomenclature : OleVariant; // Usage itération standard Delphi IenumVariant
  IEnum                : IEnumVARIANT; // Usage itération standard Delphi IenumVariant
  Nombre               : LongWord; // Usage itération standard Delphi IenumVariant
  ArticleNomenclature  : IBOArticleNomenclature3;
begin

  try
    LigneDocVente   := ADocVente.FactoryDocumentLigne.Create as
                       IBODocumentVenteLigne3;
    BaseCial        := ADocVente.Stream as IBSCIALApplication3;
    LArticleCompose := BaseCial.FactoryArticle.ReadReference(ARefArticle);

    { Insertion de la ligne du composé : }
    with LigneDocVente do
    begin
      SetDefaultArticleReference(ARefArticle, AQteCompose);
      ArticleCompose := LArticleCompose;
      SetDefaultRemise;
      WriteDefault;
    end;

    { Parcours des composants de la nomenclature et insertion des lignes :  }
	{ TODO: ATTENTION déplacer l'appel jusqu'à la collection (.List) dans une
		variable sinon chaque appel il doit recréer la liste => peut etre long }
    //IEnum := IUnknown(LArticleCompose.FactoryArticleNomenclature.List._NewEnum) as IEnumVARIANT;
    IEnum := LArticleCompose.FactoryArticleNomenclature.List._NewEnum as IEnumVARIANT;
    while IEnum.Next(1, iArticleNomenclature, Nombre) = S_OK do
    begin
      ArticleNomenclature := IUnknown(iArticleNomenclature) as IBOArticleNomenclature3;
      LigneDocVente       := ADocVente.FactoryDocumentLigne.Create as
                             IBODocumentVenteLigne3;

      { Calcul de la qté du composant en fonction du type (fixe ou variable) : }
      if ArticleNomenclature.NO_Type = ComposantTypeVariable then
      begin
          QteComposant := ArticleNomenclature.NO_Qte * AQteCompose
      end
      else
      begin
          QteComposant := ArticleNomenclature.NO_Qte
      end;
      with LigneDocVente do
      begin
        SetDefaultArticle(ArticleNomenclature.ArticleComposant, QteComposant);
        ArticleCompose := LArticleCompose;
        SetDefaultRemise;
        WriteDefault;
      end;
    end;

    Result := LigneDocVente;

  except on E: Exception do
    begin
      Writeln(E.ClassName, ': ', E.Message);
      Result := nil;
    end;
  end;

end;

function GetNumSerieDispo(
  var AArticle                 : IBOArticle3;
  ADepot                       : IBODepot3): IBOArticleDepotLot;
var
  ArticlesSerialisesNonEpuises : IBICollection;
  iArticleSerialise            : OleVariant; // Usage itération standard Delphi IenumVariant
  IEnum                        : IEnumVARIANT; // Usage itération standard Delphi IenumVariant
  Nombre                       : LongWord; // Usage itération standard Delphi IenumVariant
  ArticleDepot                 : IBOArticleDepot3;
  iArticleDepot                : OleVariant; // Usage itération standard Delphi IenumVariant
begin

  try

    { Parcours des dépôts de l'article : }
	{ TODO: ATTENTION déplacer l'appel jusqu'à la collection (.List) dans une
		variable sinon chaque appel il doit recréer la liste => peut être long }
    //IEnum := IUnknown(AArticle.FactoryArticleDepot.List._NewEnum) as IEnumVARIANT;
    IEnum := AArticle.FactoryArticleDepot.List._NewEnum as IEnumVARIANT;
    while IEnum.Next(1, iArticleDepot, Nombre) = S_OK do
    begin

      { Identification du dépôt de l'article correspondant à l'intitulé : }
      ArticleDepot := IUnknown(iArticleDepot) as IBOArticleDepot3;
      if ArticleDepot.Depot.DE_Intitule = ADepot.DE_Intitule then
      begin

        { Récupération d'une collection de N° série non épuisés : }
        ArticlesSerialisesNonEpuises :=
          ArticleDepot.FactoryArticleDepotLot.QueryNonEpuise;

        { Parcours de la collection et retour du 1er lot non réservé : }
        //IEnum := IUnknown(ArticlesSerialisesNonEpuises._NewEnum) as IEnumVARIANT;
        IEnum := ArticlesSerialisesNonEpuises._NewEnum as IEnumVARIANT;
        while IEnum.Next(1, iArticleSerialise, Nombre) = S_OK do
        begin
          if (IUnknown(iArticleSerialise) as IBOArticleDepotLot).StockATerme = 1 then
            Result := IUnknown(iArticleSerialise) as IBOArticleDepotLot;
        end;

      end;

    end;

  except on E: Exception do
    begin
      Writeln(E.ClassName, ': ', E.Message);
      Result := nil;
    end;
  end;

end;

function CreeLigneVenteArtSerialise(
  ADocVente         : IBODocumentVente3;
  ARefArticle       : string): IBODocumentVenteLigne3;
var
  LigneDocVente     : IBODocumentVenteLigne3;
  BaseCial          : IBSCIALApplication3;
  LArticle          : IBOArticle3;
  LArticleSerialise : IBOArticleDepotLot;
begin

  try
    LigneDocVente   := ADocVente.FactoryDocumentLigne.Create as
                       IBODocumentVenteLigne3;
    BaseCial        := ADocVente.Stream as IBSCIALApplication3;
    LArticle        := BaseCial.FactoryArticle.ReadReference(ARefArticle);

    { Obtention d'un N° de série non épuisé et non réservé : }
    LArticleSerialise := GetNumSerieDispo(LArticle, ADocVente.DepotStockage);

    if LArticleSerialise = nil then
    begin
      raise Exception.Create('Stock insuffisant pour l''article ' + ARefArticle
                             + ' !');
    end
    else
    begin
      with LigneDocVente do
      begin
        SetDefaultLot(LArticleSerialise, 1);
        SetDefaultRemise;
        WriteDefault;
      end;
    end;

    Result := LigneDocVente;

  except on E: Exception do
    begin
      Writeln(E.ClassName, ': ', E.Message);
      Result := nil;
    end;
  end;

end;

function ModifieInfoLibreLigne(
  ALigneDoc     : IBODocumentLigne3;
  ANomInfoLibre : string;
  AValInfoLibre : OleVariant): Boolean;
var
  temp          : IBIValues;
begin

  { Modifie une info libre d'une ligne existante : }
  try
    temp := ALigneDoc.InfoLibre;
    //ALigneDoc.InfoLibre.Item[ANomInfoLibre] := AValInfoLibre;
    temp.Item[ANomInfoLibre] := AValInfoLibre;
    ALigneDoc.Write_;
    Result := true;
  except on E: Exception do
    begin
      Writeln(E.ClassName, ': ', E.Message);
      Result := false;
    end;
  end;
end;

function CreeAcompteVente(
  var ADocVente   : IBODocumentVente3;
  AMontant        : double;
  AModeRegl       : string;
  ADateRegl       : TDateTime;
  ALibelle        : string): Boolean;
var
  Acompte         : IBODocumentAcompte3;
  BaseCpta        : IBSCPTAApplication3;
begin
  try
    Acompte       := ADocVente.FactoryDocumentAcompte.Create
                     as IBODocumentAcompte3;
    BaseCpta      := (ADocVente.Stream as BSCIALApplication100c).CptaApplication;
    with Acompte do
    begin
      DR_Date     := ADateRegl;
      DR_Montant  := AMontant;
      DR_Libelle  := ALibelle;
      Reglement   := BaseCpta.FactoryReglement.ReadIntitule(AModeRegl);
      WriteDefault;
    end;
    Result := true;
  except on E: Exception do
    begin
      Writeln(E.ClassName, ': ', E.Message);
      Result := false;
    end;
  end;
end;

function CreeReglementVente(
  var ADocVte   : IBODocumentVente3;
  AReference    : string;
  ALibelle      : string;
  AMontant      : double;
  AJournal      : string;
  AModeRegl     : string): Boolean;
var
  iReglt  : IBODocumentReglement;
  Client  : IBOTiersPart3;
  pRegler : IPMReglerEcheances;
  BaseCial : BSCIALApplication100c;
  iEcheance : IUnknown; // Usage itération L.DARDENNE IenumVariant
  IEnum   : TEnumVARIANT; // Usage itération L.DARDENNE IenumVariant

begin

  Result := False;
  try
      // Objet Base Gestion commerciale et Client

    BaseCial := ADocVte.Stream as BSCIALApplication100c;
    Client := ADocVte.TiersPayeur;

      //Création du règlement

      iReglt := BaseCial.FactoryDocumentReglement.Create as IBODocumentReglement;
      with iReglt do
      begin
        TiersPayeur := Client;
        RG_Date := Now;
        RG_Reference := AReference;
        RG_Libelle := ALibelle;
        RG_Montant := AMontant;
        //RG_Montant := ADocVte.DO_NetAPayer - ADocVte.DO_MontantRegle;
        Reglement   := BaseCial.CptaApplication.FactoryReglement.ReadIntitule(AModeRegl);
        Journal := BaseCial.CptaApplication.FactoryJournal.ReadNumero(AJournal);
        CompteG := Client.CompteGPrinc;
        WriteDefault;
      end;

      // Création du Processus régler les échéances

      pRegler := BaseCial.CreateProcess_ReglerEcheances;
      pRegler.Reglement := iReglt;

      { Parcours des échéances de la facture : }

  	{ TODO: ATTENTION déplacer l'appel jusqu'à la collection (.List) dans une
  		variable sinon à chaque appel il doit recréer la liste => peut être long }

      // Itération L.DARDENNE IenumVariant
      IEnum := TEnumVariant.Create(ADocVte.FactoryDocumentEcheance.List);
      For iEcheance in IEnum do;
      begin
        pRegler.AddDocumentEcheanceMontant(iEcheance as IBODocumentEcheance3, AMontant);
      end;

      pRegler.Process;

      Result := True;

  except on E: Exception do
    begin
      Writeln(E.ClassName, ': ', E.Message);
      Result := False;
    end;
  end;

end;


procedure AfficheEcheances(var ADocVente: IBODocumentVente3);
var
  TotMontantTTC   : Double;
  MontantTTC      : Double;
  TotPourcentTTC  : Double;
  PourcentTTC     : Double;
  iEcheance       : OleVariant; // Usage itération standard Delphi IenumVariant
  IEnum           : IEnumVARIANT; // Usage itération standard Delphi IenumVariant
  Nombre          : LongWord; // Usage itération standard Delphi IenumVariant
  Echeance        : IBODocumentEcheance3;
begin

  TotPourcentTTC  := 0;
  Writeln(sLineBreak);
  Writeln(UTF8ToAnsi('Echéances :'));
  Writeln(UTF8ToAnsi('Date         % TTC       Montant     Mode règlement'));
  try

    { Calul du TTC du document, acomptes déduits : }
    TotMontantTTC := ADocVente.Valorisation.TotalTTC
                   - ADocVente.Valorisation.TotalAcomptes;

    { Parcours des échéances du document : }
	{ TODO: ATTENTION déplacer l'appel jusqu'à la collection (.List) dans une
		variable sinon chaque appel il doit recréer la liste => peut etre long }
    //IEnum := IUnknown(ADocVente.FactoryDocumentEcheance.List._NewEnum) as IEnumVARIANT;
    IEnum := ADocVente.FactoryDocumentEcheance.List._NewEnum as IEnumVARIANT;
    while IEnum.Next(1, iEcheance, Nombre) = S_OK do
    begin

      { Calcul du montant TTC des échéances et du % de la ligne d'équilibre : }
      Echeance := IUnknown(iEcheance) as IBODocumentEcheance3;
      with Echeance do
      begin

        { Si échéance d'équilibre, calcul du pourcentage : }
        if DR_Equil then
        begin
          PourcentTTC := 100 - TotPourcentTTC
        end

        { Sinon cumul du % (calcul ultérieur % équilibre) : }
        else
        begin
          PourcentTTC     := DR_Pourcent;
          TotPourcentTTC  := TotPourcentTTC + DR_Pourcent;
        end;

        { Calcul de TTC de l'échéance }
        MontantTTC := TotMontantTTC / 100 * PourcentTTC;

        Writeln(
          DateToStr(DR_Date),
          '      ',
          FloatToStr(PourcentTTC),
          '%     ',
          FloatToStr(MontantTTC),
          '     ',
          Reglement.R_Intitule);
      end;
    end;

  except on E: Exception do
    Writeln(E.ClassName, ': ', E.Message);
  end;

end;

procedure AfficheValorisation(var ADocVente: IBODocumentVente3);
begin

  Writeln(sLineBreak);
  Writeln('Valorisation :');
  Writeln('Total HT         Total HT net       Total TTC');

  try
    with ADocVente.Valorisation do
    begin
      Writeln(
        FloatToStr(TotalHT),
        '      ',
        FloatToStr(TotalHTNet),
        '     ',
        FloatToStr(TotalTTC));
    end;
  except on E: Exception do
    Writeln(E.ClassName, ': ', E.Message);
  end;
end;

procedure AfficheTaxes(var ADocVente: IBODocumentVente3);
var
  iTaxe  : OleVariant; // Usage itération standard Delphi ET Marcov IenumVariant
  oEnum  : oEnumIterator; // Usage itération Marcov IenumVariant
  Taxe   : IDocValoTaxe;
begin

  Writeln(sLineBreak);
  Writeln('Taxes :');
  Writeln('Base         Taux       Montant');

  try

    { Affiche chacune des taxes de la ligne : }
    // Itération MARCOV IenumVariant
    for iTaxe in oEnum.Enumerate(ADocVente.Valorisation.Taxes) do
    begin
      Taxe := IUnknown(iTaxe) as IDocValoTaxe;
      Writeln(
          FormatFloat('#0.##', Taxe.BaseCalcul),
          '      ',
          FormatFloat('#0.##', Taxe.Taux),
          '%     ',
          FormatFloat('#0.##', Taxe.Montant));
    end;
  except on E: Exception do
    Writeln(E.ClassName, ': ', E.Message);
  end;
end;

begin
  // Initialize COM. ------------------------------------------
  CoInitializeEx(nil, COINIT_MULTITHREADED);

  StreamCial  := TAxcBSCIALApplication100c.Create(nil);
  BaseCial    := StreamCial.OleServer;

  try
    try
      // Si on utilise l'ouverture SQL
      if OuvreBaseCialSql(BaseCial,
        '(local)\SAGE2017',
        'BIJOU_V7',
        '<Administrateur>'
        ) then
      // Si on utilise l'ouverture .gcm
      //if OuvreBaseCial(BaseCial,
      //  'E:\DATA\Gestion\BIJOU-SQL2017\V7\BIJOU_V7.gcm',
      //  '<Administrateur>') then
      begin

        //{ Création d'une Mouvement d'entrée en stock : }

        DocStock := CreeEnteteStockME(BaseCial, 'Bijou SA', Now);
        if not (DocStock = nil) then
        begin

          { Entrée en stock sur l'emplacement principal : }
          CreeLigneStock(DocStock, 'BRAAR10', 50);

          { Entrée en stock sur l'emplacement indiqué : }
          CreeLigneStock(DocStock, 'BAAR01', 100, 'A2T1N2P3');

          { Entrée en stock d'un article géré par lot : }
          CreeLigneStockLot(
              DocStock,
              'LINGOR18',
              5,
              'LOT001',
              StrToDate('31/12/19'),
              '12345678',
              'A3T1N2P1');
        end;

         { Création d'un Bon de commande client : }

        DocVente := CreeEnteteVenteBC(BaseCial, 'CARAT', Now);
        if not (DocVente = nil) then
        begin

          { Article géré au CMUP : }
          CreeLigneVente(DocVente, 'BRAAR10', 5);

          { Article à conditionnement : }
          CreeLigneVente(DocVente, 'EM040/24', 2);

          { Article à double gamme : }
          CreeLigneVente(DocVente, 'CHAARVARC34', 2);

          { Affectation d'une valeur à une info libre ligne : }
          { #todo 1 : EXCEPTION "Type de variable incorrecte" }
          //ModifieInfoLibreLigne(
          //  CreeLigneVente(DocVente, 'BRAAR10', 1) as IBODocumentLigne3,
          //  'Commentaires',
          //  'Extension de garantie : 3 ans');

          { Article à nomenclature commerciale : }
          CreeLigneVenteArtNomenclature(DocVente, 'ENSHF', 2);

          { Article géré par N° de série : }
          CreeLigneVenteArtSerialise(DocVente, 'MOBWAC01');

          writeln(CreeAcompteVente(DocVente, 1, Utf8ToAnsi('Chèque'), Now, Utf8ToAnsi('Libellé')));
          //writeln(CreeReglementVente(DocVente, Utf8ToAnsi('Référence'), Utf8ToAnsi('Libellé'), 3.47 * 10000, Utf8ToAnsi('BEU'), Utf8ToAnsi('Espèces')));
          AfficheValorisation(DocVente);
          AfficheTaxes(DocVente);
          { #todo : Si Loi AntiFraude affiche "EOleException : L'échéance doit provenir d'une facture validée !" }
          AfficheEcheances(DocVente);
        end;

        { Création d'une Facture de Retour client : }

        DocVente := CreeEnteteVenteFR(BaseCial, 'CARAT', Now);
        if not (DocVente = nil) then
        begin

          { Article géré au CMUP : }
          CreeLigneVente(DocVente, 'BRAAR10', -5);

          { Article à conditionnement : }
          //CreeLigneVente(DocVente, 'EM040/24', 2);

          { Article à double gamme : }
          //CreeLigneVente(DocVente, 'CHAARVARC34', 2);

          { Affectation d'une valeur à une info libre ligne : }
          {
           ModifieInfoLibreLigne(
               CreeLigneVente(DocVente, 'BRAAR10', 1) as IBODocumentLigne3,
               'Commentaires',
               'Extension de garantie : 3 ans');
          }

          { Article à nomenclature commerciale : }
          //CreeLigneVenteArtNomenclature(DocVente, 'ENSHF', 2);

          { Article géré par N° de série : }
          //CreeLigneVenteArtSerialise(DocVente, 'MOBWAC01');

          //CreeAcompteVente(DocVente, 2.47, Utf8ToAnsi('Chèque'), Now, Utf8ToAnsi('Libellé'));
          writeln(CreeReglementVente(DocVente, Utf8ToAnsi('Référence'), Utf8ToAnsi('Libellé'), -10, Utf8ToAnsi('BEU'), Utf8ToAnsi('Espèces')));
          AfficheValorisation(DocVente);
          AfficheTaxes(DocVente);
          AfficheEcheances(DocVente);
        end;
      end;
    except on E: Exception do
      Writeln(E.ClassName, ': ', E.Message);
    end;

  finally
    FermeBaseCial(BaseCial);
    FreeAndNil(StreamCial);
    CoUnInitialize;
    Readln;
  end;
end.

