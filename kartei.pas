unit kartei;

{$mode objfpc}{$H+}

interface

uses
  Classes,
  SysUtils,
  FileUtil,
  Forms,
  Controls,
  Graphics,
  Dialogs,
  DBCtrls,
  StdCtrls,
  ComCtrls,
  ExtCtrls,
  DBGrids,
  EditBtn,
  Menus,
  types;

type

  { TfrmKartei }

  TfrmKartei = class(TForm)
    btnCopyBesuch: TButton;
    btnDelKind: TButton;
    btnDel: TButton;
    btnNew: TButton;
    btnNewBesuch: TButton;
    btnNewKind: TButton;
    btnNewKomm: TButton;
    btnDelKomm: TButton;
    btnDelBesuch: TButton;
    btnEnde: TButton;
    btnPrint: TButton;
    cbAbgaenge: TCheckBox;
    cbFilterMark: TCheckBox;
    cbBearbeiten: TCheckBox;
    cbMarkiert: TCheckBox;
    cbGemeinde: TComboBox;
    DateEditKomm: TDateEdit;
    DBCBAnrede: TDBComboBox;
    DBCBGeschlecht: TDBComboBox;
    DBCBGemeinde: TDBComboBox;
    dbcbUebertrittAus: TDBComboBox;
    dbcbKindKirche: TDBComboBox;
    dbcbMutterKirche: TDBComboBox;
    dbcbEhegatteKirche: TDBComboBox;
    dbcbUebertrittNach: TDBComboBox;
    dbcbVaterKirche: TDBComboBox;
    DBCB_AktivesMitglied: TDBCheckBox;
    DBCB_AnZahlungErinnern: TDBCheckBox;
    DBCB_GD_Besuch: TDBCheckBox;
    DBCB_Ausbildung: TDBCheckBox;
    DBCB_AktiverZahler: TDBCheckBox;
    DBCB_Ruhestand: TDBCheckBox;
    DBCBAbgaenge: TDBCheckBox;
    DBCBFamstand: TDBComboBox;
    dbcbKirche: TDBComboBox;
    dbediAktivitaet1: TDBEdit;
    dbediAktivitaet2: TDBEdit;
    dbediAktivitaet3: TDBEdit;
    dbediAnfangF2: TDBEdit;
    dbediAnfangF3: TDBEdit;
    dbediAnmerkungenZuF2: TDBEdit;
    dbediAnmerkungenZuF3: TDBEdit;
    DBEdiAufgewachsenIn: TDBEdit;
    DBEdiBestattungsDatum: TDBEdit;
    DBEdiBestattungsPfarrer: TDBEdit;
    DBEdiFriedhof: TDBEdit;
    DBEdiBestattungsRegNr: TDBEdit;
    DBEdiBestattungsText: TDBEdit;
    dbediEMail1: TDBEdit;
    dbediEMail2: TDBEdit;
    DbEdiEheGatteGeburtstag: TDBEdit;
    dbediGebName: TDBEdit;
    dbediGebOrt: TDBEdit;
    dbediGebStADatum: TDBEdit;
    dbediGebStAOrt: TDBEdit;
    dbediGebStARegNr: TDBEdit;
    DBediGeburtstag: TDBEdit;
    dbediInternet1: TDBEdit;
    dbediInternet2: TDBEdit;
    DBEdiKindTaufdatum: TDBEdit;
    dbediStrasse: TDBEdit;
    DBEdiTodesort: TDBEdit;
    DBEdKindVorname: TDBEdit;
    DBEdKindVorname2: TDBEdit;
    DBEdKindNachname: TDBEdit;
    DBEdiKindGebName: TDBEdit;
    DBEdiKindGeburtstag: TDBEdit;
    DBEdiBesuchsDatum: TDBEdit;
    DBEdiMitbringsel: TDBEdit;
    dbEdiTerminabsprache: TDBEdit;
    DBEdiBibeltext: TDBEdit;
    DBEdiBesuchsgrund: TDBEdit;
    dbediEndeF2: TDBEdit;
    dbediEndeF3: TDBEdit;
    dbediFunktion1: TDBEdit;
    dbediAnfangF1: TDBEdit;
    dbediEndeF1: TDBEdit;
    dbediAnmerkungenZuF1: TDBEdit;
    dbediEhrungen: TDBEdit;
    dbediAnmerkungenZuA1: TDBEdit;
    dbediAnmerkungenZuA2: TDBEdit;
    dbediAnmerkungenZuA3: TDBEdit;
    DBEdiAusschlussDatum: TDBEdit;
    DBEdiAusschlussGrund: TDBEdit;
    dbediAktivSeit: TDBEdit;
    dbediFunktion2: TDBEdit;
    dbediFunktion3: TDBEdit;
    DBEdiUebertrittsAbDatum: TDBEdit;
    DBEdiUebertrittsAbGrund: TDBEdit;
    DBEdiTodStARegNr: TDBEdit;
    DBEdiUebertrittsZuDatum: TDBEdit;
    DBEdiAustrittsDatum: TDBEdit;
    DBEdiTodStAOrt: TDBEdit;
    DBEdiTodStADatum: TDBEdit;
    DBEdiUeberwiesen_nach: TDBEdit;
    DBEdiUeberwiesen_von_Datum: TDBEdit;
    DBEdiUebertrittsZuGrund: TDBEdit;
    DBEdiAustrittsGrund: TDBEdit;
    DBEdiUeberwiesen_von: TDBEdit;
    dbediBerufAusgeuebt: TDBEdit;
    dbediBerufErlernt: TDBEdit;
    dbediFreiFeld1: TDBEdit;
    dbediArbeitgeber: TDBEdit;
    DBEdiFreiFeld3: TDBEdit;
    DBEdiFreiFeld2: TDBEdit;
    DBEdiTrauDatum: TDBEdit;
    DBEdiEheStADatum: TDBEdit;
    DBEdiTodesDatum: TDBEdit;
    DBEdiEintrittsDatum: TDBEdit;
    DBEdiTrauOrt: TDBEdit;
    DBEdiEheStAOrt: TDBEdit;
    DBEdiEintrittsGrund: TDBEdit;
    DBEdiTrauPfarrer: TDBEdit;
    DBEdiEhegatteIn: TDBEdit;
    DBEdiTrauRegNr: TDBEdit;
    DBEdiEheStARegNr: TDBEdit;
    DBEdiTrauSpruch: TDBEdit;
    DBEdiEheGebName: TDBEdit;
    DBEdiTaufDatum: TDBEdit;
    dbediInternet: TDBEdit;
    dbediLand: TDBEdit;
    dbediName: TDBEdit;
    dbediOrt: TDBEdit;
    dbediOrtsteil: TDBEdit;
    dbediPLZ: TDBEdit;
    dbedisAdresszusatz: TDBEdit;
    DBEdiFreiFeld4: TDBEdit;
    DBEdiMutterGebName: TDBEdit;
    DBEdiTaufOrt: TDBEdit;
    DBEdiKonfDatum: TDBEdit;
    DBEdiKonfOrt: TDBEdit;
    DBEdiTaufPate1: TDBEdit;
    DBEdiTrauZeuge4: TDBEdit;
    DBEdiTrauZeuge5: TDBEdit;
    DBEdiTrauZeuge6: TDBEdit;
    DBEdiTaufPate2: TDBEdit;
    DBEdiTaufPate3: TDBEdit;
    DBEdiTaufPate6: TDBEdit;
    DBEdiTaufPate5: TDBEdit;
    DBEdiTaufPate4: TDBEdit;
    DBEdiTrauZeuge1: TDBEdit;
    DBEdiTrauZeuge2: TDBEdit;
    DBEdiTrauZeuge3: TDBEdit;
    DBEdiTaufPfarrer: TDBEdit;
    DBEdiKonfPfarrer: TDBEdit;
    DBEdiKonfRegNr: TDBEdit;
    DBEdiTaufSpruch: TDBEdit;
    DBEdiTaufRegNr: TDBEdit;
    DBEdiKonfSpruch: TDBEdit;
    DBEdiUeberwiesen_nach_Datum: TDBEdit;
    DBEdiVaterName: TDBEdit;
    dbediTelMobil: TDBEdit;
    dbediTelFax: TDBEdit;
    DBEdiVaterGebName: TDBEdit;
    DBEdiMutterName: TDBEdit;
    dbediVersand: TDBEdit;
    dbediTelPrivat: TDBEdit;
    dbediTelDienst: TDBEdit;
    dbediEMail: TDBEdit;
    dbediVorname: TDBEdit;
    dbediTitel: TDBEdit;
    dbediVorname2: TDBEdit;
    DBGridOverView: TDBGrid;
    DBGridDetails: TDBGrid;
    DBMemo1: TDBMemo;
    DBMemBesuchMemo: TDBMemo;
    DBMemoGeschwister: TDBMemo;
    DBNavigator1: TDBNavigator;
    DBNavigator2: TDBNavigator;
    DBNavigator3: TDBNavigator;
    DBNavigatorBesuch: TDBNavigator;
    DBNavigator5: TDBNavigator;
    DBTextAbendmahlsdatum: TDBText;
    DBText_PersonenID: TDBText;
    DBT_LastChange: TDBText;
    gbFilter: TGroupBox;
    gbSchnellsuche: TGroupBox;
    labAlter: TLabel;
    Label1: TLabel;
    Label10: TLabel;
    Label100: TLabel;
    Label101: TLabel;
    Label102: TLabel;
    Label103: TLabel;
    Label104: TLabel;
    Label105: TLabel;
    Label106: TLabel;
    Label107: TLabel;
    Label108: TLabel;
    Label109: TLabel;
    Label11: TLabel;
    Label110: TLabel;
    Label111: TLabel;
    Label112: TLabel;
    Label113: TLabel;
    Label114: TLabel;
    Label115: TLabel;
    Label116: TLabel;
    Label117: TLabel;
    Label118: TLabel;
    Label119: TLabel;
    Label12: TLabel;
    Label120: TLabel;
    Label121: TLabel;
    Label122: TLabel;
    Label123: TLabel;
    Label124: TLabel;
    Label125: TLabel;
    Label126: TLabel;
    Label127: TLabel;
    Label128: TLabel;
    Label129: TLabel;
    Label13: TLabel;
    Label130: TLabel;
    Label131: TLabel;
    Label132: TLabel;
    Label133: TLabel;
    Label134: TLabel;
    Label135: TLabel;
    Label136: TLabel;
    Label137: TLabel;
    Label138: TLabel;
    Label139: TLabel;
    Label14: TLabel;
    Label140: TLabel;
    Label141: TLabel;
    Label142: TLabel;
    Label143: TLabel;
    Label144: TLabel;
    ediSSNachname: TLabeledEdit;
    ediSSVorname: TLabeledEdit;
    Label145: TLabel;
    Label146: TLabel;
    Label147: TLabel;
    Label148: TLabel;
    Label149: TLabel;
    Label21: TLabel;
    Label25: TLabel;
    Label26: TLabel;
    Label27: TLabel;
    Label28: TLabel;
    Label29: TLabel;
    Label3: TLabel;
    Label30: TLabel;
    Label43: TLabel;
    Label44: TLabel;
    Label45: TLabel;
    Label46: TLabel;
    Label47: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label60: TLabel;
    Label89: TLabel;
    Label90: TLabel;
    Label91: TLabel;
    Label92: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label2: TLabel;
    Label20: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    Label24: TLabel;
    Label4: TLabel;
    labFFeld4: TLabel;
    Label49: TLabel;
    Label50: TLabel;
    Label51: TLabel;
    labFFeld1: TLabel;
    Label53: TLabel;
    Label54: TLabel;
    Label55: TLabel;
    Label56: TLabel;
    Label59: TLabel;
    labFFeld3: TLabel;
    labFFeld2: TLabel;
    Label63: TLabel;
    Label64: TLabel;
    Label65: TLabel;
    Label66: TLabel;
    Label67: TLabel;
    Label68: TLabel;
    Label69: TLabel;
    Label7: TLabel;
    Label70: TLabel;
    Label71: TLabel;
    Label72: TLabel;
    Label73: TLabel;
    Label74: TLabel;
    Label75: TLabel;
    Label76: TLabel;
    Label77: TLabel;
    Label78: TLabel;
    Label79: TLabel;
    Label8: TLabel;
    Label80: TLabel;
    Label81: TLabel;
    Label82: TLabel;
    Label83: TLabel;
    Label84: TLabel;
    Label85: TLabel;
    Label86: TLabel;
    Label87: TLabel;
    Label88: TLabel;
    Label9: TLabel;
    Label93: TLabel;
    Label94: TLabel;
    Label95: TLabel;
    Label96: TLabel;
    Label97: TLabel;
    Label98: TLabel;
    Label99: TLabel;
    labRecCnt: TLabel;
    MemoKomm: TMemo;
    mnuMarkThis: TMenuItem;
    mnuDelMark: TMenuItem;
    Panel1: TPanel;
    panNavi: TPanel;
    panDatenbankOverview: TPanel;
    Panel2: TPanel;
    panShowOverview: TPanel;
    pcDetails: TPageControl;
    PopupMenuMarkiert: TPopupMenu;
    RG_Sortierung: TRadioGroup;
    scBarKartei: TScrollBar;
    Shape1: TShape;
    Shape10: TShape;
    Shape11: TShape;
    Shape12: TShape;
    Shape13: TShape;
    Shape15: TShape;
    Shape2: TShape;
    Shape3: TShape;
    Shape4: TShape;
    Shape5: TShape;
    Shape6: TShape;
    Shape7: TShape;
    Shape8: TShape;
    Shape9: TShape;
    TimerBerechnung: TTimer;
    TSMitarbeit: TTabSheet;
    TSMemo: TTabSheet;
    TSBesuch: TTabSheet;
    TSKommunionen: TTabSheet;
    TSAllgemein: TTabSheet;
    pcAdr: TPageControl;
    TSZuAbGang: TTabSheet;
    TSKasualien: TTabSheet;
    tsAdr: TTabSheet;
    TSEltern: TTabSheet;
    procedure btnCopyBesuchClick(Sender: TObject);
    procedure btnDelBesuchClick(Sender: TObject);
    procedure btnDelClick(Sender: TObject);
    procedure btnDelContextPopup(Sender: TObject; MousePos: TPoint; var Handled: boolean);
    procedure btnDelKindClick(Sender: TObject);
    procedure btnDelKommClick(Sender: TObject);
    procedure btnEndeClick(Sender: TObject);
    procedure btnNewBesuchClick(Sender: TObject);
    procedure btnNewClick(Sender: TObject);
    procedure btnNewContextPopup(Sender: TObject; MousePos: TPoint; var Handled: boolean);
    procedure btnNewKindClick(Sender: TObject);
    procedure btnNewKindContextPopup(Sender: TObject; MousePos: TPoint; var Handled: boolean);
    procedure btnNewKommClick(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure cbAbgaengeContextPopup(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure cbBearbeitenChange(Sender: TObject);
    procedure cbMarkiertChange(Sender: TObject);
    procedure DatumDblClick(Sender: TObject);
    procedure DBCBGeschlechtChange(Sender: TObject);
    procedure dbcbKindKircheMouseEnter(Sender: TObject);
    procedure dbcbOnMouseEnter(Sender: TObject);
    procedure dbediEMail1DblClick(Sender: TObject);
    procedure dbediEMail2DblClick(Sender: TObject);
    procedure dbediEMailDblClick(Sender: TObject);
    procedure DatumExit(Sender: TObject);
    procedure dbediInternet1DblClick(Sender: TObject);
    procedure dbediInternet2DblClick(Sender: TObject);
    procedure dbediInternetDblClick(Sender: TObject);
    procedure DBEdiMutterNameDblClick(Sender: TObject);
    procedure DBEdiVaterNameDblClick(Sender: TObject);
    procedure DBGridOverViewCellClick(Column: TColumn);
    procedure ediSchnellSucheKeyUp(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure ediSchnellSucheDblClick(Sender: TObject);
    procedure ediSSNachnameClick(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: char);
    procedure mnuMarkThisClick(Sender: TObject);
    procedure mnuDelMarkClick(Sender: TObject);
    procedure OnFilterChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure FormMouseWheelDown(Sender: TObject; Shift: TShiftState; MousePos: TPoint; var Handled: boolean);
    procedure FormMouseWheelUp(Sender: TObject; Shift: TShiftState; MousePos: TPoint; var Handled: boolean);
    procedure FormShow(Sender: TObject);
    procedure panDatenbankOverviewMouseLeave(Sender: TObject);
    procedure panShowOverviewMouseEnter(Sender: TObject);
    procedure pcDetailsChange(Sender: TObject);
    procedure RG_SortierungClick(Sender: TObject);
    procedure scBarKarteiChange(Sender: TObject);
    procedure TimerBerechnungTimer(Sender: TObject);
    procedure tsAdrContextPopup(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
  private
    { private declarations }
    SetLastChange: boolean;
    berechnen: boolean;
    procedure SetSortPersonen(typ : integer);
  public
    { public declarations }
    function BeforePost: boolean;
    procedure AfterScroll;
    procedure SetScrollBar;
  end;

var
  frmKartei: TfrmKartei;

implementation

uses
  dm,
  help,
  main,
  DBGrid,
  DB,      //Dataset.state
  LCLType, //VK_Keys
  LCLIntf, //Openurl
  variants,//VarArrayOf
  Clipbrd,
  DateUtils,
  global;

{$R *.lfm}

{ TfrmKartei }

var
  weiter_mit_enter : boolean;

procedure TfrmKartei.SetSortPersonen(typ : integer);

begin
  case typ of
    //Umlautsortierung in Version 7.2.0.5 umgebaut
    0: begin
         // dsetPersonen.SortedFields:='Nachname,Vorname'; //Alte Version
         frmDM.ExecSQL('Update PERSONEN SET TempString='+SQL_UTF8UmlautReplace('Nachname')+'||'' ''||'+SQL_UTF8UmlautReplace('Vorname'));
         frmDM.dsetPersonen.SortedFields:='TempString';
         frmDM.dsetPersonen.Refresh;
       end;
    1: begin
         frmDM.ExecSQL('Update PERSONEN SET TempString='+SQL_UTF8UmlautReplace('Nachname')+'||'' ''||'+SQL_UTF8UmlautReplace('Vorname'));
         frmDM.dsetPersonen.SortedFields:='TempString';
         frmDM.dsetPersonen.Refresh;
         frmDM.dsetPersonen.SortedFields:='Markiert,TempString';
         frmDM.dsetPersonen.IndexFieldNames := StringReplace(frmDM.dsetPersonen.IndexFieldNames, 'Desc', 'Desc', []);
       end;
    2: begin
         //frmDM.ExecSQL('Update PERSONEN SET Tempinteger=strftime(''%j'',geburtstag)');   //j = day of year: 001-366 setzen (geht nicht bei Schaltjahren)
         frmDM.ExecSQL('Update PERSONEN SET Tempinteger=strftime(''%m'', geburtstag)*31+strftime(''%d'', geburtstag)');
         frmDM.dsetPersonen.SortedFields:='Tempinteger';
         frmDM.dsetPersonen.Refresh;
       end;
    3: begin
         //frmDM.dsetPersonen.SortedFields:='Ort,Strasse,Nachname,Vorname'; //Alte Version
         frmDM.ExecSQL('Update PERSONEN SET TempString='+
                 SQL_UTF8UmlautReplace('Ort')     +'||'' ''||'+SQL_UTF8UmlautReplace('Strasse')+'||'' ''||'+
                 SQL_UTF8UmlautReplace('Nachname')+'||'' ''||'+SQL_UTF8UmlautReplace('Vorname'));
         frmDM.dsetPersonen.SortedFields:='TempString';
         frmDM.dsetPersonen.Refresh;
       end;
  end;
  myDebugLN('dsetPERSONEN.SortedFields:    '+frmDM.dsetPersonen.SortedFields);
  myDebugLN('dsetPERSONEN.IndexFieldNames: '+frmDM.dsetPersonen.IndexFieldNames);
  frmDM.dsetPersonen.First;
end;

procedure TfrmKartei.FormShow(Sender: TObject);
begin
  //myDebugLN('frmKartei.FormShow');
  frmDM.dbStatus(true, true);
  OnFilterChange(Sender);
  berechnen := True;
  DBGridOverView.Columns.Clear;
  DBGridOverView.Columns.Add.FieldName := 'Vorname';      DBGridOverView.Columns.Items[0].Width := 115;
  DBGridOverView.Columns.Add.FieldName := 'Nachname';     DBGridOverView.Columns.Items[1].Width := 115;
  DBGridOverView.Columns.Add.FieldName := 'Ort';          DBGridOverView.Columns.Items[2].Width := 130;
  DBGridOverView.Columns.Add.FieldName := 'Strasse';      DBGridOverView.Columns.Items[3].Width := 140;
  DBGridOverView.Columns.Add.FieldName := 'Geburtstag';   DBGridOverView.Columns.Items[4].Width :=  80;
  panShowOverview.Visible              := frmMain.mnuShowOverView.Checked;
  cbGemeinde.Items.Text   := sGemeinden;
  cbGemeinde.Text         := sGemeindenAlle;
  DBCBGemeinde.Items.Text := sGemeinden;
  DBCBGemeinde.Items.Delete(0);  //sGemeindenAlle
end;

procedure TfrmKartei.panDatenbankOverviewMouseLeave(Sender: TObject);
begin
  panDatenbankOverview.SendToBack;
  panDatenbankOverview.Visible := false;
end;

procedure TfrmKartei.panShowOverviewMouseEnter(Sender: TObject);
begin
  if frmMain.mnuShowOverView.Checked
    then
      begin
        panDatenbankOverview.BringToFront;
        panDatenbankOverview.Visible := True;
        DBGridOverView.SetFocus;
      end;
end;

procedure TfrmKartei.pcDetailsChange(Sender: TObject);

begin
  case pcDetails.TabIndex of
    1: //TSEltern / Kinder
      begin
        DBGridDetails.Visible                  := True;
        DBGridDetails.Columns.Clear;
        DBGridDetails.DataSource               := frmDM.dsKinder;
        DBGridDetails.Columns.Items[0].Visible := False;                                    //KinderID
        DBGridDetails.Columns.Items[1].Width   := Round(ScaleX(145,nDefDPI)*ScaleFactor);   //Vorname
        DBGridDetails.Columns.Items[2].Width   := Round(ScaleX(145,nDefDPI)*ScaleFactor);   //Vorname2
        DBGridDetails.Columns.Items[3].Width   := Round(ScaleX(145,nDefDPI)*ScaleFactor);   //Nachname
        DBGridDetails.Columns.Items[4].Width   := Round(ScaleX(145,nDefDPI)*ScaleFactor);   //GebName
        DBGridDetails.Columns.Items[5].Width   := Round(ScaleX(80,nDefDPI)*ScaleFactor);    //Geburtstag
        DBGridDetails.Columns.Items[6].Width   := Round(ScaleX(80,nDefDPI)*ScaleFactor);    //Taufdatum
        DBGridDetails.Columns.Items[7].Width   := Round(ScaleX(80,nDefDPI)*ScaleFactor);    //Kirche
        DBGridDetails.Columns.Items[8].Visible := False;                                    //PersonenID
      end;
    4: //TSKommunionen
      begin
        DBGridDetails.Visible                  := True;
        DBGridDetails.Columns.Clear;
        DBGridDetails.DataSource               := frmDM.dsKomm;
        DBGridDetails.Columns.Items[0].Visible := False;                                    //KommID
        DBGridDetails.Columns.Items[1].Width   := Round(ScaleX(130,nDefDPI)*ScaleFactor);   //AbendmahlsDatum
        DBGridDetails.Columns.Items[2].Visible := False;                                    //PersonenID
      end;
    5: //TSBesuch
      begin
        DBGridDetails.Visible                  := True;
        DBGridDetails.Columns.Clear;
        DBGridDetails.DataSource               := frmDM.dsBesuch;
        DBGridDetails.Columns.Items[0].Visible := False;                                    //BesuchID
        DBGridDetails.Columns.Items[1].Width   := Round(ScaleX( 80,nDefDPI)*ScaleFactor);   //Datum
        DBGridDetails.Columns.Items[2].Width   := Round(ScaleX(400,nDefDPI)*ScaleFactor);   //Grund
        DBGridDetails.Columns.Items[3].Width   := Round(ScaleX(100,nDefDPI)*ScaleFactor);   //Mitbringsel
        DBGridDetails.Columns.Items[4].Width   := Round(ScaleX(150,nDefDPI)*ScaleFactor);   //Bibeltext
        DBGridDetails.Columns.Items[5].Visible := False;                                    //Memo
        DBGridDetails.Columns.Items[6].Width   := Round(ScaleX(100,nDefDPI)*ScaleFactor);   //Terminabsprache
        DBGridDetails.Columns.Items[7].Visible := False;                                    //PersonenID
      end;
    else
      begin
        DBGridDetails.Visible := False;
      end;
  end;

  case pcDetails.TabIndex of
    1, 5: pcDetails.Height := Round(ScaleY(342,nDefDPI)*ScaleFactor); //Eltern / Kinder, Besuch
    4:    pcDetails.Height := Round(ScaleY(160,nDefDPI)*ScaleFactor); //Kommunionen
    else  pcDetails.Height := Round(ScaleY(488,nDefDPI)*ScaleFactor);
  end;

  if DBGridDetails.Visible
    then
      begin
        DBGridDetails.Top    := pcDetails.Top+pcDetails.Height+4;
        DBGridDetails.Height := panNavi.Top-DBGridDetails.Top-4;
      end;

  berechnen := True;
end;

procedure TfrmKartei.RG_SortierungClick(Sender: TObject);
begin
  SetSortPersonen(RG_Sortierung.ItemIndex);
  scBarKartei.Position := 0;
end;

procedure TfrmKartei.scBarKarteiChange(Sender: TObject);
begin
  frmDM.dsetPERSONEN.RecNo := scBarKartei.Position;
end;

procedure TfrmKartei.TimerBerechnungTimer(Sender: TObject);

var
  i: integer;

begin
  if berechnen and frmDM.dsetPERSONEN.Active
    then
      begin
        berechnen := False;
        //myDebugLN('Berechnung aufgerufen');

        //Alter
        i := AgeNow(frmDM.dsetPERSONEN.FieldByName('Geburtstag').AsDateTime);
        if (i > 111) or (i = -1)
          then labAlter.Caption := '??'
          else labAlter.Caption := IntToStr(i);

        cbMarkiert.Checked := frmDM.dsetPERSONEN.FieldByName('Markiert').AsBoolean;

        case pcDetails.TabIndex of
          4: //TSKommunionen
          begin
            MemoKomm.Clear;
            frmDM.dsetHelp.SQL.Text := Format(global.sSQL_KOMM_PER_JAHR, [frmDM.dsetPERSONEN.FieldByName('PersonenID').AsString]);
            frmDM.dsetHelp.Open;
            frmDM.dsetHelp.First;
            if frmDM.dsetHelp.EOF
              then
                begin
                  MemoKomm.Text := 'Keine';
                end
              else
                begin
                  MemoKomm.Lines.BeginUpdate;
                  while not frmDM.dsetHelp.EOF do
                    begin
                      MemoKomm.Lines.Add(frmDM.dsetHelp.FieldByName('D').AsString + ': ' + frmDM.dsetHelp.FieldByName('C').AsString);
                      frmDM.dsetHelp.Next;
                    end;
                  MemoKomm.SelStart  := 0;
                  MemoKomm.SelLength := 0;
                  MemoKomm.Lines.EndUpdate;
                end;
            frmDM.dsetHelp.Close;
            frmDM.dsetKomm.First;
            btnDelKomm.Visible := frmDM.dsetKomm.FieldByName('KommID').AsString <> '';
          end;
          5: //TSBesuch
          begin
            frmDM.dsetBesuch.First;
            btnDelBesuch.Visible  := frmDM.dsetBesuch.FieldByName('BesuchID').AsString <> '';
            btnCopyBesuch.Visible := btnDelBesuch.Visible;
          end;
        end;
      end;

  if weiter_mit_enter
    then
      begin
        while not((activeControl is Tdbedit) or (activeControl is Tdbcombobox) or (activeControl is Tdbcheckbox)) do SelectNext(ActiveControl,True,True);
        weiter_mit_enter := false;
      end;
end;

procedure TfrmKartei.tsAdrContextPopup(Sender: TObject; MousePos: TPoint; var Handled: Boolean);

var sAdr : string;

begin
  sAdr := trim(frmDM.dsetPERSONEN.FieldByName('BriefAnrede').AsString +  ' ' + frmDM.dsetPERSONEN.FieldByName('Titel').AsString);
  if sAdr <> '' then sAdr := sAdr + #13#10;
  sAdr := sAdr + frmDM.dsetPERSONEN.FieldByName('Vorname').AsString + ' ' + frmDM.dsetPERSONEN.FieldByName('Nachname').AsString+#13#10;
  if frmDM.dsetPERSONEN.FieldByName('sRes1').AsString <> '' then sAdr := sAdr + frmDM.dsetPERSONEN.FieldByName('sRes1').AsString+#13#10;
  sAdr := sAdr + frmDM.dsetPERSONEN.FieldByName('Strasse').AsString+#13#10;
  sAdr := sAdr + trim(frmDM.dsetPERSONEN.FieldByName('Land').AsString + ' ' + frmDM.dsetPERSONEN.FieldByName('PLZ').AsString + ' ' + frmDM.dsetPERSONEN.FieldByName('Ort').AsString);
  Clipboard.AsText := sAdr;
  Handled := true;
end;

procedure TfrmKartei.AfterScroll;

begin
  if frmKartei.Visible
    then
      begin        
        berechnen := True;
        scBarKartei.position := frmDM.dsetPERSONEN.RecNo;
      end;
end;

function TfrmKartei.BeforePost: boolean;

  //Gibt true zurück wenn gespeichert werden soll

begin
  Result := True;
  if frmKartei.Visible
    then
      begin
        if SetLastChange
          then
            begin
              frmDM.dsetPERSONEN.FieldByName('LastChange').AsDateTime := date;
              Result := MessageDlg('Sollen die Daten gespeichert werden?',  mtConfirmation, [mbYes, mbNo], 0) = mrYes;
            end;
        SetLastChange := True;
        //Prüfen auf Geburtstag < Taufe < Konfirmation
        if (frmDM.dsetPERSONEN.FieldByName('Geburtstag').AsString <> '') and
           (frmDM.dsetPERSONEN.FieldByName('TaufDatum').AsString <> '') and
           (frmDM.dsetPERSONEN.FieldByName('TaufDatum').AsDateTime < frmDM.dsetPERSONEN.FieldByName('Geburtstag').AsDateTime)
          then ShowMessage('Achtung: TaufDatum ist kleiner als Geburtstag!');
        if (frmDM.dsetPERSONEN.FieldByName('KonfDatum').AsString <> '') and
           (frmDM.dsetPERSONEN.FieldByName('TaufDatum').AsString <> '') and
           (frmDM.dsetPERSONEN.FieldByName('KonfDatum').AsDateTime < frmDM.dsetPERSONEN.FieldByName('TaufDatum').AsDateTime)
          then ShowMessage('Achtung: KonfDatum ist kleiner als TaufDatum!');
      end;
end;

procedure TfrmKartei.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  frmDM.dbStatus(false); // DB schliessen
  pcDetails.ActivePage := TSAllgemein; //Für den nächsten Aufruf
  cbBearbeiten.Checked := False;
end;

procedure TfrmKartei.btnNewKommClick(Sender: TObject);

var
  dtValue: TDateTime;

begin
  if TryStrToDate(DateEditKomm.Text, dtValue) then
    begin

      frmDM.ExecSQL(Format(global.sSQL_KOMM_ADD, [FormatDateTime('yyyy-mm-dd', DateEditKomm.Date), frmDM.dsetPERSONEN.FieldByName('PersonenID').AsString]));
      frmDM.dsetKomm.refresh;
      berechnen := True;
    end
  else
    begin
      ShowMessage('Ungültiges Datum');
      DateEditKomm.Date := now();
    end;
end;

procedure TfrmKartei.btnPrintClick(Sender: TObject);

begin
  //DB öffnen
  frmDM.dsetHelp.SQL.Text:='select * from personen where PersonenID='+frmDM.dsetPersonen.FieldByName('PersonenID').AsString;
  frmDM.dsetHelp.Open;

  frmDM.dsetHelp1.IndexFieldNames := 'Geburtstag Asc';
  frmDM.dsetHelp1.LinkedFields    := 'PersonenID';
  frmDM.dsetHelp1.MasterFields    := 'PersonenID';
  frmDM.dsetHelp1.MasterSource    := frmDM.dsHelp;
  frmDM.dsetHelp1.SortedFields    := 'Geburtstag';
  frmDM.dsetHelp1.SQL.Text        := 'select * from Kinder';
  frmDM.dsetHelp1.Open;

  frmMain.frReport.LoadFromFile(sReportFile);
  frmMain.frReport.ShowReport;
  frmMain.frReport.Clear;
  //DB wieder schliessen
  frmDM.dsetHelp.Close;
  frmDM.dsetHelp1.Close;

  //aufräumen
  frmDM.dsetHelp1.IndexFieldNames := '';
  frmDM.dsetHelp1.LinkedFields    := '';
  frmDM.dsetHelp1.MasterFields    := '';
  frmDM.dsetHelp1.SortedFields    := '';
  frmDM.dsetHelp1.SQL.Text        := '';
end;

procedure TfrmKartei.cbAbgaengeContextPopup(Sender: TObject; MousePos: TPoint; var Handled: Boolean);

var
  sWhere : String;

begin
  sWhere := '';
  sWhere := SQL_Where_Add_OR(sWhere, 'Ueberwiesen_nach_Datum<>''''');
  sWhere := SQL_Where_Add_OR(sWhere, 'AustrittsDatum<>''''');
  sWhere := SQL_Where_Add_OR(sWhere, 'AusschlussDatum<>''''');
  sWhere := SQL_Where_Add_OR(sWhere, 'TodesDatum<>''''');
  sWhere := SQL_Where_Add_OR(sWhere, 'UebertrittsAbDatum<>''''');
  ExecSQL('Update Personen set Abgang=''0''', frmDM.dsetHelp, false); // Zuerst alles löschen
  ExecSQL('Update Personen set Abgang=''Y'' where '+sWhere, frmDM.dsetHelp, true);
  SetScrollbar;
end;

procedure TfrmKartei.cbBearbeitenChange(Sender: TObject);
begin
  frmDM.AutoEdit(cbBearbeiten.Checked);
  btnNew.Enabled        := cbBearbeiten.Checked;
  btnNewKind.Enabled    := cbBearbeiten.Checked;
  btnNewKomm.Enabled    := cbBearbeiten.Checked;
  btnNewBesuch.Enabled  := cbBearbeiten.Checked;
  btnDel.Enabled        := cbBearbeiten.Checked;
  btnDelKind.Enabled    := cbBearbeiten.Checked;
  btnDelKomm.Enabled    := cbBearbeiten.Checked;
  btnDelBesuch.Enabled  := cbBearbeiten.Checked;
  btnCopyBesuch.Enabled := cbBearbeiten.Checked;
  if cbBearbeiten.Checked and frmMain.mnuColoredEditMode.Checked
    then frmKartei.Color := clInactiveCaption
    else frmKartei.Color := clDefault;
end;

procedure TfrmKartei.cbMarkiertChange(Sender: TObject);
begin
  SetLastChange := False;
  if not (frmDM.dsetPERSONEN.State in [dsEdit, dsInsert]) then frmDM.dsetPERSONEN.Edit;
  frmDM.dsetPERSONEN.FieldByName('Markiert').AsBoolean := cbMarkiert.Checked;
  frmDM.dsetPERSONEN.Post;
  frmDM.dsetPERSONEN.Refresh;
  SetLastChange := true;
end;

procedure TfrmKartei.btnDelKommClick(Sender: TObject);
begin
  if MessageDlg('Löschen?', 'Soll das Datum ' + frmDM.dsetKomm.FieldByName('Abendmahlsdatum').AsString +
                            ' gelöscht werden', mtConfirmation, [mbYes, mbNo], 0) = mrYes
    then
      begin
        frmDM.ExecSQL(Format(sSQL_Komm_DEL, [frmDM.dsetPERSONEN.FieldByName('PersonenID').AsString, frmDM.dsetKomm.FieldByName('KommID').AsString]));
        frmDM.dsetKomm.Refresh;
        berechnen := True;
      end;
end;

procedure TfrmKartei.btnEndeClick(Sender: TObject);
begin
  if frmDM.dsetPERSONEN.State in [dsEdit, dsInsert]
    then frmDM.dsetPERSONEN.Post; {speicheret die Änderungen}
  Close;
end;

procedure TfrmKartei.btnNewBesuchClick(Sender: TObject);
begin
  frmDM.ExecSQL(Format(global.sSQL_BESUCH_ADD, [FormatDateTime('yyyy-mm-dd', now()), frmDM.dsetPERSONEN.FieldByName('PersonenID').AsString]));
  frmDM.dsetBesuch.refresh;
  berechnen := True;
end;

procedure TfrmKartei.btnNewClick(Sender: TObject);

var
  wHelp: word;

begin
  frmDM.ExecSQL(Format(sSQL_Pers_New, [sDefaultGemeinde]));
  frmDM.dsetPERSONEN.Refresh;
  // auf dS gehen (Suchfunktion wird genutzt)
  ediSSNachname.Text := '?';
  wHelp := 0;
  ediSchnellSucheKeyUp(Sender, wHelp, []);
  SetScrollbar;
  ShowMessage('Person angelegt');
end;

procedure TfrmKartei.btnNewContextPopup(Sender: TObject; MousePos: TPoint; var Handled: boolean);

var
  wHelp: word;
  sHelp: string;

begin
  //Neue Person mit Übernahme
  sHelp := frmDM.dsetPERSONEN.FieldByName('Nachname').AsString;
  frmDM.ExecSQL('insert into ' + global.sPersTablename +
    ' (Vorname, Nachname, Strasse, Land, PLZ, Ort, Ortsteil, TelPrivat, Kirche, Gemeinde, Markiert, Abgang) VALUES ('+
    QuotedStr('???') + ', ' +
    QuotedStr(sHelp) + ', ' +
    QuotedStr(frmDM.dsetPERSONEN.FieldByName('Strasse').AsString) + ', ' +
    QuotedStr(frmDM.dsetPERSONEN.FieldByName('Land').AsString) + ', ' +
    QuotedStr(frmDM.dsetPERSONEN.FieldByName('PLZ').AsString) + ', ' +
    QuotedStr(frmDM.dsetPERSONEN.FieldByName('Ort').AsString) + ', ' +
    QuotedStr(frmDM.dsetPERSONEN.FieldByName('Ortsteil').AsString) + ', ' +
    QuotedStr(frmDM.dsetPERSONEN.FieldByName('TelPrivat').AsString) + ', ' +
    QuotedStr(frmDM.dsetPERSONEN.FieldByName('Kirche').AsString) + ', ' +
    QuotedStr(frmDM.dsetPERSONEN.FieldByName('Gemeinde').AsString) + ',0,0)');
  frmDM.dsetPERSONEN.Refresh;
  // auf dS gehen (Suchfunktion wird genutzt)
  ediSSNachname.Text := sHelp;
  ediSSVorname.Text := '???';
  wHelp := 0;
  ediSchnellSucheKeyUp(Sender, wHelp, []);
  SetScrollbar;
  ShowMessage('Person angelegt');
end;

procedure TfrmKartei.btnNewKindClick(Sender: TObject);
begin
  frmDM.ExecSQL(Format(global.sSQL_KIND_ADD, [frmDM.dsetPERSONEN.FieldByName('PersonenID').AsString]));
  frmDM.dsetKinder.refresh;
  berechnen := True;
end;

procedure TfrmKartei.btnNewKindContextPopup(Sender: TObject; MousePos: TPoint; var Handled: boolean);

var
  gebjahr: integer;
  sSQL: string;

begin
  try
    gebjahr := StrToInt(year(frmDM.dsetPERSONEN.FieldByName('Geburtstag').AsString));
  except
    on E: Exception do
      begin
        screen.cursor := crdefault;
        LogAndShowError('Bei dieser Person ist kein oder ein ungültiges Geburtsdatum eingegeben!');
        Exit; //Nicht weitermachen
      end;
  end;
  frmDBGrid.DBGrid.DataSource := frmDM.dsHelp;
  frmDM.dsetHelp.SQL.Text := format(sSQL_Auswahlliste, [IntToStr(gebjahr + 15), IntToStr(gebjahr + 60), '']);
  frmDM.dsetHelp.Open;
  frmDBGrid.DBGrid.Columns.Items[0].Visible := False; //PersonenID
  frmDBGrid.DBGrid.Columns.Items[1].Width := 150;     //Vorname
  frmDBGrid.DBGrid.Columns.Items[2].Width := 150;     //Vorname2
  frmDBGrid.DBGrid.Columns.Items[3].Width := 150;     //Nachname
  frmDBGrid.DBGrid.Columns.Items[4].Width := 150;     //GebName
  frmDBGrid.DBGrid.Columns.Items[5].Width := 80;      //Gebtag
  frmDBGrid.DBGrid.Columns.Items[6].Width := 80;      //Taufdatum
  frmDBGrid.DBGrid.Columns.Items[7].Width := 80;      //Kirche
  if frmDBGrid.ShowModal = mrOk
    then
      begin
        sSQL := 'insert into ' + global.sKindTablename +
          ' (Vorname, Vorname2, Nachname, GebName, Geburtstag, Taufdatum, Kirche, PersonenID) VALUES (' +
          QuotedStr(frmDM.dsetHelp.FieldByName('Vorname').AsString) + ', ' +
          QuotedStr(frmDM.dsetHelp.FieldByName('Vorname2').AsString) + ', ' +
          QuotedStr(frmDM.dsetHelp.FieldByName('Nachname').AsString) + ', ' +
          QuotedStr(frmDM.dsetHelp.FieldByName('Gebname').AsString) + ', ' +
          '''' + FormatDateTime('yyyy-mm-dd', frmDM.dsetHelp.FieldByName('Geburtstag').AsDateTime) + ''', ' +
          '''' + FormatDateTime('yyyy-mm-dd', frmDM.dsetHelp.FieldByName('Taufdatum').AsDateTime) + ''', ' +
          QuotedStr(frmDM.dsetHelp.FieldByName('Kirche').AsString) + ', ' +
          frmDM.dsetPERSONEN.FieldByName('PersonenID').AsString + ')';

        frmDM.dsetHelp.Close;
        //Daten sind aus dsetHelp gelesen, schließen, dann kann damit das insert Kommando genutzt werden
        frmDM.ExecSQL(sSQL);
        frmDM.dsetKinder.Refresh;
      end
    else
      frmDM.dsetHelp.Close;
end;

procedure TfrmKartei.btnDelBesuchClick(Sender: TObject);
begin
  if MessageDlg('Löschen?', 'Soll der Besuch vom ' + frmDM.dsetBesuch.FieldByName('Besuchsdatum').AsString + ' gelöscht werden', mtConfirmation, [mbYes, mbNo], 0) = mrYes
    then
      begin
        frmDM.ExecSQL(Format(sSQL_Besuch_DEL, [frmDM.dsetPERSONEN.FieldByName('PersonenID').AsString, frmDM.dsetBesuch.FieldByName('BesuchID').AsString]));
        frmDM.dsetBesuch.Refresh;
        berechnen := True;
      end;
end;

procedure TfrmKartei.btnCopyBesuchClick(Sender: TObject);

var
  sSQL,
  sHelp: string;

begin
  frmDBGrid.DBGrid.DataSource := frmDM.dsHelp;
  frmDM.dsetHelp.SQL.Text := 'select PersonenID, Vorname, Nachname, Strasse, Geburtstag from Personen order by nachname, strasse, vorname';
  frmDM.dsetHelp.Open;
  frmDBGrid.DBGrid.Columns.Items[0].Visible := False; //PersonenID
  frmDBGrid.DBGrid.Columns.Items[1].Width := 160;     //Vorname
  frmDBGrid.DBGrid.Columns.Items[2].Width := 160;     //Nachname
  frmDBGrid.DBGrid.Columns.Items[3].Width := 160;     //Strasse
  frmDBGrid.DBGrid.Columns.Items[4].Width := 100;     //Gebtag
  if frmDBGrid.ShowModal = mrOk
    then
      begin
        sHelp := 'Besuch wurde zu ' + frmDM.dsetHelp.FieldByName('Vorname').AsString + ' ' + frmDM.dsetHelp.FieldByName('Nachname').AsString +' kopiert';
        sSQL  := 'insert into ' + global.sBesuchTablename +
          ' (PersonenID, BesuchsDatum, BesuchsGrund, Mitbringsel, Bibeltext, Memo, Terminabsprache) VALUES ('
          + frmDM.dsetHelp.FieldByName('PersonenID').AsString + ', ' +
          '''' + FormatDateTime('yyyy-mm-dd', frmDM.dsetBesuch.FieldByName('BesuchsDatum').AsDateTime) + ''', '+
          '''' + frmDM.dsetBesuch.FieldByName('BesuchsGrund').AsString + ''', ' +
          '''' + frmDM.dsetBesuch.FieldByName('Mitbringsel').AsString + ''', ' +
          '''' + frmDM.dsetBesuch.FieldByName('Bibeltext').AsString + ''', ' +
          '''' + frmDM.dsetBesuch.FieldByName('Memo').AsString + ''', ' +
          '''' + frmDM.dsetBesuch.FieldByName('Terminabsprache').AsString + ''')';

        frmDM.dsetHelp.Close;
        //Daten sind aus dsetHelp gelesen, schließen, dann kann damit das insert Kommando genutzt werden
        try
          frmDM.ExecSQL(sSQL);
          frmDM.dsetBesuch.Refresh;
          Showmessage(sHelp);
        except
          on E: Exception do LogAndShowError(E.Message);
        end;
      end
    else
      frmDM.dsetHelp.Close;
end;

procedure TfrmKartei.btnDelClick(Sender: TObject);
begin
  ShowMessage('Sicherheitsfunktion: Zum Löschen machen Sie bitte einen Rechtsklick');
end;

procedure TfrmKartei.btnDelContextPopup(Sender: TObject; MousePos: TPoint; var Handled: boolean);

begin
  if MessageDlg('Löschen?', 'Soll die Person "' +
    frmDM.dsetPERSONEN.FieldByName('Vorname').AsString + ' ' +
    frmDM.dsetPERSONEN.FieldByName('Nachname').AsString +
    ' " gelöscht werden?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    frmDM.ExecSQL(Format(sSQL_Komm_Pes_DEL, [frmDM.dsetPERSONEN.FieldByName('PersonenID').AsString]));
    frmDM.ExecSQL(Format(sSQL_Besuch_Pes_DEL, [frmDM.dsetPERSONEN.FieldByName('PersonenID').AsString]));
    frmDM.ExecSQL(Format(sSQL_Kind_Pes_DEL, [frmDM.dsetPERSONEN.FieldByName('PersonenID').AsString]));
    frmDM.ExecSQL(Format(sSQL_Pers_DEL, [frmDM.dsetPERSONEN.FieldByName('PersonenID').AsString]));
    frmDM.dsetPERSONEN.Refresh;
    berechnen := True;
    SetScrollbar;
  end;
end;

procedure TfrmKartei.btnDelKindClick(Sender: TObject);
begin
  if MessageDlg('Löschen?', 'Soll das Kind "' +
    frmDM.dsetKinder.FieldByName('Vorname').AsString + ' ' +
    frmDM.dsetKinder.FieldByName('Nachname').AsString + ' ' +
    '" gelöscht werden?', mtConfirmation, [mbYes, mbNo], 0) = mrYes
    then
      begin
        frmDM.ExecSQL(Format(sSQL_KIND_DEL, [frmDM.dsetPERSONEN.FieldByName('PersonenID').AsString, frmDM.dsetKinder.FieldByName('KinderID').AsString]));
        frmDM.dsetKinder.Refresh;
        berechnen := True;
      end;
end;

procedure TfrmKartei.DatumDblClick(Sender: TObject);
begin
  if MessageDlg('Löschen?', 'Soll das Datum gelöscht werden', mtConfirmation, [mbYes, mbNo], 0) = mrYes
    then
      begin
        myDebugLN('Loesche Datum fuer Feld: ' + TDBEdit(Sender).datafield);
        frmDM.dsetPERSONEN.Edit;
        frmDM.dsetPERSONEN.FieldByName(TDBEdit(Sender).datafield).AsString := '';
        frmDM.dsetPERSONEN.Post;
        berechnen := True;
      end;
end;

procedure TfrmKartei.DBCBGeschlechtChange(Sender: TObject);
begin
  if not ((dbcbGeschlecht.Text = 'M') or (dbcbGeschlecht.Text = 'W') or (dbcbGeschlecht.Text = 'D') or (dbcbGeschlecht.Text = ' '))
    then
      begin
        Messagedlg('Nur "W", "M", "D" oder " " erlaubt', mtInformation, [mbOK], 0);
        dbcbGeschlecht.SetFocus;
      end;
end;

procedure TfrmKartei.dbcbKindKircheMouseEnter(Sender: TObject);
begin
  ///??? Workaround für Bug in Lazarus 1.8.0 CB geht nicht richtig in EDIT Mode
  if cbBearbeiten.Checked then frmDM.dsKinder.Edit;
end;

procedure TfrmKartei.dbcbOnMouseEnter(Sender: TObject);
begin
  ///??? Workaround für Bug in Lazarus 1.8.0 CB geht nicht richtig in EDIT Mode
  if cbBearbeiten.Checked then frmDM.dsetPERSONEN.Edit;
end;

procedure TfrmKartei.dbediEMail1DblClick(Sender: TObject);
begin
  Openurl('MailTo:' + frmDM.dsetPERSONEN.FieldByName('eMail2').AsString);
end;

procedure TfrmKartei.dbediEMail2DblClick(Sender: TObject);
begin
  Openurl('MailTo:' + frmDM.dsetPERSONEN.FieldByName('eMail3').AsString);
end;

procedure TfrmKartei.dbediEMailDblClick(Sender: TObject);
begin
  Openurl('MailTo:' + frmDM.dsetPERSONEN.FieldByName('eMail').AsString);
end;

procedure TfrmKartei.DatumExit(Sender: TObject);
begin
  if frmDM.dsetPERSONEN.State in [dsEdit, dsInsert]
    then
      begin
        if frmDM.dsetPERSONEN.FieldByName(TDBEdit(Sender).datafield).AsDateTime > now
          then frmDM.dsetPERSONEN.FieldByName(TDBEdit(Sender).datafield).AsDateTime := IncYear(frmDM.dsetPERSONEN.FieldByName(TDBEdit(Sender).datafield).AsDateTime, -100);
      end;
end;

procedure TfrmKartei.dbediInternet1DblClick(Sender: TObject);
begin
  Openurl(frmDM.dsetPERSONEN.FieldByName('Internet2').AsString);
end;

procedure TfrmKartei.dbediInternet2DblClick(Sender: TObject);
begin
  Openurl(frmDM.dsetPERSONEN.FieldByName('Internet3').AsString);
end;

procedure TfrmKartei.dbediInternetDblClick(Sender: TObject);
begin
  Openurl(frmDM.dsetPERSONEN.FieldByName('Internet').AsString);
end;

procedure TfrmKartei.DBEdiMutterNameDblClick(Sender: TObject);

var
  gebjahr: integer;
  sSQL: string;
  nHelp: longint;

begin
  try
    gebjahr := StrToInt(year(frmDM.dsetPERSONEN.FieldByName('Geburtstag').AsString));
  except
    on E: Exception do
    begin
      screen.cursor := crdefault;
      LogAndShowError('Bei dieser Person ist kein oder ein ungültiges Geburtsdatum eingegeben!');
      Exit; //Nicht weitermachen
    end;
  end;
  frmDBGrid.DBGrid.DataSource := frmDM.dsHelp;
  frmDM.dsetHelp.SQL.Text := format(sSQL_Auswahlliste,
    [IntToStr(gebjahr - 60), IntToStr(gebjahr - 15), 'and geschlecht =''W''']);
  frmDM.dsetHelp.Open;
  frmDBGrid.DBGrid.Columns.Items[0].Visible := False; //PersonenID
  frmDBGrid.DBGrid.Columns.Items[1].Width := 160;     //Vorname
  frmDBGrid.DBGrid.Columns.Items[2].Width := 160;     //Vorname2
  frmDBGrid.DBGrid.Columns.Items[3].Width := 160;     //Nachname
  frmDBGrid.DBGrid.Columns.Items[4].Width := 160;     //GebName
  frmDBGrid.DBGrid.Columns.Items[5].Width := 100;     //Gebtag
  frmDBGrid.DBGrid.Columns.Items[6].Width := 100;     //Kirche
  if frmDBGrid.ShowModal = mrOk
    then
      begin
        sSQL := 'update ' + global.sPersTablename + ' set MutterName=''' +
          frmDM.dsetHelp.FieldByName('Vorname').AsString + ' ' +
          frmDM.dsetHelp.FieldByName('Vorname2').AsString + ' ' +
          frmDM.dsetHelp.FieldByName('Nachname').AsString + ''',' +
          'MutterGebName=''' + frmDM.dsetHelp.FieldByName('Gebname').AsString +
          ''',' + 'MutterKirche=''' + frmDM.dsetHelp.FieldByName('Kirche').AsString +
          '''' + ' where PersonenID=' + frmDM.dsetPERSONEN.FieldByName(
          'PersonenID').AsString;

        frmDM.dsetHelp.Close;
        frmDM.ExecSQL(sSQL);
        nHelp := frmDM.dsetPERSONEN.RecNo;
        frmDM.dsetPERSONEN.Refresh;
        frmDM.dsetPERSONEN.RecNo := nHelp;
      end
  else
    frmDM.dsetHelp.Close;
end;

procedure TfrmKartei.DBEdiVaterNameDblClick(Sender: TObject);

var
  gebjahr: integer;
  sSQL: string;
  nHelp: longint;

begin
  try
    gebjahr := StrToInt(year(frmDM.dsetPERSONEN.FieldByName('Geburtstag').AsString));
  except
    on E: Exception do
    begin
      screen.cursor := crdefault;
      LogAndShowError('Bei dieser Person ist kein oder ein ungültiges Geburtsdatum eingegeben!');
      Exit; //Nicht weitermachen
    end;
  end;
  frmDBGrid.DBGrid.DataSource := frmDM.dsHelp;
  frmDM.dsetHelp.SQL.Text := format(sSQL_Auswahlliste,
    [IntToStr(gebjahr - 60), IntToStr(gebjahr - 15), 'and geschlecht =''M''']);
  frmDM.dsetHelp.Open;
  frmDBGrid.DBGrid.Columns.Items[0].Visible := False; //PersonenID
  frmDBGrid.DBGrid.Columns.Items[1].Width := 160;     //Vorname
  frmDBGrid.DBGrid.Columns.Items[2].Width := 160;     //Vorname2
  frmDBGrid.DBGrid.Columns.Items[3].Width := 160;     //Nachname
  frmDBGrid.DBGrid.Columns.Items[4].Width := 160;     //GebName
  frmDBGrid.DBGrid.Columns.Items[5].Width := 100;     //Gebtag
  frmDBGrid.DBGrid.Columns.Items[6].Width := 100;     //Kirche
  if frmDBGrid.ShowModal = mrOk
    then
      begin
        sSQL := 'update ' + global.sPersTablename + ' set VaterName=''' +
          frmDM.dsetHelp.FieldByName('Vorname').AsString + ' ' +
          frmDM.dsetHelp.FieldByName('Vorname2').AsString + ' ' +
          frmDM.dsetHelp.FieldByName('Nachname').AsString + ''',' +
          'VaterGebName=''' + frmDM.dsetHelp.FieldByName('Gebname').AsString +
          ''',' + 'VaterKirche=''' + frmDM.dsetHelp.FieldByName('Kirche').AsString +
          '''' + ' where PersonenID=' + frmDM.dsetPERSONEN.FieldByName(
          'PersonenID').AsString;

        frmDM.dsetHelp.Close;
        frmDM.ExecSQL(sSQL);
        nHelp := frmDM.dsetPERSONEN.RecNo;
        frmDM.dsetPERSONEN.Refresh;
        frmDM.dsetPERSONEN.RecNo := nHelp;
      end
  else
    frmDM.dsetHelp.Close;
end;

procedure TfrmKartei.DBGridOverViewCellClick(Column: TColumn);
begin
  panDatenbankOverview.SendToBack;
  panDatenbankOverview.Visible := False;
end;

procedure TfrmKartei.ediSchnellSucheDblClick(Sender: TObject);
begin
  ediSSNachname.Text := '';
  ediSSVorname.Text := '';
end;

procedure TfrmKartei.ediSSNachnameClick(Sender: TObject);
begin
  ediSSVorname.Text:='';
  ediSSNachName.SelectAll;
end;

procedure TfrmKartei.FormKeyPress(Sender: TObject; var Key: char);
begin
  if (key = #13) and ((activeControl is Tdbedit) or (activeControl is Tdbcombobox) or (activeControl is Tdbcheckbox))
    then
      begin
        key := #0;
        SelectNext(activeControl, True, True);
        Weiter_mit_enter := true;
      end;
end;

procedure TfrmKartei.mnuMarkThisClick(Sender: TObject);
begin
  frmDM.ExecSQL(global.sSQL_DelMark);
  cbMarkiert.Checked:=true;
  cbMarkiertChange(sender);
  frmDM.dsetPERSONEN.Refresh;
end;

procedure TfrmKartei.mnuDelMarkClick(Sender: TObject);
begin
  frmDM.ExecSQL(global.sSQL_DelMark);
  frmDM.dsetPERSONEN.Refresh;
end;

procedure TfrmKartei.ediSchnellSucheKeyUp(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  if ediSSNachname.Text <> ''
    then
      frmDM.dsetPERSONEN.Locate('Nachname, Vorname', VarArrayOf([ediSSNachname.Text, ediSSVorname.Text]), [loCaseInsensitive, loPartialKey])
    else
      frmDM.dsetPERSONEN.Locate('Vorname', VarArrayOf([ediSSVorname.Text]), [loCaseInsensitive, loPartialKey]);
end;

procedure TfrmKartei.SetScrollBar;

var
  nAnzahl: longint;
  nPos: longint;
  sAnzahl: string;

begin
  //Anzahl ermitteln
  nAnzahl := frmDM.dsetPERSONEN.RecordCount;
  sAnzahl := IntToStr(nAnzahl);
  nPos := frmDM.dsetPERSONEN.RecNo;

  //anzeigen
  myDebugLN('RecordCount: ' + sAnzahl + ', RecPos: ' + IntToStr(nPos));
  labRecCnt.Caption := sAnzahl;
  scBarKartei.Hint := 'Anzahl = ' + sAnzahl;

  //Scrollbar setzen
  if nAnzahl = 0 then
    begin
      scBarKartei.Min := 0;
      scBarKartei.position := 0;
    end
  else
    begin
      scBarKartei.Min := 1;
      if (nPos > scBarKartei.Max) then scBarKartei.position := 1; //erst einmal auf eine sichere Position
      scBarKartei.Max := nAnzahl;
      scBarKartei.position := nPos; //Jetzt auf die richtige
    end;
end;

procedure TfrmKartei.OnFilterChange(Sender: TObject);

var
  sFilter: string;

begin
  sFilter := '';
  if cbGemeinde.Text = ''
    then sFilter := SQL_Where_Add(sFilter, SQL_Where_IsNull('Gemeinde'))
    else if cbGemeinde.Text <> sGemeindenAlle
      then sFilter := SQL_Where_Add(sFilter, 'Gemeinde=''' + cbGemeinde.Text + '''');
  if not cbAbgaenge.Checked then sFilter := SQL_Where_Add(sFilter, 'Abgang=''0''');
  if cbFilterMark.Checked   then sFilter := SQL_Where_Add(sFilter, 'Markiert=''1''');
  myDebugLN('Setze Filter: ' + sFilter);
  frmDM.dsetPERSONEN.Filter   := sFilter;
  frmDM.dsetPERSONEN.Filtered := (sFilter <> '');
  frmDM.dsetPERSONEN.Refresh;
  frmDM.dsetPERSONEN.First;

  SetScrollBar;
end;

procedure TfrmKartei.FormCreate(Sender: TObject);

begin
  pcDetails.ActivePage          := TSAllgemein;
  DateEditKomm.Date             := now();
  berechnen                     := False;
  SetLastChange                 := True;
  dbcbKirche.Items.Text         := sKirchenEintraege;
  dbcbKindKirche.Items.Text     := sKirchenEintraege;
  DBcbVaterKirche.Items.Text    := sKirchenEintraege;
  DBcbMutterKirche.Items.Text   := sKirchenEintraege;
  dbcbEhegatteKirche.Items.Text := sKirchenEintraege;
  dbcbUebertrittAus.Items.Text  := sKirchenEintraege;
  dbcbUebertrittNach.Items.Text := sKirchenEintraege;
  DBCBAnrede.Items.Text         := sAnredenEintraege;
  labFFeld1.Caption             := help.ReadIniVal(sIniFile, 'Defaults','labFFeld1', 'Freies Feld 1', true);
  labFFeld2.Caption             := help.ReadIniVal(sIniFile, 'Defaults','labFFeld2', 'Freies Feld 2', true);
  labFFeld3.Caption             := help.ReadIniVal(sIniFile, 'Defaults','labFFeld3', 'Freies Feld 3', true);
  labFFeld4.Caption             := help.ReadIniVal(sIniFile, 'Defaults','labFFeld4', 'Freies Feld 4', true);
end;

procedure TfrmKartei.FormKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);

{DB 08.06.96
 ausserdem in Form frmKartei Eigenschaft KeyPreview auf true gesetzt
 Blättern mit Bild-Up und Bild-down in Karteikarte
 Beschreibung unter virtuell Tastencodes in der WIN API}

begin
  if (Shift = [])
    then
      begin
        case key of
          VK_Prior: begin
                      frmDM.dsetPERSONEN.Prior;
                      ediSchnellSucheDblClick(sender);
                    end;
          VK_Next:  begin
                      frmDM.dsetPERSONEN.Next;
                      ediSchnellSucheDblClick(sender);
                    end;
          VK_F1:    ShowMessage('Hilfe' + #13#10 + global.sFehltNoch);
        end;
      end
    else
      begin
        if (Shift = [ssCtrl])
          then
            begin
              case key of
                VK_End  : frmDM.dsetPERSONEN.Last;
                VK_Home : frmDM.dsetPERSONEN.First;
                Ord('F'): ediSSNachname.SetFocus;
                Ord('P'): btnPrintClick(self);
                Ord('S'): if frmDM.dsetPERSONEN.State in [dsEdit, dsInsert] then frmDM.dsetPERSONEN.Post;
              end;
            end;
      end;
end;

procedure TfrmKartei.FormMouseWheelDown(Sender: TObject; Shift: TShiftState; MousePos: TPoint; var Handled: boolean);
begin
  frmDM.dsetPERSONEN.Next;
  //if scBarKartei.position < scBarKartei.max then scBarKartei.position := scBarKartei.position + 1;
end;

procedure TfrmKartei.FormMouseWheelUp(Sender: TObject; Shift: TShiftState; MousePos: TPoint; var Handled: boolean);
begin
  frmDM.dsetPERSONEN.prior;
  //if scBarKartei.position > 1 then scBarKartei.position := scBarKartei.position - 1;
end;

end.

