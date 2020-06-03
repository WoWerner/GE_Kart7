unit suche;

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
  StdCtrls,
  ComCtrls, DbCtrls;

type

  { TfrmSuche }

  TfrmSuche = class(TForm)
    btnStarteSuche: TButton;
    btnCancel: TButton;
    btnReset: TButton;
    cbDelMark: TCheckBox;
    cbAbgaenge: TCheckBox;
    cbNotGem: TCheckBox;
    cbNotPLZ: TCheckBox;
    cbNotKirche: TCheckBox;
    cbKomm: TCheckBox;
    CBFamstand: TComboBox;
    CBGeschlecht: TComboBox;
    cbGemeinde: TComboBox;
    ediGebBis: TEdit;
    ediKommJ: TEdit;
    ediTaufM: TEdit;
    ediPLZ: TEdit;
    ediName: TEdit;
    ediKirche: TEdit;
    ediStrasse: TEdit;
    ediLand: TEdit;
    ediOrt: TEdit;
    ediTaufJ: TEdit;
    ediKonfJ: TEdit;
    ediGebVon: TEdit;
    ediGebM: TEdit;
    ediVers: TEdit;
    ediTrauJ: TEdit;
    ediVorName: TEdit;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    GroupBox4: TGroupBox;
    GroupBox5: TGroupBox;
    Label1: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    labAlter: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label63: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    StatusBar: TStatusBar;
    procedure btnResetClick(Sender: TObject);
    procedure btnStarteSucheClick(Sender: TObject);
    procedure ediGebChange(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: char);
  private
    { private declarations }
  public
    { public declarations }
  end; 

var
  frmSuche: TfrmSuche;

implementation

uses
  global,
  help,
  dm;

const sNotUsedYear = '1900';

{$R *.lfm}

{ TfrmSuche }

procedure TfrmSuche.btnStarteSucheClick(Sender: TObject);

var
  gut           : boolean;
  sWhere        : String;
  summ_gesamt   : integer;

begin
  screen.cursor := crhourglass;
  summ_gesamt   := 0;
  sWhere        := '';

  if cbDelMark.Checked
    then
      begin
        StatusBar.SimpleText:='Lösche alte Markierungen';
        frmDM.ExecSQL(global.sSQL_DelMark);
      end;

  frmDM.dsetHelp.SQL.Text:='select * from '+global.sPersTablename+' where';
  if cbAbgaenge.Checked
    then
      begin
        sWhere := SQL_Where_Add(sWhere, SQL_Where_IsNull('Ueberwiesen_nach_Datum'));
        sWhere := SQL_Where_Add(sWhere, SQL_Where_IsNull('AustrittsDatum'));
        sWhere := SQL_Where_Add(sWhere, SQL_Where_IsNull('AusschlussDatum'));
        sWhere := SQL_Where_Add(sWhere, SQL_Where_IsNull('TodesDatum'));
        sWhere := SQL_Where_Add(sWhere, SQL_Where_IsNull('UebertrittsAbDatum'));
      end;

  //NachName
  if ediName.text    <> '' then sWhere := SQL_Where_Add(sWhere,SQL_Where_Contains(ediName.text,    'Nachname'));
  //Strasse
  if ediStrasse.text <> '' then sWhere := SQL_Where_Add(sWhere,SQL_Where_Contains(ediStrasse.text, 'Strasse'));
  //Land
  if ediLand.text    <> '' then sWhere := SQL_Where_Add(sWhere,SQL_Where_Contains(ediLand.text,    'Land'));
  //PLZ
  if ediPLZ.text <> '00000' then
    begin
      if cbNotPLZ.Checked
        then sWhere := SQL_Where_Add(sWhere, 'NOT '+SQL_Where_Contains(ediPLZ.text, 'PLZ'))
        else sWhere := SQL_Where_Add(sWhere,        SQL_Where_Contains(ediPLZ.text, 'PLZ'));
    end;
  //Kirche
  if ediKirche.text <> '' then
    begin
      if cbNotKirche.Checked
        then sWhere := SQL_Where_Add(sWhere, 'NOT '+SQL_Where_Contains(ediKirche.text, 'Kirche'))
        else sWhere := SQL_Where_Add(sWhere,        SQL_Where_Contains(ediKirche.text, 'Kirche'));
    end;
  //Gemeinde
  if cbGemeinde.text <> '' then
    begin
      if cbNotGem.Checked
        then sWhere := SQL_Where_Add(sWhere, 'NOT '+SQL_Where_Contains(cbGemeinde.text, 'Gemeinde'))
        else sWhere := SQL_Where_Add(sWhere,        SQL_Where_Contains(cbGemeinde.text, 'Gemeinde'));
    end;
  //Taufmonat
  if ediTaufM.text <> '00'          then sWhere := SQL_Where_Add(sWhere, 'strftime(''%m'',TaufDatum)='''+ediTaufM.text+'''');
  //TaufJahr
  if ediTaufJ.text <>  sNotUsedYear then sWhere := SQL_Where_Add(sWhere, 'strftime(''%Y'',TaufDatum)='''+ediTaufJ.text+'''');
  //KonfJahr
  if ediKonfJ.text <>  sNotUsedYear then sWhere := SQL_Where_Add(sWhere, 'strftime(''%Y'',KonfDatum)='''+ediKonfJ.text+'''');
  //TaufJahr
  if ediTrauJ.text <>  sNotUsedYear then sWhere := SQL_Where_Add(sWhere, 'strftime(''%Y'',TrauDatum)='''+ediTrauJ.text+'''');
  //Versand
  if ediVers.text  <> ''            then sWhere := SQL_Where_Add(sWhere,SQL_Where_Contains(ediVers.text, 'Versand'));
  //Geburtstag
  if ediGebVon.text <> sNotUsedYear then sWhere := SQL_Where_Add(sWhere, 'strftime(''%Y'',Geburtstag)>='''+ediGebVon.text+'''');
  if ediGebBis.text <> sNotUsedYear then sWhere := SQL_Where_Add(sWhere, 'strftime(''%Y'',Geburtstag)<='''+ediGebBis.text+'''');
  //Geburtsmonat
  if ediGebM.text  <> '00'          then sWhere := SQL_Where_Add(sWhere, 'strftime(''%m'',Geburtstag)='''+ediGebM.text+'''');
  //Geschlecht
  if CBGeschlecht.Text<> ''         then sWhere := SQL_Where_Add(sWhere,SQL_Where_Contains(CBGeschlecht.text, 'Geschlecht'));
  //FamStand
  if CBFamstand.Text  <> ''         then sWhere := SQL_Where_Add(sWhere,SQL_Where_Contains(copy(CBFamstand.text,5,4), 'Familienstand'));

  frmDM.dsetHelp.sql.add(sWhere);
  frmDM.dsetHelp.sql.add('order by PersonenID');

  //Datenbank öffnen
  frmDM.dbStatus(true);
  frmDM.dsetHelp.open;
  while not frmDM.dsetHelp.eof do
    begin
      gut := true;
      inc(summ_gesamt);

      { ** Vergleiche ** ** Vergleiche ** ** Vergleiche ** ** Vergleiche ** }

      //Nachname in Where
      {Vorname}
      if ediVorname.text <> '' then          {mit pos einen Teilstring suchen}
        if pos(uppercase(ediVorname.text),uppercase(frmDM.dsetHelp.FieldByName('Vorname').asstring+' '+frmDM.dsetHelp.FieldByName('Vorname2').asstring)) = 0
          then gut := false;
      //Strasse in Where
      //Land in Where
      //PLZ in Where
      //Ort
      if ediOrt.text <> '' then          {mit pos einen Teilstring suchen}
        if pos(uppercase(ediOrt.text),uppercase(frmDM.dsetHelp.FieldByName('Ort').asstring+' '+frmDM.dsetHelp.FieldByName('Ortsteil').asstring)) = 0
          then gut := false;
      //Kirche in Where
      //und viele andere in Where Klausel

      //Kommunikanten
      if cbKomm.Checked
        then
          begin
            frmDM.dsetHelp1.SQL.Text:='select count(*) as c from Kommunionen where PersonenID='+frmDM.dsetHelp.FieldByName('PersonenID').asstring+
                                      ' and strftime(''%Y'',Abendmahlsdatum) = '''+ediKommJ.Text+'''';
            frmDM.dsetHelp1.Open;
            if frmDM.dsetHelp1.FieldByName('c').asinteger = 0 then gut := false;
            frmDM.dsetHelp1.Close;
          end;

      if gut {Alle Bedingungen erfüllt ?}
        then
          begin
            if not frmDM.dsetHelp.FieldByName('Markiert').asboolean
              then //Nur die unmarkierten dazu markieren, wegen der Geschwindigkeit
                begin
                  frmDM.dsetHelp.edit;
                  frmDM.dsetHelp.FieldByName('Markiert').asboolean := true;
                  frmDM.dsetHelp.post;
                end;
          end;

      frmDM.dsetHelp.next;
      Statusbar.SimpleText := 'Ich bin an Person ' +IntToStr(summ_gesamt);
      Statusbar.Update;
    end;
  frmDM.dsetHelp.close;

  frmDM.dsetHelp.SQL.Text:='select count(PersonenID) as m from personen where Markiert<>''0''';
  frmDM.dsetHelp.open;

  screen.cursor := crdefault;
  MessageDlg('Es sind jetzt insgesamt '+inttostr(frmDM.dsetHelp.FieldByName('m').AsInteger)+' Personen markiert', mtInformation, [mbOK],0);
  frmDM.dsetHelp.close;
  frmDM.dbStatus(false); //Datenbank schliessen

  if MessageDlg('Weiter suchen?', mtConfirmation, [mbYES, mbNo],0) = mrNo
    then close
    else cbDelMark.Checked := false;
end;

procedure TfrmSuche.ediGebChange(Sender: TObject);
begin
  if (ediGebVon.Text <> sNotUsedYear) and
     (ediGebBis.Text <> sNotUsedYear)
    then
      begin
        try
          labAlter.Caption := Inttostr(strtoInt(FormatDateTime('YYYY', now))-strtoInt(ediGebBis.Text))+' bis '+
                              Inttostr(strtoInt(FormatDateTime('YYYY', now))-strtoInt(ediGebVon.Text))+' Jahre';
        except
          labAlter.Caption := 'Error';
        end;
        exit;
      end;
  if (ediGebBis.Text <> sNotUsedYear)
    then
      begin
        try
          labAlter.Caption :=  'Ab '+Inttostr(strtoInt(FormatDateTime('YYYY', now))-strtoInt(ediGebbis.Text))+' Jahre';
        except
          labAlter.Caption := 'Error';
        end;
        exit;
      end;
  if (ediGebVon.Text <> sNotUsedYear)
    then
      begin
        try
          labAlter.Caption :=  '0 bis '+ Inttostr(strtoInt(FormatDateTime('YYYY', now))-strtoInt(ediGebvon.Text))+' Jahre';
        except
          labAlter.Caption := 'Error';
        end;
        exit;
      end;
  labAlter.Caption := '';
end;

procedure TfrmSuche.FormActivate(Sender: TObject);
begin
  ediName.SetFocus;
  cbGemeinde.Items.Text := sGemeinden;
end;

procedure TfrmSuche.FormCreate(Sender: TObject);

begin
  btnResetClick(Sender);
end;

procedure TfrmSuche.FormKeyPress(Sender: TObject; var Key: char);
begin
  if (key = #13) and ((activeControl is Tdbedit) or
                      (activeControl is Tdbcombobox) or
                      (activeControl is Tdbcheckbox))
    then
      begin
        key := #0;
        SelectNext(activeControl, True, True);
      end;
end;

procedure TfrmSuche.btnResetClick(Sender: TObject);
begin
  cbDelMark.Checked     := false;
  cbAbgaenge.Checked    := true;
  ediName.Text          := '';
  ediVorName.Text       := '';
  ediStrasse.Text       := '';
  ediLand.Text          := '';
  cbNotPLZ.Checked      := false;
  ediPLZ.Text           := '00000';
  ediOrt.Text           := '';
  ediKirche.Text        := '';
  cbNotKirche.Checked   := false;
  cbNotGem.Checked      := false;
  cbGemeinde.Text       := '';
  ediTaufM.Text         := '00';
  ediTaufJ.Text         := sNotUsedYear;
  ediKonfJ.Text         := sNotUsedYear;
  ediTrauJ.Text         := sNotUsedYear;
  ediVers.Text          := '';
  ediGebVon.Text        := sNotUsedYear;
  ediGebBis.Text        := sNotUsedYear;
  ediGebM.Text          := '00';
  CBGeschlecht.Text     := '';
  CBFamstand.Text       := '';
  cbKomm.Checked        := false;
  ediKommJ.Text         := year(datetostr(date));
end;

end.

