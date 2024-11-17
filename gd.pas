unit gd;

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
  DBGrids,
  ExtCtrls,
  DbCtrls,
  StdCtrls,
  db;

type

  { TfrmGD }

  TfrmGD = class(TForm)
    btnEnde: TButton;
    btnKomDel: TButton;
    btnNew: TButton;
    btnDel: TButton;
    btnKomAdd: TButton;
    cbGemeinde: TComboBox;
    DBCBGottesdienstForm1: TDBComboBox;
    DBCBGottesdienstForm2: TDBComboBox;
    DBCBGottesdienstForm3: TDBComboBox;
    DBCBTag: TDBComboBox;
    DBEdiKollekte: TDBEdit;
    DBEdiKollekteFuer: TDBEdit;
    DBEdiKommunikanten: TDBEdit;
    DBEdiBesucherKiGo: TDBEdit;
    DBEdiGastkommunikanten: TDBEdit;
    DBEdiPfarrer: TDBEdit;
    DBEdiKuester: TDBEdit;
    DBEdiPredigttext: TDBEdit;
    DBEdiBesucher: TDBEdit;
    DBEdiLiturgGesang2: TDBEdit;
    DBEdiLiturgGesang3: TDBEdit;
    DBEdiIntroitus: TDBEdit;
    DBEdiName: TDBEdit;
    DBEdiGottesdienstDatum: TDBEdit;
    DBEdiGemeinde: TDBEdit;
    DBEdiEingangslied: TDBEdit;
    DBEdiHauptlied: TDBEdit;
    DBEdiLiedVorPredigt: TDBEdit;
    DBEdiLiedNachPredigt: TDBEdit;
    DBEdiLiedZurBereitung: TDBEdit;
    DBEdiLiedZurAusteilung: TDBEdit;
    DBEdiSchlusslied: TDBEdit;
    DBEdiLiturgGesang1: TDBEdit;
    DBGrid1: TDBGrid;
    DBGridPersonen: TDBGrid;
    DBGridKommunikanten: TDBGrid;
    DBMemo: TDBMemo;
    DBNavigator1: TDBNavigator;
    ediSSNachname: TLabeledEdit;
    Label1: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label149: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label2: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    Label26: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    labGottesdienstForm2: TLabel;
    labGottesdienstForm3: TLabel;
    labTag: TLabel;
    labGottesdienstForm1: TLabel;
    Panel1: TPanel;
    procedure btnDelClick(Sender: TObject);
    procedure btnEndeClick(Sender: TObject);
    procedure btnKomAddClick(Sender: TObject);
    procedure btnKomDelClick(Sender: TObject);
    procedure btnNewClick(Sender: TObject);
    procedure cbGemeindeChange(Sender: TObject);
    procedure DBCBChange(Sender: TObject);
    procedure ediSSNachnameDblClick(Sender: TObject);
    procedure ediSSNachnameKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: char);
    procedure FormShow(Sender: TObject);
  private
    UpdateKommunikaten: boolean;
  public
    procedure AfterScroll;
  end;

var
  frmGD: TfrmGD;

implementation

uses
  global,
  help,
  variants,//VarArrayOf
  dm;

{$R *.lfm}

{ TfrmGD }

procedure TfrmGD.FormShow(Sender: TObject);
begin
  frmDM.dsetGD.Open;
  cbGemeinde.Items.Text := sGemeinden;
  cbGemeinde.Text       := sGemeindenAlle;
  dbGrid1.SetFocus;
end;

procedure TfrmGD.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  if (frmDM.dsetGD.state in [dsEdit, dsInsert]) and
     (MessageDlg('Sollen die Änderungen gesichert werden?', mtConfirmation, [mbYES, mbNo],0) = mrYes)
    then frmDM.dsetGD.post
    else frmDM.dsetGD.CancelUpdates;
  frmDM.dsetGD.Close;
end;

procedure TfrmGD.FormCreate(Sender: TObject);

var sHelp: string;

begin
  sHelp := 'A: Andacht'#13+
           'B: Beichte'#13+
           'F: Frei'#13+
           'H: HauptGD'#13+
           'K: KinderGD'#13+
           'L: LeseGD'#13+
           'M: Matutin'#13+
           'P: PredigtGD'#13+
           'S: Sonstige'#13+
           'V: Vesper';
  DBCBGottesdienstForm1.Hint := sHelp;
  DBCBGottesdienstForm2.Hint := sHelp;
  DBCBGottesdienstForm3.Hint := sHelp;
  UpdateKommunikaten         := false;
end;

procedure TfrmGD.FormKeyPress(Sender: TObject; var Key: char);

begin
  if (key = #13) and ((activeControl is Tdbedit) or (activeControl is Tdbcombobox) or (activeControl is Tdbcheckbox))
    then
      begin
        key := #0;
        SelectNext(activeControl, True, True);
      end;
end;

procedure TfrmGD.FormActivate(Sender: TObject);

var i : integer;

begin
  for i := 0 to frmDM.dsetGD.FieldCount - 1
    do frmDM.dsetGD.fieldbyname(frmDM.dsetGD.FieldDefs[i].name).visible := false;
  frmDM.dsetGD.fieldbyname('GottesdienstDatum').visible := true;
  //frmDM.dsetGD.fieldbyname('Kollekte').visible          := true;
end;

procedure TfrmGD.btnEndeClick(Sender: TObject);
begin
  close;
end;

procedure TfrmGD.btnKomAddClick(Sender: TObject);

var
  i : Integer;
  DataSet : TDataSet;

begin
  DataSet := DBGridPersonen.DataSource.DataSet;
  if DBGridPersonen.SelectedRows.Count > 0 then
    begin
      for i := 0 to DBGridPersonen.SelectedRows.Count-1 do
         begin
           DataSet.GotoBookmark(Pointer(DBGridPersonen.SelectedRows.Items[i]));
           frmDM.ExecSQL(Format(global.sSQL_KOMM_ADD, [SQLiteDateFormat(frmDM.dsetGD.FieldByName('GottesdienstDatum').AsDateTime), DataSet.FieldByName('PersonenID').AsString]));
         end;
    end;
  UpdateKommunikaten := true;
  AfterScroll();
end;

procedure TfrmGD.btnKomDelClick(Sender: TObject);

var
  i : Integer;
  DataSet : TDataSet;

begin
  DataSet := DBGridKommunikanten.DataSource.DataSet;
  if DBGridKommunikanten.SelectedRows.Count > 0 then
    begin
      for i := 0 to DBGridKommunikanten.SelectedRows.Count-1 do
         begin
           DataSet.GotoBookmark(Pointer(DBGridKommunikanten.SelectedRows.Items[i]));
           frmDM.ExecSQL('Delete from '+global.sKommTablename+' where KommID='+DataSet.FieldByName('KommID').AsString);
         end;
    end;
  UpdateKommunikaten := true;
  AfterScroll();
end;

procedure TfrmGD.btnDelClick(Sender: TObject);
begin
  if MessageDlg('Löschen?', 'Soll der Gottesdienst vom ' + frmDM.dsetGD.FieldByName('GottesdienstDatum').AsString + ' gelöscht werden', mtConfirmation, [mbYes, mbNo], 0) = mrYes
    then
      begin
        frmDM.ExecSQL(Format(global.sSQL_GD_DEL, [frmDM.dsetGD.FieldByName('GDID').AsString]));
        frmDM.dsetGD.Refresh;
        frmDM.dsetGD.First;
      end;
end;

procedure TfrmGD.btnNewClick(Sender: TObject);
begin
  if (frmDM.dsetGD.state in [dsEdit, dsInsert])
    then frmDM.dsetGD.post;
  frmDM.ExecSQL(Format(global.sSQL_GD_ADD, [FormatDateTime('yyyy-mm-dd', date())]));
  frmDM.dsetGD.refresh;
  frmDM.dsetGD.First;
  DBEdiGottesdienstDatum.SetFocus;
end;

procedure TfrmGD.cbGemeindeChange(Sender: TObject);

var
  sFilter: string;

begin
  sFilter := '';
  if cbGemeinde.Text = ''
    then sFilter := SQL_Where_Add(sFilter, SQL_Where_IsNull('Gemeinde'))
    else if cbGemeinde.Text <> sGemeindenAlle
      then sFilter := SQL_Where_Add(sFilter, 'Gemeinde=''' + cbGemeinde.Text + '''');
  myDebugLN('Setze Filter: ' + sFilter);
  frmDM.dsetGD.Filter   := sFilter;
  frmDM.dsetGD.Filtered := (sFilter <> '');
  frmDM.dsetGD.Refresh;
  frmDM.dsetGD.First;
end;

procedure TfrmGD.DBCBChange(Sender: TObject);

var help_chr : Char;
    help_str : string;

    function DB_To_Label(s : String):String;

    begin
      s := uppercase(s);
      if s = '' then s := ' ';
      help_chr := s[1];
      case help_chr of
        'A' : result := 'Andacht';
        'B' : result := 'Beichte';
        'F' : result := '';
        'H' : result := 'HauptGD';
        'K' : result := 'KinderGD';
        'L' : result := 'LeseGD';
        'M' : result := 'Matutin';
        'P' : result := 'PredigtGD';
        'S' : result := 'Sonstige';
        'V' : result := 'Vesper';
      else
        result := '';
      end;
    end;

begin
  if frmDM.dsetGD.state in [dsEdit, dsInsert]
    then
      begin
        frmDM.dsetGD.post;
        frmDM.dsetGD.refresh;
      end;
  LabGottesdienstForm1.caption := DB_To_Label(frmDM.dsetGD.FieldByName('GottesdienstForm1').asstring);
  LabGottesdienstForm2.caption := DB_To_Label(frmDM.dsetGD.FieldByName('GottesdienstForm2').asstring);
  LabGottesdienstForm3.caption := DB_To_Label(frmDM.dsetGD.FieldByName('GottesdienstForm3').asstring);

  help_str := uppercase(frmDM.dsetGD.FieldByName('Tag').asstring);
  if Help_Str <> ''
    then
      begin
        help_chr := help_str[1];
        case help_chr of
          'F' : LabTag.caption := 'Feiertag';
          'S' : LabTag.caption := 'Sonntag';
          'W' : LabTag.caption := 'Wochentag';
        else
          LabTag.caption := '';
        end;
      end;
end;

procedure TfrmGD.ediSSNachnameDblClick(Sender: TObject);
begin
  ediSSNachname.Text:='';
end;

procedure TfrmGD.ediSSNachnameKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if ediSSNachname.Text <> ''
    then frmDM.dsetHelp1.Locate('Nachname', VarArrayOf([ediSSNachname.Text]), [loCaseInsensitive, loPartialKey]);
end;

 procedure TfrmGD.AfterScroll;
 begin
   if frmGD.Visible then
     begin
       DBCBChange(self);

       if frmDM.dsetHelp2.Active then frmDM.dsetHelp2.Close;
       frmDM.dsetHelp2.sql.Clear;
       frmDM.dsetHelp2.sql.add('select Vorname, Nachname, '+global.sKommTablename+'.KommID from '+global.sPersTablename);
       frmDM.dsetHelp2.sql.add('left join '+global.sKommTablename+' on '+global.sKommTablename+'.PERSONENID = '+global.sPersTablename+'.PERSONENID');
       frmDM.dsetHelp2.sql.add('where AbendmahlsDatum = '+SQLiteDateFormat(frmDM.dsetGD.FieldByName('GottesdienstDatum').AsDateTime)+' and ');
       frmDM.dsetHelp2.sql.add('Gemeinde="'+frmDM.dsetGD.FieldByName('Gemeinde').AsString+'"');
       frmDM.dsetHelp2.sql.add('order by nachname, vorname');
       frmDM.dsetHelp2.open;
       DBGridKommunikanten.Columns.Items[0].Width   := 100;
       DBGridKommunikanten.Columns.Items[1].Width   := 200;
       DBGridKommunikanten.Columns.Items[2].Visible := false;

       if frmDM.dsetHelp1.Active then frmDM.dsetHelp1.Close;
       frmDM.dsetHelp1.sql.Clear;
       frmDM.dsetHelp1.sql.add('select Vorname, Nachname, Geburtstag, Gemeinde, PERSONENID from '+global.sPersTablename);
       frmDM.dsetHelp1.sql.add('where Gemeinde="'+frmDM.dsetGD.FieldByName('Gemeinde').AsString+'" and ');
       frmDM.dsetHelp1.sql.add('PERSONENID not in (select PERSONENID from '+global.sKommTablename+' where AbendmahlsDatum = '+SQLiteDateFormat(frmDM.dsetGD.FieldByName('GottesdienstDatum').AsDateTime)+')');
       frmDM.dsetHelp1.sql.add('order by nachname, vorname');
       frmDM.dsetHelp1.open;
       DBGridPersonen.Columns.Items[0].Width:=100;
       DBGridPersonen.Columns.Items[1].Width:=100;
       DBGridPersonen.Columns.Items[2].Width:=80;
       DBGridPersonen.Columns.Items[3].Width:=70;
       DBGridPersonen.Columns.Items[4].Visible:=false;
       if UpdateKommunikaten
         then
           begin
             if frmDM.dsetGD.FieldByName('GdID').AsString <> ''
               then
                 frmDM.ExecSQL('update '+global.sGDTablename+' set Kommunikanten = '+ inttostr(frmDM.dsetHelp2.RecordCount)+' where GdID = '+frmDM.dsetGD.FieldByName('GdID').AsString);
             UpdateKommunikaten := false;
           end;
     end;
 end;

end.

