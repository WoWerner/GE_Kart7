unit dm;

{$mode objfpc}{$H+}

interface

uses
  Classes, 
  SysUtils, 
  db, 
  dbf, 
  FileUtil, 
  Forms, 
  Controls, 
  Graphics,
  Dialogs,
  ZConnection, 
  ZDataset, 
  ZSqlMonitor;

type

  { TfrmDM }

  TfrmDM = class(TForm)
    dsetHelp2: TZQuery;
    dsGD: TDatasource;
    dsetHelp1: TZQuery;
    dsetGD: TZQuery;
    dsHelp: TDatasource;
    dsHelp1: TDatasource;
    dsHelp2: TDataSource;
    dsKinder: TDatasource;
    DBF_Help: TDbf;
    dsPersonen: TDatasource;
    dsKomm: TDatasource;
    dsBesuch: TDatasource;
    dsetPersonen: TZQuery;
    dsetHelp: TZQuery;
    dsetKinder: TZQuery;
    ZConGE_Kart7: TZConnection;
    dsetKomm: TZQuery;
    dsetBesuch: TZQuery;
    ZSQLMonitor: TZSQLMonitor;
    procedure dsetGDAfterScroll(DataSet: TDataSet);
    procedure dsetPostError(DataSet: TDataSet; E: EDatabaseError; var DataAction: TDataAction);
    procedure dsetPersonenAfterScroll(DataSet: TDataSet);
    procedure dsetPersonenBeforePost(DataSet: TDataSet);
    procedure dsPersonenStateChange(Sender: TObject);
    procedure ZSQLMonitorLogTrace(Sender: TObject; Event: TZLoggingEvent);
    procedure ZSQLMonitorTrace(Sender: TObject; Event: TZLoggingEvent; var LogTrace: Boolean);
  private
    { private declarations }
  public
    { public declarations }
    Procedure dbStatus(Activ : boolean; OpenDatabases: boolean = false);
    procedure SetDBPath(Path: string);
    Procedure AutoEdit(Activ : boolean);
    Procedure ExecSQL(SQL: string; ExternProtected: boolean = false);
  end;

var
  frmDM: TfrmDM;

implementation

uses 
  kartei,
  GD,
  help,
  LCLproc,
  ZDbcIntfs,
  ZDbcLogging,
  global, 
  appsettings;

{$R *.lfm}

Procedure TfrmDM.AutoEdit(Activ : boolean);

begin
  dsPersonen.AutoEdit := Activ;
  dsKinder.AutoEdit   := Activ;
  dsHelp.AutoEdit     := Activ;
end;

Procedure TfrmDM.ExecSQL(SQL: string; ExternProtected: boolean = false);

var
  nRow : integer;

begin
  if not ZConGE_KART7.Connected then dbStatus(true);

  if ExternProtected
    then ZConGE_KART7.ExecuteDirect(SQL, nRow)
    else
      begin
        screen.Cursor:=crSQLWait;
        try
          ZConGE_KART7.ExecuteDirect(SQL, nRow);
        except
          // Log Exception..
          on E: Exception do
            begin
              LogAndShowError(E.Message);
              raise;
            end;
        end;
        screen.Cursor:=crdefault;
      end;
  myDebugLN(inttostr(nRow) +' Zeilen aktualisiert');
end;

procedure TfrmDM.SetDBPath(Path: string);

begin
  ZConGE_KART7.Connected := false;
  ZConGE_KART7.Database  := Path;
end;

Procedure TfrmDM.dbStatus(Activ : boolean; OpenDatabases: boolean = false);

var sHelp : string;

begin
  if Activ
    then
      begin
        try
          ZConGE_KART7.Connected := true;
          ZSQLMonitor.Active     := bSQLDebug;

          if not bDatabaseVersionChecked
            then
              begin
                myDebugLN('Prüfe DB');

                //SQLite Version auslesen
                dsetHelp.SQL.Text := 'select sqlite_version() as version';
                dsetHelp.Open;
                myDebugLN('Aktuelle SQLite Version ist: '+dsetHelp.FieldByName('Version').asstring);
                dsetHelp.Close;
                myDebugLN('Aktuelle Zeos   Version ist: '+ZConGE_KART7.Version);

                //Zur Sicherheit Tabelle anlegen (if not exists)
                ExecSQL('create table if not exists [Version] ([V] varchar (10) null);');
                try
                  dsetHelp.SQL.Text := 'select count(*) as C from Version';
                  dsetHelp.Open;
                  if Trim(dsetHelp.FieldByName('C').AsString) = '0'
                    then //Update auf DB-Format 7.1
                      begin
                        dsetHelp.Close;
                        ExecSQL('insert into Version (V) values(''7.1'')');
                        ExecSQL('alter table Gottesdienst rename to GD');
                        ExecSQL(sSQL_GD_Table_New);
                        ExecSQL('insert into gottesdienst select * from GD');
                        ExecSQL('alter table Gottesdienst add column [KollekteFuer] VARCHAR(100) NULL');
                        ExecSQL('drop table if exists GD;');
                      end;
                finally
                  if dsetHelp.Active then dsetHelp.Close;
                end;

                //Version auslesen
                dsetHelp.SQL.Text := 'select * from Version';
                dsetHelp.Open;
                sHelp := dsetHelp.FieldByName('V').asstring;
                dsetHelp.Close;
                myDebugLN('Aktuelle DB-Version: '+sHelp);

                //Update auf DB-Format 7.2
                if sHelp = '7.1'
                  then
                    begin
                      ExecSQL('alter table Personen add column [email2] VARCHAR(100) NULL');
                      ExecSQL('alter table Personen add column [email3] VARCHAR(100) NULL');
                      ExecSQL('alter table Personen add column [Internet2] VARCHAR(100) NULL');
                      ExecSQL('alter table Personen add column [Internet3] VARCHAR(100) NULL');
                      ExecSQL('alter table Personen add column [Todesort] VARCHAR(100) NULL');
                      ExecSQL('alter table Personen add column [sRes1] VARCHAR(100) NULL');      //Adresszusatz
                      ExecSQL('alter table Personen add column [sRes2] VARCHAR(100) NULL');
                      ExecSQL('alter table Personen add column [sRes3] VARCHAR(100) NULL');
                      ExecSQL('alter table Personen add column [sRes4] VARCHAR(100) NULL');
                      ExecSQL('alter table Personen add column [dRes1] DATE NULL');               //Ehegatten Geburtsdatum
                      ExecSQL('alter table Personen add column [dRes2] DATE NULL');
                      ExecSQL('alter table Personen add column [nRes1] INTEGER NULL');
                      ExecSQL('alter table Personen add column [nRes2] INTEGER NULL');
                      ExecSQL('update Version set V=''7.2''');
                      sHelp := '7.2';
                      myDebugLN('DB-Version now: '+sHelp);
                    end;

                //Update auf DB-Format 7.5
                if sHelp = '7.2'
                  then
                    begin
                      ExecSQL('create table if not exists [Gemeinde] ([Adr1] varchar (150) null, [Adr2] varchar (150) null);');
                      //Defaults aus Ini übernehmen
                      ExecSQL('insert into Gemeinde (Adr1, Adr2) values('''+
                               help.ReadIniVal(sIniFile, 'Gemeinde','Zeile1','Mustergemeinde', false) +''','''+
                               help.ReadIniVal(sIniFile, 'Gemeinde','Zeile2','Musterstadt', false) +''')');
                      ExecSQL('update Version set V=''7.5''');
                      sHelp := '7.5';
                      myDebugLN('DB-Version now: '+sHelp);
                    end;

                //Update auf DB-Format 7.5.1
                if sHelp = '7.5'
                  then
                    begin
                      ExecSQL('alter table Kinder add column [Taufdatum] DATE NULL');
                      ExecSQL('update Version set V=''7.5.1''');
                      sHelp := '7.5.1';
                      myDebugLN('DB-Version now: '+sHelp);
                    end;

                //Update auf DB-Format 7.5.2
                if sHelp = '7.5.1'
                  then
                    begin
                      ExecSQL('update Personen set sRes2 = Gemeinde');
                      ExecSQL('alter table Personen drop column [Gemeinde]');
                      ExecSQL('alter table Personen add column [Gemeinde] VARCHAR(10) NULL');
                      ExecSQL('update Personen set Gemeinde = sRes2');
                      ExecSQL('update Personen set sRes2 = ""');
                      ExecSQL('update Version set V=''7.5.2''');
                      sHelp := '7.5.2';
                      myDebugLN('DB-Version now: '+sHelp);
                    end;

                //Update auf DB-Format 7.5.3
                if sHelp = '7.5.2'
                  then
                    begin
                      ExecSQL('create table if not exists [Users] ([Name] varchar (200) null);');
                      ExecSQL('update Version set V=''7.5.3''');
                      sHelp := '7.5.3';
                      myDebugLN('DB-Version now: '+sHelp);
                    end;
                //Update auf DB-Format 7.5.4
                if sHelp = '7.5.3'
                  then
                    begin
                      ExecSQL('drop table if exists Users');
                      //Bereinigen der Tabelle KOMMUNIONEN
                      ExecSQL('delete from '+global.sKommTablename+' where '+SQL_Where_IsNull('Abendmahlsdatum'));
                      ExecSQL('delete from '+global.sKommTablename+' where '+SQL_Where_IsNull('PersonenID'));
                      //doppelte DS löschen
                      dsetHelp2.sql.Clear;
                      dsetHelp2.sql.add('SELECT KommID, Abendmahlsdatum, PersonenID, COUNT(*) c FROM '+global.sKommTablename+' GROUP BY Abendmahlsdatum, PersonenID HAVING c > 1');
                      dsetHelp2.open;
                      while not dsetHelp2.EOF do
                        begin
                          ExecSQL('delete from '+global.sKommTablename+' where KommID='+dsetHelp2.fieldByName('KommID').asstring);
                          frmDM.dsetHelp2.Next;
                        end;
                      dsetHelp2.Close;
                      //Index erstellen um doppelte DS zu verhindern
                      ExecSQL('create unique index if not exists komm_unique on '+global.sKommTablename+' (PersonenID, Abendmahlsdatum)');
                      //Schreibfehler beseitigen
                      ExecSQL('alter table '+global.sGDTablename+' rename column Kommunikaten to Kommunikanten');
                      ExecSQL('alter table '+global.sGDTablename+' rename column Gastkommunikaten to Gastkommunikanten');
                      //Umwandlung Feld Gemeinde
                      ExecSQL('alter table '+global.sGDTablename+' add column [Temp] VARCHAR(20) NULL');
                      ExecSQL('update '     +global.sGDTablename+' set Temp = Gemeinde');
                      ExecSQL('alter table '+global.sGDTablename+' drop column [Gemeinde]');
                      ExecSQL('alter table '+global.sGDTablename+' add column [Gemeinde] VARCHAR(10) NULL');
                      ExecSQL('update '     +global.sGDTablename+' set Gemeinde = Temp');
                      ExecSQL('update '     +global.sGDTablename+' set Temp = ""');
                      //Umwandlung Feld Kollekte
                      ExecSQL('update '     +global.sGDTablename+' set Temp = Kollekte');
                      ExecSQL('alter table '+global.sGDTablename+' drop column [Kollekte]');
                      ExecSQL('alter table '+global.sGDTablename+' add column [Kollekte] VARCHAR(20) NULL');
                      ExecSQL('update '     +global.sGDTablename+' set Kollekte = Temp');
                      ExecSQL('update '     +global.sGDTablename+' set Temp = ""');
                      //Clean up
                      ExecSQL('alter table '+global.sGDTablename+' drop column [Temp]');

                      ExecSQL('update Version set V=''7.5.4''');
                      sHelp := '7.5.4';
                      myDebugLN('DB-Version now: '+sHelp);
                    end;
                bDatabaseVersionChecked := true;
              end;

          if OpenDatabases
            then
              begin
                myDebugLN('Öffne DB');
                dsetPersonen.Active := true;
                dsetKinder.Active   := true;
                dsetKomm.Active     := true;
                dsetBesuch.Active   := true;
              end;
        except
          // Log Exception..
          on E: Exception do
            begin
              LogAndShowError(E.Message);
              raise;
            end;
        end;
      end
    else
      begin
        myDebugLN('Schliesse DB');
        dsetPersonen.Active    := false;
        dsetKinder.Active      := false;
        dsetKomm.Active        := false;
        dsetBesuch.Active      := false;
        ZConGE_KART7.Connected := false;
      end;
end;

procedure TfrmDM.dsetPersonenBeforePost(DataSet: TDataSet);
begin  
  if frmKartei.BeforePost //Macht eine Abfrage ob gespeichert werden soll
    then
      begin
        //Speichern, nix zu tun
      end
    else
      begin
        dsetPersonen.CancelUpdates;
      end;
end;

procedure TfrmDM.dsPersonenStateChange(Sender: TObject);
var
  sMes: String;

begin
  if bSQLDebug
    then
      begin
        sMes := 'dsPersonen.State: ';
        case dsPersonen.State of
          dsInactive:     sMes := sMes + 'dsInactive:     The dataset is not active. No data is available.';
          dsBrowse:       sMes := sMes + 'dsBrowse:       The dataset is active, and the cursor can be used to navigate the data.';
          dsEdit:         sMes := sMes + 'dsEdit:         The dataset is in editing mode: the current record can be modified.';
          dsInsert:       sMes := sMes + 'dsInsert:       The dataset is in insert mode: the current record is a new record which can be edited.';
          dsSetKey:       sMes := sMes + 'dsSetKey:       The dataset is calculating the primary key.';
          dsCalcFields:   sMes := sMes + 'dsCalcFields:   The dataset is calculating it''s calculated fields.';
          dsFilter:       sMes := sMes + 'dsFilter:       The dataset is filtering records.';
          dsNewValue:     sMes := sMes + 'dsNewValue:     The dataset is showing the new values of a record.';
          dsOldValue:     sMes := sMes + 'dsOldValue:     The dataset is showing the old values of a record.';
          dsCurValue:     sMes := sMes + 'dsCurValue:     The dataset is showing the current values of a record.';
          dsBlockRead:    sMes := sMes + 'dsBlockRead:    The dataset is open, but no events are transferred to datasources.';
          dsInternalCalc: sMes := sMes + 'dsInternalCalc: The dataset is calculating it''s internally calculated fields.';
          dsOpening:      sMes := sMes + 'dsOpening:      The dataset is currently opening, but is not yet completely open.';
          dsRefreshFields:sMes := sMes + 'dsRefreshFields: Dataset is refreshing field values from server after an update.';
        else
                          sMes := sMes + 'Other';
        end;
        myDebugLN(sMes);
      end;
end;

procedure TfrmDM.ZSQLMonitorLogTrace(Sender: TObject; Event: TZLoggingEvent);

var
  sMes: String;

begin
  if bSQLDebug
    then
      begin
        sMes := '';
        case Event.Category of
          lcConnect:      sMes := sMes + 'Connect           ';
          lcDisconnect:   sMes := sMes + 'Disconnect        ';
          lcTransaction:  sMes := sMes + 'Transaction       ';
          lcExecute:      sMes := sMes + 'Execute           ';
          lcPrepStmt:     sMes := sMes + 'Prepare           ';
          lcBindPrepStmt: sMes := sMes + 'Bind prepared     ';
          lcExecPrepStmt: sMes := sMes + 'Execute prepared  ';
          lcUnprepStmt:   sMes := sMes + 'Unprepare prepared';
        else
                          sMes := sMes + 'Other             ';
        end;

        sMes := sMes +' ' + Event.Message;

        if Event.Error <> '' then sMes := sMes +', Err: ' + Event.error;

         myDebugLN(sMes);
      end;
end;

procedure TfrmDM.ZSQLMonitorTrace(Sender: TObject; Event: TZLoggingEvent; var LogTrace: Boolean);
begin
  //myDebugLN(Event.Message);
end;

procedure TfrmDM.dsetPersonenAfterScroll(DataSet: TDataSet);
begin
  frmKartei.AfterScroll;
end;

procedure TfrmDM.dsetGDAfterScroll(DataSet: TDataSet);
begin
  frmGD.AfterScroll;
end;

procedure TfrmDM.dsetPostError(DataSet: TDataSet; E: EDatabaseError; var DataAction: TDataAction);

var
  sMes: String;

begin
  if bSQLDebug
    then
      begin
        sMes := 'Fehler in '+DataSet.Name+': ' + E.Message;
        myDebugLN(sMes);
      end;
end;

end.

