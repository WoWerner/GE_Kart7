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
    dsetVersion: TZQuery;
    dsetGemeinde: TZQuery;
    dsGD: TDatasource;
    dsetHelp1: TZQuery;
    dsetKomm: TZQuery;
    dsetBesuch: TZQuery;
    dsetGD: TZQuery;
    dsHelp: TDatasource;
    dsHelp1: TDatasource;
    dsKinder: TDatasource;
    DBF_Help: TDbf;
    dsPersonen: TDatasource;
    dsKomm: TDatasource;
    dsBesuch: TDatasource;
    dsetPersonen: TZQuery;
    dsetHelp: TZQuery;
    dsetKinder: TZQuery;
    ZConGE_Kart7: TZConnection;
    ZSQLMonitor: TZSQLMonitor;
    procedure dsetGDAfterScroll(DataSet: TDataSet);
    procedure dsetPersonenAfterScroll(DataSet: TDataSet);
    procedure dsetPersonenBeforePost(DataSet: TDataSet);
    procedure dsPersonenStateChange(Sender: TObject);
    procedure ZSQLMonitorLogTrace(Sender: TObject; Event: TZLoggingEvent);
  private
    { private declarations }
  public
    { public declarations }
    Procedure dbStatus(Activ : boolean; OpenDatabases: boolean = false);
    procedure SetDBPath(Path: string);
    Procedure AutoEdit(Activ : boolean);
    Procedure ExecSQL(SQL: string; ExternProtected: boolean = false);
    procedure SetSortPersonen(typ : integer);
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
  dsKomm.AutoEdit     := Activ;
  dsBesuch.AutoEdit   := Activ;
  dsHelp.AutoEdit     := Activ;
end;

procedure TfrmDM.SetSortPersonen(typ : integer);

begin
  case typ of
    //Umlautsortierung in Version 7.2.0.5 umgebaut
    0: begin
         // dsetPersonen.SortedFields:='Nachname,Vorname'; //Alte Version
         ExecSQL('Update PERSONEN SET TempString='+SQL_UTF8UmlautReplace('Nachname')+'||'' ''||'+SQL_UTF8UmlautReplace('Vorname'));
         dsetPersonen.SortedFields:='TempString';
         dsetPersonen.Refresh;
       end;
    1: begin
         ExecSQL('Update PERSONEN SET TempString='+SQL_UTF8UmlautReplace('Nachname')+'||'' ''||'+SQL_UTF8UmlautReplace('Vorname'));
         dsetPersonen.SortedFields:='TempString';
         dsetPersonen.Refresh;
         dsetPersonen.SortedFields:='Markiert,TempString';
         dsetPersonen.IndexFieldNames := StringReplace(dsetPersonen.IndexFieldNames, 'Desc', 'Desc', []);
       end;
    2: begin
         //ExecSQL('Update PERSONEN SET Tempinteger=strftime(''%j'',geburtstag)');   //j = day of year: 001-366 setzen (geht nicht bei Schaltjahren)
         ExecSQL('Update PERSONEN SET Tempinteger=strftime(''%m'', geburtstag)*31+strftime(''%d'', geburtstag)');
         dsetPersonen.SortedFields:='Tempinteger';
         dsetPersonen.Refresh;
       end;
    3: begin
         //dsetPersonen.SortedFields:='Ort,Strasse,Nachname,Vorname'; //Alte Version
         ExecSQL('Update PERSONEN SET TempString='+
                 SQL_UTF8UmlautReplace('Ort')     +'||'' ''||'+SQL_UTF8UmlautReplace('Strasse')+'||'' ''||'+
                 SQL_UTF8UmlautReplace('Nachname')+'||'' ''||'+SQL_UTF8UmlautReplace('Vorname'));
         dsetPersonen.SortedFields:='TempString';
         dsetPersonen.Refresh;
       end;
  end;
  myDebugLN('dsetPERSONEN.SortedFields:    '+dsetPersonen.SortedFields);
  myDebugLN('dsetPERSONEN.IndexFieldNames: '+dsetPersonen.IndexFieldNames);
  dsetPersonen.First;
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
                dsetVersion.Open;
                sHelp := dsetVersion.FieldByName('V').asstring;
                dsetVersion.Close;
                myDebugLN('DB-Version: '+sHelp);

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
                dsetGemeinde.Active := true;
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
        dsetGemeinde.Active    := false;
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
          dsInactive:     sMes := sMes + 'The dataset is not active. No data is available.';
          dsBrowse:       sMes := sMes + 'The dataset is active, and the cursor can be used to navigate the data.';
          dsEdit:         sMes := sMes + 'The dataset is in editing mode: the current record can be modified.';
          dsInsert:       sMes := sMes + 'The dataset is in insert mode: the current record is a new record which can be edited.';
          dsSetKey:       sMes := sMes + 'The dataset is calculating the primary key.';
          dsCalcFields:   sMes := sMes + 'The dataset is calculating it''s calculated fields.';
          dsFilter:       sMes := sMes + 'The dataset is filtering records.';
          dsNewValue:     sMes := sMes + 'The dataset is showing the new values of a record.';
          dsOldValue:     sMes := sMes + 'The dataset is showing the old values of a record.';
          dsCurValue:     sMes := sMes + 'The dataset is showing the current values of a record.';
          dsBlockRead:    sMes := sMes + 'The dataset is open, but no events are transferred to datasources.';
          dsInternalCalc: sMes := sMes + 'The dataset is calculating it''s internally calculated fields.';
          dsOpening:      sMes := sMes + 'The dataset is currently opening, but is not yet completely open.';
          dsRefreshFields:sMes := sMes + 'Dataset is refreshing field values from server after an update.';
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

procedure TfrmDM.dsetPersonenAfterScroll(DataSet: TDataSet);
begin
  frmKartei.AfterScroll;
end;


procedure TfrmDM.dsetGDAfterScroll(DataSet: TDataSet);
begin
  frmGD.AfterScroll;
end;


end.

