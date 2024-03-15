unit Main;

{$mode objfpc}{$H+}

interface

uses
  Classes,
  SysUtils,
  fileutil,
  LazUTF8,
  Forms,
  Controls,
  Graphics,
  Dialogs,
  Menus,
  StdCtrls,
  ExtCtrls,
  ComCtrls,
  kartei,
  LR_DBSet,
  LR_Class,
  LR_E_TXT,
  LR_E_HTM,
  LR_E_CSV,
  Windows;

type

  { TfrmMain }

  TfrmMain = class(TForm)
    btnKartei: TButton;
    btnSuchen: TButton;
    btnEnde: TButton;
    btnAktuelles: TButton;
    frCSVExport: TfrCSVExport;
    frDBDataSet: TfrDBDataSet;
    frDBDataSetDetail: TfrDBDataSet;
    frHTMExport: TfrHTMExport;
    frReport: TfrReport;
    frTextExport: TfrTextExport;
    imgSELK: TImage;
    Label1: TLabel;
    Label2: TLabel;
    LabGemeinde2: TLabel;
    labMyCity: TLabel;
    LabGemeinde1: TLabel;
    labMyWeb: TLabel;
    LabMyName: TLabel;
    labMyStreet: TLabel;
    labMyTelNr: TLabel;
    labMyMail: TLabel;
    labVersionNeu: TLabel;
    lb_Name: TLabel;
    MainMenu: TMainMenu;
    memGeburtstag: TMemo;
    MenuItem1: TMenuItem;
    MenuItem10: TMenuItem;
    MenuItem11: TMenuItem;
    MenuItem17: TMenuItem;
    MenuItem18: TMenuItem;
    MenuItem19: TMenuItem;
    MenuItem20: TMenuItem;
    MenuItem21: TMenuItem;
    MenuItem22: TMenuItem;
    MenuItem23: TMenuItem;
    MenuItem27: TMenuItem;
    Separator1: TMenuItem;
    mnuEMailVersand: TMenuItem;
    mnuDefineAnredeEntries: TMenuItem;
    mnuGebListSplit: TMenuItem;
    mnuExport: TMenuItem;
    mnuUserDefSQL: TMenuItem;
    mnuOpenDebug: TMenuItem;
    mnuInvertMark: TMenuItem;
    MenuItem24: TMenuItem;
    MenuItem25: TMenuItem;
    MenuItem26: TMenuItem;
    mnuExecSQLBatch: TMenuItem;
    mnuSelectDatabase: TMenuItem;
    mnuShowTaufTag: TMenuItem;
    mnuIncludeAbgang: TMenuItem;
    mnuGebList: TMenuItem;
    mnuDatenRestore: TMenuItem;
    mnuSQLDebug: TMenuItem;
    mnuDebug: TMenuItem;
    mnuColoredEditMode: TMenuItem;
    mnuShowOverView: TMenuItem;
    MenuItem3: TMenuItem;
    mnuDelMarkPers: TMenuItem;
    MenuSetAnrede: TMenuItem;
    mnuCSVImport: TMenuItem;
    mnuPrintDBStrukt: TMenuItem;
    mnuJubi: TMenuItem;
    mnuSortGebAlter: TMenuItem;
    mnuDelMark: TMenuItem;
    mnuGemeindeDefault: TMenuItem;
    mnuDefineKircheEntries: TMenuItem;
    mnuDatensicherung: TMenuItem;
    mnuFreeList: TMenuItem;
    mnuSelectReportFile: TMenuItem;
    MnuSortTauf: TMenuItem;
    MenuItem12: TMenuItem;
    MenuItem13: TMenuItem;
    MenuItem14: TMenuItem;
    MenuItem15: TMenuItem;
    MenuItem16: TMenuItem;
    mnuExecSQLSingle: TMenuItem;
    mnuImportGD: TMenuItem;
    mnuDruckerrnderanzeigen: TMenuItem;
    mnuAbsToEti: TMenuItem;
    mnuBuechersendung: TMenuItem;
    mnuDruckWarensendung: TMenuItem;
    mnuDruckNix: TMenuItem;
    mnuFehlerInKarteiSuchen: TMenuItem;
    mnuGDStatistik: TMenuItem;
    mnuGDBearbeiten: TMenuItem;
    mnuVolltextsuche: TMenuItem;
    MenuItem2: TMenuItem;
    mnuDrucken: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem7: TMenuItem;
    MenuItem8: TMenuItem;
    MenuItem9: TMenuItem;
    mnuStatistik: TMenuItem;
    mnuReportDesigner: TMenuItem;
    mnuAlle: TMenuItem;
    mnuMarkierte: TMenuItem;
    mnuSortOrt: TMenuItem;
    mnuSortGeb: TMenuItem;
    mnuSortName: TMenuItem;
    mnuPrintFormular: TMenuItem;
    mnuChangePW: TMenuItem;
    mnuGemeindeAdresse: TMenuItem;
    mnuImport: TMenuItem;
    mnuEnde: TMenuItem;
    mnuDatei: TMenuItem;
    labVersion: TLabel;
    OpenDialog: TOpenDialog;
    SaveDialog: TSaveDialog;
    StatusBar: TStatusBar;
    procedure btnAktuellesClick(Sender: TObject);
    procedure btnKarteiClick(Sender: TObject);
    procedure btnSuchenClick(Sender: TObject);
    procedure btnSuchenContextPopup(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure frReportGetValue(const ParName: String; var ParValue: Variant);
    procedure imgSELKClick(Sender: TObject);
    procedure labMyMailClick(Sender: TObject);
    procedure labMyWebClick(Sender: TObject);
    procedure labVersionNeuClick(Sender: TObject);
    procedure mnuEMailVersandClick(Sender: TObject);
    procedure MenuSetAnredeClick(Sender: TObject);
    procedure mnuAbsToEtiClick(Sender: TObject);
    procedure mnuChangePWClick(Sender: TObject);
    procedure mnuColoredEditModeClick(Sender: TObject);
    procedure mnuCSVImportClick(Sender: TObject);
    procedure mnuDatenRestoreClick(Sender: TObject);
    procedure mnuDatensicherungClick(Sender: TObject);
    procedure mnuDebugClick(Sender: TObject);
    procedure mnuDefineAnredeEntriesClick(Sender: TObject);
    procedure mnuDefineKircheEntriesClick(Sender: TObject);
    procedure mnuDelMarkClick(Sender: TObject);
    procedure mnuDelMarkPersClick(Sender: TObject);
    procedure mnuDruckerrnderanzeigenClick(Sender: TObject);
    procedure mnuEndeClick(Sender: TObject);
    procedure mnuExecSQLBatchClick(Sender: TObject);
    procedure mnuExecSQLSingleClick(Sender: TObject);
    procedure mnuExportClick(Sender: TObject);
    procedure mnuFehlerInKarteiSuchenClick(Sender: TObject);
    procedure mnuFreeListClick(Sender: TObject);
    procedure mnuGDBearbeitenClick(Sender: TObject);
    procedure mnuGDStatistikClick(Sender: TObject);
    procedure mnuGebListClick(Sender: TObject);
    procedure mnuGebListSplitClick(Sender: TObject);
    procedure mnuGemeindeAdresseClick(Sender: TObject);
    procedure mnuGemeindeDefaultClick(Sender: TObject);
    procedure mnuImportClick(Sender: TObject);
    procedure mnuImportGDClick(Sender: TObject);
    procedure mnuIncludeAbgangClick(Sender: TObject);
    procedure mnuInvertMarkClick(Sender: TObject);
    procedure mnuJubiClick(Sender: TObject);
    procedure mnuOpenDebugClick(Sender: TObject);
    procedure mnuPrintDBStruktClick(Sender: TObject);
    procedure mnuPrintFormularClick(Sender: TObject);
    procedure mnuReportDesignerClick(Sender: TObject);
    procedure mnuRBClick(Sender: TObject);
    procedure mnuSelectDatabaseClick(Sender: TObject);
    procedure mnuSelectReportFileClick(Sender: TObject);
    procedure mnuShowOverViewClick(Sender: TObject);
    procedure mnuShowTaufTagClick(Sender: TObject);
    procedure mnuSQLDebugClick(Sender: TObject);
    procedure mnuStatistikClick(Sender: TObject);
    procedure mnuUserDefSQLClick(Sender: TObject);
    procedure mnuVolltextsucheClick(Sender: TObject);
  private
    { private declarations }
    slAusgabe : TStringList;
    procedure PerparePrint(Formular: string; CallDesigner: boolean);
    procedure PerpareDatabaseForPrint(AskForAge: boolean = true);
    procedure CloseDatabaseAfterPrint;
    procedure ShowDruckenDlg;
    procedure ShowDruckenMnu;
    procedure GetGemeindenFromDB;
    Procedure Datensicherung(Auto: boolean);
    function  HeutigeGebTage(datum:TDateTime): String;
    procedure HandleException(Sender: TObject; E: Exception);
  public
    { public declarations }
    slHelp   : TStringlist;
  end; 

var
  frmMain: TfrmMain;

implementation

{$R *.lfm}

uses
  global,
  help,
  input,
  LConvEncoding,  // CP1252ToUTF8
  LazFileUtils,
  inifiles,
  Ausgabe,
  ssl_openssl,   // die dll Files gehen nur bis 1.0.2u. Die Version 1.1.1 geht nicht
  httpsend,
  Statinfo,
  suche,
  freelist,
  Translations,
  PgmUpdate,
  appsettings,
  Clipbrd,
  LCLIntf,       // Openurl
  FileCtrl,      // MinimizeName
  dm,
  gd,
  db;

{ TfrmMain }

procedure TfrmMain.HandleException(Sender: TObject; E: Exception);

var
  CloseAction : TCloseAction = caNone;

  begin
  LogAndShowError('Unbehandelter Fehler.'+#13+
                  'Nachricht: '+e.Message+#13#13+
                  'Das Programm wird beendet.');

  frmDM.dbStatus(false); // DB schliessen

  FormClose(sender, CloseAction);  //Lock entfernen
  
  FlushDebug;
  
  Halt;                  // End of program execution
end;

function TfrmMain.HeutigeGebTage(datum:TDateTime): String;

var
  sWhere : String;

begin
  frmDM.dbStatus(true);
  result := '';
  sWhere := '';
  sWhere := SQL_Where_Add(sWhere, SQL_Where_IsNull('Ueberwiesen_nach_Datum'));
  sWhere := SQL_Where_Add(sWhere, SQL_Where_IsNull('AustrittsDatum'));
  sWhere := SQL_Where_Add(sWhere, SQL_Where_IsNull('AusschlussDatum'));
  sWhere := SQL_Where_Add(sWhere, SQL_Where_IsNull('TodesDatum'));
  sWhere := SQL_Where_Add(sWhere, SQL_Where_IsNull('UebertrittsAbDatum'));

  frmDM.dsetHelp.sql.Clear;
  frmDM.dsetHelp.sql.add('select vorname, Nachname, Geburtstag, TelPrivat, TaufDatum from '+global.sPersTablename);
  frmDM.dsetHelp.sql.add('where ((strftime(''%m'',geburtstag) = '''+FormatDateTime('mm', datum)+''' and strftime(''%d'',geburtstag) = '''+FormatDateTime('dd', datum)+''')  or');
  frmDM.dsetHelp.sql.add('       (strftime(''%m'',TaufDatum)  = '''+FormatDateTime('mm', datum)+''' and strftime(''%d'',TaufDatum)  = '''+FormatDateTime('dd', datum)+''')) and');
  frmDM.dsetHelp.sql.add(sWhere);
  frmDM.dsetHelp.sql.add('order by nachname, vorname');

  frmDM.dsetHelp.open;
  while not frmDM.dsetHelp.eof do
    begin
      if FormatDateTime('ddmm', datum) = FormatDateTime('ddmm', frmDM.dsetHelp.fieldByName('Geburtstag').asdateTime)
        then
          result := result +
                    DateToStr(frmDM.dsetHelp.fieldByName('Geburtstag').asdateTime)+' '+
                    AppendChar(frmDM.dsetHelp.fieldByName('vorname').asstring+' '+frmDM.dsetHelp.fieldByName('Nachname').asstring, ' ', 18)+' '+
                    inttostr(AgeAtDate(frmDM.dsetHelp.fieldByName('Geburtstag').asdateTime,datum))+', '+
                    frmDM.dsetHelp.fieldByName('TelPrivat').asstring+#13#10;
      if mnuShowTaufTag.Checked and
         (FormatDateTime('ddmm', datum) = FormatDateTime('ddmm', frmDM.dsetHelp.fieldByName('TaufDatum').asdateTime))
        then
          result := result +
                    DateToStr(frmDM.dsetHelp.fieldByName('TaufDatum').asdateTime)+' '+
                    AppendChar(frmDM.dsetHelp.fieldByName('vorname').asstring+' '+frmDM.dsetHelp.fieldByName('Nachname').asstring, ' ', 18)+' '+
                    inttostr(AgeAtDate(frmDM.dsetHelp.fieldByName('TaufDatum').asdateTime,datum))+'ter Tauftag, '+
                    frmDM.dsetHelp.fieldByName('TelPrivat').asstring+#13#10;
      frmDM.dsetHelp.next;
    end;
  frmDM.dsetHelp.close;
  frmDM.dbStatus(false);  //Datenbank schliessen
end;

procedure TfrmMain.FormCreate(Sender: TObject);

var
  HTTP     : THTTPSend;
  sRelease : String;
  INI      : TINIFile;

begin
  Application.OnException    := @HandleException;
  slHelp                     := TStringlist.Create;
  slAusgabe                  := TStringList.create;
  sAppDir                    := vConfigurations.MyDirectory;
  sIniFile                   := vConfigurations.ConfigFile;
  sSavePath                  := sAppDir+'Sicherung';
  sSavePath                  := help.ReadIniVal(sIniFile, 'Sicherung', 'Verzeichnis', sAppDir+'Sicherung\', true);
  sPrintPath                 := help.ReadIniVal(sIniFile, 'Ausgaben' , 'Verzeichnis', sAppDir+'Ausgaben\', true);
  sDatabase                  := help.ReadIniVal(sIniFile, 'Datenbank', 'Datei'      , sAppDir+'ge_kart7.db', true);
  sDatabaseLock              := ChangeFileExt(sDatabase, '.lock');
  sDebugFile                 := help.ReadIniVal(sIniFile, 'Debug', 'Name', sAppDir+'debug\debug.txt', true);
  bDebug                     := 'TRUE' = Uppercase(help.ReadIniVal(sIniFile, 'Debug', 'Debug', 'true', true));
  bSQLDebug                  := 'TRUE' = Uppercase(help.ReadIniVal(sIniFile, 'Debug', 'SQLDebug', 'true', true));
  mnuSQLDebug.Checked        := bSQLDebug;
  mnuDebug.Checked           := bDebug;
  sUserAndPCName             := replacechar(GetComputerName, ' ', '_')+';'+replacechar(GetUserName, ' ', '_');    //PC und User Name ermitteln
  frmMain.caption            := 'GE_Kart '+GetProductVersionString;
  labVersion.Caption         := labVersion.Caption + ' ' + frmMain.caption;
  bDatabaseVersionChecked    := false;
  labMyMail.Caption          := global.sEMailAdr;
  labMyWeb.Caption           := global.sHomePage;
  sReportFile                := help.ReadIniVal(sIniFile, 'Reports','Karteikarte',sAppDir+'ReportsFertig\KarteiKarte.lrf', true);
  mnuShowOverView.Checked    := ('1' = help.ReadIniVal(sIniFile, 'Defaults','ShowOverView', '1', true));
  mnuColoredEditMode.Checked := ('1' = help.ReadIniVal(sIniFile, 'Defaults','ColoredEditMode', '1', true));
  mnuShowTaufTag.Checked     := ('1' = help.ReadIniVal(sIniFile, 'Defaults','ShowTaufTag', '1', true));
  mnuGebListSplit.Checked    := ('1' = help.ReadIniVal(sIniFile, 'Defaults','GebListSplit', '1', true));
  mnuUserDefSQL.Visible      := FileExists(UTF8ToSys(sUserDefSQLFile));

  if mnuUserDefSQL.Visible
    then
      begin
        try
          slHelp.LoadFromFile(sUserDefSQLFile);
        except
          slHelp.Text:='';
        end;
        if slHelp.Text <> '' then mnuUserDefSQL.Hint:=slHelp.Text;
      end;

  sPrintPath  := IncludeTrailingPathDelimiter(sPrintPath);
  ForceDirectories(UTF8ToSys(sPrintPath));
  ForceDirectories(UTF8ToSys(ExtractFilePath(sDebugFile)));

  myDebugLN('Starte GE_Kart '+labVersion.Caption);
  myDebugLN('AppDir    : '+sAppDir);
  myDebugLN('sIniFile  : '+sIniFile);
  myDebugLN('sSavePath : '+sSavePath);
  myDebugLN('sPrintPath: '+sPrintPath);
  myDebugLN('Database  : '+sDatabase);

  //Dialoge übersetzen
  TranslateUnitResourceStrings('LCLStrConsts','lclstrconsts.%s.po','de','en');

  try
    imgSELK.Picture.LoadFromFile(sAppDir+'\selk_ohne.png');
  except

  end;

  //Laden der Kircheneinträge
  slHelp.Text       := sKirchenEintraegeDef;
  slHelp.CommaText  := help.ReadIniVal(sIniFile, 'Defaults','Kirchen', slHelp.CommaText, true);
  sKirchenEintraege := slHelp.Text;
  slHelp.Clear;

  //Laden der Anredeneinträge
  slHelp.Text       := sAnredeEintraegeDef;
  slHelp.CommaText  := help.ReadIniVal(sIniFile, 'Defaults','Anreden', slHelp.CommaText, true);
  sAnredenEintraege := slHelp.Text;
  slHelp.Clear;

  sDefaultGemeinde  := help.ReadIniVal(sIniFile, 'Defaults','Gemeinde', '', true);

  //Skalierung ermitteln für ALLE Fenster
  ScaleFactor  := 1;
  myDebugLN('Screen.PixelsPerInch: '+inttostr(Screen.PixelsPerInch)+', Designwert: '+inttostr(nDefDPI));

  WORKAREA      := GETMaxWindowsSize; //Verfügbarer Platz
  ScaleFactorX  := 1;
  ScaleFactorY  := 1;

  if WORKAREA.Right  < MaxWindWidth  then ScaleFactorX := WORKAREA.Right  / MaxWindWidth;
  if WORKAREA.Bottom < MaxWindHeight then ScaleFactorY := WORKAREA.Bottom / MaxWindHeight;

  if ScaleFactorX < ScaleFactor then ScaleFactor := ScaleFactorX;
  if ScaleFactorY < ScaleFactor then ScaleFactor := ScaleFactorY;

  myDebugLN('ScaleFactorX '+floattostr(ScaleFactorX));
  myDebugLN('ScaleFactorY '+floattostr(ScaleFactorY));
  myDebugLN('ScaleFactor  '+floattostr(ScaleFactor));

  //Prüfung auf neue Version
  HTTP := THTTPSend.Create;
  try
    if not HTTP.HTTPMethod('GET', 'https://'+sHomePage+'/GE_KART/version.txt?;V'+GetProductVersionString+';'+sUserAndPCName)
      then
        begin
	  myDebugLN('ERROR HTTPGET, Resultcode: '+inttostr(Http.Resultcode));
          labVersionNeu.Font.Size := -10;
          labVersionNeu.Caption   := 'Keine Verbindung zum Updateserver';
          labVersionNeu.Color     := clSkyBlue;
          labVersionNeu.OnClick   := nil;
        end
      else
        begin
          slHelp.loadfromstream(Http.Document);
          slHelp.Text := CP1252ToUTF8(slHelp.Text);
          myDebugLN('HTTPGET, Resultcode: '+inttostr(Http.Resultcode)+' '+Http.Resultstring);
          myDebugLN('Http.headers.text  : '+#13#10+Http.headers.text);
          myDebugLN('Http.Document      : '+#13#10+slHelp.Text);

          if Http.Resultcode = 200
            then
              begin
                sNewVers := slHelp.Strings[0];
                if GetProductVersionString < sNewVers
                  then
                    begin
                      slHelp.Delete(0);
                      labVersionNeu.Caption := labVersionNeu.Caption+sNewVers;
                      labVersionNeu.Hint    := slHelp.Text;
                      labVersionNeu.Cursor  := crHandPoint;
                    end
                  else
                    begin
                      labVersionNeu.Font.Size := -10;
                      labVersionNeu.OnClick   := nil;
                      if GetProductVersionString = sNewVers
                        then
                          begin
                            labVersionNeu.Caption   := 'Das Programm ist aktuell';
                            labVersionNeu.Color     := clMoneyGreen;
                          end
                        else
                          begin
                            labVersionNeu.Caption   := 'Das Programm ist neuer. Aktuelle Version: '+sNewVers;
                            labVersionNeu.Color     := clAqua;
                          end;
                    end;
              end
            else
              begin
                labVersionNeu.Caption := 'Fehler bei der Abfrage'
              end;
       end;
  finally
    HTTP.Free;
  end;

  HTTP := THTTPSend.Create;
  try
    if HTTP.HTTPMethod('GET', 'https://'+sHomePage+'/GE_KART/aktuelles.txt')
      then
        begin
          slHelp.loadfromstream(Http.Document);
          sAktuelles := CP1252ToUTF8(slHelp.Text);
          myDebugLN('HTTPGET, Resultcode: '+inttostr(Http.Resultcode)+' '+Http.Resultstring);
          myDebugLN('Http.headers.text  : '+#13#10+Http.headers.text);
          myDebugLN('Http.Document      : '+#13#10+sAktuelles);

          if (Http.Resultcode = 200) and (sAktuelles <> '')
            then
              begin
                btnAktuelles.Visible := true;
                btnAktuelles.Caption := 'Aktuelles vom '+slHelp.Strings[0];
              end;
        end;
  finally
    HTTP.Free;
    slHelp.Clear;
  end;

  //Datensicherung
  try
    dtLastSave := StrToDate(help.ReadIniVal(sIniFile, 'Sicherung', 'Datum', '01.01.1980', true));
  except
    dtLastSave := StrToDate('01.01.1980');
  end;
  
  if dtLastSave = StrToDate('01.01.1980')
    then StatusBar.Panels[1].Text := 'Noch keine Datensicherung gemacht'
    else StatusBar.Panels[1].Text := 'Letzte Datensicherung: '+DateToStr(dtLastSave);
	
  randomize;
  if (dtLastSave = StrToDate('01.01.1980')) or
     (dtLastSave+30 < now)
    then
      Datensicherung(true)  //Automatische Datensicherung
    else
      if (trunc(random(12)) = 1) then Showmessage('Denken Sie an eine regelmäßige Datensicherung'+#13+StatusBar.Panels[1].Text);

  //Releasenote anzeigen
  sRelease := help.ReadIniVal(sIniFile, 'Programm', 'Version', '0', true);
  if GetProductVersionString <> sRelease
    then
      begin
        try
          slHelp.LoadFromFile('releasenote_'+GetProductVersionString+'.txt');
        except
          slHelp.Text:='';
        end;
        if slHelp.Text <> ''
          then
            if MessageDlg(slHelp.Text, mtConfirmation, [mbYes, mbNo],0) = mrNo
              then help.WriteIniVal(sIniFile, 'Programm', 'Version', GetProductVersionString);
        slHelp.Clear;
      end;

  //Clean up old entrys
  INI := TINIFile.Create(UTF8ToSys(sIniFile));
  INI.EraseSection('Gemeinde');
  INI.Free;
end;

procedure tfrmMain.ShowDruckenDlg;

var weiter : boolean;
    i      : integer;

begin
  screen.cursor := crhourglass;
  frmAusgabe.Memo.clear;
  try
    frmAusgabe.Memo.Text := slAusgabe.Text;
  except
    frmAusgabe.Memo.Clear;
    frmAusgabe.Memo.lines.Add('Achtung');
    frmAusgabe.Memo.lines.Add('Aus technischen Gründen können nicht alle Daten angezeigt werden');
    frmAusgabe.Memo.lines.Add('----------------------------------------------------------------');
    frmAusgabe.Memo.lines.Add('');
    //Jetzt laden, was möglich ist
    weiter := true;
    i      := 0;
    repeat
      try
        frmAusgabe.Memo.lines.Add(slAusgabe.Strings[i]);
        inc(i);
      except
        Weiter := false;
      end;
    until (not Weiter) or (i = slAusgabe.Count);
  end;
  screen.cursor := crdefault;
  frmAusgabe.btnOK.Visible:=false;
  frmAusgabe.btnClose.Caption:='Schliessen';
  frmAusgabe.showmodal;
end;

procedure TfrmMain.FormDestroy(Sender: TObject);

begin
  slAusgabe.Free;
  slHelp.Free;
  myDebugLN('Beende GE_Kart');
  myDebugLN('************************************************************************************');
  FlushDebug;
end;

procedure TfrmMain.FormShow(Sender: TObject);

var
  sHelp  : string;
  bHelp  : boolean;

begin
  slHelp.Clear;
  bHelp  := (ParamCount = 2) and (lowercase(ParamStr(2)) = 'popup');

  if FileExists(UTF8ToSys(sDatabase)) then
    begin
      myDebugLN('FormShow: '+sDatabase);
      if FileExists(UTF8ToSys(sDatabaseLock)) then
        begin
           slHelp.LoadFromFile(sDatabaseLock);
           MessageDlg('Die Datenbank wird von '+RemoveLastCRLF(slHelp.Text)+' benutzt.'+#13+
           'Das Programm wird beendet.'+#13#13+
           'Wenn sie sicher sind, das sie der einzige Benutzer sind, können Sie die Datei'+#13+
           sDatabaseLock+#13+
           'löschen' , mtWarning, [mbOK],0);
           slHelp.Clear;
           close;  //??? blitzt noch kurz auf
        end
      else
        begin
          slHelp.Text := sUserAndPCName;
          slHelp.SaveToFile(sDatabaseLock);
          slHelp.Clear;
          frmDM.SetDBPath({UTF8ToSys}(sDatabase));
          Statusbar.Hint           := 'Database : '+sDatabase;
          Statusbar.Panels[0].Text := MinimizeName(sDatabase, Statusbar.Canvas, Statusbar.Panels[0].Width);

          frmDM.dbStatus(true);   //Damit wird die Datenbankstruktur geprüft und ggf. angepasst
          frmDM.dbStatus(false);  //DB schliessen

          memGeburtstag.Clear;
          sHelp := HeutigeGebTage(now-2);
          if sHelp <> '' then memGeburtstag.Text := memGeburtstag.Text+'Vorgestern:'+#13#10+sHelp+#13#10;
          sHelp := HeutigeGebTage(now-1);
          if sHelp <> '' then memGeburtstag.Text := memGeburtstag.Text+'Gestern:'+#13#10+sHelp+#13#10;
          sHelp := HeutigeGebTage(now);
          if sHelp <> '' then memGeburtstag.Text := memGeburtstag.Text+'Heute:'+#13#10+sHelp+#13#10;
          sHelp := HeutigeGebTage(now+1);
          if sHelp <> '' then memGeburtstag.Text := memGeburtstag.Text+'Morgen:'+#13#10+sHelp+#13#10;
          sHelp := HeutigeGebTage(now+2);
          if sHelp <> '' then memGeburtstag.Text := memGeburtstag.Text+'Übermorgen:'+#13#10+sHelp;

          memGeburtstag.visible := (memGeburtstag.Text <> '');

          if bHelp
            then
              begin
                if memGeburtstag.Text <> '' then MessageDlg(memGeburtstag.Text, mtInformation, [mbOK],0);
                close;  //??? blitzt noch kurz auf
              end;

          frmDM.dbStatus(true);

          //Sortierreihenfolge Karteikarte vorbereiten
          frmDM.ExecSQL('Update PERSONEN SET TempString='+SQL_UTF8UmlautReplace('Nachname')+'||'' ''||'+SQL_UTF8UmlautReplace('Vorname'));
          //DB bereinigen
          frmDM.ExecSQL(sSQL_ClearMark1);
          frmDM.ExecSQL(sSQL_ClearMark2);
          frmDM.ExecSQL(sSQL_ClearAbgang1);
          frmDM.ExecSQL(sSQL_ClearAbgang2);
          //NULL Werte aus Gemeinde entfernen
          frmDM.ExecSQL('Update PERSONEN SET Gemeinde=''''               where '+ SQL_Where_IsNull('Gemeinde'));
          frmDM.ExecSQL('Update PERSONEN SET Ueberwiesen_nach_Datum='''' where Ueberwiesen_nach_Datum=="30.12.1899"');
          frmDM.ExecSQL('Update PERSONEN SET AustrittsDatum=''''         where AustrittsDatum="30.12.1899"');
          frmDM.ExecSQL('Update PERSONEN SET AusschlussDatum=''''        where AusschlussDatum="30.12.1899"');
          frmDM.ExecSQL('Update PERSONEN SET TodesDatum=''''             where TodesDatum="30.12.1899"');
          frmDM.ExecSQL('Update PERSONEN SET UebertrittsAbDatum=''''     where UebertrittsAbDatum="30.12.1899"');

          frmDM.dsetHelp.SQL.Text:='Select * from Gemeinde';
          frmDM.dsetHelp.Open;
          labGemeinde1.caption := frmDM.dsetHelp.FieldByName('Adr1').AsString;
          labGemeinde2.caption := frmDM.dsetHelp.FieldByName('Adr2').AsString;
          frmDM.dsetHelp.close;
          frmDM.dbStatus(false); // DB schliessen

          GetGemeindenFromDB;
        end;
    end
  else
    begin
      StatusBar.Panels[0].Text := 'Datenbankdatei unter "'+sDatabase+'" nicht gefunden.';
      ShowMessage(StatusBar.Panels[0].Text+#13'Bitte über Datei-Datenbank-Auswählen eine gültige Datenbank auswählen.!');
      frmDM.SetDBPath('');
    end;
end;

procedure TfrmMain.GetGemeindenFromDB;
begin
  //Suchen der in der Datenbank verwendeten Gemeinden
  slHelp.Clear;
  slHelp.Add(sGemeindenAlle);
  frmDM.dsetHelp2.sql.Clear;
  frmDM.dsetHelp2.sql.add('select Gemeinde from personen group by gemeinde order by gemeinde');
  frmDM.dsetHelp2.open;
  while not frmDM.dsetHelp2.eof do
    begin
      slHelp.Add(frmDM.dsetHelp2.fieldByName('Gemeinde').asstring);
      frmDM.dsetHelp2.Next;
    end;
  frmDM.dsetHelp2.Close;
  sGemeinden := slHelp.Text;
end;

procedure TfrmMain.frReportGetValue(const ParName: String; var ParValue: Variant);
begin
  //myDebugLN('LazReport ask for:'+ParName);
  if mnuAbsToEti.Checked
    then
      begin
        if (ParName = 'Adr1') then ParValue := labGemeinde1.caption;
        if (ParName = 'Adr2') then ParValue := labGemeinde2.caption;
      end
    else
      begin
        ParValue := '';
      end;
  if (ParName = 'Kennzeichnung')
    then
      begin
        if mnuBuechersendung.Checked    then ParValue := 'Büchersendung';
        if mnuDruckWarensendung.Checked then ParValue := 'Warensendung';
        if mnuDruckNix.Checked          then ParValue := '';
      end;
  if (ParName = 'Adr1x') then ParValue := labGemeinde1.caption;
  if (ParName = 'Adr2x') then ParValue := labGemeinde2.caption;
end;

procedure TfrmMain.imgSELKClick(Sender: TObject);
begin
  Openurl('www.selk.de');
end;

procedure TfrmMain.btnKarteiClick(Sender: TObject);
begin
  GetGemeindenFromDB;
  kartei.frmKartei.ShowModal;
end;

procedure TfrmMain.btnAktuellesClick(Sender: TObject);
begin
  frmAusgabe.SetDefaults(btnAktuelles.Caption, sAktuelles, '', '', 'Schliessen', false);
  frmAusgabe.ShowModal;
end;

procedure TfrmMain.btnSuchenClick(Sender: TObject);
begin
  frmSuche.showmodal;
end;

procedure TfrmMain.btnSuchenContextPopup(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
begin
  mnuVolltextSucheClick(Sender);
end;

procedure TfrmMain.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  if FileExists(UTF8ToSys(sDatabaseLock)) then
    begin
      slHelp.LoadFromFile(sDatabaseLock);
      if (RemoveLastCRLF(slHelp.Text) = sUserAndPCName) then SysUtils.DeleteFile(UTF8ToSys(sDatabaseLock));
    end;
end;

procedure TfrmMain.labMyMailClick(Sender: TObject);
begin
  Openurl('MailTo:'+sEMailAdr+'?subject=GE_KART');
end;

procedure TfrmMain.labMyWebClick(Sender: TObject);
begin
  Openurl('https://'+sHomePage+'/GE_Kart.html');
end;

procedure TfrmMain.labVersionNeuClick(Sender: TObject);

begin
  frmPgmUpdate.URL      := 'https://'+sHomePage+'/GE_KART/v'+sNewVers+'.zip';
  frmPgmUpdate.FileName := sAppDir+'v'+sNewVers+'.zip';
  frmPgmUpdate.showmodal;
end;

procedure TfrmMain.mnuEMailVersandClick(Sender: TObject);

var
  sEMail: String;

begin
  PerpareDatabaseForPrint(false);
  sEMail := '';
  while not frmDM.dsetHelp.eof do
    begin
      if frmDM.dsetHelp.fieldByName('eMail').asstring <> ''  then sEMail := sEMail + Trim(LowerCase(frmDM.dsetHelp.fieldByName('eMail').asstring))+';';
      if frmDM.dsetHelp.fieldByName('email2').asstring <> '' then sEMail := sEMail + Trim(LowerCase(frmDM.dsetHelp.fieldByName('email2').asstring))+';';
      if frmDM.dsetHelp.fieldByName('email3').asstring <> '' then sEMail := sEMail + Trim(LowerCase(frmDM.dsetHelp.fieldByName('email3').asstring))+';';
      frmDM.dsetHelp.next;
    end;

  if sEMail <> ''
    then
      begin
        Clipboard.AsText := sEMail;
        Showmessage('Die EMail-Adressen wurden in die Zwischenablage kopiert');
      end
    else
      begin
        Showmessage('Keine EMail-Adressen gefunden');
      end;
  CloseDatabaseAfterPrint;
end;

procedure TfrmMain.MenuSetAnredeClick(Sender: TObject);

begin
  frmDM.ExecSQL('Update personen set Briefanrede=''Herrn'' where '+SQL_Where_IsNull('Briefanrede')+' and (Geschlecht=''M'')');
  frmDM.ExecSQL('Update personen set Briefanrede=''Frau'' where '+SQL_Where_IsNull('Briefanrede')+' and (Geschlecht=''W'')');
end;

procedure TfrmMain.mnuAbsToEtiClick(Sender: TObject);
begin
  mnuAbsToEti.Checked:= not mnuAbsToEti.Checked;
  ShowDruckenMnu;
end;

procedure TfrmMain.mnuChangePWClick(Sender: TObject);
begin
  frmInput.SetDefaults('Altes Passwort eingeben', 'Passwort', '', '', '', true);
  if frmInput.ShowModal = mrOK
    then
      begin
        //myDebugLN('Codiertes Passwort '+frmInput.Edit1.Text+' = 'help.verschluessele(frmInput.Edit1.Text));
        if help.verschluessele(frmInput.Edit1.Text) = sPW
          then
            begin
              frmInput.SetDefaults('Neues Passwort, nichts eingeben für keine Passwortabfrage', 'Passwort', '', 'Wiederholung', '', true);
              if (frmInput.ShowModal = mrOK) and (frmInput.Edit1.Text = frmInput.Edit2.Text)
                then
                  begin
                    help.WriteIniVal(sIniFile,'Passwort', 'Passwort', help.verschluessele(frmInput.Edit1.Text));
                    sPW := frmInput.Edit1.Text;
                  end
                else
                  begin
                    ShowMessage('Abbruch oder Passwörter verschieden');
                  end;
            end
          else
            begin
              ShowMessage('Passwort ist falsch');
            end;
      end;
end;

procedure TfrmMain.mnuColoredEditModeClick(Sender: TObject);
begin
  mnuColoredEditMode.Checked := not mnuColoredEditMode.Checked;
  if mnuColoredEditMode.Checked
    then help.WriteIniVal(sIniFile, 'Defaults', 'ColoredEditMode', '1')
    else help.WriteIniVal(sIniFile, 'Defaults', 'ColoredEditMode', '0');
end;

procedure TfrmMain.mnuCSVImportClick(Sender: TObject);

var Line       : String;
    Item       : String;
    ErrorText  : string;
    FeldNamen  : String;
    sSQL       : string;
    i          : integer;
    ActLine    : Integer;
    Lines      : integer;
    ItemNo     : integer;

begin
  if MessageDlg('Sie müssen jetzt eine CSV (comma seperated value) Datei auswählen.'#13+
             'Als Trennzeichen sind "'+CSV_Delimiter+'" und TAB erlaubt'#13+
             'Format der Datei:'#13#13+
             '1.Zeile: Feldnamen. Die gültigen Feldnamen können Sie ermitteln mit einen Datenexport oder mit einen Ausdruck der Struktur'#13#13+
             'Folgende Zeilen: Daten. Die Daten dürfen kein "'+CSV_Delimiter+'" und TAB enthalten'#13#13+
             'Alle " als Feldbegrenzer werden gelöscht'#13+
             'Die Tabelle darf kein Feld PersonenID, Markiert oder Abgang enthalten'#13#13+
             'Beispiel:'#13+
             'NAME;STRASSE;LAND;PLZ;ORT;'#13+
             'Adam;Bergstraße 9;;12345;Musterstadt;'#13+
             'Brecht;"Sample street";GB;98765;London;'#13#13+
             'Sicherheitsfunktion: Zum Fortfahren "Wiederholen" drücken!', mtConfirmation, [mbYes, mbRetry, mbNo],0) = mrRetry
    then
      begin
        Lines     := 0;
        ItemNo    := 1;
        ErrorText := '';
        slHelp.Clear;
        try
          openDialog.Title       := 'Selektiere eine CSV Import Datei';
          openDialog.InitialDir  := UTF8ToSys(sAppDir);          // Set up the starting directory to be the current one
          openDialog.Options     := [ofFileMustExist];           // Only allow existing files to be selected
          openDialog.Filter      := 'Alle|*.*|CSV-Datei |*.csv';
          openDialog.FilterIndex := 2;                           // Select report files as the starting filter type

          // Display the open file dialog
          if openDialog.Execute then
            begin
              myDebugLN('Import: '+ OpenDialog.Filename);

              slHelp.loadfromfile(UTF8ToSys(OpenDialog.Filename));
              slHelp.text := RemoveBOM(slHelp.text);

              FeldNamen := 'INSERT INTO '+sPersTablename+  ' (';
              //Header lesen
              if slHelp.count > 1
                then
                  begin
                    Line := slHelp.strings[0];
                    repeat
                      Item := UPPERCASE(GetCSVRecordItem(ItemNo, Line, [CSV_Delimiter, #9], '"'));

                      if (Item = 'PERSONENID') or
                         (Item = 'MARKIERT') or
                         (Item = 'ABGANG')
                        then Raise Exception.CreateFmt('Ungültiger Spaltenname "%s" in der Tabelle. Die Funktion wird abgebrochen.', [Item]);

                      if Item <> '' then
                        begin
                          FeldNamen := FeldNamen + Item +  ', ';
                          inc(ItemNo);
                        end;
                    until Item = '';
                  end;
              FeldNamen := FeldNamen + ' Markiert, Abgang) VALUES (';
              myDebugLN(FeldNamen);

              //Daten einfügen
              if slHelp.count > 1 then
                for ActLine := 1 to slHelp.count-1 do
                  begin
                    Line := slHelp.strings[ActLine];
                    Line := Trim(Line);
                    if Line <> ''
                      then
                        begin
                          sSQL := FeldNamen;
                          for i := 1 to ItemNo-1 do
                            begin
                              Item := GetCSVRecordItem(i, Line, [CSV_Delimiter, #9], '"');
                              //Auf Datum prüfen
                              if (Item <> '') and IsDateFormat(Item, DefaultFormatSettings.DateSeparator)
                                then Item := FormatDateTime('yyyy-mm-dd', StrToDate(Item));
                              sSQL := sSQL + '''' + Item + ''', ';
                            end;
                          sSQL := sSQL + '0, 0)';     //Markiert und Abgang auf false setzen
                          try
                            frmDM.ExecSQL(sSQL, true);
                            inc(Lines);
                          except
                            on E: Exception
                              do
                                begin
                                  ErrorText := ErrorText + 'In Zeile: '+inttostr(ActLine)+
                                               ' ist folgender Fehler aufgetreten: '+e.Message+'. Die Zeile wird NICHT importiert!'#13;
                                  break;
                                end;
                          end;
                        end;
                  end;
              if ErrorText <> '' then LogAndShowError(ErrorText);
              MessageDlg(Inttostr(Lines)+' Zeilen importiert', mtInformation, [mbOK],0);
            end;
        except
          on E: Exception do LogAndShowError(E.Message);
        end;
      end
    else
      Showmessage('Funktion abgebrochen.');
end;

procedure TfrmMain.mnuDatenRestoreClick(Sender: TObject);

Var
  sSelected,
  sOldDB,
  sOldDBnewName,
  sNewDBName : string;

begin
  //Showmessage('Zum Zurücklesen einer Datenbank beenden Sie bitte GE_kart. Dann kopieren Sie ihre Datensicherung in das GE_Kart Verzeichnis. Dort löschen Sie die Datei "GE_KART7.db". Als letztes benennen sie ihre Datensicherung in "GE_KART7.db" um.');
  openDialog.Title       := 'Selektiere eine Datensicherung';
  openDialog.InitialDir  := UTF8ToSys(sAppDir+'Sicherung\');
  openDialog.Options     := [ofFileMustExist];
  openDialog.Filter      := 'Alle|*.*|DB-Datei |*.db';
  openDialog.FilterIndex := 2;
  if openDialog.Execute
    then
      begin
        sSelected     := openDialog.FileName;
        sOldDB        := sDatabase;
        sOldDBnewName := sDatabase+'.bak';
        sNewDBName    := sAppDir+'GE_Kart7.db';
        if MessageDlg('Die Datei '+sOldDB+' wird in '+sOldDBnewName+' umbenannt.'#13+
                      'Die Datei '+sSelected+' wird nach '+sNewDBName+' kopiert.'#13#13+
                      'Sicherheitsfunktion: Zum Fortfahren "Wiederholen" drücken!', mtConfirmation, [mbYes, mbRetry, mbNo],0) = mrRetry
          then
            begin
              frmDM.dbStatus(false); //schliesse DB

              if RenameFileUTF8(sOldDB, sOldDBnewName)
                then if fileutil.CopyFile(sSelected, sNewDBName, true)
                  then
                    begin
                      sDatabase := sNewDBName;
                      myDebugLN('New Database : '+sDatabase);
                      bDatabaseVersionChecked := false;
                      FormShow(sender);                               // Prüft die Datei und stellt das Main Formular richtig dar.
                      help.WriteIniVal(sIniFile, 'Datenbank','Datei', sDatabase);
                    end
                  else LogAndShowError(SysToUTF8(SysErrorMessage(GetLastOSError())))
                else LogAndShowError(SysToUTF8(SysErrorMessage(GetLastOSError())));
            end
          else
            Showmessage('Funktion abgebrochen.');
      end;
end;

Procedure TfrmMain.Datensicherung(Auto: boolean);

var
  shelp,
  sDest : String;

begin
  shelp := 'ge_kart_'+FormatDateTime('yyyymmdd_hhnnss', now())+'.db';
  sDest := '';

  if Auto
    then sDest := sAppDir+'Sicherung\'+shelp
    else
      begin
        frmDM.dbStatus(false); //schliesse DB
        ForceDirectories(UTF8ToSys(sSavePath));
        SaveDialog.InitialDir := sSavePath;
        SaveDialog.FileName   := shelp;
        if SaveDialog.Execute
          then sDest := SaveDialog.FileName;
      end;

  if sDest <> ''
    then
      begin
        myDebugLN('Savesrc:  '+sDatabase);
        myDebugLN('Savedest: '+sDest);

        sSavePath := extractFilePath(sDest);
        ForceDirectories(UTF8ToSys(sSavePath));

        if fileutil.CopyFile(sDatabase, sDest, true)  //Erwartet UTF8 Strings
          then
            begin
              dtLastSave := now;
              help.WriteIniVal(sIniFile, 'Sicherung', 'Datum', DateToStr(dtLastSave));
              StatusBar.Panels[1].Text := 'Letzte Datensicherung: '+DateToStr(dtLastSave);
              help.WriteIniVal(sIniFile, 'Sicherung', 'Verzeichnis', sSavePath);
            end
          else
            LogAndShowError(SysToUTF8(SysErrorMessage(GetLastOSError())));
      end;
end;

procedure TfrmMain.mnuDatensicherungClick(Sender: TObject);

begin
  Datensicherung(false);
end;

procedure TfrmMain.mnuDebugClick(Sender: TObject);
begin
  bDebug              := not bDebug;
  mnuSQLDebug.Checked := bDebug;
  if bDebug
    then help.WriteIniVal(sIniFile, 'Debug', 'Debug', 'true')
    else help.WriteIniVal(sIniFile, 'Debug', 'Debug', 'false');
end;

procedure TfrmMain.mnuDefineAnredeEntriesClick(Sender: TObject);
begin
  frmAusgabe.SetDefaults('Anredeneinträge', sAnredenEintraege, '', 'Speichern', 'Abbrechen', true);
  if frmAusgabe.ShowModal = mrOK
    then
      begin
        help.WriteIniVal(sIniFile, 'Defaults','Anreden', frmAusgabe.Memo.Lines.CommaText);
        Showmessage('Zum Übernehmen der neuen Einträge bitte GE_Kart neu starten.')
      end;
end;

procedure TfrmMain.mnuDefineKircheEntriesClick(Sender: TObject);
begin
  frmAusgabe.SetDefaults('Kircheneinträge', sKirchenEintraege, '', 'Speichern', 'Abbrechen', true);
  if frmAusgabe.ShowModal = mrOK
    then
      begin
        help.WriteIniVal(sIniFile, 'Defaults','Kirchen', frmAusgabe.Memo.Lines.CommaText);
        Showmessage('Zum Übernehmen der neuen Einträge bitte GE_Kart neu starten.')
      end;
end;

procedure TfrmMain.mnuDelMarkClick(Sender: TObject);
begin
  frmDM.ExecSQL(global.sSQL_DelMark);
end;

procedure TfrmMain.mnuDelMarkPersClick(Sender: TObject);

var
  sNamen : string;

begin
  //Datenbank öffnen
  frmDM.dbStatus(true);

  //DB bereinigen
  frmDM.ExecSQL(sSQL_ClearMark1);
  frmDM.ExecSQL(sSQL_ClearMark2);

  frmDM.dsetHelp.sql.Text := 'select vorname, Nachname from '+global.sPersTablename+' where markiert <> ''0'' order by nachname, vorname';
  frmDM.dsetHelp.open;

  sNamen := '';
  while not frmDM.dsetHelp.eof do
    begin
      sNamen := sNamen + frmDM.dsetHelp.fieldByName('vorname').asstring+' '+frmDM.dsetHelp.fieldByName('Nachname').asstring+#13;
      frmDM.dsetHelp.next;
    end;
  frmDM.dsetHelp.close;
  frmDM.dbStatus(false);  //Datenbank schliessen

  if sNamen <> ''
    then
      begin
        if MessageDlg('Folgene Personen wurden gefunden'+#13+
                      sNamen+#13+
                      'Sollen Sie gelöscht werden?'+#13#13+
                      'Sicherheitsfunktion: Zum Fortfahren "Wiederholen" drücken!', mtConfirmation, [mbYes, mbRetry, mbNo],0) = mrRetry
          then
            begin
              frmDM.ExecSQL('delete from personen where markiert <> ''0''');
              frmDM.ExecSQL('VACUUM');
            end;
      end
    else
      begin
        Showmessage('Keine Personen gefunden');
      end;
end;

procedure TfrmMain.mnuDruckerrnderanzeigenClick(Sender: TObject);

begin
  GetPrinterMargins();
end;

procedure TfrmMain.mnuEndeClick(Sender: TObject);
begin
  close;
end;

procedure TfrmMain.mnuExecSQLBatchClick(Sender: TObject);

var
  i    : integer;
  rows : longint;
  sHelp: String;

begin
  frmAusgabe.SetDefaults('SQL Kommandos ausführen', '--Hier SQL Kommandos eingeben'+#10#13+
                                                    '--Je Zeile ein Kommando', '', 'SQL ausführen', 'Abbrechen', true);
  if frmAusgabe.ShowModal = mrOK
    then
      begin
        rows := 0;
        for i := 0 to frmAusgabe.Memo.Lines.count-1 do
          begin
            sHelp := frmAusgabe.Memo.Lines[i];
            sHelp := SQL_DeleteComment(sHelp);
            if sHelp <> '' then rows := rows + ExecSQL(sHelp, frmDM.dsetHelp, false);
          end;
        if rows > 0 then Showmessage(inttostr(rows)+' Zeile(n) bearbeitet');
      end;
end;

procedure TfrmMain.mnuExecSQLSingleClick(Sender: TObject);

var
  rows : integer;

begin
  frmDM.dbStatus(true);
  frmAusgabe.SetDefaults('SQL Kommando ausführen', '--Hier SQL Kommando eingeben', '', 'SQL ausführen', 'Abbrechen', true);
  if frmAusgabe.ShowModal = mrOK
    then
      begin
        rows := ExecSQL(frmAusgabe.Memo.Text, frmDM.dsetHelp, false);
        if rows > 0 then Showmessage(inttostr(rows)+' Zeilen bearbeitet');
      end;
  frmDM.dbStatus(false); // DB schliessen
end;

procedure TfrmMain.mnuExportClick(Sender: TObject);

begin
  frmDM.dbStatus(true);
  frmDM.dsetHelp.SQL.Text := 'select * from Personen';
  frmDM.dsetHelp.Open;
  ExportQueToCSVFile(frmDM.dsetHelp, 'Export.csv', ';', '"', true, false);
  frmDM.dsetHelp.Close;
  frmDM.dbStatus(false); // DB schliessen
end;

procedure TfrmMain.mnuFehlerInKarteiSuchenClick(Sender: TObject);
var
  sHelp : string;
  i     : integer;

    procedure einfuegen(text,text2:string);                   {textformatierung}

    begin
       slAusgabe.Add(AppendChar(text, ' ', 33)+text2);
    end;

begin
  screen.cursor := crhourglass;
  try
    slAusgabe.clear;
    slAusgabe.Add('Fehlerhafte Datensätze');
    slAusgabe.add('');
    slAusgabe.add('Folgende Prüfungen sind eingebaut:');
    slAusgabe.add('    Kein Geburtsdatum');
    slAusgabe.add('    Kein Taufdatum');
    slAusgabe.add('    Kein oder falsches Geschlecht');
    slAusgabe.add('    Keine Anrede');
    slAusgabe.add('    Kein Vorname');
    slAusgabe.add('    Kein Familienstand');
    slAusgabe.add('    Keine Konfession');
    slAusgabe.add('    Taufdatum < Geb.Datum');
    slAusgabe.add('    Konf.Datum < Taufdatum');
    slAusgabe.add('    Alter > 65 und kein Ruhestand');
    slAusgabe.add('');
    slAusgabe.add('Alle Personen, die unter "Kirche" INFO eingetragen haben, werden bei der Suche NICHT berücksichtigt.');
    slAusgabe.add('');
    slAusgabe.add('Längenüberprüfung');
    slAusgabe.add('');
    slAusgabe.add('Falls Sie weitere Prüfungen wünschen, teilen Sie es mit mit!');
    slAusgabe.add('PS.: bitte mit Formel.');
    slAusgabe.add(' ');

    frmDM.dsetHelp.SQL.Text:='select * from Personen where (kirche is null) or (upper(Kirche) <> ''INFO'') order by Nachname, Vorname';

    frmDM.dbStatus(true); //Datenbank öffnen
    frmDM.dsetHelp.Open;
    frmDM.dsetHelp.first;
    while not frmDM.dsetHelp.eof do
      begin
        sHelp := frmDM.dsetHelp.fieldByName('vorname').asstring+' '+frmDM.dsetHelp.fieldByName('Nachname').asstring;

        if frmDM.dsetHelp.fieldByName('Geburtstag').asstring = ''
          then einfuegen(sHelp, ' Kein Geburtstag')
          else
            if (StrToInt(year(DateToStr(date))) - StrToInt(year(frmDM.dsetHelp.FieldByName('Geburtstag').asstring)) > 65) and not frmDM.dsetHelp.fieldByName('ruhestand').asboolean
              then einfuegen(sHelp, ' älter als 65 und nicht im Ruhestand?');
        if trim(frmDM.dsetHelp.fieldByName('TaufDatum').asstring) = ''  then einfuegen(sHelp, ' Kein Taufdatum');
        if (frmDM.dsetHelp.fieldByName('Geburtstag').asstring <> '') and (frmDM.dsetHelp.fieldByName('TaufDatum').asstring <> '')
          then
            if frmDM.dsetHelp.fieldByName('TaufDatum').asdatetime < frmDM.dsetHelp.fieldByName('Geburtstag').asdatetime
              then einfuegen(sHelp, ' TaufDatum < Geburtstag');
        if (frmDM.dsetHelp.fieldByName('TaufDatum').asstring <> '') and (frmDM.dsetHelp.fieldByName('KonfDatum').asstring <> '')
          then
            if frmDM.dsetHelp.fieldByName('KonfDatum').asdatetime < frmDM.dsetHelp.fieldByName('TaufDatum').asdatetime
              then einfuegen(sHelp, ' Konf.Datum < Taufdatum');

        if frmDM.dsetHelp.fieldByName('Familienstand').asstring = ''  then einfuegen(sHelp, ' Kein Familienstand');
        if frmDM.dsetHelp.fieldByName('Geschlecht').asstring = ''     then einfuegen(sHelp, ' Kein Geschlecht');

        if not ((frmDM.dsetHelp.fieldByName('Geschlecht').asstring = 'W') or
                (frmDM.dsetHelp.fieldByName('Geschlecht').asstring = 'M') or
                (frmDM.dsetHelp.fieldByName('Geschlecht').asstring = 'D'))
        	  then einfuegen(sHelp, ' kein oder falsches Geschlecht');
        if frmDM.dsetHelp.fieldByName('BriefAnrede').asstring = ''    then einfuegen(sHelp, ' Keine Anrede');
        if frmDM.dsetHelp.fieldByName('vorname').asstring = ''        then einfuegen(sHelp, ' Kein Vorname');
        if frmDM.dsetHelp.fieldByName('Kirche').asstring = ''         then einfuegen(sHelp, ' Keine Kirche');

        //Längenüberprüfung
        for i := 0 to frmDM.dsetHelp.FieldDefs.Count-1 do
          if frmDM.dsetHelp.FieldDefs.Items[i].DataType = ftString then
            begin
              if length(frmDM.dsetHelp.fieldByName(frmDM.dsetHelp.FieldDefs.Items[i].Name).asstring) > frmDM.dsetHelp.FieldDefs.Items[i].Size then
                begin
                  einfuegen(sHelp, ' Text im Feld "'+frmDM.dsetHelp.FieldDefs.Items[i].Name+'" zu lang');
                end;
            end;
        frmDM.dsetHelp.next;
      end;
    frmDM.dsetHelp.close;
    frmDM.dbStatus(false); //Datenbank schliessen
  finally
    screen.cursor := crdefault;
  end;
  ShowDruckenDlg;
end;

procedure TfrmMain.mnuFreeListClick(Sender: TObject);

var
  i,
  len : integer;
  s   : String;
  FieldSize: array[0..199] of integer;

begin
  slAusgabe.Clear;  //slAusgabe als Zwichenspeicher

  PerpareDatabaseForPrint;
  for i := 0 to frmDM.dsetHelp.FieldDefs.Count-1 do
    begin
      if frmDM.dsetHelp.FieldDefs[i].DataType in [ftDate, ftString, ftInteger]
        then slAusgabe.Add(frmDM.dsetHelp.FieldDefs[i].Name);
    end;

  frmFreieListe.SrcList.Items.Text := slAusgabe.Text;
  frmFreieListe.SrcList.Hint := 'Adresszusatz = sRes1'+#13#10+'Ehegatten Geburtsdatum = dRes1';
  frmFreieListe.SrcList.ShowHint:=true;
  frmFreieListe.DstList.Items.Clear;
  slAusgabe.Clear;  //slAusgabe wieder leeren
  if frmFreieListe.ShowModal = mrOK
    then
      begin
        if frmFreieListe.DstList.Items.Count = 0
          then Showmessage('Keine Felder gewählt')
          else
            begin
              screen.cursor := crhourglass;
              //Header ausgeben
              s := '';
              for i := 0 to frmFreieListe.DstList.Items.Count-1 do
                begin
                  if frmFreieListe.rbCSV.Checked
                  then
                    begin
                      if i > 0 then s := s + ';';
                      s := s+'"'+DeleteChar(frmFreieListe.DstList.Items[i], '"')+'"';
                    end
                  else
                    begin
                      if i > 0 then s := s + ' ';
                      //Prüfen ob die Textlänge größer als Size ist
                      //Integer und Datum behandeln
                      case frmDM.dsetHelp.FieldByName(frmFreieListe.DstList.Items[i]).DataType of
                        ftDate    : len := 10;
                        ftString  : len := frmDM.dsetHelp.FieldByName(frmFreieListe.DstList.Items[i]).Size div 4;
                        ftInteger : len := 10;
                      end;

                      len := max(len, length(frmFreieListe.DstList.Items[i]));
                      FieldSize[i] := len;

                      s := s + AppendChar(frmFreieListe.DstList.Items[i], ' ', len);
                    end;
                end;
              slAusgabe.add(s);

              //Daten ausgeben
              frmDM.dsetHelp.First;
              while not frmDM.dsetHelp.EOF do
                begin
                  s := '';
                  for i := 0 to frmFreieListe.DstList.Items.Count-1 do
                    begin
                      if frmFreieListe.rbCSV.Checked then
                        begin
                          if i > 0 then s := s + ';';
                          s := s+'"'+DeleteChar(frmDM.dsetHelp.FieldByName(frmFreieListe.DstList.Items[i]).AsString , '"')+'"';
                        end
                      else
                        begin
                          if i > 0 then s := s + ' ';
                          len := FieldSize[i];
                          s := s+AppendChar(frmDM.dsetHelp.FieldByName(frmFreieListe.DstList.Items[i]).AsString , ' ', len);
                        end;
                    end;
                  slAusgabe.add(s);

                  frmDM.dsetHelp.Next;
                end;
              screen.cursor := crDefault;
              ShowDruckenDlg;
            end;
      end;
  CloseDatabaseAfterPrint;
end;

procedure TfrmMain.mnuGDBearbeitenClick(Sender: TObject);
begin
  frmGD.Showmodal;
  if frmDM.dsetHelp1.Active then frmDM.dsetHelp1.close;
  if frmDM.dsetHelp2.Active then frmDM.dsetHelp2.Close;
end;

procedure TfrmMain.mnuGDStatistikClick(Sender: TObject);

  procedure einfuegen(text,text2:string);                   {textformatierung}

  begin
    slAusgabe.Add(AppendChar(text,' ',32)+text2);
  end;

var a_s,
    b_s,
    h_s,
    k_s,
    l_s,
    m_s,
    p_s,
    s_s,
    v_s,
    a_f,
    b_f,
    h_f,
    k_f,
    l_f,
    m_f,
    p_f,
    s_f,
    v_f,
    a_w,
    b_w,
    h_w,
    k_w,
    l_w,
    m_w,
    p_w,
    s_w,
    v_w,
    s,
    w,
    f,
    besucher_s,
    besucher_f,
    besucher_w,
    kommunikanten,
    gastkomm        : integer;
    sWhere          : String;
    dt              : TDateTime;

begin
  //Datenbank öffnen
  frmDM.dbStatus(true);

  slAusgabe.clear;

  a_s           := 0;
  b_s           := 0;
  h_s           := 0;
  k_s           := 0;
  l_s           := 0;
  m_s           := 0;
  p_s           := 0;
  s_s           := 0;
  v_s           := 0;
  a_f           := 0;
  b_f           := 0;
  h_f           := 0;
  k_f           := 0;
  l_f           := 0;
  m_f           := 0;
  p_f           := 0;
  s_f           := 0;
  v_f           := 0;
  a_w           := 0;
  b_w           := 0;
  h_w           := 0;
  k_w           := 0;
  l_w           := 0;
  m_w           := 0;
  p_w           := 0;
  s_w           := 0;
  v_w           := 0;
  s             := 0;
  w             := 0;
  f             := 0;
  besucher_s    := 0;
  besucher_f    := 0;
  besucher_w    := 0;
  kommunikanten := 0;
  gastkomm      := 0;

  sWhere        := '';

  dt := now();
  if FormatDateTime('MM', dt) = '01' then dt := dt-365;
  frmStatInfo.stat_jahr.Text        := FormatDateTime('YYYY', dt);
  frmStatInfo.stat_jahr.Enabled     := true;
  frmStatInfo.stat_Kirche.Enabled   := false;
  frmStatInfo.stat_gemeinde.Enabled := true;
  frmStatInfo.stat_gemeinde.Text    := sDefaultGemeinde;
  if frmStatInfo.ShowModal = mrOK
    then
      begin
        try
          screen.cursor := crhourglass;

	        //Filter auf Gemeinde
          if frmStatInfo.stat_gemeinde.Text = ''
            then sWhere := SQL_Where_Add(sWhere, SQL_Where_IsNull('Gemeinde'))
            else if frmStatInfo.stat_gemeinde.Text <> sGemeindenAlle
              then sWhere := SQL_Where_Add(sWhere, 'Gemeinde=''' + frmStatInfo.stat_gemeinde.Text + '''');

          //und Jahr
          sWhere := SQL_Where_Add(sWhere, 'strftime(''%Y'',GottesdienstDatum)='''+frmStatInfo.stat_jahr.text+'''');

	        frmDM.dsetHelp.SQL.Text := 'select * from Gottesdienst where '+sWhere;
          frmDM.dsetHelp.open;
          frmDM.dsetHelp.first;
          while not frmDM.dsetHelp.eof do
            begin
              kommunikanten := kommunikanten + frmDM.dsetHelp.FieldByName('Kommunikanten').asinteger;
              gastkomm := gastkomm + frmDM.dsetHelp.FieldByName('GastKommunikanten').asinteger;

              if frmDM.dsetHelp.FieldByName('tag').asstring = 'S' then            { sonntag }
                begin
                  besucher_s := besucher_s + frmDM.dsetHelp.FieldByName('besucher').asinteger;
                  inc(s);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'A') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'A') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'A')
                    then inc(a_s);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'B') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'B') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'B')
                    then inc(b_s);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'H') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'H') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'H')
                    then inc(H_s);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'K') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'K') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'K')
                    then inc(k_s);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'L') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'L') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'L')
                    then inc(l_s);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'M') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'M') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'M')
                    then inc(m_s);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'P') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'P') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'P')
                    then inc(p_s);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'S') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'S') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'S')
                    then inc(s_s);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'V') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'V') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'V')
                    then inc(v_s);
                end;

              if frmDM.dsetHelp.FieldByName('tag').asstring = 'F' then            { Feiertag }
                begin
                  besucher_f := besucher_f + frmDM.dsetHelp.FieldByName('besucher').asinteger;
                  inc(f);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'A') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'A') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'A')
                    then inc(a_f);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'B') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'B') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'B')
                    then inc(b_f);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'H') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'H') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'H')
                    then inc(H_f);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'K') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'K') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'K')
                    then inc(k_f);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'L') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'L') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'L')
                    then inc(l_f);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'M') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'M') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'M')
                    then inc(m_f);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'P') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'P') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'P')
                    then inc(p_f);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'S') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'S') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'S')
                    then inc(s_f);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'V') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'V') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'V')
                    then inc(v_f);
                end;

              if frmDM.dsetHelp.FieldByName('tag').asstring = 'W' then            { Feiertag }
                begin
                  besucher_w := besucher_w + frmDM.dsetHelp.FieldByName('besucher').asinteger;
                  inc(w);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'A') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'A') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'A')
                    then inc(a_w);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'B') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'B') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'B')
                    then inc(b_w);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'H') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'H') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'H')
                    then inc(H_w);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'K') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'K') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'K')
                    then inc(k_w);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'L') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'L') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'L')
                    then inc(l_w);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'M') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'M') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'M')
                    then inc(m_w);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'P') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'P') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'P')
                    then inc(p_w);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'S') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'S') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'S')
                    then inc(s_w);
                  if (frmDM.dsetHelp.FieldByName('GottesdienstForm1').asstring = 'V') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm2').asstring = 'V') or
                     (frmDM.dsetHelp.FieldByName('GottesdienstForm3').asstring = 'V')
                    then inc(v_w);
                end;
            frmDM.dsetHelp.next;
          end;

        slAusgabe.add('');
        einfuegen('Gottesdienststatistik','');
        slAusgabe.add('');
        einfuegen('Sonntag Andachten:     '+inttostr(a_s),' ');
        einfuegen('Sonntag Beichten:      '+inttostr(b_s),' ');
        einfuegen('Sonntag Haupt GD:      '+inttostr(h_s),' ');
        einfuegen('Sonntag Kinder GD:     '+inttostr(k_s),' ');
        einfuegen('Sonntag Lese GD:       '+inttostr(l_s),' ');
        einfuegen('Sonntag Matutin:       '+inttostr(m_s),' ');
        einfuegen('Sonntag Predigt GD:    '+inttostr(p_s),' ');
        einfuegen('Sonntag Sonstige:      '+inttostr(s_s),' ');
        einfuegen('Sonntag Vesper:        '+inttostr(v_s),' ');
        slAusgabe.add(' ');
        einfuegen('Besucher Sonntag:      '+inttostr(besucher_s),' ');
        einfuegen('Eingegebene Sonntage:  '+inttostr(s),' ');
        if s <> 0 then einfuegen('Duchschnittl. Besuch:  '+RealToStr(besucher_s/s,4,1),' ');
        slAusgabe.add(' ');
        einfuegen('Feiertag Andachten:    '+inttostr(a_f),' ');
        einfuegen('Feiertag Beichten:     '+inttostr(b_f),' ');
        einfuegen('Feiertag Haupt GD:     '+inttostr(h_f),' ');
        einfuegen('Feiertag Kinder GD:    '+inttostr(k_f),' ');
        einfuegen('Feiertag Lese GD:      '+inttostr(l_f),' ');
        einfuegen('Feiertag Matutin:      '+inttostr(m_f),' ');
        einfuegen('Feiertag Predigt GD:   '+inttostr(p_f),' ');
        einfuegen('Feiertag Sonstige:     '+inttostr(s_f),' ');
        einfuegen('Feiertag Vesper:       '+inttostr(v_f),' ');
        slAusgabe.add(' ');
        einfuegen('Besucher Feiertag:     '+inttostr(besucher_f),' ');
        einfuegen('Eingegebene Feiertage: '+inttostr(f),' ');
        if f <> 0 then einfuegen('Duchschnittl. Besuch:  '+RealToStr(besucher_f/f,4,1),' ');
        slAusgabe.add(' ');
        einfuegen('Werktag Andachten:     '+inttostr(a_w),' ');
        einfuegen('Werktag Beichten:      '+inttostr(b_w),' ');
        einfuegen('Werktag Haupt GD:      '+inttostr(h_w),' ');
        einfuegen('Werktag Kinder GD:     '+inttostr(k_w),' ');
        einfuegen('Werktag Lese GD:       '+inttostr(l_w),' ');
        einfuegen('Werktag Matutin:       '+inttostr(m_w),' ');
        einfuegen('Werktag Predigt GD:    '+inttostr(p_w),' ');
        einfuegen('Werktag Sonstige:      '+inttostr(s_w),' ');
        einfuegen('Werktag Vesper:        '+inttostr(v_w),' ');
        slAusgabe.add(' ');
        einfuegen('Besucher Werktag:      '+inttostr(besucher_w),' ');
        einfuegen('Eingegebene Werktage:  '+inttostr(w),' ');
        if w <> 0 then einfuegen('Duchschnittl. Besuch:  '+RealToStr(besucher_w/w,4,1),' ');
        slAusgabe.add(' ');
        slAusgabe.add(' ');
        einfuegen('Kommunikanten:         '+inttostr(kommunikanten),' ');
        slAusgabe.add(' ');
        einfuegen('Gastkommunikanten:     '+inttostr(gastkomm),' ');
        slAusgabe.add(' ');
        slAusgabe.add('*****************************************************');
        slAusgabe.add(' ');
        slAusgabe.add('Eingabeparameter Jahr    : '+frmStatInfo.stat_jahr.text);
        slAusgabe.add('Eingabeparameter Gemeinde: '+frmStatInfo.stat_gemeinde.Text);
        slAusgabe.add(' ');
        slAusgabe.add(#13#10+'Erzeugt von GE_Kart '+ GetProductVersionString +' am: '+datetostr(date));
        screen.cursor := crdefault;
        ShowDruckenDlg;
      finally
        frmDM.dsetHelp.close;
      end;
    end;
  frmDM.dbStatus(false); //Datenbank schliessen
end;

procedure TfrmMain.mnuGebListClick(Sender: TObject);
var
  sWhere1,
  sWhere2,
  sWhere3,
  sWhere4,
  sWhere5,
  sGebAb   : String;
  i        : integer;

  Function FormatAge(GebTag: TDateTime; Age: integer):String;

  var
    YY ,MM ,DD  : Word;
    YYG,MMG,DDG : Word;
    iAge        : integer;

  begin
    iAge := Age;
    DecodeDate(Date,YY,MM,DD);
    DecodeDate(GebTag,YYG,MMG,DDG);
    if MM > MMG then inc(iAge);
    result := DateToStr(GebTag)+Format(#9'%3d'#9,[iAge])
  end;

  procedure Ausgabe;

  begin
    frmDM.dsetHelp.sql.Clear;
    frmDM.dsetHelp.sql.add('select vorname, Nachname, Geburtstag, (strftime(''%Y'', ''now'') - strftime(''%Y'', geburtstag)) as Age from '+global.sPersTablename+' where');
    frmDM.dsetHelp.sql.add(sWhere1);
    frmDM.dsetHelp.sql.add(sWhere2);
    frmDM.dsetHelp.sql.add(sWhere3);
    frmDM.dsetHelp.sql.add(sWhere4);
    frmDM.dsetHelp.sql.add(sWhere5);
    frmDM.dsetHelp.sql.add('order by strftime(''%j'',geburtstag), nachname, vorname');

    frmDM.dsetHelp.open;

    slAusgabe.add('Geburtstage ab '+sGebAb);
    while not frmDM.dsetHelp.eof do
      begin
        if not (frmDM.dsetHelp.fieldByName('Geburtstag').asstring = '') then
          slAusgabe.add(FormatAge(frmDM.dsetHelp.fieldByName('Geburtstag').asdateTime, frmDM.dsetHelp.fieldByName('Age').asinteger)+
                        frmDM.dsetHelp.fieldByName('vorname').asstring+' '+frmDM.dsetHelp.fieldByName('Nachname').asstring);
        frmDM.dsetHelp.Next;
      end;
    frmDM.dsetHelp.Close;

    if sGebAb <> '0' then
      begin
        slAusgabe.add('');
        slAusgabe.add('Geburtstage bis 14');

        frmDM.dsetHelp.sql.Clear;
        frmDM.dsetHelp.sql.add('select vorname, Nachname, Geburtstag, (strftime(''%Y'', ''now'') - strftime(''%Y'', geburtstag)) as Age from '+global.sPersTablename+' where');
        frmDM.dsetHelp.sql.add(sWhere1);
        frmDM.dsetHelp.sql.add('((strftime(''%Y'', ''now'') - strftime(''%Y'', geburtstag)) <= 14) and');
        frmDM.dsetHelp.sql.add(sWhere3);
        frmDM.dsetHelp.sql.add(sWhere4);
        frmDM.dsetHelp.sql.add(sWhere5);
        frmDM.dsetHelp.sql.add('order by strftime(''%j'',geburtstag), nachname, vorname');

        frmDM.dsetHelp.open;
        while not frmDM.dsetHelp.eof do
          begin
            slAusgabe.add(FormatAge(frmDM.dsetHelp.fieldByName('Geburtstag').asdateTime, frmDM.dsetHelp.fieldByName('Age').asinteger)+
                          frmDM.dsetHelp.fieldByName('vorname').asstring+' '+frmDM.dsetHelp.fieldByName('Nachname').asstring);
            frmDM.dsetHelp.Next;
          end;
        frmDM.dsetHelp.Close;
      end;
  end;

const
  AnzahlMonate = 5;

begin
  frmStatInfo.stat_jahr.Enabled     := false;
  frmStatInfo.stat_Kirche.Enabled   := true;
  frmStatInfo.stat_gemeinde.Enabled := not mnuGebListSplit.Checked;
  frmStatInfo.stat_gemeinde.Text    := sDefaultGemeinde;
  if frmStatInfo.ShowModal = mrOK
    then
      begin
        sWhere1 := '';
        slHelp.Clear;
        for i := 0 to AnzahlMonate do
          sWhere1 := SQL_Where_Add_OR(sWhere1, '(strftime(''%m'',geburtstag) = '''+FormatDateTime('mm', IncMonth(now(), i))+''')');
        sWhere1 := '('+sWhere1+') and';

        sGebAb  := help.ReadIniVal(sIniFile, 'Defaults','JubiListAlleGebTageAb', '65', true);
        sWhere2 := '((strftime(''%Y'', ''now'') - strftime(''%Y'', geburtstag)) >= '+sGebAb+') and';

        sWhere3 := '';
        sWhere3 := SQL_Where_Add(sWhere3, SQL_Where_IsNull('Ueberwiesen_nach_Datum'));
        sWhere3 := SQL_Where_Add(sWhere3, SQL_Where_IsNull('AustrittsDatum'));
        sWhere3 := SQL_Where_Add(sWhere3, SQL_Where_IsNull('AusschlussDatum'));
        sWhere3 := SQL_Where_Add(sWhere3, SQL_Where_IsNull('TodesDatum'));
        sWhere3 := SQL_Where_Add(sWhere3, SQL_Where_IsNull('UebertrittsAbDatum'));

        sWhere5 := '';
        if frmStatInfo.stat_Kirche.text <> ''
          then sWhere5 := SQL_Where_Add(sWhere5, 'and Kirche=''' + frmStatInfo.stat_Kirche.Text + '''');

        slAusgabe.Clear;
        slAusgabe.add('Geburtstage für die Monate '+FormatDateTime('mm', now())+' bis '+FormatDateTime('mm', IncMonth(now(), AnzahlMonate)));
        slAusgabe.add('');
        if mnuGebListSplit.Checked
          then
            begin
              slHelp.Text := sGemeinden;
              for i := 0 to slHelp.Count-1 do
                begin
                  if (slHelp.strings[i] <> sGemeindenAlle) then
                    begin
                      sWhere4 := 'and (Gemeinde = ''' + slHelp.strings[i] +''')';
                      slAusgabe.add('Geburtstage für die Gemeinde "'+slHelp.strings[i]+'"');
                      slAusgabe.add('');
                      Ausgabe;
                      slAusgabe.add('');
                    end;
                end;
            end
          else
            begin
              sWhere4 := '';
              if frmStatInfo.stat_gemeinde.Text = ''
                then sWhere4 := SQL_Where_Add(sWhere4, 'and '+SQL_Where_IsNull('Gemeinde'))
                else if frmStatInfo.stat_gemeinde.Text <> sGemeindenAlle
                  then sWhere4 := SQL_Where_Add(sWhere4, 'and Gemeinde=''' + frmStatInfo.stat_gemeinde.Text + '''');
              slAusgabe.add('Geburtstage für die Gemeinde "'+frmStatInfo.stat_gemeinde.Text+'"');
              slAusgabe.add('');
              Ausgabe;
              slAusgabe.add('');
            end;

        slAusgabe.add('');
        slAusgabe.add('Die "Geburtstage ab '+sGebAb+'" können in der Datei '+sIniFile+', Variable: "JubiListAlleGebTageAb" eingestellt werden.');
        if sGebAb <> '0' then slAusgabe.add('Um alle Geburtstage zu bekommen, dort eine "0" eintragen');
        slAusgabe.add('');
        slAusgabe.add('Erzeugt von GE_Kart '+ GetProductVersionString +' am: '+datetostr(date));
        ShowDruckenDlg;
      end;
end;

procedure TfrmMain.mnuGebListSplitClick(Sender: TObject);
begin
  mnuGebListSplit.Checked := not mnuGebListSplit.Checked;
  if mnuGebListSplit.Checked
    then help.WriteIniVal(sIniFile, 'Defaults', 'GebListSplit', '1')
    else help.WriteIniVal(sIniFile, 'Defaults', 'GebListSplit', '0');
end;

procedure TfrmMain.mnuGemeindeAdresseClick(Sender: TObject);
begin
  frmInput.SetDefaults('Adresseingabe', 'Zeile 1', labGemeinde1.caption, 'Zeile 2', labGemeinde2.caption, false);
  if frmInput.ShowModal = mrOK
    then
      begin
        labGemeinde1.caption := frmInput.Edit1.Text;
        labGemeinde2.caption := frmInput.Edit2.Text;
        frmDM.ExecSQL('Update Gemeinde SET Adr1="'+frmInput.Edit1.Text+'", Adr2="'+frmInput.Edit2.Text+'"');
      end;
end;

procedure TfrmMain.mnuGemeindeDefaultClick(Sender: TObject);
begin
  frmInput.SetDefaults('Standard für "Gemeinde" bei neuer Person festlegen', 'Gemeinde', sDefaultGemeinde, '', '', false);
  if frmInput.ShowModal = mrOK then
    begin
      sDefaultGemeinde := frmInput.Edit1.Text;
      help.WriteIniVal(sIniFile, 'Defaults','Gemeinde', sDefaultGemeinde);
    end;
end;

procedure TfrmMain.mnuImportClick(Sender: TObject);
var
  sNam,
  sVal,
  sGE_Kart5Fieldname,
  sGE_Kart7Fieldname,
  sGE_Kart5Val       : string;
  i,
  cnt                : integer;
  bOhneMemo          : boolean;
  sl,
  slNum              : TStringList;

begin
  frmDM.dbStatus(true); //Datenbank öffnen
  sl    := TStringlist.Create;
  slNum := TStringlist.Create;
  cnt   := 0;
  if MessageDlg('Importieren?', 'ACHTUNG: Diese Funktion löscht alle Daten und liest eine alte Datenbank neu ein.'+#13#13+'Sicherheitsfunktion: Zum Fortfahren "Wiederholen" drücken!', mtConfirmation, [mbYes, mbRetry, mbNo],0) = mrRetry
    then
      begin
        screen.cursor := crhourglass;
        sl.LoadFromFile('NewDB.SQL');
        //Zeilen einzeln bearbeiten, das Tranactionen nocht mehr gehen.
        for i := 0 to sl.count-1 do
          frmDM.ExecSQL(sl.strings[i]); //Datenbank neu anlegen damit die Nummer bei 1 beginnt

        //Personentabelle
        frmDM.DBF_Help.TableName:=UTF8ToSys(sAppDir)+'Datenbank6\Ge_kart5.dbf';
        sl.Text:=global.sGE5TOGE7;  //Feldnamenaliasliste
        try
          frmDM.DBF_Help.Open;
          //for i := 0 to frmDM.DBF_Help.FieldDefs.Count-1 do myDebugLN(frmDM.DBF_Help.FieldDefs.Items[i].Name);
          bOhneMemo := false;

          while not frmDM.DBF_Help.EOF do
            begin
              sNam               := '';
              sVal               := '';

              for i := 0 to sl.Count-1 do
                begin
                  sGE_Kart5Fieldname := sl.Names[i];
                  sGE_Kart7Fieldname := sl.ValueFromIndex[i];
                  try
                    //Damit wird geprüft ob es das Feld gibt
                    sGE_Kart5Val := frmDM.DBF_Help.FieldByName(sGE_Kart5Fieldname).AsString;
                  except
                    sGE_Kart5Val := '';
                  end;
                  if sGE_Kart5Val <> ''
                    then
                      begin
                        sNam := sNam + '"'+sGE_Kart7Fieldname+'", ';
                        //Sonderzeichenumwandlung und Datumsformat beachten
                        if frmDM.DBF_Help.FieldByName(sGE_Kart5Fieldname).DataType = ftDate
                          then sVal := sVal + '"'+FormatDateTime('yyyy-mm-dd', frmDM.DBF_Help.FieldByName(sGE_Kart5Fieldname).AsDateTime)+'", '
                          else
                            begin
                              if bOhneMemo and (sGE_Kart5Fieldname = 'MEMO')
                                then sVal := sVal + '" ",'
                                else sVal := sVal + '"'+CP850toUTF8(ReplaceChar(DeleteChars(trim(sGE_Kart5Val), ['"', #13]), #10, ' '))+'", ';
                            end;
                      end;
                end;

              sNam := sNam + '"Abgang", "Markiert"';
              sVal := sVal + '0, 0';
              try
                frmDM.ExecSQL('Insert into '+global.sPersTablename+#13' ('+sNam+') VALUES'#13' ('+sVal+')', true);
                bOhneMemo := false;
              except
                on E: Exception do
                  begin
                    myDebugLN(E.Message);
                    if bOhneMemo
                      then raise E
                      else bOhneMemo := true;
                  end;
              end;
              if bOhneMemo = false
                then
                  begin
                    frmDM.DBF_Help.Next;
                    inc(cnt);
                    if (cnt mod 10) = 0
                      then
                        begin
                          StatusBar.Panels[0].Text:=Format('%u Datensätze eingefügt', [cnt]);
                          StatusBar.Update;
                          application.processmessages;
                        end;
                  end;
            end;
          frmDM.DBF_Help.Close;
        except
          on E: Exception do LogAndShowError(E.Message);
        end;

        //Liste für die Zuordnung dBase <--> SQLite IDs
        frmDM.dsetHelp.SQL.Text := 'Select PersonenID, TempInteger from '+global.sPersTablename;
        frmDM.dsetHelp.Open;
        while not frmDM.dsetHelp.EOF do
          begin
            slNum.Append(frmDM.dsetHelp.FieldByName('TempInteger').AsString+'='+frmDM.dsetHelp.FieldByName('PersonenID').AsString);
            frmDM.dsetHelp.Next;
          end;
        frmDM.dsetHelp.Close;
        myDebugLN(slNum.Text);

        //Besuchstabelle
        frmDM.DBF_Help.TableName:=UTF8ToSys(sAppDir)+'Datenbank6\Ge_besu5.dbf';
        sl.Text:=global.sBE5TOBE7;  //Feldnamenaliasliste
        try
          frmDM.DBF_Help.Open;

          while not frmDM.DBF_Help.EOF do
            begin
              sNam := '';
              sVal := '';
              for i := 0 to sl.Count-1 do
                if frmDM.DBF_Help.FieldByName(sl.Names[i]).AsString <> '' then
                  begin
                    sNam := sNam + '"'+sl.ValueFromIndex[i]+'", ';
                    //Sonderzeichenumwandlung und Datumsformat beachten
                    if frmDM.DBF_Help.FieldByName(sl.Names[i]).DataType = ftDate
                      then sVal := sVal + '"'+FormatDateTime('yyyy-mm-dd', frmDM.DBF_Help.FieldByName(sl.Names[i]).AsDateTime)+'", '
                      else sVal := sVal + '"'+CP850toUTF8(DeleteChar(trim(frmDM.DBF_Help.FieldByName(sl.Names[i]).AsString), '"'))+'", ';
                  end;
              //Verknüpfung zur Personentabelle
              sNam := sNam + 'PersonenID';
              sVal := sVal + sLNum.Values[frmDM.DBF_Help.FieldByName('BENUMMER').AsString];
              if copy(sVal, length(sVAL), 1) <> ','
                then //BENummer gefunden
                  frmDM.ExecSQL('Insert into '+global.sBesuchTablename+' ('+sNam+') VALUES ('+sVal+')')
                else
                  myDebugLN('Invalid SQL: '+sNam+ ' '+sVal);
              frmDM.DBF_Help.Next;
              inc(cnt);
              if (cnt mod 10) = 0
                then
                  begin
                    StatusBar.Panels[0].Text:=Format('%u Datensätze eingefügt', [cnt]);
                    StatusBar.Update;
                    application.processmessages;
                  end;
            end;
          frmDM.DBF_Help.Close;
        except
          on E: Exception do LogAndShowError(E.Message);
        end;

        //Kinder
        frmDM.DBF_Help.TableName:=UTF8ToSys(sAppDir)+'Datenbank6\GE_Kind5.dbf';
        sl.Text:=global.sKI5TOKI7;  //Feldnamenaliasliste
        try
          frmDM.DBF_Help.Open;

          while not frmDM.DBF_Help.EOF do
            begin
              sNam := '';
              sVal := '';
              for i := 0 to sl.Count-1 do
                if frmDM.DBF_Help.FieldByName(sl.Names[i]).AsString <> '' then
                  begin
                    sNam := sNam + '"'+sl.ValueFromIndex[i]+'",';
                    //Sonderzeichenumwandlung und Datumsformat beachten
                    if frmDM.DBF_Help.FieldByName(sl.Names[i]).DataType = ftDate
                      then sVal := sVal + '"'+FormatDateTime('yyyy-mm-dd', frmDM.DBF_Help.FieldByName(sl.Names[i]).AsDateTime)+'",'
                      else sVal := sVal + '"'+CP850toUTF8(DeleteChar(trim(frmDM.DBF_Help.FieldByName(sl.Names[i]).AsString), '"'))+'",';
                  end;
              //Verknüpfung zur Personentabelle
              sNam := sNam + 'PersonenID';
              sVal := sVal + sLNum.Values[frmDM.DBF_Help.FieldByName('KINUMMER').AsString];
              if copy(sVal, length(sVAL), 1) <> ','
                then //KINummer gefunden
                  frmDM.ExecSQL('Insert into '+global.sKindTablename+' ('+sNam+') VALUES ('+sVal+')')
                else
                  myDebugLN('Invalid SQL: '+sNam+ ' '+sVal);
              frmDM.DBF_Help.Next;
              inc(cnt);
              if (cnt mod 10) = 0
                then
                  begin
                    StatusBar.Panels[0].Text:=Format('%u Datensätze eingefügt', [cnt]);
                    StatusBar.Update;
                    application.processmessages;
                  end;
            end;
          frmDM.DBF_Help.Close;
        except
          on E: Exception do LogAndShowError(E.Message);
        end;

        //Kommunionen
        frmDM.DBF_Help.TableName:=UTF8ToSys(sAppDir)+'Datenbank6\Ge_komm5.dbf';
        sl.Text:=global.sKO5TOKO7;  //Feldnamenaliasliste
        try
          frmDM.DBF_Help.Open;

          while not frmDM.DBF_Help.EOF do
            begin
              sNam := '';
              sVal := '';
              for i := 0 to sl.Count-1 do
                if frmDM.DBF_Help.FieldByName(sl.Names[i]).AsString <> '' then
                  begin
                    sNam := sNam + '"'+sl.ValueFromIndex[i]+'",';
                    //Sonderzeichenumwandlung und Datumsformat beachten
                    if frmDM.DBF_Help.FieldByName(sl.Names[i]).DataType = ftDate
                      then sVal := sVal + '"'+FormatDateTime('yyyy-mm-dd', frmDM.DBF_Help.FieldByName(sl.Names[i]).AsDateTime)+'",'
                      else sVal := sVal + '"'+CP850toUTF8(DeleteChar(trim(frmDM.DBF_Help.FieldByName(sl.Names[i]).AsString), '"'))+'",';
                  end;
              //Verknüpfung zur Personentabelle
              sNam := sNam + 'PersonenID';
              sVal := sVal + sLNum.Values[frmDM.DBF_Help.FieldByName('KONUMMER').AsString];
              if copy(sVal, length(sVAL), 1) <> ','
                then //KONummer gefunden
                  frmDM.ExecSQL('Insert into '+global.sKommTablename+' ('+sNam+') VALUES ('+sVal+')')
                else
                  myDebugLN('Invalid SQL: '+sNam+ ' '+sVal);
              frmDM.DBF_Help.Next;
              inc(cnt);
              if (cnt mod 10) = 0
                then
                  begin
                    StatusBar.Panels[0].Text:=Format('%u Datensätze eingefügt', [cnt]);
                    StatusBar.Update;
                    application.processmessages;
                  end;
            end;
          //Endgültige Zahl anzeigen
          StatusBar.Panels[0].Text:=Format('%u Datensätze eingefügt', [cnt]);
          StatusBar.Update;
          application.processmessages;

          frmDM.DBF_Help.Close;
        except
          on E: Exception do LogAndShowError(E.Message);
        end;
        screen.cursor := crdefault;
        Showmessage('Daten erfolgreich importiert');
      end;
  sl.Free;
  sLNum.Free;
  StatusBar.Panels[0].Text:='';
  frmDM.dbStatus(false);  //Datenbank schliessen
end;

procedure TfrmMain.mnuImportGDClick(Sender: TObject);

var
  sNam,
  sVal  : string;
  i,cnt : integer;
  bOhneMemo : boolean;
  sl    : TStringList;

begin
  sl  := TStringlist.Create;
  cnt := 0;
  if MessageDlg('Importieren?', 'Diese Funktion löscht alle der Gottesdienttabelle liest eine alte Datenbank neu ein.'+#13#13+'Sicherheitsfunktion: Zum Fortfahren "Wiederholen" drücken!', mtConfirmation, [mbYes, mbRetry, mbNo],0) = mrRetry
    then
      begin
        frmDM.ExecSQL('Delete from '+sGDTablename);
        frmDM.ExecSQL('VACUUM');
        bOhneMemo := false;

        //Gottesdiensttabelle
        frmDM.DBF_Help.TableName:=UTF8ToSys(sAppDir)+'Datenbank6\GD5.dbf';
        sl.Text:=global.sGD5TOGD7;  //Feldnamenaliasliste
        try
          frmDM.DBF_Help.Open;
          //for i := 0 to frmDM.DBF_Help.FieldDefs.Count-1 do myDebugLN(frmDM.DBF_Help.FieldDefs.Items[i].Name);

          while not frmDM.DBF_Help.EOF do
            begin
              sNam := '';
              sVal := '';
              for i := 0 to sl.Count-1 do
                if frmDM.DBF_Help.FieldByName(sl.Names[i]).AsString <> '' then
                  begin
                    sNam := sNam + '"'+sl.ValueFromIndex[i]+'",';
                    //Sonderzeichenumwandlung und Datumsformat beachten
                    if frmDM.DBF_Help.FieldByName(sl.Names[i]).DataType = ftDate
                      then sVal := sVal + '"'+FormatDateTime('yyyy-mm-dd', frmDM.DBF_Help.FieldByName(sl.Names[i]).AsDateTime)+'",'
                      else
                         begin
                          if bOhneMemo and (sl.Names[i] = 'MEMO')
                            then sVal := sVal + '" ",'
                            else sVal := sVal + '"'+CP850toUTF8(DeleteChar(trim(frmDM.DBF_Help.FieldByName(sl.Names[i]).AsString), '"'))+'",';
                         end;
                  end;
              delete(sNam, length(sNam), 1);
              delete(sVal, length(sVal), 1);
              try
                frmDM.ExecSQL('Insert into '+global.sGDTablename+#13' ('+sNam+') VALUES'#13' ('+sVal+')');
                bOhneMemo := false;
              except
                on E: Exception do
                  begin
                    myDebugLN(E.Message);
                    if bOhneMemo
                      then raise E
                      else bOhneMemo := true;
                  end;
              end;
              if bOhneMemo = false
                then
                  begin
                    frmDM.DBF_Help.Next;
                    inc(cnt);
                    StatusBar.Panels[0].Text:=Format('%u Datensätze eingefügt', [cnt]);
                    StatusBar.Update;
                    application.processmessages;
                  end;
            end;
          frmDM.DBF_Help.Close;
        except
          on E: Exception do LogAndShowError(E.Message);
        end;
        Showmessage('Daten erfolgreich importiert');
      end;
  sl.Free;
  StatusBar.Panels[0].Text:='';
end;

procedure TfrmMain.mnuIncludeAbgangClick(Sender: TObject);
begin
  mnuIncludeAbgang.Checked:= not mnuIncludeAbgang.Checked;
  ShowDruckenMnu;
end;

procedure TfrmMain.mnuInvertMarkClick(Sender: TObject);
begin
  //DB bereinigen
  frmDM.ExecSQL(sSQL_ClearMark1);
  frmDM.ExecSQL(sSQL_ClearMark2);
  //Invertieren
  frmDM.ExecSQL(sSQL_InvertMark);
end;

procedure TfrmMain.mnuJubiClick(Sender: TObject);

var
  sGeb,
  sTrau,
  sKonf,
  sName,
  sHelp,
  sWhere,
  sMess  : String;
  dtGeb,
  dtTrau,
  dTKonf : TDateTime;
  i      : integer;
  y      : integer;
  SetGebTage : Set of 0..255;


begin
  sWhere := '';
  sWhere := SQL_Where_Add(sWhere, SQL_Where_IsNull('Ueberwiesen_nach_Datum'));
  sWhere := SQL_Where_Add(sWhere, SQL_Where_IsNull('AustrittsDatum'));
  sWhere := SQL_Where_Add(sWhere, SQL_Where_IsNull('AusschlussDatum'));
  sWhere := SQL_Where_Add(sWhere, SQL_Where_IsNull('TodesDatum'));
  sWhere := SQL_Where_Add(sWhere, SQL_Where_IsNull('UebertrittsAbDatum'));

  SetGebTage := [];   {bis V 7.4.0.1: [40, 50, 60, 65..110]}

  sHelp := help.ReadIniVal(sIniFile, 'Defaults','JubiListRundeGebTageAb', '50', true);
  try
    y := strtoint(sHelp) div 10;
    if y >= 11 then y := 5;
  except
    y := 5;
    help.WriteIniVal(sIniFile, 'Defaults','JubiListRundeGebTageAb', '50');
  end;
  sMess := 'Runde Geburtstage ab '+inttostr(y*10);
  for i := y to 11 do SetGebTage := SetGebTage + [i*10];

  sHelp := help.ReadIniVal(sIniFile, 'Defaults','JubiListAlleGebTageAb', '70', true);
  try
    y := strtoint(sHelp);
    if y >= 110 then y := 70;
  except
    y := 70;
    help.WriteIniVal(sIniFile, 'Defaults','JubiListRundeGebTageAb', '50');
  end;
  sMess := sMess + ', alle Geburtstage ab '+inttostr(y);
  for i := y to 110 do SetGebTage := SetGebTage + [i];

  frmDM.dbStatus(true);  //Datenbank öffnen

  slAusgabe.clear;

  y := myval(year(datetostr(date)));

  frmDM.dsetHelp.sql.Text := 'select vorname, nachname, Geburtstag, KonfDatum, TrauDatum, Gemeinde from ' + sPersTablename + ' where ';
  frmDM.dsetHelp.sql.add(sWhere);
  frmDM.dsetHelp.open;
  while not frmDM.dsetHelp.eof do
    begin
      //Vorbereitung
      sName := frmDM.dsetHelp.fieldByName('vorname').asstring+' '+ frmDM.dsetHelp.fieldByName('nachname').asstring;
      dtGeb := frmDM.dsetHelp.fieldByName('Geburtstag').asdateTime;
      sGeb  := AppendChar(FormatDateTime('mmddyyyy',dtGeb)+' '+FormatDateTime('dd.mm.yyyy',dtGeb)+' '+sName, ' ', 50);
      dtKonf:= frmDM.dsetHelp.fieldByName('KonfDatum').asdateTime;
      sKonf := AppendChar(FormatDateTime('mmddyyyy',dtKonf)+' '+FormatDateTime('dd.mm.yyyy',dtKonf)+' '+sName, ' ', 50);
      dtTrau:= frmDM.dsetHelp.fieldByName('TrauDatum').asdateTime;
      sTrau := AppendChar(FormatDateTime('mmddyyyy',dtTrau)+' '+FormatDateTime('dd.mm.yyyy',dtTrau)+' '+sName, ' ', 50);

      //Auswertung
      if y-(myval(FormatDateTime('yyyy',dtGeb))) in SetGebTage then slAusgabe.Add(sGeb+ ' ' +  inttostr(Y-myval(FormatDateTime('yyyy',frmDM.dsetHelp.fieldByName('Geburtstag').asdateTime))) + '. Geburtstag            '+frmDM.dsetHelp.fieldByName('Gemeinde').AsString);

      if y-(myval(FormatDateTime('yyyy',dtKonf))) = 25 then slAusgabe.Add(sKonf+ ' Silberne Konfirmation     '+frmDM.dsetHelp.fieldByName('Gemeinde').AsString);
      if y-(myval(FormatDateTime('yyyy',dtKonf))) = 50 then slAusgabe.Add(sKonf+ ' Goldene Konfirmation      '+frmDM.dsetHelp.fieldByName('Gemeinde').AsString);

      case y-(myval(FormatDateTime('yyyy',dtTrau))) of
        25: slAusgabe.Add(sTrau+ ' Silberhochzeit (25)       '+frmDM.dsetHelp.fieldByName('Gemeinde').AsString);
//        40: slAusgabe.Add(sTrau+ ' Rubinhochzeit (40)');     //ab V7.4.1.0
        50: slAusgabe.Add(sTrau+ ' Goldene Hochzeit (50)     '+frmDM.dsetHelp.fieldByName('Gemeinde').AsString);
        60: slAusgabe.Add(sTrau+ ' Diamantene Hochzeit (60)) '+frmDM.dsetHelp.fieldByName('Gemeinde').AsString);
        65: slAusgabe.Add(sTrau+ ' Eiserne Hochzeit (65)     '+frmDM.dsetHelp.fieldByName('Gemeinde').AsString);
        70: slAusgabe.Add(sTrau+ ' Gnadenhochzeit (70))      '+frmDM.dsetHelp.fieldByName('Gemeinde').AsString);
        75: slAusgabe.Add(sTrau+ ' Kronjuwelenhochzeit (75)  '+frmDM.dsetHelp.fieldByName('Gemeinde').AsString);
      end;

      frmDM.dsetHelp.next;
    end;
  frmDM.dsetHelp.close;
  frmDM.dbStatus(false); //Datenbank schliessen

  slAusgabe.Sort;

  //Löschen der Sortinfo
  For i := 0 to slAusgabe.count-1 do
    begin
      sHelp := slAusgabe.Strings[i];
      delete(sHelp, 1, 9);
      slAusgabe.Strings[i] := sHelp;
    end;

  //Header and Footer
  slAusgabe.Insert(0,'Jubiläumsliste für '+year(datetostr(date))+#13#10);
  slAusgabe.add(#13#10+sMess+
                #13#10+'Einstellbar in '+sIniFile+#13#10+
                #13#10+'Erzeugt von GE_Kart '+ GetProductVersionString +' am: '+datetostr(date));

  ShowDruckenDlg;        //Ergebnis anzeigen

end;

procedure TfrmMain.mnuOpenDebugClick(Sender: TObject);
begin
  opendocument(sDebugFile);
end;

procedure TfrmMain.mnuPrintDBStruktClick(Sender: TObject);
var
  i     : integer;

begin
  slAusgabe.Clear;
  //Datenbank öffnen
  frmDM.dbStatus(true);
  frmDM.dsetHelp.sql.Text := 'select * from '+global.sPersTablename;
  frmDM.dsetHelp.open;
  for i := 0 to frmDM.dsetHelp.FieldDefs.Count-1 do
    slAusgabe.add(appendChar(frmDM.dsetHelp.FieldDefs.Items[i].Name,' ', 25)+
                  Format('%5s ',[inttostr(frmDM.dsetHelp.FieldDefs.Items[i].Size)])+
                  FieldTypeToString(frmDM.dsetHelp.FieldDefs.Items[i].DataType));
  frmDM.dsetHelp.close;
  frmDM.dbStatus(false); //Datenbank schliessen

  slAusgabe.add(#13#10+'Erzeugt von GE_Kart '+ GetProductVersionString +' am: '+datetostr(date));

  ShowDruckenDlg;        //Ergebnis anzeigen
end;


Procedure TfrmMain.PerpareDatabaseForPrint(AskForAge: boolean = true);

var sOrder, sWhere:  String;

begin
  sWhere := '';
  sOrder := '';

  //Order by ermitteln
  if      mnuSortName.checked     then sOrder := ' order by Upper(tempstring)'
  else if mnuSortOrt.checked      then sOrder := ' order by PLZ, Ort, Ortsteil, Strasse, Upper(tempstring)'
  else if mnuSortGeb.checked      then sOrder := ' order by strftime(''%m%d'',geburtstag), Upper(tempstring)'  // %j day of year: 001-366 //Geht nicht bei Schaltjahr
  else if mnuSortGebAlter.checked then sOrder := ' order by strftime(''%J'',geburtstag), Upper(tempstring)'    // %J Julian day number
  else if mnuSortTauf.checked     then sOrder := ' order by strftime(''%m%d'',TaufDatum), Upper(tempstring)';

  //Filter
  if mnuMarkierte.checked          then sWhere := SQL_Where_Add(sWhere, 'Markiert<>0');
  if not mnuIncludeAbgang.Checked  then sWhere := SQL_Where_Add(sWhere, 'Abgang=0');
  if sWhere <> '' then sWhere := ' where ' + sWhere;

  //DB bereinigen
  frmDM.ExecSQL(sSQL_ClearMark1);
  frmDM.ExecSQL(sSQL_ClearMark2);
  frmDM.ExecSQL(sSQL_ClearAbgang1);
  frmDM.ExecSQL(sSQL_ClearAbgang2);

  // Altersberechnung:
  if AskForAge and
     (StrToInt(FormatDateTime('MM',Now)) > 10) and
     (MessageDlg('Für Listen mit Alter: Ausgeben für das Folgejahr?', mtConfirmation, [mbYes, mbNo],0) = mrYes) //Ab November für das Folgejahr?
    then frmDM.dsetHelp.SQL.Text:='select *, (strftime(''%Y'', ''now'') - strftime(''%Y'', geburtstag)) + 1 as Age from personen'
    else frmDM.dsetHelp.SQL.Text:='select *, (strftime(''%Y'', ''now'') - strftime(''%Y'', geburtstag)) as Age from personen';

  frmDM.dsetHelp.SQL.Add(sWhere+sOrder);

  frmDM.dsetHelp.Open;

  frmDM.dsetHelp1.IndexFieldNames := 'Geburtstag Asc';
  frmDM.dsetHelp1.LinkedFields    := 'PersonenID';
  frmDM.dsetHelp1.MasterFields    := 'PersonenID';
  frmDM.dsetHelp1.MasterSource    := frmDM.dsHelp;
  frmDM.dsetHelp1.SortedFields    := 'Geburtstag';
  frmDM.dsetHelp1.SQL.Text        := 'select * from Kinder';
  frmDM.dsetHelp1.Open;
end;

Procedure TfrmMain.PerparePrint(Formular: string; CallDesigner: boolean);
begin
  PerpareDatabaseForPrint;
  if Formular <> '' then frReport.LoadFromFile(Formular);

  if frReport.Variables.IndexOf('Adresse')         <0 then frReport.Variables.Add('Adresse');
  if frReport.Variables.IndexOf(' Adr1')           <0 then frReport.Variables.Add(' Adr1');
  if frReport.Variables.IndexOf(' Adr2')           <0 then frReport.Variables.Add(' Adr2');
  if frReport.Variables.IndexOf(' Kennzeichnung')  <0 then frReport.Variables.Add(' Kennzeichnung');

  if CallDesigner
    then frReport.DesignReport
    else frReport.ShowReport;
  frReport.Clear;
  CloseDatabaseAfterPrint;
end;

Procedure TfrmMain.CloseDatabaseAfterPrint;
begin
  //DB wieder schliessen
  frmDM.dsetHelp1.Close;
  frmDM.dsetHelp.Close;

  //aufräumen
  frmDM.dsetHelp1.IndexFieldNames := '';
  frmDM.dsetHelp1.LinkedFields    := '';
  frmDM.dsetHelp1.MasterFields    := '';
  frmDM.dsetHelp1.SortedFields    := '';
  frmDM.dsetHelp1.SQL.Text        := '';

  frmDM.dbStatus(false); // DB schliessen
end;

procedure TfrmMain.mnuPrintFormularClick(Sender: TObject);

begin
  openDialog.Title       := 'Selektiere eine Reportdatei Datei';
  openDialog.InitialDir  := UTF8ToSys(sAppDir+'ReportsFertig\');
  openDialog.Options := [ofFileMustExist];
  openDialog.Filter := 'Alle|*.*|Reportdatei |*.lrf';
  openDialog.FilterIndex := 2;
  if openDialog.Execute then PerparePrint(openDialog.FileName, false);
end;

procedure TfrmMain.mnuReportDesignerClick(Sender: TObject);
begin
  PerparePrint('', true);
end;

procedure tfrmMain.ShowDruckenMnu;

begin
  //Druckmenu anzeigen
  postmessage(handle,wm_syscommand,sc_keymenu, 0);  {Idee by Mattias Leonhardt}
  postmessage(handle,wm_keydown,vk_right,0);
  postmessage(handle,wm_keydown,vk_return,0);
end;

procedure TfrmMain.mnuRBClick(Sender: TObject);

begin
  if Sender is TMenuItem then TMenuItem(Sender).checked := true;
  ShowDruckenMnu;
end;

procedure TfrmMain.mnuSelectDatabaseClick(Sender: TObject);
var
  CloseAction : TCloseAction = caNone;
begin
  openDialog.Title        := 'Selektiere eine Datenbank';
  openDialog.InitialDir   := UTF8ToSys(ExtractFilePath(sDatabase));  // Set up the starting directory to be the current one
  openDialog.Options      := [ofFileMustExist];          // Only allow existing files to be selected
  openDialog.Filter       := 'Alle|*.*|DB-Datei |*.db';
  openDialog.FilterIndex  := 2;                          // Select report files as the starting filter type
  if openDialog.Execute                                 // Display the open file dialog
    then
      begin
        FormClose(sender, CloseAction);
        sDatabase := openDialog.FileName;
        bDatabaseVersionChecked := false;
        myDebugLN('New Database : '+sDatabase);
        FormShow(sender);                               // Prüft die Datei und stellt das Main Formular richtig dar.
        help.WriteIniVal(sIniFile, 'Datenbank','Datei', sDatabase);
      end;
end;

procedure TfrmMain.mnuSelectReportFileClick(Sender: TObject);
begin
  openDialog.Title       := 'Selektiere eine Reportdatei';
  openDialog.InitialDir  := UTF8ToSys(sAppDir+'ReportsFertig\');
  openDialog.Options     := [ofFileMustExist];
  openDialog.Filter      := 'Alle|*.*|Reportdatei |*.lrf';
  openDialog.FilterIndex := 2;
  if openDialog.Execute
    then
      begin
        sReportFile := openDialog.FileName;
        help.WriteIniVal(sIniFile, 'Reports','Karteikarte',sReportFile);
      end;
end;

procedure TfrmMain.mnuShowOverViewClick(Sender: TObject);

begin
  mnuShowOverView.Checked := not mnuShowOverView.Checked;
  if mnuShowOverView.Checked
    then help.WriteIniVal(sIniFile, 'Defaults', 'ShowOverView', '1')
    else help.WriteIniVal(sIniFile, 'Defaults', 'ShowOverView', '0');
end;

procedure TfrmMain.mnuShowTaufTagClick(Sender: TObject);

begin
  mnuShowTaufTag.Checked := not mnuShowTaufTag.Checked;
  if mnuShowTaufTag.Checked
    then help.WriteIniVal(sIniFile, 'Defaults', 'ShowTaufTag', '1')
    else help.WriteIniVal(sIniFile, 'Defaults', 'ShowTaufTag', '0');
end;

procedure TfrmMain.mnuSQLDebugClick(Sender: TObject);

begin
  bSQLDebug                 := not bSQLDebug;
  mnuSQLDebug.Checked       := bSQLDebug;
  frmDM.ZSQLMonitor.Active  := bSQLDebug;
  if bSQLDebug
    then help.WriteIniVal(sIniFile, 'Debug', 'SQLDebug', 'TRUE')
    else help.WriteIniVal(sIniFile, 'Debug', 'SQLDebug', 'FALSE');
end;

procedure TfrmMain.mnuStatistikClick(Sender: TObject);

  procedure einfuegen(text,text2:string);                   {textformatierung}

  begin
    slAusgabe.Add(AppendChar(text,' ',32)+text2);
  end;

  procedure einfuegen2(text,text2:string);                  {textformatierung}

  begin
    if text2 <> '' then slAusgabe.add(#13#10+text+#13#10+text2);
  end;

  Procedure SetAbgang;

  begin
    frmDM.dsetPersonen.edit;
    frmDM.dsetPersonen.fieldByName('Abgang').asboolean   := true;
    frmDM.dsetPersonen.fieldByName('Markiert').asboolean := true;
    frmDM.dsetPersonen.post;
  end;

  function Balken (len: real) : string;

  var i : integer;
      s : string;

  begin
    i := trunc(len);
    str(i, s);
    s := '%'+s+'s';
    s := Format(s, [' ']);
    s := replacechar(s, ' ', '*');
    Balken := ' '+s;
  end;

var
    ok           : boolean;
    nRecords,
    erw,
    kinder,
    n_konf_erw,
    erw_taufe,
    nw,
    nm,
    nw10,      //Weiblich  0-10 Jahre
    nw20,      //Weiblich 11-20 Jahre
    nw30,      //Weiblich 21-30 Jahre
    nw40,      //Weiblich 31-40 Jahre
    nw50,      //Weiblich 41-50 Jahre
    nw60,      //Weiblich 51-60 Jahre
    nw70,      //Weiblich 61-70 Jahre
    nw80,      //Weiblich 71-80 Jahre
    nw90,      //Weiblich 81-90 Jahre
    nw100,     //Weiblich 91-100 Jahre
    nw101,     //Weiblich 101.. Jahre
    nm10,      //Männlich  0-10 Jahre
    nm20,      //Männlich 11-20 Jahre
    nm30,      //Männlich 21-30 Jahre
    nm40,      //Männlich 31-40 Jahre
    nm50,      //Männlich 41-50 Jahre
    nm60,      //Männlich 51-60 Jahre
    nm70,      //Männlich 61-70 Jahre
    nm80,      //Männlich 71-80 Jahre
    nm90,      //Männlich 81-90 Jahre
    nm100,     //Männlich 91-100 Jahre
    nm101,     //Männlich 101.. Jahre
    n6,
    n13,
    n17,
    n29,
    n39,
    n49,
    n65,
    n74,
    n75,
    alter,
    GeburtsJahr,
    konf,
    k_tauf,
    e_tauf,
    TaufJahr,
    summe,
    trau,
    zugang,
    uebergetr,
    ueberweis,
    uebern,
    EinTritt,
    ausgesch,
    abgang,
    abgang_old,
    verst,
    austr,
    konvertzu,
    komm,
    restan,
    SummKomm,
    PersSummKomm   : integer;
    sKommListe,
    sRestListe,
    sFilter,
    sAbgaenge,
    sZugang,
    sPerson,
    sOldYearAbgang,
    sN_Konf_Erw,
    sErw_Getaufte,
    sNewYearZugang,
    sInvalidGeschlecht,
    sInvalidKomm,
    sInvalidGebTag,
    sInvalidTaufTag : string;
    dt : TDateTime;

begin
  //Datenbank öffnen
  frmDM.dbStatus(true, true);

  //Init
  slAusgabe.clear;
  nRecords           := 0;
  erw                := 0;
  n_konf_erw         := 0;
  erw_taufe          := 0;
  kinder             := 0;
  konf               := 0;
  k_tauf             := 0;
  e_tauf             := 0;
  summe              := 0;
  trau               := 0;
  zugang             := 0;
  uebergetr          := 0;
  ueberweis          := 0;
  uebern             := 0;
  EinTritt           := 0;
  ausgesch           := 0;
  abgang             := 0;
  abgang_old         := 0;
  verst              := 0;
  austr              := 0;
  konvertzu          := 0;
  komm               := 0;
  restan             := 0;
  SummKomm           := 0;
  nw                 := 0;
  nm                 := 0;
  nw10               := 0;
  nw20               := 0;
  nw30               := 0;
  nw40               := 0;
  nw50               := 0;
  nw60               := 0;
  nw70               := 0;
  nw80               := 0;
  nw90               := 0;
  nw100              := 0;
  nw101              := 0;
  nm10               := 0;
  nm20               := 0;
  nm30               := 0;
  nm40               := 0;
  nm50               := 0;
  nm60               := 0;
  nm70               := 0;
  nm80               := 0;
  nm90               := 0;
  nm100              := 0;
  nm101              := 0;
  n6                 := 0;
  n13                := 0;
  n17                := 0;
  n29                := 0;
  n39                := 0;
  n49                := 0;
  n65                := 0;
  n74                := 0;
  n75                := 0;
  sKommListe         := '';
  sRestListe         := '';
  sInvalidGeschlecht := '';
  sInvalidGebTag     := '';
  sInvalidTaufTag    := '';
  sInvalidKomm       := '';
  sAbgaenge          := '';
  sZugang            := '';
  sN_Konf_Erw        := '';
  sErw_Getaufte      := '';
  sFilter            := '';

  dt := now();
  if FormatDateTime('MM', dt) = '01' then dt := dt-365;
  frmStatInfo.stat_jahr.Text        := FormatDateTime('YYYY', dt);
  frmStatInfo.stat_jahr.Enabled     := true;
  frmStatInfo.stat_Kirche.Enabled   := true;
  frmStatInfo.stat_gemeinde.Enabled := true;
  frmStatInfo.stat_gemeinde.Text    := sDefaultGemeinde;
  if frmStatInfo.ShowModal = mrOK
    then
      begin
        try
          screen.cursor := crhourglass;

          frmDM.ExecSQL(global.sSQL_DelMark);
          frmDM.ExecSQL(global.sSQL_DelAbgang);

          //Filter auf Gemeinde und Kirche
          if frmStatInfo.stat_gemeinde.Text = ''
            then sFilter := SQL_Where_Add(sFilter, SQL_Where_IsNull('Gemeinde'))
            else if frmStatInfo.stat_gemeinde.Text <> sGemeindenAlle
              then sFilter := SQL_Where_Add(sFilter, 'Gemeinde=''' + frmStatInfo.stat_gemeinde.Text + '''');
          if frmStatInfo.stat_Kirche.text <> ''
            then sFilter := SQL_Where_Add(sFilter, 'Kirche=''' + frmStatInfo.stat_Kirche.Text + '''');
          myDebugLN('Setze Filter: ' + sFilter);
          frmDM.dsetPERSONEN.Filter   := sFilter;
          frmDM.dsetPERSONEN.Filtered := (sFilter <> '');
          frmDM.dsetPERSONEN.Refresh;
          frmDM.dsetPERSONEN.First;
          myDebugLN(inttostr(frmDM.dsetPersonen.RecordCount));

          while not frmDM.dsetPersonen.EOF do
            begin
              inc(nRecords);
              StatusBar.Panels[0].Text:='Bearbeite Datensatz: '+inttostr(nRecords) + ' von ' + inttostr(frmDM.dsetPersonen.RecordCount);
              statusbar.Update;

              sPerson := frmDM.dsetPersonen.fieldByName('Vorname').asstring +' ' + frmDM.dsetPersonen.fieldByName('Nachname').asstring;
              myDebugLN('Bearbeite Person: ' + sPerson);

              try
                TaufJahr        := strtoint(year(frmDM.dsetPersonen.fieldByName('TaufDatum').asstring));
              except
                TaufJahr        := -1;
                sInvalidTaufTag += sPerson + #13#10;
              end;

              try
                GeburtsJahr    := strtoint(year(frmDM.dsetPersonen.fieldByName('Geburtstag').asstring));
                alter          := strtoint(frmStatInfo.stat_jahr.text) - GeburtsJahr;
              except
                GeburtsJahr    := 0;
                alter          := -1;
                sInvalidGebTag += sPerson + #13#10;
              end;

              {die eigentliche statistik}
              if year(frmDM.dsetPersonen.fieldByName('KonfDatum').asstring) = frmStatInfo.stat_jahr.text then inc(konf);
              if year(frmDM.dsetPersonen.fieldByName('TrauDatum').asstring) = frmStatInfo.stat_jahr.text then inc(trau);

              //Zugänge
              //Zugänge aus Folgejahren NICHT mitzählen
              if year(frmDM.dsetPersonen.fieldByName('TaufDatum').asstring) = frmStatInfo.stat_jahr.text then
                begin
                  if alter < 15
                    then inc(k_tauf)
                    else inc(e_Tauf);
                  inc(zugang);
                  sZugang += sPerson+#13#10;
                end;
              if year(frmDM.dsetPersonen.fieldByName('UebertrittsZuDatum').asstring) = frmStatInfo.stat_jahr.text then
                begin
                  inc(uebergetr);
                  inc(zugang);
                  sZugang += sPerson+#13#10
                end;
              if year(frmDM.dsetPersonen.fieldByName('Ueberwiesen_von_Datum').asstring) = frmStatInfo.stat_jahr.text then
                begin
                  inc(ueberweis);
                  inc(zugang);
                  sZugang += sPerson+#13#10
                end;
              if year(frmDM.dsetPersonen.fieldByName('EintrittsDatum').asstring) = frmStatInfo.stat_jahr.text then
                begin
                  inc(EinTritt);
                  inc(zugang);
                  sZugang += sPerson+#13#10
                end;

              //Zugänge aus Folgejahren
              sNewYearZugang := '';
              if (year(frmDM.dsetPersonen.fieldByName('TaufDatum').asstring) <> '') and
                 (year(frmDM.dsetPersonen.fieldByName('TaufDatum').asstring) >    frmStatInfo.stat_jahr.text)
                then
                  begin
                    sNewYearZugang := year(frmDM.dsetPersonen.fieldByName('TaufDatum').asstring);
                  end;

              if (year(frmDM.dsetPersonen.fieldByName('UebertrittsZuDatum').asstring) <> '') and
                 (year(frmDM.dsetPersonen.fieldByName('UebertrittsZuDatum').asstring) > frmStatInfo.stat_jahr.text)
                then
                  begin
                    sNewYearZugang := year(frmDM.dsetPersonen.fieldByName('UebertrittsZuDatum').asstring);
                  end;
              if (year(frmDM.dsetPersonen.fieldByName('Ueberwiesen_von_Datum').asstring) <> '') and
                 (year(frmDM.dsetPersonen.fieldByName('Ueberwiesen_von_Datum').asstring) > frmStatInfo.stat_jahr.text)
                then
                  begin
                    sNewYearZugang := year(frmDM.dsetPersonen.fieldByName('Ueberwiesen_von_Datum').asstring);
                  end;
              if (year(frmDM.dsetPersonen.fieldByName('EintrittsDatum').asstring) <> '') and
                 (year(frmDM.dsetPersonen.fieldByName('EintrittsDatum').asstring) > frmStatInfo.stat_jahr.text)
                then
                  begin
                    sNewYearZugang := year(frmDM.dsetPersonen.fieldByName('EintrittsDatum').asstring);
                  end;

              //Abgänge
              if year(frmDM.dsetPersonen.fieldByName('Ueberwiesen_nach_Datum').asstring) = frmStatInfo.stat_jahr.text
                then
                  begin
                    inc(uebern);
                    inc(abgang);
                    SetAbgang;
                  end;
              if year(frmDM.dsetPersonen.fieldByName('AustrittsDatum').asstring) = frmStatInfo.stat_jahr.text
                then
                  begin
                    inc(austr);
                    inc(abgang);
                    SetAbgang;
                  end;
              if year(frmDM.dsetPersonen.fieldByName('AusschlussDatum').asstring) = frmStatInfo.stat_jahr.text
                then
                  begin
                    inc(ausgesch);
                    inc(abgang);
                    SetAbgang;
                  end;
              if year(frmDM.dsetPersonen.fieldByName('TodesDatum').asstring) = frmStatInfo.stat_jahr.text
                then
                  begin
                    inc(verst);
                    inc(abgang);
                    SetAbgang;
                  end;
              if year(frmDM.dsetPersonen.fieldByName('UebertrittsAbDatum').asstring) = frmStatInfo.stat_jahr.text
                then
                  begin
                    inc(KONVERTZU);
                    inc(abgang);
                    SetAbgang;
                  end;

              //Abgänge aus älteren Jahren
              sOldYearAbgang := '';
              if (year(frmDM.dsetPersonen.fieldByName('Ueberwiesen_nach_Datum').asstring) <> '') and
                 (year(frmDM.dsetPersonen.fieldByName('Ueberwiesen_nach_Datum').asstring) < frmStatInfo.stat_jahr.text) then
                begin
                  inc(abgang_old);
                  SetAbgang;
                  sOldYearAbgang := year(frmDM.dsetPersonen.fieldByName('Ueberwiesen_nach_Datum').asstring);
                end;
              if (year(frmDM.dsetPersonen.fieldByName('UebertrittsAbDatum').asstring) <> '') and
                 (year(frmDM.dsetPersonen.fieldByName('UebertrittsAbDatum').asstring) < frmStatInfo.stat_jahr.text) then
                begin
                  inc(abgang_old);
                  SetAbgang;
                  sOldYearAbgang := year(frmDM.dsetPersonen.fieldByName('UebertrittsAbDatum').asstring);
                end;
              if (year(frmDM.dsetPersonen.fieldByName('AustrittsDatum').asstring) <> '') and
                 (year(frmDM.dsetPersonen.fieldByName('AustrittsDatum').asstring) < frmStatInfo.stat_jahr.text) then
                begin
                  inc(abgang_old);
                  SetAbgang;
                  sOldYearAbgang := year(frmDM.dsetPersonen.fieldByName('AustrittsDatum').asstring);
                end;
              if (year(frmDM.dsetPersonen.fieldByName('AusschlussDatum').asstring) <> '') and
                 (year(frmDM.dsetPersonen.fieldByName('AusschlussDatum').asstring) < frmStatInfo.stat_jahr.text) then
                begin
                  inc(abgang_old);
                  SetAbgang;
                  sOldYearAbgang := year(frmDM.dsetPersonen.fieldByName('AusschlussDatum').asstring);
                end;
              if (year(frmDM.dsetPersonen.fieldByName('TodesDatum').asstring) <> '') and
                 (year(frmDM.dsetPersonen.fieldByName('TodesDatum').asstring) < frmStatInfo.stat_jahr.text) then
                begin
                  inc(abgang_old);
                  SetAbgang;
                  sOldYearAbgang := year(frmDM.dsetPersonen.fieldByName('TodesDatum').asstring);
                end;

              {Kommunikantenberechnung}
              ok           := false;
              PersSummKomm := 0;
              frmDM.dsetKomm.first;
              while not frmDM.dsetKomm.eof do
                begin
                  if year(frmDM.dsetKomm.FieldByName('AbendmahlsDatum').asstring) = frmStatInfo.stat_jahr.text
                    then
                      begin
                        ok := true;
                        inc(SummKomm);
                        inc(PersSummKomm)
                      end;
                  frmDM.dsetKomm.next;
                end;

              if not frmDM.dsetPersonen.fieldByName('Abgang').asboolean
                then
                  begin
                    if (frmDM.dsetPersonen.fieldByName('KonfDatum').asstring <> '') or //Konfimirte
                       ((TaufJahr - GeburtsJahr) > 14) // Erwachsenentaufe
                      then
                        begin
                          if ok
                            then
                              begin
                                inc(komm);
                                sKommListe += AppendChar(sPerson,' ',32) + Format('%3.0d', [PersSummKomm])+#13#10;
                              end
                            else
                              begin
                                inc(restan);
                                sRestListe += AppendChar(sPerson,' ',32) + Format('%3.0d', [PersSummKomm])+#13#10;
                              end;
                          end
                        else if ok then sInvalidKomm += sPerson+#13#10;
                  end;

              if not frmDM.dsetPersonen.fieldByName('Abgang').asboolean
                then // zu zählende Personen
                  begin
                    if sNewYearZugang = ''
                      then
                        begin
                          if (frmDM.dsetPersonen.fieldByName('KonfDatum').asstring <> '') or
                             ((TaufJahr - GeburtsJahr) > 14)  // Erwachsenentaufe
                            then
                              begin
                                inc(erw);
                                if ((TaufJahr - GeburtsJahr) > 14)
                                  then
                                    begin
                                      inc(erw_taufe);
                                      sErw_Getaufte += sPerson+ #13#10;
                                    end;
                              end
                            else
                              if alter > 17
                                then
                                  begin
                                    inc(n_konf_erw);
                                    sN_Konf_Erw += sPerson+ #13#10;
                                  end
                                else inc(kinder);
                          inc(summe);

                          if uppercase(frmDM.dsetPersonen.fieldByName('Geschlecht').asstring) = 'W'
                            then
                              begin
                                inc(nw);
                                case alter of
                                   0 ..10:  inc(nw10 );
                                  11 ..20:  inc(nw20 );
                                  21 ..30:  inc(nw30 );
                                  31 ..40:  inc(nw40 );
                                  41 ..50:  inc(nw50 );
                                  51 ..60:  inc(nw60 );
                                  61 ..70:  inc(nw70 );
                                  71 ..80:  inc(nw80 );
                                  81 ..90:  inc(nw90 );
                                  91 ..100: inc(nw100);
                                  101..120: inc(nw101);
                                end;
                              end
                            else
                              begin
                                if uppercase(frmDM.dsetPersonen.fieldByName('Geschlecht').asstring) = 'M'
                                  then
                                    begin
                                      inc(nm);
                                      case alter of
                                         0 ..10:  inc(nm10 );
                                        11 ..20:  inc(nm20 );
                                        21 ..30:  inc(nm30 );
                                        31 ..40:  inc(nm40 );
                                        41 ..50:  inc(nm50 );
                                        51 ..60:  inc(nm60 );
                                        61 ..70:  inc(nm70 );
                                        71 ..80:  inc(nm80 );
                                        81 ..90:  inc(nm90 );
                                        91 ..100: inc(nm100);
                                        101..120: inc(nm101);
                                      end;
                                    end
                                  else
                                    begin
                                       sInvalidGeschlecht += sPerson+ #13#10;
                                    end;
                              end;
                          case alter of
                             0 .. 6:  inc(n6 );
                             7 ..13:  inc(n13);
                            14 ..17:  inc(n17);
                            18 ..29:  inc(n29);
                            30 ..39:  inc(n39);
                            40 ..49:  inc(n49);
                            50 ..65:  inc(n65);
                            66 ..74:  inc(n74);
                            75..120:  inc(n75);
                          end;
                        end
                      else
                        begin
                          sZugang += sPerson+ ' (Zugang im Jahr '+sNewYearZugang+'; wird NICHT mitgezählt!)'+#13#10;
                        end;
                  end
                else {Abgänge}
                  begin
                    sAbgaenge += sPerson;
                    if sOldYearAbgang <> '' then sAbgaenge += ' (Abgang im Jahr '+sOldYearAbgang+'); wird NICHT mitgezählt!';
                    sAbgaenge += #13#10;
                  end;

              frmDM.dsetPersonen.Next;
            end;

          //und jetzt die Ausgabe
          slAusgabe.add('*****************************************************');
          slAusgabe.add('Eingabeparameter Jahr       : "'+frmStatInfo.stat_jahr.text+'"');
          slAusgabe.add('                 Konfession : "'+frmStatInfo.stat_Kirche.text+'"');
          slAusgabe.add('                 Gemeinde   : "'+frmStatInfo.stat_gemeinde.text+'"');
          slAusgabe.add('*****************************************************');
          If Summe = 0
            then
              begin
                slAusgabe.add(' ');
                slAusgabe.add('Keine Personen gefunden');
              end
            else
              begin
                slAusgabe.add(' ');
                einfuegen('Jahresstatistik','');
                slAusgabe.add(' ');
                einfuegen('Konfirmierte    : '+inttostr(erw),RealToStr(erw/summe*100,2,2)+ ' %');
                einfuegen('   davon Erwachsengetaufte: '+inttostr(erw_taufe),'');
                einfuegen('Nicht-Konfirm.  : '+inttostr(kinder),RealToStr(kinder/summe*100,2,2)+ ' %');
                einfuegen('Nicht-Konf. Erw : '+inttostr(n_konf_erw),RealToStr(n_konf_erw/summe*100,2,2)+ ' %');
                einfuegen('Summe           : '+inttostr(summe),'');
                einfuegen('Bemerkungen zu dieser Summe: Sie enthält KEINE Abgänge.','');
                slAusgabe.add(' ');
                slAusgabe.add('Altersstruktur');
                einfuegen('Weiblich             : '+inttostr(nw)   ,RealToStr(nw    /summe*100,5,2)+ ' %'); //Hier könnte noch ein Balken ausgegeben werden
                einfuegen('Weiblich  0-10 Jahre : '+inttostr(nw10 ),RealToStr(nw10  /summe*100,5,2)+ ' %'+Balken(nw10  /summe*300));
                einfuegen('Weiblich 11-20 Jahre : '+inttostr(nw20 ),RealToStr(nw20  /summe*100,5,2)+ ' %'+Balken(nw20  /summe*300));
                einfuegen('Weiblich 21-30 Jahre : '+inttostr(nw30 ),RealToStr(nw30  /summe*100,5,2)+ ' %'+Balken(nw30  /summe*300));
                einfuegen('Weiblich 31-40 Jahre : '+inttostr(nw40 ),RealToStr(nw40  /summe*100,5,2)+ ' %'+Balken(nw40  /summe*300));
                einfuegen('Weiblich 41-50 Jahre : '+inttostr(nw50 ),RealToStr(nw50  /summe*100,5,2)+ ' %'+Balken(nw50  /summe*300));
                einfuegen('Weiblich 51-60 Jahre : '+inttostr(nw60 ),RealToStr(nw60  /summe*100,5,2)+ ' %'+Balken(nw60  /summe*300));
                einfuegen('Weiblich 61-70 Jahre : '+inttostr(nw70 ),RealToStr(nw70  /summe*100,5,2)+ ' %'+Balken(nw70  /summe*300));
                einfuegen('Weiblich 71-80 Jahre : '+inttostr(nw80 ),RealToStr(nw80  /summe*100,5,2)+ ' %'+Balken(nw80  /summe*300));
                einfuegen('Weiblich 81-90 Jahre : '+inttostr(nw90 ),RealToStr(nw90  /summe*100,5,2)+ ' %'+Balken(nw90  /summe*300));
                einfuegen('Weiblich 91-100 Jahre: '+inttostr(nw100),RealToStr(nw100 /summe*100,5,2)+ ' %'+Balken(nw100 /summe*300));
                einfuegen('Weiblich 101.. Jahre : '+inttostr(nw101),RealToStr(nw101 /summe*100,5,2)+ ' %'+Balken(nw101 /summe*300));
                einfuegen('Männlich             : '+inttostr(nm)   ,RealToStr(nm    /summe*100,5,2)+ ' %');
                einfuegen('Männlich  0-10 Jahre : '+inttostr(nm10 ),RealToStr(nm10  /summe*100,5,2)+ ' %'+Balken(nm10  /summe*300));
                einfuegen('Männlich 11-20 Jahre : '+inttostr(nm20 ),RealToStr(nm20  /summe*100,5,2)+ ' %'+Balken(nm20  /summe*300));
                einfuegen('Männlich 21-30 Jahre : '+inttostr(nm30 ),RealToStr(nm30  /summe*100,5,2)+ ' %'+Balken(nm30  /summe*300));
                einfuegen('Männlich 31-40 Jahre : '+inttostr(nm40 ),RealToStr(nm40  /summe*100,5,2)+ ' %'+Balken(nm40  /summe*300));
                einfuegen('Männlich 41-50 Jahre : '+inttostr(nm50 ),RealToStr(nm50  /summe*100,5,2)+ ' %'+Balken(nm50  /summe*300));
                einfuegen('Männlich 51-60 Jahre : '+inttostr(nm60 ),RealToStr(nm60  /summe*100,5,2)+ ' %'+Balken(nm60  /summe*300));
                einfuegen('Männlich 61-70 Jahre : '+inttostr(nm70 ),RealToStr(nm70  /summe*100,5,2)+ ' %'+Balken(nm70  /summe*300));
                einfuegen('Männlich 71-80 Jahre : '+inttostr(nm80 ),RealToStr(nm80  /summe*100,5,2)+ ' %'+Balken(nm80  /summe*300));
                einfuegen('Männlich 81-90 Jahre : '+inttostr(nm90 ),RealToStr(nm90  /summe*100,5,2)+ ' %'+Balken(nm90  /summe*300));
                einfuegen('Männlich 91-100 Jahre: '+inttostr(nm100),RealToStr(nm100 /summe*100,5,2)+ ' %'+Balken(nm100 /summe*300));
                einfuegen('Männlich 101.. Jahre : '+inttostr(nm101),RealToStr(nm101 /summe*100,5,2)+ ' %'+Balken(nm101 /summe*300));
                slAusgabe.add(' ');
                slAusgabe.add('SELK Struktur');
                einfuegen(' 0-6   Jahre: '+inttostr(n6 ),RealToStr(n6 /summe*100,5,2)+ ' %'+Balken(n6 /summe*100));
                einfuegen(' 7-13  Jahre: '+inttostr(n13),RealToStr(n13/summe*100,5,2)+ ' %'+Balken(n13/summe*100));
                einfuegen(' 14-17 Jahre: '+inttostr(n17),RealToStr(n17/summe*100,5,2)+ ' %'+Balken(n17/summe*100));
                einfuegen(' 18-29 Jahre: '+inttostr(n29),RealToStr(n29/summe*100,5,2)+ ' %'+Balken(n29/summe*100));
                einfuegen(' 30-39 Jahre: '+inttostr(n39),RealToStr(n39/summe*100,5,2)+ ' %'+Balken(n39/summe*100));
                einfuegen(' 40-49 Jahre: '+inttostr(n49),RealToStr(n49/summe*100,5,2)+ ' %'+Balken(n49/summe*100));
                einfuegen(' 50-65 Jahre: '+inttostr(n65),RealToStr(n65/summe*100,5,2)+ ' %'+Balken(n65/summe*100));
                einfuegen(' 66-74 Jahre: '+inttostr(n74),RealToStr(n74/summe*100,5,2)+ ' %'+Balken(n74/summe*100));
                einfuegen(' 75..  Jahre: '+inttostr(n75),RealToStr(n75/summe*100,5,2)+ ' %'+Balken(n75/summe*100));
                slAusgabe.add(' ');

                if sInvalidGeschlecht <> '' then slAusgabe.add('WARNUNG Ungültiges Geschlecht bei folgenden Personen:'+#13#10+sInvalidGeschlecht);
                if sInvalidGebTag <> ''     then slAusgabe.add('WARNUNG Ungültiges Geburtsdatum bei folgenden Personen:'+#13#10+sInvalidGebTag);
                if sInvalidTaufTag <> ''    then slAusgabe.add('WARNUNG Ungültiges Taufdatum bei folgenden Personen:'+#13#10+sInvalidTaufTag);

                einfuegen('Kommunikanten  : '+inttostr(komm),'');
                einfuegen('Restanten      : '+inttostr(restan),'');
                einfuegen('Abendmahlsgänge: '+inttostr(SummKomm),'');
                if sInvalidKomm <> ''       then slAusgabe.add('WARNUNG Abendmahlsgänge bei nicht konf. Personen:'+#13#10+sInvalidKomm);

                slAusgabe.add('');
                einfuegen('Zugänge        : '+inttostr(zugang),'');
                einfuegen('Kindertaufen   : '+inttostr(k_tauf),'');
                einfuegen('Erw. Taufen    : '+inttostr(e_tauf)+' Formel: (Statistikjahr-Geb.Jahr)>15','');
                einfuegen('Eintritte      : '+inttostr(EinTritt),'');
                einfuegen('Übergetreten   : '+inttostr(uebergetr),'');
                einfuegen('Überweisungen  : '+inttostr(ueberweis),'');
                slAusgabe.add('');
                einfuegen('Abgänge        : '+inttostr(Abgang),'');
                einfuegen('Verstorben     : '+inttostr(verst),'');
                einfuegen('Austritte      : '+inttostr(austr),'');
                einfuegen('Ausschlüsse    : '+inttostr(ausgesch),'');
                einfuegen('Überweisungen  : '+inttostr(uebern),'');
                einfuegen('Übertritte     : '+inttostr(Konvertzu),'');
                slAusgabe.add('');
                einfuegen('Konfirmationen : '+inttostr(konf),'');
                if trau > 0
                  then
                    begin
                      slAusgabe.add('');
                      einfuegen('Bemerkung zu Trauungen: Wenn die Eheleute beide aus Ihrer','');
                      einfuegen('Gemeinde kommen, werden Sie doppelt gezählt!','');
                    end;
                einfuegen('Trauungen      : '+inttostr(trau),'');
                slAusgabe.add(#13#10);

                //Besuche
                sFilter := 'where strftime(''%Y'',besuch.Besuchsdatum) = '''+frmStatInfo.stat_jahr.text+''''+
                           ' and personen.Kirche='''+frmStatInfo.stat_Kirche.Text+''' ';
                if frmStatInfo.stat_gemeinde.Text <> '' then sFilter := sFilter +'and personen.Gemeinde='''+frmStatInfo.stat_gemeinde.Text+''' ';

                frmDM.dsetHelp.SQL.Text:='select count(besuch.PersonenID) as C from besuch '+
                                         'join personen on besuch.personenID=personen.personenID ' + sFilter;
                frmDM.dsetHelp.Open;
                einfuegen('Besuche ges.   : '+frmDM.dsetHelp.FieldByName('C').AsString + ' Alle eingetragenen Besuche','');
                frmDM.dsetHelp.Close;

                frmDM.dsetHelp.SQL.Text:='select besuch.PersonenID from besuch '+
                                         'join personen on besuch.personenID=personen.personenID '+
                                         sFilter+' '+
                                         'group by besuch.Besuchsdatum, besuch.Besuchsgrund';
                frmDM.dsetHelp.Open;
                einfuegen('Besuche        : '+inttostr(frmDM.dsetHelp.RecordCount) + ' Besuche, gruppiert nach Besuchsdatum und Besuchsgrund','');
                frmDM.dsetHelp.Close;

                frmDM.dsetHelp.SQL.Text:='select besuch.PersonenID from besuch '+
                                         'join personen on besuch.personenID=personen.personenID ' +
                                         sFilter+' '+
                                         'group by besuch.PersonenID';
                frmDM.dsetHelp.Open;
                einfuegen('Besuchte Pers. : '+inttostr(frmDM.dsetHelp.RecordCount) + ' Wie viele verschiedene Personen wurden besucht.','');
                frmDM.dsetHelp.Close;
                slAusgabe.add('');

                einfuegen2('Zugänge:',sZugang);

                if (Abgang > 0) or
                   (Abgang_old > 0)
                  then
                    begin
                      slAusgabe.add(#13#10+'Es wurden ' +inttostr(Abgang)+' Abgang / Abgänge gefunden. Bei ihnen wurde das Feld "Abgang" gesetzt.');
                      if Abgang_old > 0 then slAusgabe.add('Es wurden ' +inttostr(Abgang_old)+' alte(r) Abgang / Abgänge gefunden. Bei ihnen wurde das Feld "Abgang" gesetzt.');
                      slAusgabe.add('Die gefundenen Personen wurden markiert.');
                      slAusgabe.add('Zur Dokumentation können die Karteikarten geduckt werden und danach können die Personen gelöscht werden.');
                      einfuegen2('Abgänge:',sAbgaenge);
                    end;
                einfuegen2('Nicht konfirmirte Erwachsene:',sN_Konf_Erw);
                einfuegen2('Erwachsenengetaufte:',sErw_Getaufte);
                if (komm>0)
                  then
                    begin
                      einfuegen2('Kommunikantenliste:',sKommListe);
                      einfuegen2('Restantenliste:',sRestListe);
                    end;

                slAusgabe.add(#13#10+'Erzeugt von GE_Kart '+ GetProductVersionString +' am: '+datetostr(date));
              end;
        finally
          screen.cursor := crdefault;
        end;
        StatusBar.Panels[0].Text:='fertig';
        statusbar.Update;
        ShowDruckenDlg; //Ergebnis anzeigen
      end;
  frmDM.dbStatus(false); //Datenbank schliessen
end;

procedure TfrmMain.mnuUserDefSQLClick(Sender: TObject);

begin
  slHelp.Clear;
  try
    slHelp.LoadFromFile(sUserDefSQLFile);
  except
    slHelp.Text := '';
  end;
  ExecSQL(slHelp.Text, frmDM.dsetHelp, false);
end;

procedure TfrmMain.mnuVolltextsucheClick(Sender: TObject);

var sSuchText,
    sFieldText,
    sMess,
    sMess2     : string;
    nRecords,
    nFound,
    i          : integer;

begin
  //Datenbank öffnen
  frmDM.dbStatus(true);
  if MessageDlg('Löschen?', 'Sollen die alten Markierungen gelöscht werden?', mtConfirmation, [mbYes, mbNo],0) = mrYes
    then frmDM.ExecSQL(global.sSQL_DelMark);
  frmInput.SetDefaults('Volltextsuche', 'Suchtext', '', '', '', false);
  if frmInput.ShowModal = mrOK then
    begin
      frmDM.dsetHelp.SQL.Text:='select * from personen';
      frmDM.dsetHelp.Open;
      frmDM.dsetHelp.First;
      nRecords := 0;
      nFound   := 0;
      sSuchtext:=uppercase(frmInput.Edit1.Text);
      sMess    := 'Gesucht wurde: "'+sSuchtext+'"'+#13#10;
      sMess2   := 'Gefunden in:'+#13#10;
      while not frmDM.dsetHelp.EOF do
      begin
        inc(nRecords);
        StatusBar.Panels[0].Text:='Suche in Datensatz: '+inttostr(nRecords);
        statusbar.Update;
        for i := 0 to frmDM.dsetHelp.Fields.Count-1 do
          begin
            case frmDM.dsetHelp.FieldDefs.Items[i].DataType of
              ftString,
              ftMemo  : sFieldText := uppercase(frmDM.dsetHelp.Fields.Fields[i].AsString);
              ftDate  : sFieldText := FormatDateTime('dd.mm.yyyy' ,frmDM.dsetHelp.Fields.Fields[i].AsDateTime);
              else      sFieldText := '';
            end;
            if pos(sSuchText, sFieldText) <> 0 then
              begin
                sMess2 := sMess2 + AppendChar(frmDM.dsetHelp.FieldByName('Vorname').AsString+ ' '+ frmDM.dsetHelp.FieldByName('Nachname').AsString, ' ', 40)+
                                   'Feld: '+frmDM.dsetHelp.FieldDefs.Items[i].Name+#13#10;
                frmDM.dsetHelp.Edit;
                frmDM.dsetHelp.FieldByName('Markiert').Asboolean := true;
                frmDM.dsetHelp.Post;
                inc(nFound);
                break;
              end;
            end;
        frmDM.dsetHelp.Next;
      end;
      frmDM.dsetHelp.close;
      sMess := sMess + 'Durchsuche Datensätze: '+inttostr(nRecords)+#13#10+
                       'Gefundene Datensätze:  '+inttostr(nFound)+#13#10;
      if nFound = 0
        then Showmessage(sMess)
        else
          begin
            frmAusgabe.SetDefaults('Suchergebnisse', sMess + sMess2, '', '', 'Schliessen', false);
            frmAusgabe.ShowModal;
          end;
    end;
  frmDM.dbStatus(false); //Datenbank schliessen
end;

end.

