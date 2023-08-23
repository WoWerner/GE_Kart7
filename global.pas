unit Global;

{$mode objfpc}{$H+}

interface

uses
  Classes,
  SysUtils;

const
  sEMailAdr          = 'mail@w-werner.de';
  sHomePage          = 'www.w-werner.de';
  CSV_Delimiter      = ';';

  sFehltNoch         = 'Diese Funktion fehlt noch.'+#13#10+'Wenn Sie sie brauchen, melden Sie sich.';
  sInBearbeitung     = 'Diese Funktion ist in Bearbeitung.'+#13#10+'Die Ergebnisse können noch falsch sein.';

  nDefDPI            =   96;
  MaxWindWidth       = 1000; //Karteifenster
  MaxWindHeight      =  675; //Karteifenster

  //db-Zugriff
  sPersTablename     = 'PERSONEN';
  sKommTablename     = 'KOMMUNIONEN';
  sKindTablename     = 'KINDER';
  sBesuchTablename   = 'BESUCH';
  sGDTablename       = 'GOTTESDIENST';
  sUsersTablename    = 'Users';

  //SQL
  sSQL_KOMM_PER_JAHR = 'select count(strftime(''%%Y'', AbendmahlsDatum)) as C,'+
                       'strftime(''%%Y'', AbendmahlsDatum) as D '+
                       'from '+sKommTablename+' '+
                       'where (PersonenID = %s) '+
                       'GROUP BY D '+
                       'ORDER BY D DESC;';
  sSQL_Komm_ADD      = 'INSERT INTO '+sKommTablename+  ' (Abendmahlsdatum, PersonenID) VALUES (''%s'', %s)';
  sSQL_BESUCH_ADD    = 'INSERT INTO '+sBESUCHTablename+' (Besuchsdatum, Besuchsgrund, PersonenID) VALUES (''%s'', ''neuer Besuch'',%s)';
  sSQL_KIND_ADD      = 'INSERT INTO '+sKindTablename+  ' (Vorname, Vorname2, Nachname, PersonenID) VALUES (''Neues'', ''Kind'', ''hier bearbeiten'',%s)';
  //Gottesdienst anlegen
  sSQL_GD_ADD        = 'INSERT INTO '+sGDTablename+  ' (GottesdienstDatum, Tag, GottesdienstForm1, GottesdienstForm2, GottesdienstForm3) VALUES (''%s'', ''S'', ''P'', ''F'', ''F'')';

  //Einzelnen DS löschen
  sSQL_Komm_DEL      = 'Delete from '+sKommTablename+  ' where PersonenID=%s and KommID=%s';
  sSQL_Besuch_DEL    = 'Delete from '+sBesuchTablename+' where PersonenID=%s and BesuchID=%s';
  sSQL_KIND_DEL      = 'Delete from '+sKINDTablename+  ' where PersonenID=%s and KinderID=%s';
  sSQL_GD_DEL        = 'Delete from '+sGDTablename+    ' where GdID=%s';

  //Person löschen
  sSQL_Komm_Pes_DEL  = 'Delete from '+sKommTablename+  ' where PersonenID=%s';
  sSQL_Besuch_Pes_DEL= 'Delete from '+sBesuchTablename+' where PersonenID=%s';
  sSQL_Kind_Pes_DEL  = 'Delete from '+sKindTablename+  ' where PersonenID=%s';
  sSQL_Pers_DEL      = 'Delete from '+sPersTablename+  ' where PersonenID=%s';

  //Person anlegen
  sSQL_Pers_New      = 'INSERT INTO '+sPersTablename+  ' (Vorname, Nachname, Markiert, Abgang, Gemeinde) VALUES (''????'', ''????'', 0, 0, ''%s'')';

  //Markierungen
  sSQL_DelMark       = 'Update '+sPersTablename+' set markiert=0';
  sSQL_ClearMark1    = 'Update '+sPersTablename+' set markiert=0 where ((markiert=''N'') or (markiert=''F''))';
  sSQL_ClearMark2    = 'Update '+sPersTablename+' set markiert=1 where ((markiert=''Y'') or (markiert=''T''))';
  sSQL_InvertMark    = 'Update '+sPersTablename+' set markiert = not markiert';

  //Abgänge
  sSQL_DelAbgang     = 'Update '+sPersTablename+' set Abgang=0';
  sSQL_ClearAbgang1  = 'Update '+sPersTablename+' set Abgang=0 where ((Abgang=''N'') or (Abgang=''F''))';
  sSQL_ClearAbgang2  = 'Update '+sPersTablename+' set Abgang=1 where ((Abgang=''Y'') or (Abgang=''T''))';

  //Auswahlliste
  sSQL_Auswahlliste = 'select personenID, Vorname, Vorname2, NachName, GebName, Geburtstag, Taufdatum, Kirche from Personen '+
                      'where geburtstag >= ''%s-01-01'' and geburtstag <= ''%s-12-31'' %s order by Nachname, Vorname';
                      //3. Parameter Blank oder "and geschlecht ='M'" oder "and geschlecht ='W'"

  //Übersetzungstabellen für Import
  sBE5TOBE7          =  'BEWANN=BesuchsDatum'+#13#10+
                        'BEWARUM=BesuchsGrund';

  sKI5TOKI7          =  'KINDNAME=Nachname'+#13#10+
                        'KINDVORNAM=Vorname'+#13#10+
                        'KINDGEBDAT=Geburtstag'+#13#10+
                        'KINDKONF=Kirche';

  sKO5TOKO7          =  'KOMMWANN=AbendmahlsDatum';

  sGE5TOGE7          =  'NUMMER=TempInteger'+#13#10+
                        'TITEL=Titel'+#13#10+
                        'NAME=Nachname'+#13#10+
                        'STRASSE=Strasse'+#13#10+
                        'LAND=Land'+#13#10+
                        'PLZ=PLZ'+#13#10+
                        'ORT=ORT'+#13#10+
                        'ORTSTEIL=ORTSTEIL'+#13#10+
                        'GEBNAME=GebName'+#13#10+
                        'GEBDAT=Geburtstag'+#13#10+
                        'GEBORT=GebOrt'+#13#10+
                        'GEBSTANDDA=GebStADatum'+#13#10+
                        'GEBSTANDOR=GebStAOrt'+#13#10+
                        'GEBSTANDRE=GebStARegNr'+#13#10+
                        'FAMSTAND=Familienstand'+#13#10+
                        'GESCHL=Geschlecht'+#13#10+
                        'VATER=VaterName'+#13#10+
                        'GEBNAVATER=VaterGebName'+#13#10+
                        'KONFVA=VaterKirche'+#13#10+
                        'MUTTER=MutterName'+#13#10+
                        'GEBNAMUTTE=MutterGebName'+#13#10+
                        'KONFMU=MutterKirche'+#13#10+
                        'ERL_BERUF=BerufErlernt'+#13#10+
                        'AUS_BERUF=BerufAusgeuebt'+#13#10+
                        'RUHESTAND=Ruhestand'+#13#10+
                        'MITARBEIT=FreiFeld4'+#13#10+
                        'KONF=Kirche'+#13#10+
                        'TAUFDAT=TaufDatum'+#13#10+
                        'TAUFORT=TaufOrt'+#13#10+
                        'TAUFPFR=TaufPfarrer'+#13#10+
                        'KONFDAT=KonfDatum'+#13#10+
                        'KONFORT=KonfOrt'+#13#10+
                        'KONFPFR=KonfPfarrer'+#13#10+
                        'GETRDAT=TrauDatum'+#13#10+
                        'GETRORT=TrauOrt'+#13#10+
                        'GETRPFR=TrauPfarrer'+#13#10+
                        'EHEGATTE=EhegatteIn'+#13#10+
                        'EHEGEBNAME=EheGebName'+#13#10+
                        'EHEGATKONF=EhegatteKirche'+#13#10+
                        'STANDESDAT=EheStADatum'+#13#10+
                        'STANDESORT=EheStAOrt'+#13#10+
                        'REG_NR=EheStARegNr'+#13#10+
                        'TRAUZEUGE1=TrauZeuge1'+#13#10+
                        'TRAUZEUGE2=TrauZeuge2'+#13#10+
                        'TRAUZEUGE3=TrauZeuge3'+#13#10+
                        'TRAUZEUGE4=TrauZeuge4'+#13#10+
                        'UEBERGETRE=UebertrittsZuDatum'+#13#10+
                        'AUSKONF=UebertrittAus'+#13#10+
                        'UEBERVONDA=Ueberwiesen_von_Datum'+#13#10+
                        'UEBERVON=Ueberwiesen_von'+#13#10+
                        'UEBERNADAT=Ueberwiesen_nach_Datum'+#13#10+
                        'UEBERNA=Ueberwiesen_nach'+#13#10+
                        'AUGETRDAT=AustrittsDatum'+#13#10+
                        'KONVERT=UebertrittsAbGrund'+#13#10+
                        'KONVERTZU=UebertrittNach'+#13#10+
                        'AUSGESCDAT=AusschlussDatum'+#13#10+
                        'VERSTORDAT=TodesDatum'+#13#10+
                        'VERSTANDDA=TodStADatum'+#13#10+
                        'VERSTREGNR=BestattungsRegNr'+#13#10+
                        'VERSTANDOR=TodStAOrt'+#13#10+
                        'VERSTANDRE=TodStARegNr'+#13#10+
                        'BESTATTEDA=BestattungsDatum'+#13#10+
                        'BESTVON=BestattungsPfarrer'+#13#10+
                        'BESTIN=Friedhof'+#13#10+
                        'TEXTTAUF=TaufSpruch'+#13#10+
                        'TEXTKONF=KonfSpruch'+#13#10+
                        'TEXTTRAU=TrauSpruch'+#13#10+
                        'TEXTVERST=BestattungsText'+#13#10+
                        'MEMO=Memo'+#13#10+
                        'PATEN1=TaufPate1'+#13#10+
                        'PATEN2=TaufPate2'+#13#10+
                        'PATEN3=TaufPate3'+#13#10+
                        'PATEN4=TaufPate4'+#13#10+
                        'VERSAND=Versand'+#13#10+
                        'GEMEINDE=Gemeinde'+#13#10+
                        'LASTUPDATE=LastChange'+#13#10+
                        'EMAIL=eMail'+#13#10+
                      //'ANREDE=BriefAnrede'+#13#10+
                        'TEL=TelPrivat'+#13#10+
                        'TEL2=TelDienst'+#13#10+
                        'FAX=Fax'+#13#10+
                        'VORNAME=Vorname'+#13#10+
                        'VORNAME2=Vorname2'+#13#10+
                        'KIBUTAUF=TaufRegNr'+#13#10+
                        'KIBUKONF=KonfRegNr'+#13#10+
                        'KIBUTRAU=TrauRegNr'+#13#10+
                        'EINTRITTDA=EintrittsDatum';

  sGD5TOGD7          =  //'=GdID'+#13#10+
                        'DATUM=GottesdienstDatum'+#13#10+
                        'NAME=Name'+#13#10+
                        'ART1=GottesdienstForm1'+#13#10+
                        'ART2=GottesdienstForm2'+#13#10+
                        'ART3=GottesdienstForm3'+#13#10+
                        //'=Eingangslied'+#13#10+
                        //'=Hauptlied'+#13#10+
                        //'=LiedVorPredigt'+#13#10+
                        //'=LiedNachPredigt'+#13#10+
                        //'=LiedZurBereitung'+#13#10+
                        //'=LiedZurAusteilung'+#13#10+
                        //'=Schlusslied'+#13#10+
                        //'=LiturgGesang1'+#13#10+
                        //'=LiturgGesang2'+#13#10+
                        //'=LiturgGesang3'+#13#10+
                        //'=Introitus'+#13#10+
                        //'=BesucherKiGo'+#13#10+
                        //'=Kollekte'+#13#10+
                        //'=Pfarrer'+#13#10+
                        'BESUCHER=Besucher'+#13#10+
                        'MEMO=Memo'+#13#10+
                        'PREDIGTTEX=Predigttext'+#13#10+
                        'KUESTER=Kuester'+#13#10+
                        'KOMM=Kommunikaten'+#13#10+
                        'GASTKOMM=Gastkommunikaten'+#13#10+
                        'TAG=Tag'+#13#10+
                        'GEMEINDE=Gemeinde';

  sKirchenEintraegeDef = ''+#13#10+
                        'SELK'+#13#10+
                        'ev.'+#13#10+
                        'röm.-kath.'+#13#10+
                        'ELKiB'+#13#10+
                        'Info'+#13#10+
                        'Leerfeld'+#13#10+
                        'ev.-meth.'+#13#10+
                        'ohne'+#13#10+
                        'andere';
  sAnredeEintraegeDef = ''+#13#10+
                        'Herrn'+#13#10+
                        'Frau'+#13#10+
                        'An'+#13#10+
                        'Fam.';
  sSQL_GD_Table_New =   'CREATE TABLE [Gottesdienst] ('+
                        '[GdID] INTEGER  NOT NULL PRIMARY KEY AUTOINCREMENT,'+
                        '[GottesdienstDatum] DATE  NULL,'+
                        '[Name] VARCHAR(100) NULL,'+
                        '[GottesdienstForm1] VARCHAR(20) NULL,'+
                        '[GottesdienstForm2] VARCHAR(20) NULL,'+
                        '[GottesdienstForm3] VARCHAR(20) NULL,'+
                        '[Eingangslied] VARCHAR(100) NULL,'+
                        '[Hauptlied] VARCHAR(100) NULL,'+
                        '[LiedVorPredigt] VARCHAR(100) NULL,'+
                        '[LiedNachPredigt] VARCHAR(100) NULL,'+
                        '[LiedZurBereitung] VARCHAR(100) NULL,'+
                        '[LiedZurAusteilung] VARCHAR(100) NULL,'+
                        '[Schlusslied] VARCHAR(100) NULL,'+
                        '[LiturgGesang1] VARCHAR(100) NULL,'+
                        '[LiturgGesang2] VARCHAR(100) NULL,'+
                        '[LiturgGesang3] VARCHAR(100) NULL,'+
                        '[Introitus] VARCHAR(100) NULL,'+
                        '[BesucherKiGo] INTEGER NULL,'+
                        '[Kollekte] CURRENCY NULL,'+
                        '[Pfarrer] VARCHAR(100) NULL,'+
                        '[Besucher] INTEGER NULL,'+
                        '[Memo] TEXT NULL,'+
                        '[Predigttext] VARCHAR(100) NULL,'+
                        '[Kuester] VARCHAR(100) NULL,'+
                        '[Kommunikaten] INTEGER NULL,'+
                        '[Gastkommunikaten] INTEGER NULL,'+
                        '[Tag] VARCHAR(20) NULL,'+
                        '[Gemeinde] VARCHAR(2) NULL'+
                        ');';

  sGemeindenAlle      = 'Alle';



type
  EMyException = class(Exception);

var
  sPW, sPW1         : string; //Verschlüsseltes PW
  sAktuelles        : String;
  sAnredenEintraege : String;
  sAppDir           : String;
  sDatabase         : string;
  sDebugFile        : String;
  sDefaultGemeinde  : string;
  sGemeinden        : String;
  sIniFile          : String;
  sKirchenEintraege : String;
  sNewVers          : String;
  sPrintPath        : String;
  sReportFile       : String;
  sSavePath         : String;
  sUserDefSQLFile   : String = 'UserDefSQL.SQL';
  sUserAndPCName    : String;
  bSQLDebug         : boolean;
  dtLastSave        : TDateTime;
  ScaleFactor,
  ScaleFactorY,
  ScaleFactorX      : double;
  bDatabaseVersionChecked : boolean;
  WORKAREA          : TRect;


implementation

end.

