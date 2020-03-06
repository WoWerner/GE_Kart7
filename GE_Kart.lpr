program GE_Kart;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms,
  Main,
  Global,
  help,
  input,
  Controls,
  Dialogs,
  dm,
  dbflaz,
  lazcontrols,
  printer4lazarus,
  zcomponent,
  lazreport,
  kartei,
  vinfo,
  appsettings,
  DBGrid,
  Ausgabe,
  STATINFO,
  suche,
  gd,
  PgmUpdate,
  freelist;

{$R *.res}

begin
  Application.Scaled:=True;
  Application.Title:='GE_Kart Verwaltungsprogramm für Kirchengemeinden';
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.CreateForm(TfrmDM, frmDM);
  Application.CreateForm(TfrmInput, frmInput);
  Application.CreateForm(TfrmAusgabe, frmAusgabe);
  Application.CreateForm(TfrmKartei, frmKartei);
  Application.CreateForm(TfrmDBGrid, frmDBGrid);
  Application.CreateForm(TfrmStatInfo, frmStatInfo);
  Application.CreateForm(TfrmSuche, frmSuche);
  Application.CreateForm(TfrmGD, frmGD);
  Application.CreateForm(TfrmFreieListe, frmFreieListe);
  Application.CreateForm(TfrmPgmUpdate,frmPgmUpdate);

  if isconsole
    then
      Application.Run
    else
      begin
        sPW  := help.ReadIniVal(global.sIniFile, 'Passwort', 'Passwort', '87663', true); //= Master
        sPW1 := '';
        frmInput.SetDefaults('Passwort eingeben', 'Passwort', '', '', '', true);
        if ParamCount >= 1 then sPW1 := help.verschluessele(ParamStr(1));
        if (sPW1 = sPW) or (sPW = '40896') // '40896' = Kein Passwort
          then
            Application.Run
          else
            begin
              if (frmInput.ShowModal = mrOK) then sPW1 := help.verschluessele(frmInput.Edit1.Text);
              if (sPW1 = sPW)
                then Application.Run
                else ShowMessage('Passwort ist falsch, Programm wird beendet');
            end;
      end;
end.

