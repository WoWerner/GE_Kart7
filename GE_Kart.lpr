program GE_Kart;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, Main, Global, Controls, Dialogs, appsettings, Ausgabe, Freelist,
  help, input, PgmUpdate, uhttpdownloader, dm, dbflaz, lazcontrols,
  printer4lazarus, zcomponent, lazreport, kartei, DBGrid, STATINFO, suche, gd,
  Windows;   //Mutex

{$R *.res}

CONST
 MutexName = 'GE_Kart MuTeX';

VAR
 hMutex: THandle;

begin
  Try
    hMutex := CreateMutex(Nil, True, MutexName);

    If (hMutex = 0) Or (GetLastError = ERROR_ALREADY_EXISTS)
      then
        begin
          ShowMessage ('GE_Kart läuft bereits');
        end
      else
        begin
          Try
	    Application.Scaled:=True;
            Application.Title:='GE_Kart Verwaltungsprogramm für Kirchengemeinden';
            Application.Initialize;
            Application.CreateForm(TfrmMain, frmMain);
            Application.CreateForm(TfrmDM, frmDM);
            Application.CreateForm(TfrmKartei, frmKartei);
            Application.CreateForm(TfrmDBGrid, frmDBGrid);
            Application.CreateForm(TfrmStatInfo, frmStatInfo);
            Application.CreateForm(TfrmSuche, frmSuche);
            Application.CreateForm(TfrmGD, frmGD);
            Application.CreateForm(TfrmAusgabe, frmAusgabe);
            Application.CreateForm(TfrmFreieListe, frmFreieListe);
            Application.CreateForm(TfrmInput, frmInput);
            Application.CreateForm(TfrmPgmUpdate, frmPgmUpdate);

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
          Finally
            If hMutex <> 0 Then CloseHandle(hMutex);
          End;
        end;
   Except
     // LOG...
     If hMutex <> 0 Then CloseHandle(hMutex);
   End;
end.

