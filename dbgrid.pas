unit DBGrid;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, DBGrids,
  Buttons, StdCtrls, ExtCtrls;

type

  { TfrmDBGrid }

  TfrmDBGrid = class(TForm)
    btnCancel: TButton;
    btnOK: TButton;
    DBGrid: TDBGrid;
    Panel1: TPanel;
  private
    { private declarations }
  public
    { public declarations }
  end; 

var
  frmDBGrid: TfrmDBGrid;

implementation

{$R *.lfm}

{ TfrmDBGrid }

end.

