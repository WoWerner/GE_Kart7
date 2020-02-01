unit statinfo;

interface

uses
  Classes,
  Graphics,
  Forms,
  Controls,
  Buttons,
  StdCtrls,
  ExtCtrls;

type

  { TfrmStatInfo }

  TfrmStatInfo = class(TForm)
    Bevel1: TBevel;
    btnOK: TButton;
    btnCancel: TButton;
    stat_gemeinde: TComboBox;
    Label1: TLabel;
    stat_jahr: TEdit;
    Label2: TLabel;
    Label3: TLabel;
    stat_Kirche: TComboBox;
    Label4: TLabel;
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var frmStatInfo : TfrmStatInfo;

implementation

{$R *.lfm}

uses
  global;

{ TfrmStatInfo }

procedure TfrmStatInfo.FormShow(Sender: TObject);
begin
  stat_Kirche.Items.Text   := sKirchenEintraege;
  stat_gemeinde.Items.Text := sGemeinden;
  stat_gemeinde.Text       := sDefaultGemeinde;
end;

end.
