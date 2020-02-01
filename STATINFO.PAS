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
    Label1: TLabel;
    stat_jahr: TEdit;
    Label2: TLabel;
    Label3: TLabel;
    stat_Kirche: TComboBox;
    Label4: TLabel;
    stat_gemeinde: TEdit;
    procedure FormCreate(Sender: TObject);
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

procedure TfrmStatInfo.FormCreate(Sender: TObject);
begin
  stat_Kirche.Items.Text:=sKirchenEintraege;
end;

end.
