unit Menadzer;

interface

uses
    System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs,
  FMX.Controls.Presentation, FMX.StdCtrls, FMX.Objects, FMX.Maps,
  UnosPodataka, Data.DB, Data.Win.ADODB, FMX.Edit, FMX.ListBox;
  type
  TPrijavljeniKorisnik = record
    ID: Integer;
    KorisnickoIme: string;
    Role: string;
  end;
  var
  TrenutniKorisnik: TPrijavljeniKorisnik;
type
  TForm3 = class(TForm)
    MapView1: TMapView;
    PrijavaBtn: TButton;
    DugmeFakture: TButton;
    Button2: TButton;
    Button3: TButton;
    btnChat: TButton;
    Text1: TText;
    SpeedButton1: TSpeedButton;
    ADOQuery1: TADOQuery;
    ADOConnection1: TADOConnection;
    ComboKorisnici: TComboBox;
    DugmeIzaberiKorisnika: TButton;
    procedure SpeedButton1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure DugmeIzaberiKorisnikaClick(Sender: TObject);

  private
 procedure PopuniComboKorisnike;
  public

    { Public declarations }
  end;

var
  Form3: TForm3;

implementation
uses Vozac;


{$R *.fmx}


procedure TForm3.DugmeIzaberiKorisnikaClick(Sender: TObject);
var
IzabraniID: Integer;
begin
  if ComboKorisnici.ItemIndex = -1 then
  begin
    ShowMessage('Molimo izaberite korisnika.');
    Exit;
  end;
  IzabraniID := Integer(ComboKorisnici.Items.Objects[ComboKorisnici.ItemIndex]);
  TrenutniKorisnik.ID := IzabraniID;
  TrenutniKorisnik.KorisnickoIme := ComboKorisnici.Items[ComboKorisnici.ItemIndex];
  TrenutniKorisnik.Role := 'Korisnik';
  Form7.Show;



end;

procedure TForm3.FormCreate(Sender: TObject);
begin
ADOConnection1.ConnectionString :=
    'Provider=Microsoft.Jet.OLEDB.4.0;' +
    'Data Source=' + ExtractFilePath(ParamStr(0)) + 'mpmtransport.mdb;';
  ADOConnection1.LoginPrompt := False;
  ADOConnection1.Connected := True;
  PopuniComboKorisnike;

end;

procedure TForm3.SpeedButton1Click(Sender: TObject);
begin
Form2.Show;
Close;
end;
procedure TForm3.PopuniComboKorisnike;
begin
  ComboKorisnici.Items.Clear;
  ADOQuery1.Close;
  ADOQuery1.SQL.Text := 'SELECT id, korisnickoime FROM login';
  ADOQuery1.Open;

  while not ADOQuery1.Eof do
  begin
    ComboKorisnici.Items.AddObject(
      ADOQuery1.FieldByName('korisnickoime').AsString,
      TObject(ADOQuery1.FieldByName('id').AsInteger)
    );
    ADOQuery1.Next;
  end;
end;

end.
