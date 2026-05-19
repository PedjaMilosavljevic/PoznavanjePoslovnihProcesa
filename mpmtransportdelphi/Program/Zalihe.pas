unit Zalihe;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs,
  FMX.Controls.Presentation, FMX.StdCtrls, FMX.Objects, FMX.Edit,
  FMX.ListBox, FMX.Layouts, FMX.Grid, FMX.Grid.Style, FMX.ScrollBox,
  Data.DB, Data.Win.ADODB, UnosPodataka;

type
  TFormZalihe = class(TForm)
    SpeedButton1: TSpeedButton;
    Text1: TText;
    // Toolbar
    btnNovaNabavka: TButton;
    btnIzdajRobu: TButton;
    btnOsvezi: TButton;
    // Filter
    txtPretraga: TEdit;
    ComboKategorija: TComboBox;
    lblPretraga: TLabel;
    lblKategorija: TLabel;
    // Grid
    GridZalihe: TStringGrid;
    ColNaziv: TStringColumn;
    ColKategorija: TStringColumn;
    ColKolicina: TStringColumn;
    ColMinKolicina: TStringColumn;
    ColLokacija: TStringColumn;
    ColStatus: TStringColumn;
    // Alert panel
    PanelAlert: TPanel;
    lblAlert: TLabel;
    // DB
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    procedure FormCreate(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure btnNovaNabavkaClick(Sender: TObject);
    procedure btnIzdajRobuClick(Sender: TObject);
    procedure btnOsveziClick(Sender: TObject);
    procedure txtPretragaChange(Sender: TObject);
    procedure ComboKategorijaChange(Sender: TObject);
  private
    procedure PoveziBazu;
    procedure UcitajZalihe(Filter: string = ''; Kategorija: string = '');
    procedure ProveriKriticneZalihe;
    function StatusZalihe(Kolicina, MinKolicina: Integer): string;
  public
    { Public declarations }
  end;

var
  FormZalihe: TFormZalihe;

implementation

{$R *.fmx}

uses NovaNabavka, IzdavanjeRobe;

procedure TFormZalihe.PoveziBazu;
begin
  ADOConnection1.ConnectionString :=
    'Provider=Microsoft.Jet.OLEDB.4.0;' +
    'Data Source=' + ExtractFilePath(ParamStr(0)) + 'mpmtransport.mdb;';
  ADOConnection1.LoginPrompt := False;
  ADOConnection1.Connected := True;
end;

procedure TFormZalihe.FormCreate(Sender: TObject);
begin
  PoveziBazu;
  // Postavljanje kolona grida
  ColNaziv.Header     := 'Naziv artikla';
  ColKategorija.Header := 'Kategorija';
  ColKolicina.Header  := 'Kolicina';
  ColMinKolicina.Header := 'Min. kolicina';
  ColLokacija.Header  := 'Lokacija';
  ColStatus.Header    := 'Status';
  // Popuni kategorije
  ComboKategorija.Items.Clear;
  ComboKategorija.Items.Add('Sve kategorije');
  ComboKategorija.Items.Add('Rezervni delovi');
  ComboKategorija.Items.Add('Maziva i goriva');
  ComboKategorija.Items.Add('Gume i felne');
  ComboKategorija.Items.Add('Alati i oprema');
  ComboKategorija.ItemIndex := 0;
  UcitajZalihe;
  ProveriKriticneZalihe;
end;

function TFormZalihe.StatusZalihe(Kolicina, MinKolicina: Integer): string;
begin
  if Kolicina = 0 then
    Result := 'NEMA NA STANJU'
  else if Kolicina < MinKolicina then
    Result := 'KRITICNO'
  else if Kolicina < MinKolicina * 1.5 then
    Result := 'NISKA ZALIHA'
  else
    Result := 'OK';
end;

procedure TFormZalihe.UcitajZalihe(Filter: string; Kategorija: string);
var
  SQL: string;
  Row: Integer;
  Kolicina, MinKolicina: Integer;
begin
  SQL := 'SELECT id_artikla, naziv, kategorija, kolicina_na_stanju, ' +
         'min_kolicina, lokacija_u_magacinu FROM zalihe WHERE 1=1';

  if Trim(Filter) <> '' then
    SQL := SQL + ' AND naziv LIKE ''%' + Trim(Filter) + '%''';

  if (Kategorija <> '') and (Kategorija <> 'Sve kategorije') then
    SQL := SQL + ' AND kategorija = ''' + Kategorija + '''';

  SQL := SQL + ' ORDER BY naziv';

  ADOQuery1.Close;
  ADOQuery1.SQL.Text := SQL;
  ADOQuery1.Open;

  GridZalihe.RowCount := ADOQuery1.RecordCount;
  Row := 0;

  while not ADOQuery1.Eof do
  begin
    Kolicina    := ADOQuery1.FieldByName('kolicina_na_stanju').AsInteger;
    MinKolicina := ADOQuery1.FieldByName('min_kolicina').AsInteger;

    GridZalihe.Cells[0, Row] := ADOQuery1.FieldByName('naziv').AsString;
    GridZalihe.Cells[1, Row] := ADOQuery1.FieldByName('kategorija').AsString;
    GridZalihe.Cells[2, Row] := IntToStr(Kolicina);
    GridZalihe.Cells[3, Row] := IntToStr(MinKolicina);
    GridZalihe.Cells[4, Row] := ADOQuery1.FieldByName('lokacija_u_magacinu').AsString;
    GridZalihe.Cells[5, Row] := StatusZalihe(Kolicina, MinKolicina);

    Inc(Row);
    ADOQuery1.Next;
  end;
end;

procedure TFormZalihe.ProveriKriticneZalihe;
var
  BrojKriticnih: Integer;
begin
  ADOQuery1.Close;
  ADOQuery1.SQL.Text :=
    'SELECT COUNT(*) AS br FROM zalihe ' +
    'WHERE kolicina_na_stanju < min_kolicina';
  ADOQuery1.Open;
  BrojKriticnih := ADOQuery1.FieldByName('br').AsInteger;

  if BrojKriticnih > 0 then
  begin
    PanelAlert.Visible := True;
    lblAlert.Text := 'UPOZORENJE: ' + IntToStr(BrojKriticnih) +
                     ' artikala ispod minimalne zalihe! Potrebna nabavka.';
  end
  else
    PanelAlert.Visible := False;
end;

procedure TFormZalihe.SpeedButton1Click(Sender: TObject);
begin
  Form2.Show;
  Close;
end;

procedure TFormZalihe.btnNovaNabavkaClick(Sender: TObject);
begin
  FormNovaNabavka.ShowModal;
  UcitajZalihe;
  ProveriKriticneZalihe;
end;

procedure TFormZalihe.btnIzdajRobuClick(Sender: TObject);
begin
  FormIzdavanjeRobe.ShowModal;
  UcitajZalihe;
  ProveriKriticneZalihe;
end;

procedure TFormZalihe.btnOsveziClick(Sender: TObject);
begin
  UcitajZalihe(txtPretraga.Text,
    ComboKategorija.Items[ComboKategorija.ItemIndex]);
  ProveriKriticneZalihe;
end;

procedure TFormZalihe.txtPretragaChange(Sender: TObject);
begin
  UcitajZalihe(txtPretraga.Text,
    ComboKategorija.Items[ComboKategorija.ItemIndex]);
end;

procedure TFormZalihe.ComboKategorijaChange(Sender: TObject);
begin
  UcitajZalihe(txtPretraga.Text,
    ComboKategorija.Items[ComboKategorija.ItemIndex]);
end;

end.
