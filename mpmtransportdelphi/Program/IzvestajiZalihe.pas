unit IzvestajiZalihe;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs,
  FMX.Controls.Presentation, FMX.StdCtrls, FMX.Objects, FMX.Edit,
  FMX.ListBox, FMX.Layouts, FMX.Grid, FMX.Grid.Style, FMX.ScrollBox,
  FMX.TabControl, FMX.DateTimeCtrls,
  Data.DB, Data.Win.ADODB, UnosPodataka;

type
  TFormIzvestaji = class(TForm)
    SpeedButton1: TSpeedButton;
    Text1: TText;
    // Tabovi
    TabControl1: TTabControl;
    TabNabavke: TTabItem;
    TabIzdavanja: TTabItem;
    TabKriticne: TTabItem;
    // Filter za nabavke
    lblOdDatuma: TLabel;
    DateOd: TDateEdit;
    lblDoDatuma: TLabel;
    DateDo: TDateEdit;
    btnFiltriraj: TButton;
    // Grid nabavke
    GridNabavke: TStringGrid;
    ColNBroj: TStringColumn;
    ColNDobavljac: TStringColumn;
    ColNArtikl: TStringColumn;
    ColNKolicina: TStringColumn;
    ColNVrednost: TStringColumn;
    ColNDatum: TStringColumn;
    ColNStatus: TStringColumn;
    // Grid izdavanja
    GridIzdavanja: TStringGrid;
    ColIBroj: TStringColumn;
    ColIVozilo: TStringColumn;
    ColIArtikl: TStringColumn;
    ColIKolicina: TStringColumn;
    ColIDatum: TStringColumn;
    ColIStatus: TStringColumn;
    // Grid kriticne zalihe
    GridKriticne: TStringGrid;
    ColKNaziv: TStringColumn;
    ColKKategorija: TStringColumn;
    ColKStanje: TStringColumn;
    ColKMinimum: TStringColumn;
    ColKRazlika: TStringColumn;
    // Sumarni panel
    PanelSumarno: TPanel;
    lblUkupnoNabavljeno: TLabel;
    lblUkupnoIzdato: TLabel;
    lblVrednostZaliha: TLabel;
    // DB
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure btnFiltrirajClick(Sender: TObject);
    procedure TabControl1Change(Sender: TObject);
  private
    procedure PoveziBazu;
    procedure UcitajNabavke;
    procedure UcitajIzdavanja;
    procedure UcitajKriticneZalihe;
    procedure IzracunajSumarno;
  public
    { Public declarations }
  end;

var
  FormIzvestaji: TFormIzvestaji;

implementation

{$R *.fmx}

procedure TFormIzvestaji.PoveziBazu;
var
  dbPath: string;
begin
  dbPath := ExtractFilePath(ParamStr(0)) + '..\..\..\Baza podataka\mpmtransport.mdb';

  if not FileExists(dbPath) then
  begin
    ShowMessage('Baza ne postoji na lokaciji: ' + dbPath);
    Exit;
  end;

  try
    ADOConnection1.ConnectionString :=
      'Provider=Microsoft.Jet.OLEDB.4.0;' +
      'Data Source=' + dbPath + ';';
    ADOConnection1.LoginPrompt := False;
    ADOConnection1.Connected := True;
  except
    on E: Exception do
      ShowMessage('Greska pri konekciji sa bazom: ' + E.Message);
  end;
end;

procedure TFormIzvestaji.FormCreate(Sender: TObject);
begin
  PoveziBazu;
  // Podesi kolone grida Nabavke
  ColNBroj.Header      := 'Broj naloga';
  ColNDobavljac.Header := 'Dobavljac';
  ColNArtikl.Header    := 'Artikal';
  ColNKolicina.Header  := 'Kolicina';
  ColNVrednost.Header  := 'Vrednost (RSD)';
  ColNDatum.Header     := 'Datum';
  ColNStatus.Header    := 'Status';
  // Kolone Izdavanja
  ColIBroj.Header     := 'Broj naloga';
  ColIVozilo.Header   := 'Vozilo';
  ColIArtikl.Header   := 'Artikal';
  ColIKolicina.Header := 'Kolicina';
  ColIDatum.Header    := 'Datum';
  ColIStatus.Header   := 'Status';
  // Kolone Kriticne
  ColKNaziv.Header      := 'Naziv';
  ColKKategorija.Header := 'Kategorija';
  ColKStanje.Header     := 'Na stanju';
  ColKMinimum.Header    := 'Minimum';
  ColKRazlika.Header    := 'Nedostaje';
  // Default datumi — poslednji mesec
  DateOd.Date := IncMonth(Now, -1);
  DateDo.Date := Now;

  UcitajNabavke;
  UcitajIzdavanja;
  UcitajKriticneZalihe;
  IzracunajSumarno;
end;

procedure TFormIzvestaji.FormShow(Sender: TObject);
begin
  UcitajNabavke;
  UcitajIzdavanja;
  UcitajKriticneZalihe;
  IzracunajSumarno;
end;

procedure TFormIzvestaji.UcitajNabavke;
var
  Row: Integer;
begin
  ADOQuery1.Close;
  ADOQuery1.SQL.Text :=
    'SELECT nn.id_naloga, d.naziv_firme, z.naziv, nn.kolicina, ' +
    '       nn.ukupna_vrednost, nn.datum_naloga, nn.status ' +
    'FROM (nalog_nabavka nn ' +
    'INNER JOIN dobavljaci d ON d.id_dobavljaca = nn.id_dobavljaca) ' +
    'INNER JOIN zalihe z     ON z.id_artikla    = nn.id_artikla ' +
    'WHERE nn.datum_naloga >= #' + FormatDateTime('mm/dd/yyyy', DateOd.Date) + '# ' +
    'AND nn.datum_naloga <= #' + FormatDateTime('mm/dd/yyyy', DateDo.Date + 1) + '# ' +
    'ORDER BY nn.datum_naloga DESC';

  ADOQuery1.Open;

  GridNabavke.RowCount := ADOQuery1.RecordCount;
  Row := 0;
  while not ADOQuery1.Eof do
  begin
    GridNabavke.Cells[0, Row] := 'NOV-' + IntToStr(ADOQuery1.FieldByName('id_naloga').AsInteger);
    GridNabavke.Cells[1, Row] := ADOQuery1.FieldByName('naziv_firme').AsString;
    GridNabavke.Cells[2, Row] := ADOQuery1.FieldByName('naziv').AsString;
    GridNabavke.Cells[3, Row] := ADOQuery1.FieldByName('kolicina').AsString;
    GridNabavke.Cells[4, Row] := FormatFloat('#,##0.00',
      ADOQuery1.FieldByName('ukupna_vrednost').AsFloat);
    GridNabavke.Cells[5, Row] := FormatDateTime('dd.mm.yyyy',
      ADOQuery1.FieldByName('datum_naloga').AsDateTime);
    GridNabavke.Cells[6, Row] := ADOQuery1.FieldByName('status').AsString;
    Inc(Row);
    ADOQuery1.Next;
  end;
end;

procedure TFormIzvestaji.UcitajIzdavanja;
var
  Row: Integer;
begin
  ADOQuery1.Close;
  ADOQuery1.SQL.Text :=
    'SELECT ni.id_izdavanja, voz.registarski_broj, z.naziv, ' +
    '       ni.kolicina, ni.datum_izdavanja, ni.status ' +
    'FROM (nalog_izdavanje ni ' +
    'INNER JOIN Vozila  voz ON voz.ID  = ni.id_vozila) ' +
    'INNER JOIN zalihe  z   ON z.id_artikla   = ni.id_artikla ' +
    'WHERE ni.datum_izdavanja >= #' + FormatDateTime('mm/dd/yyyy', DateOd.Date) + '# ' +
    'AND ni.datum_izdavanja <= #' + FormatDateTime('mm/dd/yyyy', DateDo.Date + 1) + '# ' +
    'ORDER BY ni.datum_izdavanja DESC';

  ADOQuery1.Open;

  GridIzdavanja.RowCount := ADOQuery1.RecordCount;
  Row := 0;
  while not ADOQuery1.Eof do
  begin
    GridIzdavanja.Cells[0, Row] := 'IZD-' + IntToStr(ADOQuery1.FieldByName('id_izdavanja').AsInteger);
    GridIzdavanja.Cells[1, Row] := ADOQuery1.FieldByName('registarski_broj').AsString;
    GridIzdavanja.Cells[2, Row] := ADOQuery1.FieldByName('naziv').AsString;
    GridIzdavanja.Cells[3, Row] := ADOQuery1.FieldByName('kolicina').AsString;
    GridIzdavanja.Cells[4, Row] := FormatDateTime('dd.mm.yyyy',
      ADOQuery1.FieldByName('datum_izdavanja').AsDateTime);
    GridIzdavanja.Cells[5, Row] := ADOQuery1.FieldByName('status').AsString;
    Inc(Row);
    ADOQuery1.Next;
  end;
end;

procedure TFormIzvestaji.UcitajKriticneZalihe;
var
  Row: Integer;
  Stanje, Minimum: Integer;
begin
  ADOQuery1.Close;
  ADOQuery1.SQL.Text :=
    'SELECT naziv, kategorija, kolicina_na_stanju, min_kolicina ' +
    'FROM zalihe ' +
    'WHERE kolicina_na_stanju < min_kolicina ' +
    'ORDER BY (min_kolicina - kolicina_na_stanju) DESC';
  ADOQuery1.Open;

  GridKriticne.RowCount := ADOQuery1.RecordCount;
  Row := 0;
  while not ADOQuery1.Eof do
  begin
    Stanje  := ADOQuery1.FieldByName('kolicina_na_stanju').AsInteger;
    Minimum := ADOQuery1.FieldByName('min_kolicina').AsInteger;

    GridKriticne.Cells[0, Row] := ADOQuery1.FieldByName('naziv').AsString;
    GridKriticne.Cells[1, Row] := ADOQuery1.FieldByName('kategorija').AsString;
    GridKriticne.Cells[2, Row] := IntToStr(Stanje);
    GridKriticne.Cells[3, Row] := IntToStr(Minimum);
    GridKriticne.Cells[4, Row] := IntToStr(Minimum - Stanje);
    Inc(Row);
    ADOQuery1.Next;
  end;
end;

procedure TFormIzvestaji.IzracunajSumarno;
begin
  // Ukupno nabavljeno
  ADOQuery1.Close;
  ADOQuery1.SQL.Text := 'SELECT SUM(kolicina) AS uk FROM nalog_nabavka';
  ADOQuery1.Open;
  if ADOQuery1.FieldByName('uk').IsNull then
    lblUkupnoNabavljeno.Text := 'Ukupno nabavljeno: 0 kom.'
  else
    lblUkupnoNabavljeno.Text := 'Ukupno nabavljeno: ' +
      ADOQuery1.FieldByName('uk').AsString + ' kom.';

  // Ukupno izdato
  ADOQuery1.Close;
  ADOQuery1.SQL.Text := 'SELECT SUM(kolicina) AS uk FROM nalog_izdavanje';
  ADOQuery1.Open;
  if ADOQuery1.FieldByName('uk').IsNull then
    lblUkupnoIzdato.Text := 'Ukupno izdato: 0 kom.'
  else
    lblUkupnoIzdato.Text := 'Ukupno izdato: ' +
      ADOQuery1.FieldByName('uk').AsString + ' kom.';

  // Vrednost zaliha — koristi PROSECNU cenu po artiklu da se izbegne
  // dupliranje vrednosti kada je artikal naruчivan vise puta
  ADOQuery1.Close;
  ADOQuery1.SQL.Text :=
    'SELECT SUM(z.kolicina_na_stanju * p.avg_cena) AS vr ' +
    'FROM zalihe z ' +
    'LEFT JOIN (SELECT id_artikla, AVG(cena_po_komadu) AS avg_cena ' +
    '           FROM nalog_nabavka GROUP BY id_artikla) p ' +
    'ON p.id_artikla = z.id_artikla';
  ADOQuery1.Open;
  if ADOQuery1.FieldByName('vr').IsNull then
    lblVrednostZaliha.Text := 'Vrednost zaliha: 0.00 RSD'
  else
    lblVrednostZaliha.Text := 'Vrednost zaliha: ' +
      FormatFloat('#,##0.00', ADOQuery1.FieldByName('vr').AsFloat) + ' RSD';
end;

procedure TFormIzvestaji.btnFiltrirajClick(Sender: TObject);
begin
  UcitajNabavke;
  UcitajIzdavanja;
end;

procedure TFormIzvestaji.TabControl1Change(Sender: TObject);
begin
  if TabControl1.ActiveTab = TabKriticne then
    UcitajKriticneZalihe;
end;

procedure TFormIzvestaji.SpeedButton1Click(Sender: TObject);
var
  F: TForm;
begin
  F := TForm(Application.FindComponent('FormZalihe'));
  if Assigned(F) then
  begin
    F.Show;
    Hide;
  end
  else
  begin
    Form2.Show;
    Hide;
  end;
end;

end.
