unit IzdavanjeRobe;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs,
  FMX.Controls.Presentation, FMX.StdCtrls, FMX.Objects, FMX.Edit,
  FMX.ListBox, FMX.Layouts,
  Data.DB, Data.Win.ADODB, UnosPodataka;

type
  TFormIzdavanjeRobe = class(TForm)
    RectHeader: TRectangle;
    SpeedButton1: TSpeedButton;
    lblTitle: TLabel;
    RectBody: TRectangle;
    lblVozilo: TLabel;
    ComboVozilo: TComboBox;
    lblVozacIme: TLabel;
    lblArtikl: TLabel;
    ComboArtikl: TComboBox;
    PanelInfo: TPanel;
    lblDostupnoKolicina: TLabel;
    lblLokacija: TLabel;
    lblKolicina: TLabel;
    txtKolicina: TEdit;
    lblKorisnik: TLabel;
    txtKorisnik: TEdit;
    lblNapomena: TLabel;
    txtNapomena: TEdit;
    btnIzdaj: TButton;
    btnOtkazi: TButton;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    ADOQuery2: TADOQuery;
    ADOQuery3: TADOQuery;
    procedure FormCreate(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure btnIzdajClick(Sender: TObject);
    procedure btnOtkaziClick(Sender: TObject);
    procedure ComboVoziloChange(Sender: TObject);
    procedure ComboArtiklChange(Sender: TObject);
  private
    procedure PoveziBazu;
    procedure PopuniVozila;
    procedure PopuniArtikle;
    procedure PrikaziVozacaZaVozilo;
    procedure PrikaziInfoArtikla;
    function ValidacijaUnosa: Boolean;
    procedure VratiSeNaPrethodnu;
  end;

var
  FormIzdavanjeRobe: TFormIzdavanjeRobe;

implementation

{$R *.fmx}

uses Zalihe;

procedure TFormIzdavanjeRobe.PoveziBazu;
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

procedure TFormIzdavanjeRobe.FormCreate(Sender: TObject);
begin
  PoveziBazu;
  PopuniVozila;
  PopuniArtikle;
  lblVozacIme.Text := 'Vozac: -';
  PanelInfo.Visible := False;
end;

procedure TFormIzdavanjeRobe.PopuniVozila;
begin
  ComboVozilo.Items.Clear;
  ADOQuery1.Close;
  ADOQuery1.SQL.Text :=
    'SELECT ID, registarski_broj FROM Vozila ORDER BY registarski_broj';
  ADOQuery1.Open;
  while not ADOQuery1.Eof do
  begin
    ComboVozilo.Items.AddObject(
      ADOQuery1.FieldByName('registarski_broj').AsString,
      TObject(ADOQuery1.FieldByName('ID').AsInteger)
    );
    ADOQuery1.Next;
  end;
end;

procedure TFormIzdavanjeRobe.PopuniArtikle;
begin
  ComboArtikl.Items.Clear;
  ADOQuery1.Close;
  ADOQuery1.SQL.Text :=
    'SELECT id_artikla, naziv, kolicina_na_stanju ' +
    'FROM zalihe WHERE kolicina_na_stanju > 0 ORDER BY naziv';
  ADOQuery1.Open;
  while not ADOQuery1.Eof do
  begin
    ComboArtikl.Items.AddObject(
      ADOQuery1.FieldByName('naziv').AsString,
      TObject(ADOQuery1.FieldByName('id_artikla').AsInteger)
    );
    ADOQuery1.Next;
  end;
end;

procedure TFormIzdavanjeRobe.ComboVoziloChange(Sender: TObject);
begin
  PrikaziVozacaZaVozilo;
end;

procedure TFormIzdavanjeRobe.ComboArtiklChange(Sender: TObject);
begin
  PrikaziInfoArtikla;
end;

procedure TFormIzdavanjeRobe.PrikaziVozacaZaVozilo;
var
  IDVozila: Integer;
begin
  if ComboVozilo.ItemIndex = -1 then
  begin
    lblVozacIme.Text := 'Vozac: -';
    Exit;
  end;
  IDVozila := Integer(ComboVozilo.Items.Objects[ComboVozilo.ItemIndex]);
  ADOQuery2.Close;
  ADOQuery2.SQL.Text :=
    'SELECT registarski_broj FROM Vozila WHERE ID = ' + IntToStr(IDVozila);
  ADOQuery2.Open;
  if not ADOQuery2.Eof then
    lblVozacIme.Text := 'Vozilo: ' + ADOQuery2.FieldByName('registarski_broj').AsString
  else
    lblVozacIme.Text := 'Vozac: -';
end;

procedure TFormIzdavanjeRobe.PrikaziInfoArtikla;
var
  IDArtikla: Integer;
begin
  if ComboArtikl.ItemIndex = -1 then
  begin
    PanelInfo.Visible := False;
    Exit;
  end;
  IDArtikla := Integer(ComboArtikl.Items.Objects[ComboArtikl.ItemIndex]);
  ADOQuery3.Close;
  ADOQuery3.SQL.Text :=
    'SELECT kolicina_na_stanju, lokacija_u_magacinu FROM zalihe WHERE id_artikla = ' +
    IntToStr(IDArtikla);
  ADOQuery3.Open;
  if not ADOQuery3.Eof then
  begin
    lblDostupnoKolicina.Text := 'Dostupno: ' +
      ADOQuery3.FieldByName('kolicina_na_stanju').AsString + ' kom.';
    lblLokacija.Text := 'Lokacija: ' +
      ADOQuery3.FieldByName('lokacija_u_magacinu').AsString;
    PanelInfo.Visible := True;
  end;
end;

function TFormIzdavanjeRobe.ValidacijaUnosa: Boolean;
var
  Kolicina, KolicinaNaStanju, IDArtikla: Integer;
begin
  Result := False;
  if ComboVozilo.ItemIndex = -1 then
  begin
    ShowMessage('Molimo izaberite vozilo.');
    Exit;
  end;
  if ComboArtikl.ItemIndex = -1 then
  begin
    ShowMessage('Molimo izaberite artikal.');
    Exit;
  end;
  if Trim(txtKorisnik.Text) = '' then
  begin
    ShowMessage('Molimo unesite ime i prezime korisnika koji preuzima robu.');
    Exit;
  end;
  if Trim(txtKolicina.Text) = '' then
  begin
    ShowMessage('Molimo unesite kolicinu.');
    Exit;
  end;
  if not TryStrToInt(txtKolicina.Text, Kolicina) or (Kolicina <= 0) then
  begin
    ShowMessage('Kolicina mora biti pozitivan ceo broj.');
    Exit;
  end;
  IDArtikla := Integer(ComboArtikl.Items.Objects[ComboArtikl.ItemIndex]);

  if not ADOConnection1.Connected then
    PoveziBazu;

  try
    ADOQuery3.Close;
    ADOQuery3.SQL.Text :=
      'SELECT kolicina_na_stanju FROM zalihe WHERE id_artikla = ' + IntToStr(IDArtikla);
    ADOQuery3.Open;
    KolicinaNaStanju := ADOQuery3.FieldByName('kolicina_na_stanju').AsInteger;
  except
    on E: Exception do
    begin
      ShowMessage('Greska pri provjeri stanja: ' + E.Message);
      Exit;
    end;
  end;
  if Kolicina > KolicinaNaStanju then
  begin
    ShowMessage('Nema dovoljno na stanju!' + sLineBreak +
      'Trazeno: ' + IntToStr(Kolicina) + sLineBreak +
      'Dostupno: ' + IntToStr(KolicinaNaStanju));
    Exit;
  end;
  Result := True;
end;

procedure TFormIzdavanjeRobe.btnIzdajClick(Sender: TObject);
var
  IDVozila, IDArtikla, Kolicina: Integer;
begin
  if not ValidacijaUnosa then
    Exit;
  IDVozila  := Integer(ComboVozilo.Items.Objects[ComboVozilo.ItemIndex]);
  IDArtikla := Integer(ComboArtikl.Items.Objects[ComboArtikl.ItemIndex]);
  Kolicina  := StrToInt(txtKolicina.Text);

  if not ADOConnection1.Connected then
    PoveziBazu;

  try
    ADOQuery1.Close;
    ADOQuery1.SQL.Text :=
      'INSERT INTO nalog_izdavanje ' +
      '(id_vozila, id_artikla, kolicina, datum_izdavanja, status, napomena, korisnik) ' +
      'VALUES (' +
      IntToStr(IDVozila) + ', ' +
      IntToStr(IDArtikla) + ', ' +
      IntToStr(Kolicina) + ', ' +
      'Now(), ' +
      '''Izdato'', ' +
      '''' + Trim(txtNapomena.Text) + ''', ' +
      '''' + Trim(txtKorisnik.Text) + ''')';
    ADOQuery1.ExecSQL;

    ADOQuery2.Close;
    ADOQuery2.SQL.Text :=
      'UPDATE zalihe SET kolicina_na_stanju = kolicina_na_stanju - ' +
      IntToStr(Kolicina) +
      ' WHERE id_artikla = ' + IntToStr(IDArtikla);
    ADOQuery2.ExecSQL;

    ShowMessage('Roba je uspesno izdata!');
    VratiSeNaPrethodnu;
  except
    on E: Exception do
      ShowMessage('Greska: ' + E.Message);
  end;
end;

procedure TFormIzdavanjeRobe.VratiSeNaPrethodnu;
var
  F: TForm;
begin
  F := TForm(Application.FindComponent(PrethodnaFormaZalihe));
  if Assigned(F) then
  begin
    F.Show;
    Hide;
  end
  else
  begin
    FormZalihe.Show;
    Hide;
  end;
end;

procedure TFormIzdavanjeRobe.btnOtkaziClick(Sender: TObject);
begin
  VratiSeNaPrethodnu;
end;

procedure TFormIzdavanjeRobe.SpeedButton1Click(Sender: TObject);
begin
  VratiSeNaPrethodnu;
end;

end.
