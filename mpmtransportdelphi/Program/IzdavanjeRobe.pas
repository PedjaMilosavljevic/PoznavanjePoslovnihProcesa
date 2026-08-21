unit IzdavanjeRobe;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs,
  FMX.Controls.Presentation, FMX.StdCtrls, FMX.Objects, FMX.Edit,
  FMX.ListBox, FMX.Layouts,
  Data.DB, Data.Win.ADODB, UnosPodataka, Zalihe;

type
  TFormIzdavanjeRobe = class(TForm)
    RectHeader: TRectangle;
    SpeedButton1: TSpeedButton;
    lblTitle: TLabel;
    RectBody: TRectangle;
    lblVozilo: TLabel;
    ComboVozilo: TComboBox;
    lblKorisnik: TLabel;
    txtKorisnik: TEdit;
    lblVozacIme: TLabel;
    Line1: TLine;
    lblDodajArtikl: TLabel;
    lblArtikl: TLabel;
    ComboArtikl: TComboBox;
    PanelInfo: TPanel;
    lblDostupnoKolicina: TLabel;
    lblLokacija: TLabel;
    lblKolicina: TLabel;
    txtKolicina: TEdit;
    btnDodajStavku: TButton;
    Line2: TLine;
    lblStavkeNaslov: TLabel;
    btnObrisiStavku: TButton;
    ListBoxStavke: TListBox;
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
    procedure btnDodajStavkuClick(Sender: TObject);
    procedure btnObrisiStavkuClick(Sender: TObject);
  private
    StavkiArtiklaID : array[0..99] of Integer;
    StavkiKolicina   : array[0..99] of Integer;
    BrojStavki       : Integer;

    procedure PoveziBazu;
    procedure PopuniVozila;
    procedure PopuniArtikle;
    procedure PrikaziInfoVozila;
    procedure PrikaziInfoArtikla;
    procedure AzurirajListu;
    procedure OcistiFormu;
    procedure VratiSeNaPrethodnu;
    function  DostupnaKolicina(IDArtikla: Integer): Integer;
    function  ValidacijaUnosa: Boolean;
  end;

var
  FormIzdavanjeRobe: TFormIzdavanjeRobe;

implementation

{$R *.fmx}

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
      ShowMessage('Greska pri konekciji: ' + E.Message);
  end;
end;

procedure TFormIzdavanjeRobe.FormCreate(Sender: TObject);
begin
  PoveziBazu;
  PopuniVozila;
  PopuniArtikle;
  lblVozacIme.Text := 'Vozilo: -';
  PanelInfo.Visible := False;
  BrojStavki := 0;
  AzurirajListu;
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
  PrikaziInfoVozila;
end;

procedure TFormIzdavanjeRobe.ComboArtiklChange(Sender: TObject);
begin
  PrikaziInfoArtikla;
end;

procedure TFormIzdavanjeRobe.PrikaziInfoVozila;
var
  IDVozila: Integer;
begin
  if ComboVozilo.ItemIndex = -1 then
  begin
    lblVozacIme.Text := 'Vozilo: -';
    Exit;
  end;
  IDVozila := Integer(ComboVozilo.Items.Objects[ComboVozilo.ItemIndex]);
  ADOQuery2.Close;
  ADOQuery2.SQL.Text :=
    'SELECT registarski_broj FROM Vozila WHERE ID = ' + IntToStr(IDVozila);
  ADOQuery2.Open;
  if not ADOQuery2.Eof then
    lblVozacIme.Text := 'Izabrano vozilo: ' +
      ADOQuery2.FieldByName('registarski_broj').AsString;
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

function TFormIzdavanjeRobe.DostupnaKolicina(IDArtikla: Integer): Integer;
begin
  ADOQuery3.Close;
  ADOQuery3.SQL.Text :=
    'SELECT kolicina_na_stanju FROM zalihe WHERE id_artikla = ' +
    IntToStr(IDArtikla);
  ADOQuery3.Open;
  if not ADOQuery3.Eof then
    Result := ADOQuery3.FieldByName('kolicina_na_stanju').AsInteger
  else
    Result := 0;
end;

procedure TFormIzdavanjeRobe.btnDodajStavkuClick(Sender: TObject);
var
  IDArtikla, Kolicina, Dostupno, VecUKosarici: Integer;
  i: Integer;
begin
  if ComboArtikl.ItemIndex = -1 then
  begin
    ShowMessage('Molimo izaberite artikal.');
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

  // Izbroji koliko je vec u kosarici za ovaj artikal
  VecUKosarici := 0;
  for i := 0 to BrojStavki - 1 do
    if StavkiArtiklaID[i] = IDArtikla then
      Inc(VecUKosarici, StavkiKolicina[i]);

  Dostupno := DostupnaKolicina(IDArtikla);

  if VecUKosarici + Kolicina > Dostupno then
  begin
    ShowMessage('Nema dovoljno na stanju!' + sLineBreak +
      'Trazeno ukupno: ' + IntToStr(VecUKosarici + Kolicina) + sLineBreak +
      'Dostupno: ' + IntToStr(Dostupno));
    Exit;
  end;

  if BrojStavki >= 100 then
  begin
    ShowMessage('Dostignut maksimalan broj stavki.');
    Exit;
  end;

  StavkiArtiklaID[BrojStavki] := IDArtikla;
  StavkiKolicina[BrojStavki]  := Kolicina;
  Inc(BrojStavki);

  ComboArtikl.ItemIndex := -1;
  txtKolicina.Text := '';
  PanelInfo.Visible := False;

  AzurirajListu;
end;

procedure TFormIzdavanjeRobe.AzurirajListu;
var
  i: Integer;
  NazivArtikla: string;
  Item: TListBoxItem;
begin
  ListBoxStavke.Clear;
  for i := 0 to BrojStavki - 1 do
  begin
    ADOQuery3.Close;
    ADOQuery3.SQL.Text :=
      'SELECT naziv FROM zalihe WHERE id_artikla = ' +
      IntToStr(StavkiArtiklaID[i]);
    ADOQuery3.Open;
    if not ADOQuery3.Eof then
      NazivArtikla := ADOQuery3.FieldByName('naziv').AsString
    else
      NazivArtikla := '#' + IntToStr(StavkiArtiklaID[i]);

    Item := TListBoxItem.Create(ListBoxStavke);
    Item.Height := 36;
    Item.Text := IntToStr(i+1) + '. ' + NazivArtikla +
      ' - ' + IntToStr(StavkiKolicina[i]) + ' kom.';
    Item.Parent := ListBoxStavke;
  end;
  lblStavkeNaslov.Text := 'Stavke za izdavanje (' + IntToStr(BrojStavki) + ')';
end;

procedure TFormIzdavanjeRobe.btnObrisiStavkuClick(Sender: TObject);
var
  Idx, i: Integer;
begin
  if ListBoxStavke.ItemIndex = -1 then
  begin
    ShowMessage('Izaberite stavku za brisanje.');
    Exit;
  end;
  Idx := ListBoxStavke.ItemIndex;
  for i := Idx to BrojStavki - 2 do
  begin
    StavkiArtiklaID[i] := StavkiArtiklaID[i+1];
    StavkiKolicina[i]  := StavkiKolicina[i+1];
  end;
  Dec(BrojStavki);
  AzurirajListu;
end;

function TFormIzdavanjeRobe.ValidacijaUnosa: Boolean;
begin
  Result := False;
  if ComboVozilo.ItemIndex = -1 then
  begin
    ShowMessage('Molimo izaberite vozilo.');
    Exit;
  end;
  if Trim(txtKorisnik.Text) = '' then
  begin
    ShowMessage('Molimo unesite ime korisnika koji preuzima robu.');
    Exit;
  end;
  if BrojStavki = 0 then
  begin
    ShowMessage('Dodajte bar jedan artikal za izdavanje.');
    Exit;
  end;
  Result := True;
end;

procedure TFormIzdavanjeRobe.btnIzdajClick(Sender: TObject);
var
  IDVozila, IDIzdavanja, i: Integer;
begin
  if not ValidacijaUnosa then
    Exit;

  IDVozila := Integer(ComboVozilo.Items.Objects[ComboVozilo.ItemIndex]);

  if not ADOConnection1.Connected then PoveziBazu;

  try
    // Korak 1: INSERT header u nalog_izdavanje
    ADOQuery1.Close;
    ADOQuery1.SQL.Text :=
      'INSERT INTO nalog_izdavanje ' +
      '(id_vozila, datum_izdavanja, status, napomena, korisnik) ' +
      'VALUES (' +
      IntToStr(IDVozila) + ', Now(), ''Izdato'', '''', ' +
      '''' + Trim(txtKorisnik.Text) + ''')';
    ADOQuery1.ExecSQL;

    // Korak 2: Dohvati ID kreiranog izdavanja
    ADOQuery1.Close;
    ADOQuery1.SQL.Text := 'SELECT @@IDENTITY AS id';
    ADOQuery1.Open;
    IDIzdavanja := ADOQuery1.FieldByName('id').AsInteger;

    // Korak 3: Za svaku stavku - INSERT u stavke_izdavanja + UPDATE zalihe
    for i := 0 to BrojStavki - 1 do
    begin
      // INSERT stavka
      ADOQuery2.Close;
      ADOQuery2.SQL.Text :=
        'INSERT INTO stavke_izdavanja (id_izdavanja, id_artikla, kolicina) ' +
        'VALUES (' +
        IntToStr(IDIzdavanja) + ', ' +
        IntToStr(StavkiArtiklaID[i]) + ', ' +
        IntToStr(StavkiKolicina[i]) + ')';
      ADOQuery2.ExecSQL;

      // UPDATE zalihe - smanji kolicinu
      ADOQuery2.Close;
      ADOQuery2.SQL.Text :=
        'UPDATE zalihe SET kolicina_na_stanju = kolicina_na_stanju - ' +
        IntToStr(StavkiKolicina[i]) +
        ' WHERE id_artikla = ' + IntToStr(StavkiArtiklaID[i]);
      ADOQuery2.ExecSQL;
    end;

    ShowMessage('Roba je uspesno izdata!' + sLineBreak +
      'Broj artikala: ' + IntToStr(BrojStavki));

    OcistiFormu;
    VratiSeNaPrethodnu;

  except
    on E: Exception do
      ShowMessage('Greska pri izdavanju robe: ' + E.Message);
  end;
end;

procedure TFormIzdavanjeRobe.OcistiFormu;
begin
  ComboVozilo.ItemIndex := -1;
  ComboArtikl.ItemIndex := -1;
  txtKolicina.Text := '';
  txtKorisnik.Text := '';
  PanelInfo.Visible := False;
  lblVozacIme.Text := 'Vozilo: -';
  BrojStavki := 0;
  AzurirajListu;
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
  OcistiFormu;
  VratiSeNaPrethodnu;
end;

procedure TFormIzdavanjeRobe.SpeedButton1Click(Sender: TObject);
begin
  OcistiFormu;
  VratiSeNaPrethodnu;
end;

end.
