unit NovaNabavka;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs,
  FMX.Controls.Presentation, FMX.StdCtrls, FMX.Objects, FMX.Edit,
  FMX.ListBox, FMX.Layouts, FMX.DateTimeCtrls,
  Data.DB, Data.Win.ADODB, UnosPodataka, Zalihe;

type
  TFormNovaNabavka = class(TForm)
    RectHeader: TRectangle;
    SpeedButton1: TSpeedButton;
    lblTitle: TLabel;
    RectBody: TRectangle;
    lblDobavljac: TLabel;
    ComboDobavljac: TComboBox;
    lblDatumIsporuke: TLabel;
    DatePickerIsporuka: TDateEdit;
    lblNapomena: TLabel;
    txtNapomena: TEdit;
    Line1: TLine;
    lblDodajArtikl: TLabel;
    lblArtikl: TLabel;
    ComboArtikl: TComboBox;
    PanelInfo: TPanel;
    lblTrenutnoStanje: TLabel;
    lblMinimum: TLabel;
    lblKolicina: TLabel;
    txtKolicina: TEdit;
    lblCena: TLabel;
    txtCena: TEdit;
    btnDodajStavku: TButton;
    Line2: TLine;
    lblStavkeNaslov: TLabel;
    btnObrisiStavku: TButton;
    ListBoxStavke: TListBox;
    lblUkupno: TLabel;
    btnSacuvaj: TButton;
    btnOtkazi: TButton;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    ADOQuery2: TADOQuery;
    procedure FormCreate(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure btnSacuvajClick(Sender: TObject);
    procedure btnOtkaziClick(Sender: TObject);
    procedure ComboArtiklChange(Sender: TObject);
    procedure btnDodajStavkuClick(Sender: TObject);
    procedure btnObrisiStavkuClick(Sender: TObject);
  private
    // In-memory cart: parallel arrays for current order items
    StavkiArtiklaID : array[0..99] of Integer;
    StavkiKolicina  : array[0..99] of Integer;
    StavkiCena      : array[0..99] of Double;
    BrojStavki      : Integer;

    procedure PoveziBazu;
    procedure PopuniDobavljace;
    procedure PopuniArtikle;
    procedure PrikaziStanjeArtikla;
    procedure AzurirajListuIUkupno;
    procedure VratiSeNaPrethodnu;
    procedure OcistiFormu;
  end;

var
  FormNovaNabavka: TFormNovaNabavka;

implementation

{$R *.fmx}

procedure TFormNovaNabavka.PoveziBazu;
var
  dbPath: string;
begin
  dbPath := ExtractFilePath(ParamStr(0)) + '..\..\..\ Baza podataka\mpmtransport.mdb';
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

procedure TFormNovaNabavka.FormCreate(Sender: TObject);
begin
  PoveziBazu;
  PopuniDobavljace;
  PopuniArtikle;
  DatePickerIsporuka.Date := Now + 7;
  PanelInfo.Visible := False;
  BrojStavki := 0;
  AzurirajListuIUkupno;
end;

procedure TFormNovaNabavka.PopuniDobavljace;
begin
  ComboDobavljac.Items.Clear;
  ADOQuery1.Close;
  ADOQuery1.SQL.Text :=
    'SELECT id_dobavljaca, naziv_firme FROM dobavljaci ORDER BY naziv_firme';
  ADOQuery1.Open;
  while not ADOQuery1.Eof do
  begin
    ComboDobavljac.Items.AddObject(
      ADOQuery1.FieldByName('naziv_firme').AsString,
      TObject(ADOQuery1.FieldByName('id_dobavljaca').AsInteger)
    );
    ADOQuery1.Next;
  end;
end;

procedure TFormNovaNabavka.PopuniArtikle;
begin
  ComboArtikl.Items.Clear;
  ADOQuery1.Close;
  ADOQuery1.SQL.Text :=
    'SELECT id_artikla, naziv FROM zalihe ORDER BY naziv';
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

procedure TFormNovaNabavka.ComboArtiklChange(Sender: TObject);
begin
  PrikaziStanjeArtikla;
end;

procedure TFormNovaNabavka.PrikaziStanjeArtikla;
var
  IDArtikla: Integer;
begin
  if ComboArtikl.ItemIndex = -1 then
  begin
    PanelInfo.Visible := False;
    Exit;
  end;
  IDArtikla := Integer(ComboArtikl.Items.Objects[ComboArtikl.ItemIndex]);
  ADOQuery2.Close;
  ADOQuery2.SQL.Text :=
    'SELECT kolicina_na_stanju, min_kolicina FROM zalihe WHERE id_artikla = ' +
    IntToStr(IDArtikla);
  ADOQuery2.Open;
  if not ADOQuery2.Eof then
  begin
    lblTrenutnoStanje.Text := 'Na stanju: ' +
      ADOQuery2.FieldByName('kolicina_na_stanju').AsString + ' kom.';
    lblMinimum.Text := 'Minimum: ' +
      ADOQuery2.FieldByName('min_kolicina').AsString + ' kom.';
    PanelInfo.Visible := True;
  end;
end;

procedure TFormNovaNabavka.btnDodajStavkuClick(Sender: TObject);
var
  IDArtikla, Kolicina: Integer;
  Cena: Double;
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
  if Trim(txtCena.Text) = '' then
  begin
    ShowMessage('Molimo unesite cenu po komadu.');
    Exit;
  end;

  if BrojStavki >= 100 then
  begin
    ShowMessage('Dostignut maksimalan broj stavki (100).');
    Exit;
  end;

  IDArtikla := Integer(ComboArtikl.Items.Objects[ComboArtikl.ItemIndex]);
  Cena := StrToFloat(StringReplace(txtCena.Text, ',', '.', [rfReplaceAll]));

  // Save to in-memory arrays
  StavkiArtiklaID[BrojStavki] := IDArtikla;
  StavkiKolicina[BrojStavki]  := Kolicina;
  StavkiCena[BrojStavki]      := Cena;
  Inc(BrojStavki);

  // Reset article fields
  ComboArtikl.ItemIndex := -1;
  txtKolicina.Text := '';
  txtCena.Text := '';
  PanelInfo.Visible := False;

  AzurirajListuIUkupno;
end;

procedure TFormNovaNabavka.AzurirajListuIUkupno;
var
  i: Integer;
  UkupnoSvega: Double;
  NazivArtikla: string;
  Item: TListBoxItem;
begin
  ListBoxStavke.Clear;
  UkupnoSvega := 0;

  for i := 0 to BrojStavki - 1 do
  begin
    // Lookup article name
    ADOQuery2.Close;
    ADOQuery2.SQL.Text :=
      'SELECT naziv FROM zalihe WHERE id_artikla = ' +
      IntToStr(StavkiArtiklaID[i]);
    ADOQuery2.Open;
    if not ADOQuery2.Eof then
      NazivArtikla := ADOQuery2.FieldByName('naziv').AsString
    else
      NazivArtikla := 'Artikal #' + IntToStr(StavkiArtiklaID[i]);

    Item := TListBoxItem.Create(ListBoxStavke);
    Item.Height := 36;
    Item.Text := IntToStr(i+1) + '. ' + NazivArtikla +
      ' x ' + IntToStr(StavkiKolicina[i]) + ' kom. = ' +
      FormatFloat('#,##0.00', StavkiKolicina[i] * StavkiCena[i]) + ' RSD';
    Item.Parent := ListBoxStavke;

    UkupnoSvega := UkupnoSvega + StavkiKolicina[i] * StavkiCena[i];
  end;

  lblStavkeNaslov.Text := 'Stavke narudzbenice (' + IntToStr(BrojStavki) + ')';
  lblUkupno.Text := 'Ukupna vrednost: ' +
    FormatFloat('#,##0.00', UkupnoSvega) + ' RSD';
end;

procedure TFormNovaNabavka.btnObrisiStavkuClick(Sender: TObject);
var
  Idx, i: Integer;
begin
  if ListBoxStavke.ItemIndex = -1 then
  begin
    ShowMessage('Izaberite stavku za brisanje.');
    Exit;
  end;
  Idx := ListBoxStavke.ItemIndex;

  // Shift arrays left
  for i := Idx to BrojStavki - 2 do
  begin
    StavkiArtiklaID[i] := StavkiArtiklaID[i+1];
    StavkiKolicina[i]  := StavkiKolicina[i+1];
    StavkiCena[i]      := StavkiCena[i+1];
  end;
  Dec(BrojStavki);

  AzurirajListuIUkupno;
end;

procedure TFormNovaNabavka.btnSacuvajClick(Sender: TObject);
var
  IDDobavljaca, IDNaloga, i: Integer;
  UkupnoSvega, Cena: Double;
begin
  // Validacija
  if ComboDobavljac.ItemIndex = -1 then
  begin
    ShowMessage('Molimo izaberite dobavljaca.');
    Exit;
  end;
  if BrojStavki = 0 then
  begin
    ShowMessage('Dodajte bar jedan artikal u narudzbenicu.');
    Exit;
  end;

  IDDobavljaca := Integer(ComboDobavljac.Items.Objects[ComboDobavljac.ItemIndex]);

  // Racunaj ukupno
  UkupnoSvega := 0;
  for i := 0 to BrojStavki - 1 do
    UkupnoSvega := UkupnoSvega + StavkiKolicina[i] * StavkiCena[i];

  try
    // Korak 1: INSERT header nalog_nabavka (bez id_artikla/kolicina/cena)
    ADOQuery1.Close;
    ADOQuery1.SQL.Text :=
      'INSERT INTO nalog_nabavka ' +
      '(id_dobavljaca, ukupna_vrednost, datum_naloga, datum_isporuke, status, napomena) ' +
      'VALUES (' +
      IntToStr(IDDobavljaca) + ', ' +
      FloatToStr(UkupnoSvega) + ', ' +
      'Now(), ' +
      '#' + FormatDateTime('mm/dd/yyyy', DatePickerIsporuka.Date) + '#, ' +
      '''U obradi'', ' +
      '''' + Trim(txtNapomena.Text) + ''')';
    ADOQuery1.ExecSQL;

    // Korak 2: Dohvati ID kreiranog naloga
    ADOQuery1.Close;
    ADOQuery1.SQL.Text := 'SELECT @@IDENTITY AS id';
    ADOQuery1.Open;
    IDNaloga := ADOQuery1.FieldByName('id').AsInteger;

    // Korak 3: INSERT sve stavke u stavke_nabavke
    for i := 0 to BrojStavki - 1 do
    begin
      Cena := StavkiCena[i];
      ADOQuery2.Close;
      ADOQuery2.SQL.Text :=
        'INSERT INTO stavke_nabavke ' +
        '(id_naloga, id_artikla, kolicina, cena_po_komadu, ukupna_vrednost) ' +
        'VALUES (' +
        IntToStr(IDNaloga) + ', ' +
        IntToStr(StavkiArtiklaID[i]) + ', ' +
        IntToStr(StavkiKolicina[i]) + ', ' +
        FloatToStr(Cena) + ', ' +
        FloatToStr(StavkiKolicina[i] * Cena) + ')';
      ADOQuery2.ExecSQL;
    end;

    ShowMessage('Narudzbenica je uspesno kreirana!' + sLineBreak +
      'Broj artikala: ' + IntToStr(BrojStavki) + sLineBreak +
      'Ukupna vrednost: ' + FormatFloat('#,##0.00', UkupnoSvega) + ' RSD');

    OcistiFormu;
    VratiSeNaPrethodnu;

  except
    on E: Exception do
      ShowMessage('Greska pri cuvanju narudzbenice: ' + E.Message);
  end;
end;

procedure TFormNovaNabavka.OcistiFormu;
begin
  ComboDobavljac.ItemIndex := -1;
  ComboArtikl.ItemIndex := -1;
  txtKolicina.Text := '';
  txtCena.Text := '';
  txtNapomena.Text := '';
  PanelInfo.Visible := False;
  BrojStavki := 0;
  AzurirajListuIUkupno;
end;

procedure TFormNovaNabavka.VratiSeNaPrethodnu;
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

procedure TFormNovaNabavka.btnOtkaziClick(Sender: TObject);
begin
  OcistiFormu;
  VratiSeNaPrethodnu;
end;

procedure TFormNovaNabavka.SpeedButton1Click(Sender: TObject);
begin
  OcistiFormu;
  VratiSeNaPrethodnu;
end;

end.
