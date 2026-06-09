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
    SpeedButton1: TSpeedButton;
    Text1: TText;
    // Polja forme
    lblVozilo: TLabel;
    ComboVozilo: TComboBox;
    lblVozac: TLabel;
    lblVozacIme: TLabel;
    lblArtikl: TLabel;
    ComboArtikl: TComboBox;
    lblKolicina: TLabel;
    txtKolicina: TEdit;
    lblNapomena: TLabel;
    txtNapomena: TEdit;
    // Info panel
    PanelInfo: TPanel;
    lblDostupnoKolicina: TLabel;
    lblLokacija: TLabel;
    // Dugmad
    btnIzdaj: TButton;
    btnOtkazi: TButton;
    // DB
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
  public
    { Public declarations }
  end;

var
  FormIzdavanjeRobe: TFormIzdavanjeRobe;

implementation

{$R *.fmx}

procedure TFormIzdavanjeRobe.PoveziBazu;
begin
  ADOConnection1.ConnectionString :=
    'Provider=Microsoft.Jet.OLEDB.4.0;' +
    'Data Source=' + ExtractFilePath(ParamStr(0)) + 'mpmtransport.mdb;';
  ADOConnection1.LoginPrompt := False;
  ADOConnection1.Connected := True;
end;

procedure TFormIzdavanjeRobe.FormCreate(Sender: TObject);
begin
  PoveziBazu;
  PopuniVozila;
  PopuniArtikle;
  lblVozacIme.Text := 'Vozac: —';
  PanelInfo.Visible := False;
end;

procedure TFormIzdavanjeRobe.PopuniVozila;
begin
  ComboVozilo.Items.Clear;
  ADOQuery1.Close;
  ADOQuery1.SQL.Text :=
    'SELECT id_vozila, registarski_broj FROM vozila ' +
    'WHERE status = ''Slobodno'' ORDER BY registarski_broj';
  ADOQuery1.Open;
  while not ADOQuery1.Eof do
  begin
    ComboVozilo.Items.AddObject(
      ADOQuery1.FieldByName('registarski_broj').AsString,
      TObject(ADOQuery1.FieldByName('id_vozila').AsInteger)
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
      ADOQuery1.FieldByName('naziv').AsString + ' (' +
      ADOQuery1.FieldByName('kolicina_na_stanju').AsString + ' kom.)',
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
    lblVozacIme.Text := 'Vozac: —';
    Exit;
  end;

  IDVozila := Integer(ComboVozilo.Items.Objects[ComboVozilo.ItemIndex]);

  ADOQuery2.Close;
  ADOQuery2.SQL.Text :=
    'SELECT v.ime, v.prezime FROM vozaci v ' +
    'INNER JOIN vozila voz ON voz.id_vozaca = v.id_vozaca ' +
    'WHERE voz.id_vozila = :id';
  ADOQuery2.Parameters.ParamByName('id').Value := IDVozila;
  ADOQuery2.Open;

  if not ADOQuery2.Eof then
    lblVozacIme.Text := 'Vozac: ' +
      ADOQuery2.FieldByName('ime').AsString + ' ' +
      ADOQuery2.FieldByName('prezime').AsString
  else
    lblVozacIme.Text := 'Vozac: nije dodeljen';
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
    'SELECT kolicina_na_stanju, lokacija_u_magacinu FROM zalihe WHERE id_artikla = :id';
  ADOQuery3.Parameters.ParamByName('id').Value := IDArtikla;
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
  IDArtikla, KolicinaZahtev, KolicinaNaStanju: Integer;
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

  if Trim(txtKolicina.Text) = '' then
  begin
    ShowMessage('Molimo unesite kolicinu.');
    Exit;
  end;

  if not TryStrToInt(txtKolicina.Text, KolicinaZahtev) or (KolicinaZahtev <= 0) then
  begin
    ShowMessage('Kolicina mora biti pozitivan ceo broj.');
    Exit;
  end;

  // Provjera da li ima dovoljno na stanju
  IDArtikla := Integer(ComboArtikl.Items.Objects[ComboArtikl.ItemIndex]);
  ADOQuery3.Close;
  ADOQuery3.SQL.Text :=
    'SELECT kolicina_na_stanju FROM zalihe WHERE id_artikla = :id';
  ADOQuery3.Parameters.ParamByName('id').Value := IDArtikla;
  ADOQuery3.Open;

  KolicinaNaStanju := ADOQuery3.FieldByName('kolicina_na_stanju').AsInteger;

  if KolicinaZahtev > KolicinaNaStanju then
  begin
    ShowMessage('Nema dovoljno na stanju!' + sLineBreak +
                'Trazeno: ' + IntToStr(KolicinaZahtev) + sLineBreak +
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

  try
    // Kreiranje naloga za izdavanje
    ADOQuery1.Close;
    ADOQuery1.SQL.Text :=
      'INSERT INTO nalog_izdavanje ' +
      '(id_vozila, id_artikla, kolicina, datum_izdavanja, status, napomena) ' +
      'VALUES (:vozilo, :artikal, :kol, :datum, :status, :napomena)';

    ADOQuery1.Parameters.ParamByName('vozilo').Value   := IDVozila;
    ADOQuery1.Parameters.ParamByName('artikal').Value  := IDArtikla;
    ADOQuery1.Parameters.ParamByName('kol').Value      := Kolicina;
    ADOQuery1.Parameters.ParamByName('datum').Value    := Now;
    ADOQuery1.Parameters.ParamByName('status').Value   := 'Izdato';
    ADOQuery1.Parameters.ParamByName('napomena').Value := Trim(txtNapomena.Text);
    ADOQuery1.ExecSQL;

    // Azuriranje stanja zaliha - smanjiti kolicinu
    ADOQuery2.Close;
    ADOQuery2.SQL.Text :=
      'UPDATE zalihe SET kolicina_na_stanju = kolicina_na_stanju - :kol ' +
      'WHERE id_artikla = :id';
    ADOQuery2.Parameters.ParamByName('kol').Value := Kolicina;
    ADOQuery2.Parameters.ParamByName('id').Value  := IDArtikla;
    ADOQuery2.ExecSQL;

    ShowMessage('Roba je uspesno izdata!' + sLineBreak +
                'Artikal: ' + ComboArtikl.Items[ComboArtikl.ItemIndex] + sLineBreak +
                'Kolicina: ' + IntToStr(Kolicina) + ' kom.');
    ModalResult := mrOk;

  except
    on E: Exception do
      ShowMessage('Greska pri izdavanju robe: ' + E.Message);
  end;
end;

procedure TFormIzdavanjeRobe.btnOtkaziClick(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

procedure TFormIzdavanjeRobe.SpeedButton1Click(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

end.
