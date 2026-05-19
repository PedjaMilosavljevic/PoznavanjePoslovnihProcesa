unit NovaNabavka;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs,
  FMX.Controls.Presentation, FMX.StdCtrls, FMX.Objects, FMX.Edit,
  FMX.ListBox, FMX.Layouts, FMX.DateTimeCtrls,
  Data.DB, Data.Win.ADODB, UnosPodataka;

type
  TFormNovaNabavka = class(TForm)
    SpeedButton1: TSpeedButton;
    Text1: TText;
    // Polja forme
    lblDobavljac: TLabel;
    ComboDobavljac: TComboBox;
    lblArtikl: TLabel;
    ComboArtikl: TComboBox;
    lblKolicina: TLabel;
    txtKolicina: TEdit;
    lblCenaPoKomadu: TLabel;
    txtCena: TEdit;
    lblDatumIsporuke: TLabel;
    DatePickerIsporuka: TDateEdit;
    lblNapomena: TLabel;
    txtNapomena: TEdit;
    // Info panel - trenutno stanje
    PanelInfo: TPanel;
    lblTrenutnoStanje: TLabel;
    lblMinimum: TLabel;
    // Dugmad
    btnSacuvaj: TButton;
    btnOtkazi: TButton;
    // DB
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    ADOQuery2: TADOQuery;
    procedure FormCreate(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure btnSacuvajClick(Sender: TObject);
    procedure btnOtkaziClick(Sender: TObject);
    procedure ComboArtiklChange(Sender: TObject);
  private
    procedure PoveziBazu;
    procedure PopuniDobavljace;
    procedure PopuniArtikle;
    procedure PrikaziStanjeArtikla;
    function ValidacijaUnosa: Boolean;
  public
    { Public declarations }
  end;

var
  FormNovaNabavka: TFormNovaNabavka;

implementation

{$R *.fmx}

procedure TFormNovaNabavka.PoveziBazu;
begin
  ADOConnection1.ConnectionString :=
    'Provider=Microsoft.Jet.OLEDB.4.0;' +
    'Data Source=' + ExtractFilePath(ParamStr(0)) + 'mpmtransport.mdb;';
  ADOConnection1.LoginPrompt := False;
  ADOConnection1.Connected := True;
end;

procedure TFormNovaNabavka.FormCreate(Sender: TObject);
begin
  PoveziBazu;
  PopuniDobavljace;
  PopuniArtikle;
  DatePickerIsporuka.Date := Now + 7;
  PanelInfo.Visible := False;
end;

procedure TFormNovaNabavka.PopuniDobavljace;
begin
  ComboDobavljac.Items.Clear;
  ADOQuery1.Close;
  ADOQuery1.SQL.Text := 'SELECT id_dobavljaca, naziv_firme FROM dobavljaci ORDER BY naziv_firme';
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
    'SELECT id_artikla, naziv, kolicina_na_stanju, min_kolicina ' +
    'FROM zalihe ORDER BY naziv';
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
    'SELECT kolicina_na_stanju, min_kolicina FROM zalihe WHERE id_artikla = :id';
  ADOQuery2.Parameters.ParamByName('id').Value := IDArtikla;
  ADOQuery2.Open;

  if not ADOQuery2.Eof then
  begin
    lblTrenutnoStanje.Text := 'Trenutno na stanju: ' +
      ADOQuery2.FieldByName('kolicina_na_stanju').AsString + ' kom.';
    lblMinimum.Text := 'Minimalna zaliha: ' +
      ADOQuery2.FieldByName('min_kolicina').AsString + ' kom.';
    PanelInfo.Visible := True;
  end;
end;

function TFormNovaNabavka.ValidacijaUnosa: Boolean;
begin
  Result := False;

  if ComboDobavljac.ItemIndex = -1 then
  begin
    ShowMessage('Molimo izaberite dobavljaca.');
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

  if not TryStrToInt(txtKolicina.Text, Integer(0)) or
     (StrToInt(txtKolicina.Text) <= 0) then
  begin
    ShowMessage('Kolicina mora biti pozitivan ceo broj.');
    Exit;
  end;

  if Trim(txtCena.Text) = '' then
  begin
    ShowMessage('Molimo unesite cenu po komadu.');
    Exit;
  end;

  Result := True;
end;

procedure TFormNovaNabavka.btnSacuvajClick(Sender: TObject);
var
  IDDobavljaca, IDArtikla, Kolicina: Integer;
  CenaPoKomadu, UkupnaVrednost: Double;
  DatumNaloga, DatumIsporuke: TDateTime;
begin
  if not ValidacijaUnosa then
    Exit;

  IDDobavljaca  := Integer(ComboDobavljac.Items.Objects[ComboDobavljac.ItemIndex]);
  IDArtikla     := Integer(ComboArtikl.Items.Objects[ComboArtikl.ItemIndex]);
  Kolicina      := StrToInt(txtKolicina.Text);
  CenaPoKomadu  := StrToFloat(StringReplace(txtCena.Text, ',', '.', [rfReplaceAll]));
  UkupnaVrednost := Kolicina * CenaPoKomadu;
  DatumNaloga   := Now;
  DatumIsporuke := DatePickerIsporuka.Date;

  try
    // Unos naloga za nabavku
    ADOQuery1.Close;
    ADOQuery1.SQL.Text :=
      'INSERT INTO nalog_nabavka ' +
      '(id_dobavljaca, id_artikla, kolicina, cena_po_komadu, ukupna_vrednost, ' +
      'datum_naloga, datum_isporuke, status, napomena) ' +
      'VALUES (:dobavljac, :artikal, :kol, :cena, :ukupno, :datum, :isporuka, :status, :napomena)';

    ADOQuery1.Parameters.ParamByName('dobavljac').Value := IDDobavljaca;
    ADOQuery1.Parameters.ParamByName('artikal').Value   := IDArtikla;
    ADOQuery1.Parameters.ParamByName('kol').Value       := Kolicina;
    ADOQuery1.Parameters.ParamByName('cena').Value      := CenaPoKomadu;
    ADOQuery1.Parameters.ParamByName('ukupno').Value    := UkupnaVrednost;
    ADOQuery1.Parameters.ParamByName('datum').Value     := DatumNaloga;
    ADOQuery1.Parameters.ParamByName('isporuka').Value  := DatumIsporuke;
    ADOQuery1.Parameters.ParamByName('status').Value    := 'U obradi';
    ADOQuery1.Parameters.ParamByName('napomena').Value  := Trim(txtNapomena.Text);
    ADOQuery1.ExecSQL;

    ShowMessage('Narudzbenica je uspesno kreirana!' + sLineBreak +
                'Ukupna vrednost: ' + FormatFloat('#,##0.00', UkupnaVrednost) + ' RSD');
    ModalResult := mrOk;

  except
    on E: Exception do
      ShowMessage('Greska pri cuvanju narudzbenice: ' + E.Message);
  end;
end;

procedure TFormNovaNabavka.btnOtkaziClick(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

procedure TFormNovaNabavka.SpeedButton1Click(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

end.
