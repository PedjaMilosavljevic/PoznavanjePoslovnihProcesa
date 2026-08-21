unit Narudzbenice;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs,
  FMX.Controls.Presentation, FMX.StdCtrls, FMX.Objects, FMX.ListBox,
  FMX.Layouts, FMX.ScrollBox,
  Data.DB, Data.Win.ADODB, UnosPodataka;

type
  TFormNarudzbenice = class(TForm)
    RectHeader: TRectangle;
    SpeedButton1: TSpeedButton;
    lblTitle: TLabel;
    Line1: TLine;
    RectTableHeader: TRectangle;
    lblHBroj: TLabel;
    lblHDob: TLabel;
    lblHDatum: TLabel;
    lblHArtikli: TLabel;
    lblHVred: TLabel;
    lblHStatus: TLabel;
    ScrollBox1: TScrollBox;
    ListBoxNabavke: TListBox;
    RectBtnBar: TRectangle;
    btnNoviNalog: TButton;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    ADOQuery2: TADOQuery;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure btnNoviNalogClick(Sender: TObject);
  private
    procedure PoveziBazu;
    procedure UcitajNabavke;
    function StatusBoja(Status: string): TAlphaColor;
    procedure DodajRed(Broj, Dobavljac, Datum, Artikli, Vrednost, Status: string;
      IDNaloga: Integer);
    procedure PotvrdiPrijemClick(Sender: TObject);
  end;

var
  FormNarudzbenice: TFormNarudzbenice;

implementation

{$R *.fmx}

uses Zalihe;



procedure TFormNarudzbenice.PoveziBazu;
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

procedure TFormNarudzbenice.FormCreate(Sender: TObject);
begin
  PoveziBazu;
  UcitajNabavke;
end;

procedure TFormNarudzbenice.FormShow(Sender: TObject);
begin
  UcitajNabavke;
end;

function TFormNarudzbenice.StatusBoja(Status: string): TAlphaColor;
begin
  if      Status = 'Primljeno'     then Result := $FF1D9E75
  else if Status = 'U obradi'      then Result := $FF2980B9
  else if Status = 'Ceka potvrdu'  then Result := $FFEF9F27
  else                                  Result := $FFE05050;
end;

procedure TFormNarudzbenice.DodajRed(Broj, Dobavljac, Datum, Artikli, Vrednost, Status: string;
  IDNaloga: Integer);
var
  Item: TListBoxItem;
  Layout: TLayout;
  lblBr, lblDob, lblDat, lblArt, lblVred: TLabel;
  rectStatus: TRectangle;
  lblS: TLabel;
  lineSep: TLine;
  btnPotvrdi: TButton;
begin
  Item := TListBoxItem.Create(ListBoxNabavke);
  Item.Height := 92;  // uniformna visina za sve redove (FMX TListBox ima problema sa promenljivom visinom)
  Item.Parent := ListBoxNabavke;

  Layout := TLayout.Create(Item);
  Layout.Parent := Item;
  Layout.Align  := TAlignLayout.Client;

  lblBr := TLabel.Create(Layout);
  lblBr.Parent := Layout;
  lblBr.Position.X := 10;
  lblBr.Position.Y := 8;
  lblBr.Width  := 70;
  lblBr.Height := 16;
  lblBr.Text   := Broj;
  lblBr.TextSettings.Font.Size  := 10;
  lblBr.TextSettings.Font.Style := [TFontStyle.fsBold];
  lblBr.TextSettings.FontColor  := $FF222222;

  lblDob := TLabel.Create(Layout);
  lblDob.Parent := Layout;
  lblDob.Position.X := 86;
  lblDob.Position.Y := 8;
  lblDob.Width    := 68;
  lblDob.Height   := 32;
  lblDob.Text     := Dobavljac;
  lblDob.WordWrap := True;
  lblDob.TextSettings.Font.Size := 10;
  lblDob.TextSettings.FontColor := $FF444444;

  lblDat := TLabel.Create(Layout);
  lblDat.Parent := Layout;
  lblDat.Position.X := 160;
  lblDat.Position.Y := 8;
  lblDat.Width  := 58;
  lblDat.Height := 16;
  lblDat.Text   := Datum;
  lblDat.TextSettings.Font.Size := 10;
  lblDat.TextSettings.FontColor := $FF666666;

  lblArt := TLabel.Create(Layout);
  lblArt.Parent := Layout;
  lblArt.Position.X := 224;
  lblArt.Position.Y := 8;
  lblArt.Width    := 44;
  lblArt.Height   := 32;
  lblArt.Text     := Artikli;
  lblArt.WordWrap := True;
  lblArt.TextSettings.Font.Size := 10;
  lblArt.TextSettings.FontColor := $FF666666;

  lblVred := TLabel.Create(Layout);
  lblVred.Parent := Layout;
  lblVred.Position.X := 272;
  lblVred.Position.Y := 8;
  lblVred.Width    := 52;
  lblVred.Height   := 32;
  lblVred.Text     := Vrednost;
  lblVred.WordWrap := True;
  lblVred.TextSettings.Font.Size  := 10;
  lblVred.TextSettings.Font.Style := [TFontStyle.fsBold];
  lblVred.TextSettings.FontColor  := $FF222222;

  rectStatus := TRectangle.Create(Layout);
  rectStatus.Parent     := Layout;
  rectStatus.Position.X := 330;
  rectStatus.Position.Y := 10;
  rectStatus.Width   := 50;
  rectStatus.Height  := 22;
  rectStatus.XRadius := 6;
  rectStatus.YRadius := 6;
  rectStatus.Fill.Color  := StatusBoja(Status);
  rectStatus.Stroke.Kind := TBrushKind.None;

  lblS := TLabel.Create(rectStatus);
  lblS.Parent := rectStatus;
  lblS.Align  := TAlignLayout.Client;
  lblS.Text   := Status;
  lblS.TextSettings.Font.Size := 9;
  lblS.TextSettings.FontColor := TAlphaColors.White;
  lblS.TextSettings.HorzAlign := TTextAlign.Center;
  lblS.TextSettings.VertAlign := TTextAlign.Center;

  // -- Dugme "Potvrdi prijem" - vidljivo samo za naloge 'U obradi' --
  btnPotvrdi := TButton.Create(Layout);
  btnPotvrdi.Parent := Layout;
  btnPotvrdi.Position.X := 10;
  btnPotvrdi.Position.Y := 50;
  btnPotvrdi.Width  := 372;
  btnPotvrdi.Height := 34;
  btnPotvrdi.Text   := 'Potvrdi prijem robe';
  btnPotvrdi.Tag    := IDNaloga;
  btnPotvrdi.OnClick := PotvrdiPrijemClick;
  btnPotvrdi.Visible := (Status = 'U obradi');

  lineSep := TLine.Create(Layout);
  lineSep.Parent := Layout;
  lineSep.Align  := TAlignLayout.Bottom;
  lineSep.Height := 1;
  lineSep.Stroke.Color := $FFEEEEEE;
end;

procedure TFormNarudzbenice.PotvrdiPrijemClick(Sender: TObject);
var
  IDNaloga, IDArtikla, Kolicina, BrojAzuriranih: Integer;
  TrenutniStatus: string;
  ImaStavki: Boolean;
begin
  IDNaloga := TButton(Sender).Tag;

  // Provjeri status naloga
  ADOQuery1.Close;
  ADOQuery1.SQL.Text :=
    'SELECT status FROM nalog_nabavka WHERE id_naloga = ' + IntToStr(IDNaloga);
  ADOQuery1.Open;

  if ADOQuery1.Eof then
  begin
    ShowMessage('Nalog nije pronadjen.');
    Exit;
  end;

  TrenutniStatus := ADOQuery1.FieldByName('status').AsString;
  if TrenutniStatus <> 'U obradi' then
  begin
    ShowMessage('Ovaj nalog je vec obradjen (status: ' + TrenutniStatus + ').');
    Exit;
  end;

  // Provjeri da li nalog ima stavke u stavke_nabavke (novi tip naloga)
  ADOQuery1.Close;
  ADOQuery1.SQL.Text :=
    'SELECT id_artikla, kolicina FROM stavke_nabavke WHERE id_naloga = ' +
    IntToStr(IDNaloga);
  ADOQuery1.Open;
  ImaStavki := not ADOQuery1.Eof;

  try
  
    if ImaStavki then
    begin
      // Novi tip naloga: azuriraj sve artikle iz stavke_nabavke
      while not ADOQuery1.Eof do
      begin
        IDArtikla := ADOQuery1.FieldByName('id_artikla').AsInteger;
        Kolicina  := ADOQuery1.FieldByName('kolicina').AsInteger;

        ADOQuery1.DisableControls;
        // Sacuvaj poziciju - koristimo drugi query za UPDATE
        ADOQuery1.BookmarkValid(ADOQuery1.Bookmark);

              ADOQuery1.Next;
      end;

      // Sad azuriraj - koristimo ADOQuery1 ponovo
      ADOQuery1.Close;
      ADOQuery1.SQL.Text :=
        'SELECT id_artikla, kolicina FROM stavke_nabavke WHERE id_naloga = ' +
        IntToStr(IDNaloga);
      ADOQuery1.Open;

      while not ADOQuery1.Eof do
      begin
        IDArtikla := ADOQuery1.FieldByName('id_artikla').AsInteger;
        Kolicina  := ADOQuery1.FieldByName('kolicina').AsInteger;
        ADOQuery1.Next;

        ADOQuery2.Close;
        ADOQuery2.SQL.Text :=
          'UPDATE zalihe SET kolicina_na_stanju = kolicina_na_stanju + ' +
          IntToStr(Kolicina) + ' WHERE id_artikla = ' + IntToStr(IDArtikla);
        ADOQuery2.ExecSQL;
      end;
    end
    else
    begin
      // Stari tip naloga: jedan artikal u nalog_nabavka
      ADOQuery1.Close;
      ADOQuery1.SQL.Text :=
        'SELECT id_artikla, kolicina FROM nalog_nabavka WHERE id_naloga = ' +
        IntToStr(IDNaloga);
      ADOQuery1.Open;

      if not ADOQuery1.Eof then
      begin
        IDArtikla := ADOQuery1.FieldByName('id_artikla').AsInteger;
        Kolicina  := ADOQuery1.FieldByName('kolicina').AsInteger;
  
        ADOQuery2.Close;
        ADOQuery2.SQL.Text :=
          'UPDATE zalihe SET kolicina_na_stanju = kolicina_na_stanju + ' +
          IntToStr(Kolicina) + ' WHERE id_artikla = ' + IntToStr(IDArtikla);
        ADOQuery2.ExecSQL;
      end;
    end;

    // Azuriraj status naloga
    ADOQuery1.Close;
    ADOQuery1.SQL.Text :=
      'UPDATE nalog_nabavka SET status = ''Primljeno'' WHERE id_naloga = ' +
      IntToStr(IDNaloga);
    ADOQuery1.ExecSQL;

    ShowMessage('Prijem robe je potvrdjen!' + sLineBreak +
      'Zalihe su azurirane.');

    UcitajNabavke;
  except
    on E: Exception do
      ShowMessage('Greska pri potvrdi prijema: ' + E.Message);
  end;
end;


procedure TFormNarudzbenice.UcitajNabavke;
begin
  ListBoxNabavke.Clear;
  ADOQuery1.Close;
  ADOQuery1.SQL.Text :=
    'SELECT nn.id_naloga, d.naziv_firme, nn.datum_naloga, ' +
    '       nn.ukupna_vrednost, nn.status, ' +
    '       (SELECT COUNT(*) FROM stavke_nabavke s WHERE s.id_naloga = nn.id_naloga) AS br_stavki ' +
    'FROM nalog_nabavka nn ' +
    'INNER JOIN dobavljaci d ON d.id_dobavljaca = nn.id_dobavljaca ' +
    'ORDER BY nn.datum_naloga DESC';
  ADOQuery1.Open;

  while not ADOQuery1.Eof do
  begin
    DodajRed(
      '#NOV-' + ADOQuery1.FieldByName('id_naloga').AsString,
      ADOQuery1.FieldByName('naziv_firme').AsString,
      FormatDateTime('dd.mm.yyyy', ADOQuery1.FieldByName('datum_naloga').AsDateTime),
      ADOQuery1.FieldByName('br_stavki').AsString + ' artikala',
      FormatFloat('#,##0', ADOQuery1.FieldByName('ukupna_vrednost').AsFloat) + ' RSD',
      ADOQuery1.FieldByName('status').AsString,
      ADOQuery1.FieldByName('id_naloga').AsInteger
    );
    ADOQuery1.Next;
  end;
end;

procedure TFormNarudzbenice.SpeedButton1Click(Sender: TObject);
begin
  FormZalihe.Show;
  Hide;
end;

procedure TFormNarudzbenice.btnNoviNalogClick(Sender: TObject);
var
  F: TForm;
begin
  PrethodnaFormaZalihe := 'FormNarudzbenice';
  F := TForm(Application.FindComponent('FormNovaNabavka'));
  if Assigned(F) then
  begin
    F.Show;
    Hide;
  end
  else
    ShowMessage('Forma za nabavku nije ucitana u projekat.');
end;

end.
