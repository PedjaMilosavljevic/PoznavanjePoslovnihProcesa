unit NaloziIzdavanje;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs,
  FMX.Controls.Presentation, FMX.StdCtrls, FMX.Objects, FMX.ListBox,
  FMX.Layouts, FMX.ScrollBox,
  Data.DB, Data.Win.ADODB, UnosPodataka;

type
  TFormNaloziIzdavanje = class(TForm)
    SpeedButton1: TSpeedButton;
    lblTitle: TLabel;
    Line1: TLine;
    RectTableHeader: TRectangle;
    lblHBroj: TLabel;
    lblHVozilo: TLabel;
    lblHDatum: TLabel;
    lblHKorisnik: TLabel;
    lblHArtikli: TLabel;
    lblHStatus: TLabel;
    ScrollBox1: TScrollBox;
    ListBoxIzdavanja: TListBox;
    RectBtnBar: TRectangle;
    btnNoviNalog: TButton;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    procedure FormCreate(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure btnNoviNalogClick(Sender: TObject);
  private
    procedure PoveziBazu;
    procedure UcitajIzdavanja;
    function StatusBoja(Status: string): TAlphaColor;
    procedure DodajRed(Broj, Vozilo, Datum, Korisnik, Artikli, Status: string);
  public
  end;

var
  FormNaloziIzdavanje: TFormNaloziIzdavanje;

implementation

{$R *.fmx}

uses IzdavanjeRobe;

procedure TFormNaloziIzdavanje.PoveziBazu;
begin
  ADOConnection1.ConnectionString :=
    'Provider=Microsoft.Jet.OLEDB.4.0;' +
    'Data Source=' + ExtractFilePath(ParamStr(0)) + 'mpmtransport.mdb;';
  ADOConnection1.LoginPrompt := False;
  ADOConnection1.Connected := True;
end;

procedure TFormNaloziIzdavanje.FormCreate(Sender: TObject);
begin
  PoveziBazu;
  UcitajIzdavanja;
end;

function TFormNaloziIzdavanje.StatusBoja(Status: string): TAlphaColor;
begin
  if Status = 'Izdato' then Result := $FF1D9E75
  else if Status = 'Ceka preuzimanje' then Result := $FFEF9F27
  else Result := $FFE05050;
end;

procedure TFormNaloziIzdavanje.DodajRed(Broj, Vozilo, Datum, Korisnik, Artikli, Status: string);
var
  Item: TListBoxItem;
  rectRow, rectStatus: TRectangle;
  lblBr, lblVoz, lblDat, lblKor, lblArt, lblS: TLabel;
  lineSep: TLine;
begin
  Item := TListBoxItem.Create(ListBoxIzdavanja);
  Item.Height := 60;
  Item.Selectable := True;

  rectRow := TRectangle.Create(Item);
  rectRow.Parent := Item;
  rectRow.Align := TAlignLayout.Client;
  rectRow.Fill.Color := TAlphaColors.White;
  rectRow.Stroke.Kind := TBrushKind.None;

  // Broj
  lblBr := TLabel.Create(rectRow);
  lblBr.Parent := rectRow;
  lblBr.Position.X := 10;
  lblBr.Position.Y := 8;
  lblBr.Width := 58;
  lblBr.Height := 16;
  lblBr.Text := Broj;
  lblBr.TextSettings.Font.Size := 10;
  lblBr.TextSettings.Font.Style := [TFontStyle.fsBold];
  lblBr.TextSettings.FontColor := $FF222222;

  // Vozilo
  lblVoz := TLabel.Create(rectRow);
  lblVoz.Parent := rectRow;
  lblVoz.Position.X := 72;
  lblVoz.Position.Y := 8;
  lblVoz.Width := 52;
  lblVoz.Height := 32;
  lblVoz.Text := Vozilo;
  lblVoz.WordWrap := True;
  lblVoz.TextSettings.Font.Size := 10;
  lblVoz.TextSettings.FontColor := $FF444444;

  // Datum
  lblDat := TLabel.Create(rectRow);
  lblDat.Parent := rectRow;
  lblDat.Position.X := 128;
  lblDat.Position.Y := 8;
  lblDat.Width := 60;
  lblDat.Height := 16;
  lblDat.Text := Datum;
  lblDat.TextSettings.Font.Size := 10;
  lblDat.TextSettings.FontColor := $FF666666;

  // Korisnik
  lblKor := TLabel.Create(rectRow);
  lblKor.Parent := rectRow;
  lblKor.Position.X := 192;
  lblKor.Position.Y := 8;
  lblKor.Width := 54;
  lblKor.Height := 32;
  lblKor.Text := Korisnik;
  lblKor.WordWrap := True;
  lblKor.TextSettings.Font.Size := 10;
  lblKor.TextSettings.FontColor := $FF444444;

  // Artikli
  lblArt := TLabel.Create(rectRow);
  lblArt.Parent := rectRow;
  lblArt.Position.X := 250;
  lblArt.Position.Y := 8;
  lblArt.Width := 62;
  lblArt.Height := 32;
  lblArt.Text := Artikli;
  lblArt.WordWrap := True;
  lblArt.TextSettings.Font.Size := 10;
  lblArt.TextSettings.FontColor := $FF444444;

  // Status badge
  rectStatus := TRectangle.Create(rectRow);
  rectStatus.Parent := rectRow;
  rectStatus.Position.X := 320;
  rectStatus.Position.Y := 10;
  rectStatus.Width := 62;
  rectStatus.Height := 22;
  rectStatus.XRadius := 6;
  rectStatus.YRadius := 6;
  rectStatus.Fill.Color := $22000000;
  rectStatus.Stroke.Kind := TBrushKind.None;

  lblS := TLabel.Create(rectStatus);
  lblS.Parent := rectStatus;
  lblS.Align := TAlignLayout.Client;
  lblS.Text := Status;
  lblS.TextSettings.Font.Size := 9;
  lblS.TextSettings.FontColor := StatusBoja(Status);
  lblS.TextSettings.HorzAlign := TTextAlign.Center;

  lineSep := TLine.Create(rectRow);
  lineSep.Parent := rectRow;
  lineSep.Align := TAlignLayout.Bottom;
  lineSep.Height := 1;
  lineSep.Stroke.Color := $FFEEEEEE;

  Item.Parent := ListBoxIzdavanja;
end;

procedure TFormNaloziIzdavanje.UcitajIzdavanja;
begin
  ListBoxIzdavanja.Clear;

  ADOQuery1.Close;
  ADOQuery1.SQL.Text :=
    'SELECT ni.id_izdavanja, v.registarski_broj, ni.datum_izdavanja, ' +
    '       voz.ime + '' '' + voz.prezime AS korisnik, ' +
    '       z.naziv + '' x '' + CStr(ni.kolicina) AS artikli, ' +
    '       ni.status ' +
    'FROM nalog_izdavanje ni ' +
    'INNER JOIN vozila  v   ON v.id_vozila   = ni.id_vozila ' +
    'LEFT  JOIN vozaci  voz ON voz.id_vozaca = v.id_vozaca ' +
    'INNER JOIN zalihe  z   ON z.id_artikla  = ni.id_artikla ' +
    'ORDER BY ni.datum_izdavanja DESC';
  ADOQuery1.Open;

  while not ADOQuery1.Eof do
  begin
    DodajRed(
      '#IZD-' + ADOQuery1.FieldByName('id_izdavanja').AsString,
      ADOQuery1.FieldByName('registarski_broj').AsString,
      FormatDateTime('dd.mm.yyyy', ADOQuery1.FieldByName('datum_izdavanja').AsDateTime),
      ADOQuery1.FieldByName('korisnik').AsString,
      ADOQuery1.FieldByName('artikli').AsString,
      ADOQuery1.FieldByName('status').AsString
    );
    ADOQuery1.Next;
  end;
end;

procedure TFormNaloziIzdavanje.SpeedButton1Click(Sender: TObject);
begin
  Form2.Show;
  Close;
end;

procedure TFormNaloziIzdavanje.btnNoviNalogClick(Sender: TObject);
begin
  FormIzdavanjeRobe.ShowModal;
  UcitajIzdavanja;
end;

end.
