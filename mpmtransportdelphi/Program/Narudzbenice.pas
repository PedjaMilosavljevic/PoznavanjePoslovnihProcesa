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
    procedure FormCreate(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure btnNoviNalogClick(Sender: TObject);
  private
    procedure PoveziBazu;
    procedure UcitajNabavke;
    function StatusBoja(Status: string): TAlphaColor;
    procedure DodajRed(Broj, Dobavljac, Datum, Artikli, Vrednost, Status: string);
  public
  end;

var
  FormNarudzbenice: TFormNarudzbenice;

implementation

{$R *.fmx}

uses NovaNabavka;

procedure TFormNarudzbenice.PoveziBazu;
begin
  ADOConnection1.ConnectionString :=
    'Provider=Microsoft.Jet.OLEDB.4.0;' +
    'Data Source=' + ExtractFilePath(ParamStr(0)) + 'mpmtransport.mdb;';
  ADOConnection1.LoginPrompt := False;
  ADOConnection1.Connected := True;
end;

procedure TFormNarudzbenice.FormCreate(Sender: TObject);
begin
  PoveziBazu;
  UcitajNabavke;
end;

function TFormNarudzbenice.StatusBoja(Status: string): TAlphaColor;
begin
  if Status = 'Primljeno' then Result := $FF1D9E75
  else if Status = 'U obradi' then Result := $FF2980B9
  else if Status = 'Ceka potvrdu' then Result := $FFEF9F27
  else Result := $FFE05050;
end;

procedure TFormNarudzbenice.DodajRed(Broj, Dobavljac, Datum, Artikli, Vrednost, Status: string);
var
  Item: TListBoxItem;
  rectRow, rectStatus: TRectangle;
  lblBr, lblDob, lblDat, lblArt, lblVred, lblS: TLabel;
  lineSep: TLine;
begin
  Item := TListBoxItem.Create(ListBoxNabavke);
  Item.Height := 60;
  Item.Selectable := True;

  rectRow := TRectangle.Create(Item);
  rectRow.Parent := Item;
  rectRow.Align := TAlignLayout.Client;
  rectRow.Fill.Color := TAlphaColors.White;
  rectRow.Stroke.Kind := TBrushKind.None;

  // Broj naloga
  lblBr := TLabel.Create(rectRow);
  lblBr.Parent := rectRow;
  lblBr.Position.X := 10;
  lblBr.Position.Y := 8;
  lblBr.Width := 70;
  lblBr.Height := 16;
  lblBr.Text := Broj;
  lblBr.TextSettings.Font.Size := 10;
  lblBr.TextSettings.Font.Style := [TFontStyle.fsBold];
  lblBr.TextSettings.FontColor := $FF222222;

  // Dobavljac
  lblDob := TLabel.Create(rectRow);
  lblDob.Parent := rectRow;
  lblDob.Position.X := 86;
  lblDob.Position.Y := 8;
  lblDob.Width := 68;
  lblDob.Height := 32;
  lblDob.Text := Dobavljac;
  lblDob.WordWrap := True;
  lblDob.TextSettings.Font.Size := 10;
  lblDob.TextSettings.FontColor := $FF444444;

  // Datum
  lblDat := TLabel.Create(rectRow);
  lblDat.Parent := rectRow;
  lblDat.Position.X := 160;
  lblDat.Position.Y := 8;
  lblDat.Width := 58;
  lblDat.Height := 16;
  lblDat.Text := Datum;
  lblDat.TextSettings.Font.Size := 10;
  lblDat.TextSettings.FontColor := $FF666666;

  // Artikli
  lblArt := TLabel.Create(rectRow);
  lblArt.Parent := rectRow;
  lblArt.Position.X := 224;
  lblArt.Position.Y := 8;
  lblArt.Width := 44;
  lblArt.Height := 32;
  lblArt.Text := Artikli;
  lblArt.WordWrap := True;
  lblArt.TextSettings.Font.Size := 10;
  lblArt.TextSettings.FontColor := $FF666666;

  // Vrednost
  lblVred := TLabel.Create(rectRow);
  lblVred.Parent := rectRow;
  lblVred.Position.X := 272;
  lblVred.Position.Y := 8;
  lblVred.Width := 52;
  lblVred.Height := 32;
  lblVred.Text := Vrednost;
  lblVred.WordWrap := True;
  lblVred.TextSettings.Font.Size := 10;
  lblVred.TextSettings.Font.Style := [TFontStyle.fsBold];
  lblVred.TextSettings.FontColor := $FF222222;

  // Status badge
  rectStatus := TRectangle.Create(rectRow);
  rectStatus.Parent := rectRow;
  rectStatus.Position.X := 330;
  rectStatus.Position.Y := 10;
  rectStatus.Width := 50;
  rectStatus.Height := 22;
  rectStatus.XRadius := 6;
  rectStatus.YRadius := 6;
  rectStatus.Fill.Color := StatusBoja(Status) and $22FFFFFF or ($22000000);
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

  Item.Parent := ListBoxNabavke;
end;

procedure TFormNarudzbenice.UcitajNabavke;
var
  BrojArtikala: Integer;
begin
  ListBoxNabavke.Clear;

  ADOQuery1.Close;
  ADOQuery1.SQL.Text :=
    'SELECT nn.id_naloga, d.naziv_firme, nn.datum_naloga, ' +
    '       nn.ukupna_vrednost, nn.status, ' +
    '       (SELECT COUNT(*) FROM nalog_nabavka_stavke s WHERE s.id_naloga = nn.id_naloga) AS br_art ' +
    'FROM nalog_nabavka nn ' +
    'INNER JOIN dobavljaci d ON d.id_dobavljaca = nn.id_dobavljaca ' +
    'ORDER BY nn.datum_naloga DESC';
  ADOQuery1.Open;

  while not ADOQuery1.Eof do
  begin
    BrojArtikala := ADOQuery1.FieldByName('br_art').AsInteger;
    DodajRed(
      '#NOV-' + ADOQuery1.FieldByName('id_naloga').AsString,
      ADOQuery1.FieldByName('naziv_firme').AsString,
      FormatDateTime('dd.mm.yyyy', ADOQuery1.FieldByName('datum_naloga').AsDateTime),
      IntToStr(BrojArtikala) + ' stavki',
      FormatFloat('#,##0', ADOQuery1.FieldByName('ukupna_vrednost').AsFloat) + ' RSD',
      ADOQuery1.FieldByName('status').AsString
    );
    ADOQuery1.Next;
  end;
end;

procedure TFormNarudzbenice.SpeedButton1Click(Sender: TObject);
begin
  Form2.Show;
  Close;
end;

procedure TFormNarudzbenice.btnNoviNalogClick(Sender: TObject);
begin
  FormNovaNabavka.ShowModal;
  UcitajNabavke;
end;

end.
