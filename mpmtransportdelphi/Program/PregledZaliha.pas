unit PregledZaliha;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs,
  FMX.Controls.Presentation, FMX.StdCtrls, FMX.Objects, FMX.Edit,
  FMX.ListBox, FMX.Layouts, FMX.ScrollBox,
  Data.DB, Data.Win.ADODB, UnosPodataka;

type
  TFormPregledZaliha = class(TForm)
    RectHeader: TRectangle;
    SpeedButton1: TSpeedButton;
    lblTitle: TLabel;
    Line1: TLine;
    RectFilter: TRectangle;
    txtPretraga: TEdit;
    ComboKategorija: TComboBox;
    RectTableHeader: TRectangle;
    lblHNaziv: TLabel;
    lblHKat: TLabel;
    lblHStanje: TLabel;
    lblHMin: TLabel;
    lblHLok: TLabel;
    lblHStatus: TLabel;
    ScrollBox1: TScrollBox;
    ListBoxZalihe: TListBox;
    RectBtnBar: TRectangle;
    btnNovaNabavka: TButton;
    btnIzdajRobu: TButton;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure btnNovaNabavkaClick(Sender: TObject);
    procedure btnIzdajRobuClick(Sender: TObject);
    procedure txtPretragaChange(Sender: TObject);
    procedure ComboKategorijaChange(Sender: TObject);
  private
    procedure PoveziBazu;
    procedure UcitajZalihe(Filter: string = ''; Kat: string = '');
    function StatusBoja(Status: string): TAlphaColor;
    function StatusZalihe(Kol, Min: Integer): string;
    procedure DodajRed(Naziv, Kat, Kolicina, Min, Lok, Status: string);
  end;

var
  FormPregledZaliha: TFormPregledZaliha;

implementation

{$R *.fmx}

uses Zalihe;



procedure TFormPregledZaliha.PoveziBazu;
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

procedure TFormPregledZaliha.FormCreate(Sender: TObject);
begin
  PoveziBazu;
  ComboKategorija.Items.Add('Sve kategorije');
  ComboKategorija.Items.Add('Rezervni delovi');
  ComboKategorija.Items.Add('Maziva i goriva');
  ComboKategorija.Items.Add('Gume i felne');
  ComboKategorija.Items.Add('Alati i oprema');
  ComboKategorija.ItemIndex := 0;
  UcitajZalihe;
end;

procedure TFormPregledZaliha.FormShow(Sender: TObject);
begin
  UcitajZalihe(txtPretraga.Text,
    ComboKategorija.Items[ComboKategorija.ItemIndex]);
end;

function TFormPregledZaliha.StatusZalihe(Kol, Min: Integer): string;
begin
  if Kol = 0 then         Result := 'Nema'
  else if Kol < Min then  Result := 'Kriticno'
  else if Kol < Min * 2 then Result := 'Niska zaliha'
  else                    Result := 'OK';
end;

function TFormPregledZaliha.StatusBoja(Status: string): TAlphaColor;
begin
  if      Status = 'OK'          then Result := $FF1D9E75
  else if Status = 'Niska zaliha' then Result := $FFEF9F27
  else                                 Result := $FFE05050;
end;

procedure TFormPregledZaliha.DodajRed(Naziv, Kat, Kolicina, Min, Lok, Status: string);
var
  Item: TListBoxItem;
  Layout: TLayout;
  lblN, lblK, lblKol, lblM, lblL: TLabel;
  rectStatus: TRectangle;
  lblS: TLabel;
  lineSep: TLine;
begin
  Item := TListBoxItem.Create(ListBoxZalihe);
  Item.Height := 52;
  Item.Parent := ListBoxZalihe;

  Layout := TLayout.Create(Item);
  Layout.Parent  := Item;
  Layout.Align   := TAlignLayout.Client;

  // Naziv
  lblN := TLabel.Create(Layout);
  lblN.Parent := Layout;
  lblN.Position.X := 10;
  lblN.Position.Y := 8;
  lblN.Width  := 100;
  lblN.Height := 36;
  lblN.Text := Naziv;
  lblN.WordWrap := True;
  lblN.TextSettings.Font.Size  := 10;
  lblN.TextSettings.FontColor  := $FF222222;

  // Kategorija
  lblK := TLabel.Create(Layout);
  lblK.Parent := Layout;
  lblK.Position.X := 118;
  lblK.Position.Y := 18;
  lblK.Width  := 68;
  lblK.Height := 16;
  lblK.Text := Kat;
  lblK.TextSettings.Font.Size := 10;
  lblK.TextSettings.FontColor := $FF555555;

  // Kolicina
  lblKol := TLabel.Create(Layout);
  lblKol.Parent := Layout;
  lblKol.Position.X := 192;
  lblKol.Position.Y := 18;
  lblKol.Width  := 48;
  lblKol.Height := 16;
  lblKol.Text := Kolicina;
  lblKol.TextSettings.Font.Size := 10;
  lblKol.TextSettings.FontColor := $FF222222;

  // Min
  lblM := TLabel.Create(Layout);
  lblM.Parent := Layout;
  lblM.Position.X := 246;
  lblM.Position.Y := 18;
  lblM.Width  := 24;
  lblM.Height := 16;
  lblM.Text := Min;
  lblM.TextSettings.Font.Size := 10;
  lblM.TextSettings.FontColor := $FF888888;

  // Lokacija
  lblL := TLabel.Create(Layout);
  lblL.Parent := Layout;
  lblL.Position.X := 276;
  lblL.Position.Y := 18;
  lblL.Width  := 52;
  lblL.Height := 16;
  lblL.Text := Lok;
  lblL.TextSettings.Font.Size := 10;
  lblL.TextSettings.FontColor := $FF555555;

  // Status badge
  rectStatus := TRectangle.Create(Layout);
  rectStatus.Parent   := Layout;
  rectStatus.Position.X := 330;
  rectStatus.Position.Y := 14;
  rectStatus.Width  := 50;
  rectStatus.Height := 22;
  rectStatus.XRadius := 6;
  rectStatus.YRadius := 6;
  rectStatus.Fill.Color   := StatusBoja(Status);
  rectStatus.Stroke.Kind  := TBrushKind.None;

  lblS := TLabel.Create(rectStatus);
  lblS.Parent := rectStatus;
  lblS.Align  := TAlignLayout.Client;
  lblS.Text   := Status;
  lblS.TextSettings.Font.Size  := 9;
  lblS.TextSettings.FontColor  := TAlphaColors.White;
  lblS.TextSettings.HorzAlign  := TTextAlign.Center;
  lblS.TextSettings.VertAlign  := TTextAlign.Center;

  // Separator
  lineSep := TLine.Create(Layout);
  lineSep.Parent := Layout;
  lineSep.Align  := TAlignLayout.Bottom;
  lineSep.Height := 1;
  lineSep.Stroke.Color := $FFEEEEEE;
end;

procedure TFormPregledZaliha.UcitajZalihe(Filter: string; Kat: string);
var
  SQL: string;
  Kol, Min: Integer;
begin
  ListBoxZalihe.Clear;
  SQL := 'SELECT naziv, kategorija, kolicina_na_stanju, min_kolicina, ' +
         'lokacija_u_magacinu FROM zalihe WHERE 1=1';
  if Trim(Filter) <> '' then
    SQL := SQL + ' AND naziv LIKE ''%' + Trim(Filter) + '%''';
  if (Kat <> '') and (Kat <> 'Sve kategorije') then
    SQL := SQL + ' AND kategorija = ''' + Kat + '''';
  SQL := SQL + ' ORDER BY naziv';

  ADOQuery1.Close;
  ADOQuery1.SQL.Text := SQL;
  ADOQuery1.Open;

  while not ADOQuery1.Eof do
  begin
    Kol := ADOQuery1.FieldByName('kolicina_na_stanju').AsInteger;
    Min := ADOQuery1.FieldByName('min_kolicina').AsInteger;
    DodajRed(
      ADOQuery1.FieldByName('naziv').AsString,
      ADOQuery1.FieldByName('kategorija').AsString,
      IntToStr(Kol) + ' kom.',
      IntToStr(Min),
      ADOQuery1.FieldByName('lokacija_u_magacinu').AsString,
      StatusZalihe(Kol, Min)
    );
    ADOQuery1.Next;
  end;
end;

procedure TFormPregledZaliha.SpeedButton1Click(Sender: TObject);
begin
  FormZalihe.Show;
  Hide;
end;

procedure TFormPregledZaliha.btnNovaNabavkaClick(Sender: TObject);
var
  F: TForm;
begin
  PrethodnaFormaZalihe := 'FormPregledZaliha';
  F := TForm(Application.FindComponent('FormNovaNabavka'));
  if Assigned(F) then
  begin
    F.Show;
    Hide;
  end;
end;

procedure TFormPregledZaliha.btnIzdajRobuClick(Sender: TObject);
var
  F: TForm;
begin
  PrethodnaFormaZalihe := 'FormPregledZaliha';
  F := TForm(Application.FindComponent('FormIzdavanjeRobe'));
  if Assigned(F) then
  begin
    F.Show;
    Hide;
  end;
end;

procedure TFormPregledZaliha.txtPretragaChange(Sender: TObject);
begin
  UcitajZalihe(txtPretraga.Text,
    ComboKategorija.Items[ComboKategorija.ItemIndex]);
end;

procedure TFormPregledZaliha.ComboKategorijaChange(Sender: TObject);
begin
  UcitajZalihe(txtPretraga.Text,
    ComboKategorija.Items[ComboKategorija.ItemIndex]);
end;

end.
