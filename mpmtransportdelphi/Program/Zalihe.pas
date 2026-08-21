unit Zalihe;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs,
  FMX.Controls.Presentation, FMX.StdCtrls, FMX.Objects, FMX.Layouts,
  FMX.ScrollBox, Data.DB, Data.Win.ADODB, UnosPodataka;

type
  TFormZalihe = class(TForm)
    RectHeader: TRectangle;
    SpeedButton1: TSpeedButton;
    lblTitle: TLabel;
    ScrollBox1: TScrollBox;
    RectMetrike: TRectangle;
    RectUkupno: TRectangle;
    lblUkupnoCaption: TLabel;
    lblUkupno: TLabel;
    lblUkupnoSub: TLabel;
    RectKriticno: TRectangle;
    lblKriticnoCaption: TLabel;
    lblKriticno: TLabel;
    lblKriticnoSub: TLabel;
    RectNabavke: TRectangle;
    lblNabavkeCaption: TLabel;
    lblNabavke: TLabel;
    lblNabavkeSub: TLabel;
    RectIzdavanja: TRectangle;
    lblIzdavanjaCaption: TLabel;
    lblIzdavanja: TLabel;
    lblIzdavanjaSub: TLabel;
    RectBrzaPristup: TRectangle;
    lblBrzPristupTitle: TLabel;
    btnArtikli: TButton;
    btnNabavke: TButton;
    btnIzdavanja: TButton;
    btnNovaNabavka: TButton;
    RectNavBar: TRectangle;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure btnArtikliClick(Sender: TObject);
    procedure btnNabavkeClick(Sender: TObject);
    procedure btnIzdavanjaClick(Sender: TObject);
    procedure btnNovaNabavkaClick(Sender: TObject);
    procedure btnIzvestajiClick(Sender: TObject);
  private
    procedure PoveziBazu;
    procedure UcitajMetrike;
  end;

var
  FormZalihe: TFormZalihe;
  PrethodnaFormaZalihe: string = 'FormZalihe';  // pamti odakle se doslo

implementation   uses Menadzer;

{$R *.fmx}

procedure TFormZalihe.PoveziBazu;
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

procedure TFormZalihe.FormCreate(Sender: TObject);
begin
  PoveziBazu;
  UcitajMetrike;
end;

procedure TFormZalihe.FormShow(Sender: TObject);
begin
  UcitajMetrike;
end;

procedure TFormZalihe.UcitajMetrike;
begin
  try
    // Ukupno artikala
    ADOQuery1.Close;
    ADOQuery1.SQL.Text := 'SELECT COUNT(*) AS br FROM zalihe';
    ADOQuery1.Open;
    lblUkupno.Text := ADOQuery1.FieldByName('br').AsString;

    // Kriticne zalihe
    ADOQuery1.Close;
    ADOQuery1.SQL.Text :=
      'SELECT COUNT(*) AS br FROM zalihe ' +
      'WHERE kolicina_na_stanju < min_kolicina';
    ADOQuery1.Open;
    lblKriticno.Text := ADOQuery1.FieldByName('br').AsString;

    // Narudzbenice u toku
    ADOQuery1.Close;
    ADOQuery1.SQL.Text :=
      'SELECT COUNT(*) AS br FROM nalog_nabavka ' +
      'WHERE status = ''U obradi''';
    ADOQuery1.Open;
    lblNabavke.Text := ADOQuery1.FieldByName('br').AsString;

    // Izdavanja danas
    ADOQuery1.Close;
    ADOQuery1.SQL.Text :=
      'SELECT COUNT(*) AS br FROM nalog_izdavanje ' +
      'WHERE datum_izdavanja >= #' +
        FormatDateTime('mm/dd/yyyy', Date) + '#';
    ADOQuery1.Open;
    lblIzdavanja.Text := ADOQuery1.FieldByName('br').AsString;

  except
    on E: Exception do
      ShowMessage('Greska pri citanju podataka: ' + E.Message);
  end;
end;

procedure TFormZalihe.SpeedButton1Click(Sender: TObject);
begin
  Form3.Show;
  Close;
end;

procedure TFormZalihe.btnArtikliClick(Sender: TObject);
var
  F: TForm;
begin
  PrethodnaFormaZalihe := 'FormZalihe';
  F := TForm(Application.FindComponent('FormPregledZaliha'));
  if Assigned(F) then F.Show;
  Hide;
end;

procedure TFormZalihe.btnNabavkeClick(Sender: TObject);
var
  F: TForm;
begin
  PrethodnaFormaZalihe := 'FormZalihe';
  F := TForm(Application.FindComponent('FormNarudzbenice'));
  if Assigned(F) then F.Show;
  Hide;
end;

procedure TFormZalihe.btnIzdavanjaClick(Sender: TObject);
var
  F: TForm;
begin
  PrethodnaFormaZalihe := 'FormZalihe';
  F := TForm(Application.FindComponent('FormNaloziIzdavanje'));
  if Assigned(F) then F.Show;
  Hide;
end;

procedure TFormZalihe.btnNovaNabavkaClick(Sender: TObject);
var
  F: TForm;
begin
  PrethodnaFormaZalihe := 'FormZalihe';
  F := TForm(Application.FindComponent('FormNovaNabavka'));
  if Assigned(F) then
  begin
    F.Show;
    Hide;
  end;
end;

procedure TFormZalihe.btnIzvestajiClick(Sender: TObject);
var
  F: TForm;
begin
  F := TForm(Application.FindComponent('FormIzvestaji'));
  if Assigned(F) then
  begin
    F.Show;
    Hide;
  end
  else
    ShowMessage('Forma za izvestaje nije ucitana u projekat.');
end;

end.
