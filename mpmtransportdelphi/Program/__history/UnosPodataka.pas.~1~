unit Unit2;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs,
  Unit1, FMX.Edit, FMX.Objects, FMX.StdCtrls, FMX.Controls.Presentation,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Error, FireDAC.UI.Intf,
  FireDAC.Phys.Intf, FireDAC.Stan.Def, FireDAC.Stan.Pool, FireDAC.Stan.Async,
  FireDAC.Phys, FireDAC.FMXUI.Wait, FireDAC.Stan.Param, FireDAC.DatS,
  FireDAC.DApt.Intf, FireDAC.DApt, Data.DB, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client, FireDAC.Phys.MySQL, FireDAC.Phys.MySQLDef, Data.Win.ADODB;

type
  TForm2 = class(TForm)
    SpeedButton1: TSpeedButton;
    Text1: TText;
    txtbKorisnickoime: TEdit;
    txtbSifra: TEdit;
    PrijavaBtn: TButton;
    Image1: TImage;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    DataSource1: TDataSource;
    procedure PrijavaBtnClick(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

{$R *.fmx}

uses Unit3, Unit4, Unit5,Unit6,Unit7;

procedure TForm2.PrijavaBtnClick(Sender: TObject);
    var
  role: string;
begin



               ADOQuery1.Close;
  ADOQuery1.SQL.Text :=
    'SELECT role FROM login ' +
    'WHERE Trim(korisnickoime)=:u AND Trim(sifra)=:p';

  ADOQuery1.Parameters.ParamByName('u').Value :=
    Trim(txtbKorisnickoime.Text);
  ADOQuery1.Parameters.ParamByName('p').Value :=
    Trim(txtbSifra.Text);

  ADOQuery1.Open;

  if not ADOQuery1.Eof then


begin
    role := ADOQuery1.FieldByName('role').AsString;

    if SameText(role, 'menadzer') then
      Form3.Show
    else if SameText(role, 'administrator') then
      Form4.Show
      else if SameText(role, 'dispecer') then
      Form5.Show
      else if SameText(role, 'klijent') then
      Form6.Show
      else if SameText(role, 'vozac') then
      Form7.Show
    else
      ShowMessage('Nepoznata uloga!');

    Hide;   // 👈 NE Close
  end
  else
    ShowMessage('Pogrešno korisničko ime ili šifra!');
end;

procedure TForm2.SpeedButton1Click(Sender: TObject);
begin
            Form1.Show;
            Close;
end;

end.
