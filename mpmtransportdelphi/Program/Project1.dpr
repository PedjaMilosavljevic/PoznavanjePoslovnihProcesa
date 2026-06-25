program Project1;

uses
  System.StartUpCopy,
  FMX.Forms,
  pocetna in 'pocetna.pas' {Form1},
  UnosPodataka in 'UnosPodataka.pas' {Form2},
  Menadzer in 'Menadzer.pas' {Form3},
  Admin in 'Admin.pas' {Form4},
  Unit5 in 'Unit5.pas' {Form5},
  Klijent in 'Klijent.pas' {Form6},
  Vozac in 'Vozac.pas' {Form7},
  Zalihe in 'Zalihe.pas' {FormZalihe},
  PregledZaliha in 'PregledZaliha.pas' {FormPregledZaliha},
  Narudzbenice in 'Narudzbenice.pas' {FormNarudzbenice},
  NaloziIzdavanje in 'NaloziIzdavanje.pas' {FormNaloziIzdavanje},
  NovaNabavka in 'NovaNabavka.pas' {FormNovaNabavka},
  IzdavanjeRobe in 'IzdavanjeRobe.pas' {FormIzdavanjeRobe},
  IzvestajiZalihe in 'IzvestajiZalihe.pas' {FormIzvestaji};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TForm2, Form2);
  Application.CreateForm(TForm3, Form3);
  Application.CreateForm(TForm4, Form4);
  Application.CreateForm(TForm5, Form5);
  Application.CreateForm(TForm6, Form6);
  Application.CreateForm(TForm7, Form7);
  Application.CreateForm(TFormZalihe, FormZalihe);
  Application.CreateForm(TFormPregledZaliha, FormPregledZaliha);
  Application.CreateForm(TFormNarudzbenice, FormNarudzbenice);
  Application.CreateForm(TFormNaloziIzdavanje, FormNaloziIzdavanje);
  Application.CreateForm(TFormNovaNabavka, FormNovaNabavka);
  Application.CreateForm(TFormIzdavanjeRobe, FormIzdavanjeRobe);
  Application.CreateForm(TFormIzvestaji, FormIzvestaji);
  Application.Run;
end.
