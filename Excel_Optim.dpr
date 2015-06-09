program Excel_Optim;

uses
  Forms,
  excel_optim_main in 'excel_optim_main.pas' {Form1},
  about in 'about.pas' {AboutBox};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'GA Excel';
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TAboutBox, AboutBox);
  Application.Run;
end.
