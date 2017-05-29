program Project1;

uses
  Vcl.Forms,
  Unit1 in 'Unit1.pas' {Form1},
  Unit2 in 'Unit2.pas' {ConfigForm},
  help in 'help.pas' {Form3};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.HelpFile := '';
  Application.Title := '';
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TConfigForm, ConfigForm);
  Application.CreateForm(TForm3, Form3);
  Application.Run;
end.
