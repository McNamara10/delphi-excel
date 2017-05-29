unit Unit2;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls,inifiles,
  Vcl.Imaging.pngimage, Vcl.ExtCtrls;

type
  TConfigForm = class(TForm)
    Edatabase: TEdit;
    Label1: TLabel;
    Bsave: TButton;
    Image1: TImage;
    procedure FormCreate(Sender: TObject);
    procedure BsaveClick(Sender: TObject);
  private
     ini:TIniFile;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ConfigForm: TConfigForm;

implementation

{$R *.dfm}

procedure TConfigForm.BsaveClick(Sender: TObject);
begin
    ini := TIniFile.Create(getcurrentdir + '\settings.ini');
    ini.WriteString('config','dirdatabase',Edatabase.Text);
    ini.Free;
    ConfigForm.Close;

end;

procedure TConfigForm.FormCreate(Sender: TObject);
begin
    ini := TIniFile.Create(getcurrentdir + '\settings.ini');
    Edatabase.Text := ini.ReadString('config','dirdatabase',Edatabase.Text);
    ini.Free;
end;

end.
