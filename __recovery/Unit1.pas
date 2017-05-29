unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Data.DB, Vcl.Grids,ComObj,
  Vcl.DBGrids, Data.Win.ADODB, uImportExcel, Datasnap.DBClient,HTMLHelpViewer,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Stan.Async, FireDAC.DApt, FireDAC.UI.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Phys, FireDAC.Phys.MSAcc, FireDAC.Phys.MSAccDef,
  FireDAC.VCLUI.Wait, FireDAC.Comp.Client, FireDAC.Comp.DataSet, Vcl.Menus,inifiles,
  Vcl.ExtCtrls,Unit2,help;

type
  TForm1 = class(TForm)
    Bimport: TButton;
    ADOConnExcel: TADOConnection;
    QueryExcel: TADOQuery;
    dsexcel: TDataSource;
    DBGrid1: TDBGrid;
    ADOConnAccess: TADOConnection;
    QueryAccess: TADOQuery;
    ComAccess: TADOCommand;
    DBGridImport: TDBGrid;
    dsaccess: TDataSource;
    Button3: TButton;
    OpenDialog1: TOpenDialog;
    ImportExcel1: TImportExcel;
    Button4: TButton;
    Grid: TStringGrid;
    DBGrid3: TDBGrid;
    Bexport: TButton;
    Button6: TButton;
    MainMenu1: TMainMenu;
    ApriExcel1: TMenuItem;
    GroupBox1: TGroupBox;
    EditX: TEdit;
    EditY: TEdit;
    Labelx: TLabel;
    LabelY: TLabel;
    ComAccessdelete: TADOCommand;
    Panel1: TPanel;
    Binsert: TButton;
    GroupBox2: TGroupBox;
    Nuovo1: TMenuItem;
    Info1: TMenuItem;
    Edit1: TEdit;
    procedure BinsertClick(Sender: TObject);
    procedure BimportClick(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure BexportClick(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure ApriExcel1Click(Sender: TObject);
    procedure EditXChange(Sender: TObject);
    procedure EditYChange(Sender: TObject);
    procedure Nuovo1Click(Sender: TObject);
    procedure Info1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
      ini:TIniFile;
      var importdir,paramx,paramy,config:string;
      procedure nuovo();
  public
    procedure Guida1Click(Sender: TObject);
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}


procedure TForm1.ApriExcel1Click(Sender: TObject);
begin
   if OpenDialog1.Execute then begin
     importdir := OpenDialog1.FileName;
     Panel1.Caption := importdir;
     EditX.Enabled  := True;
     EditY.Enabled  := True;
   end;
end;

procedure TForm1.BimportClick(Sender: TObject);
var i:integer;
begin
     try
         Bimport.Enabled := False;
         Binsert.Enabled := True;
         //AdoConnExcel.Connected := True;
         QueryExcel.ConnectionString:=Format('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=%s;Extended Properties="Excel 12.0 Xml;HDR=YES;IMEX=1"',[importdir]);
         QueryExcel.ParamCheck := False;
         paramx := EditX.Text;
         paramy := EditY.Text;
         QueryExcel.SQL.Clear;
         QueryExcel.Close;
         QueryExcel.SQL.Add('SELECT * FROM [LISTAPANNELLI$'+'A'+paramx+':'+'Y'+paramy+']');
         QueryExcel.Open;
         for i := 0 to DBGrid1.Columns.Count -1 do begin
         DBGrid1.Columns[i].Width := 100;
         end;
          showmessage('Excel Importato');
     except
           Showmessage('Importazione Fallita');
           BImport.Enabled := True;
     end;

end;

procedure TForm1.BinsertClick(Sender: TObject);
var i:integer;
begin
    try
          AdoConnAccess.Connected := False;
          AdoConnAccess.ConnectionString := 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source='+config+';Persist Security Info=False;';
          BInsert.Enabled := False;
          Bexport.Enabled := True;
          // Pulisco la tabella
          ComAccessdelete.CommandText := 'DELETE FROM mytable';
          // Inserisco i Dati nel Database
          ComAccess.CommandText := 'insert into mytable(Descrizione,Extra3) values(:descrizione,:note)';

          while not  QueryExcel.eof do
              begin

                ComAccess.Parameters.ParamValues['descrizione']      := QueryExcel.FieldByname('DESCRIZIONE').AsString;
                ComAccess.Parameters.ParamValues['note']             := QueryExcel.FieldByname('MOBILE').AsString;
                ComAccess.Execute;
                QueryExcel.Next;

              end;
         QueryAccess.Close;
         QueryAccess.Open;
         // Setting colonne DBGrid
         for i := 0 to DBGridimport.Columns.Count -1 do begin
         DBGridImport.Columns[i].Width := 100;
         end;
         showmessage('Excel inserito nel database');



    except
        on E : Exception do begin
        ShowMessage(E.ClassName+E.Message+'Errore di connessione al database');
         BInsert.Enabled := True;
        end;
    end;
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
 ImportExcel1.ExportarParaExcel(QueryAccess,'prova');
end;

procedure TForm1.Button4Click(Sender: TObject);
 var col,row:integer;

begin
    AdoConnAccess.Connected := False;
    AdoConnAccess.ConnectionString :='Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source='+config;
    QueryAccess.close;
    QueryAccess.Open;
    Grid.RowCount:=2;
    Grid.ColCount:=QueryAccess.FieldCount;
    Grid.FixedRows:=1;
    Grid.FixedCols:=0;

    { set column names }
    for col:=0 to Grid.ColCount-1 do
    Grid.Cells[col,0]:=QueryAccess.Fields[col].FieldName;

        { fill grid with query results }
        while not QueryAccess.Eof do begin
        for col:=0 to Grid.ColCount-1 do
        Grid.Cells[col,Grid.RowCount-1]:=QueryAccess.Fields[col].AsString;
        for row := 0 to QueryAccess.RowsAffected do begin

        Grid.RowCount:=Grid.RowCount+1;
        QueryAccess.Next;

        end;
        end;

Grid.RowCount:=Grid.RowCount-1;
QueryAccess.Close;

end;


procedure TForm1.BexportClick(Sender: TObject);
var linha, coluna : integer;
var planilha : variant;
var valorcampo : string;
begin
  planilha:= CreateoleObject('Excel.Application');
  planilha.WorkBooks.add(1);
  planilha.visible := False;

  AdoConnAccess.Connected := False;
  AdoConnAccess.ConnectionString :='Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source='+config;
  QueryAccess.Close;
  QueryAccess.Open;
  QueryAccess.First;

  try
     QueryAccess.DisableControls;
    for linha := 0 to  QueryAccess.RecordCount - 1 do begin
      Application.ProcessMessages;
      for coluna := 1 to  QueryAccess.FieldCount do begin
        valorcampo :=  QueryAccess.Fields[coluna - 1].AsString;
        planilha.cells[linha + 2,coluna] := valorCampo;
      end;
       QueryAccess.Next;
    end;
    for coluna := 1 to  QueryAccess.FieldCount do begin
      valorcampo :=  QueryAccess.Fields[coluna - 1].DisplayLabel;
      planilha.cells[1,coluna] := valorcampo;
    end;
    planilha.columns.Autofit;
  finally
     QueryAccess.EnableControls;
     ShowMessage('Archivio Creato');
  end;
  planilha.visible := True;
  // pulisco la tabella
  ComAccess.CommandText := 'DELETE FROM mytable';
  ComAccess.Execute;
  showmessage('Record eliminati');
  //ComAccess.Free;
  //ADOConnAccess.Connected := False;
    planilha := NULL;


end;

procedure TForm1.Button6Click(Sender: TObject);
begin
  ConfigForm.Show;
end;

procedure TForm1.EditXChange(Sender: TObject);
begin
  if (EditX.Text > '') and (EditY.Text > '') then begin
  Bimport.Enabled := true;
  paramx := EditX.Text;
  paramy := EditY.Text;
  end;

  if (EditX.Text > '') and (EditY.Text = '') then begin
  Bimport.Enabled := false;
  paramx := EditX.Text;
  paramy := EditY.Text;
  end;

  if (EditX.Text = '') and (EditY.Text > '') then begin
  Bimport.Enabled := false;
  paramx := EditX.Text;
  paramy := EditY.Text;
  end;
end;

procedure TForm1.EditYChange(Sender: TObject);
begin
  if (EditX.Text > '') and (EditY.Text > '') then begin
  Bimport.Enabled := true;
  paramx := EditX.Text;
  paramy := EditY.Text;
  end;

  if (EditX.Text > '') and (EditY.Text = '') then begin
  Bimport.Enabled := false;
  paramx := EditX.Text;
  paramy := EditY.Text;
  end;

  if (EditX.Text = '') and (EditY.Text > '') then begin
  Bimport.Enabled := false;
  paramx := EditX.Text;
  paramy := EditY.Text;
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
    ini := TIniFile.Create(getcurrentdir + '\settings.ini');
    config := ini.ReadString('config','dirdatabase',config);

end;

procedure TForm1.Guida1Click(Sender: TObject);
begin

end;

procedure TForm1.Info1Click(Sender: TObject);
begin
          ConfigForm.Show;
end;

procedure TForm1.nuovo;
begin
     try
        // Reset variabili
        importdir       := '';
        paramx          := '';
        paramy          := '';
        // Reset Componenti
        EditX.Text      := '';
        EditY.Text      := '';
        Binsert.Enabled := False;
        Bexport.Enabled := False;
        Panel1.Caption := '';
        // Reset Query e Datasource
       // AdoConnAccess.Connected := False;
       // AdoConnExcel.Connected := False;
        //AdoConnAccess.ConnectionString :='Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source='+config;
        //QueryExcel.Close;
        //QueryExcel.SQL.Clear;
        //QueryAccess.Close;
        //QueryAccess.SQL.Clear;
        //dsaccess.Free;
        //dsexcel.Free;
          QueryExcel.Close;
          QueryAccess.Close;





     except
         ShowMessage('Errore della connessione al database');
     end;
end;

procedure TForm1.Nuovo1Click(Sender: TObject);
begin
nuovo;
end;

end.
