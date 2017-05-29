object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'EffeEffe Excel'
  ClientHeight = 687
  ClientWidth = 1385
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object GroupBox2: TGroupBox
    Left = 679
    Top = 48
    Width = 714
    Height = 631
    Caption = 'Importazione File Excel per  Gestionale'
    TabOrder = 12
    object Edit1: TEdit
      Left = 8
      Top = 416
      Width = 121
      Height = 21
      TabOrder = 0
      Text = 'Edit1'
      Visible = False
    end
  end
  object GroupBox1: TGroupBox
    Left = 8
    Top = 47
    Width = 665
    Height = 632
    Caption = 'Excel Originale'
    TabOrder = 9
    object Labelx: TLabel
      Left = 159
      Top = 27
      Width = 100
      Height = 13
      Caption = 'Parametro colonna A'
    end
    object LabelY: TLabel
      Left = 408
      Top = 27
      Width = 99
      Height = 13
      Caption = 'Parametro colonna Y'
    end
    object EditX: TEdit
      Left = 159
      Top = 46
      Width = 243
      Height = 21
      Enabled = False
      TabOrder = 0
      TextHint = 'Numero della colonna di Excel DESCRIZIONE'
      OnChange = EditXChange
    end
    object EditY: TEdit
      Left = 408
      Top = 46
      Width = 249
      Height = 21
      Enabled = False
      TabOrder = 1
      TextHint = 
        'Inserire il numero della colonna finale della Colonna Descrizion' +
        'e'
      OnChange = EditYChange
    end
  end
  object DBGrid1: TDBGrid
    Left = 8
    Top = 120
    Width = 657
    Height = 559
    Align = alCustom
    DataSource = dsexcel
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
  end
  object Bimport: TButton
    Left = 8
    Top = 81
    Width = 121
    Height = 33
    Caption = 'IMPORT EXCEL'
    Enabled = False
    TabOrder = 0
    OnClick = BimportClick
  end
  object DBGridImport: TDBGrid
    Left = 814
    Top = 120
    Width = 563
    Height = 559
    Align = alCustom
    DataSource = dsaccess
    Enabled = False
    TabOrder = 2
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
    Columns = <
      item
        Expanded = False
        Visible = True
      end>
  end
  object Button3: TButton
    Left = 1207
    Top = 17
    Width = 113
    Height = 25
    Caption = 'Export Excel'
    TabOrder = 3
    OnClick = Button3Click
  end
  object Button4: TButton
    Left = 1183
    Top = 89
    Width = 146
    Height = 25
    Caption = 'Load Data StringGrid'
    TabOrder = 4
    Visible = False
    OnClick = Button4Click
  end
  object Grid: TStringGrid
    Left = 687
    Top = 499
    Width = 121
    Height = 57
    TabOrder = 5
    Visible = False
    ColWidths = (
      64
      64
      64
      64
      64)
    RowHeights = (
      24
      24
      24
      24
      24)
  end
  object DBGrid3: TDBGrid
    Left = 679
    Top = 584
    Width = 106
    Height = 41
    TabOrder = 6
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
    Visible = False
  end
  object Bexport: TButton
    Left = 1023
    Top = 89
    Width = 113
    Height = 25
    Caption = 'Crea Nuovo Excel'
    Enabled = False
    TabOrder = 7
    OnClick = BexportClick
  end
  object Button6: TButton
    Left = 830
    Top = 89
    Width = 146
    Height = 25
    Caption = 'Query GroupBy'
    TabOrder = 8
    Visible = False
    OnClick = Button6Click
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 1385
    Height = 41
    Align = alTop
    Caption = 'Effeffe Excel'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -20
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 10
  end
  object Binsert: TButton
    Left = 687
    Top = 336
    Width = 121
    Height = 57
    Caption = 'TRASFERISCI DATI'
    Enabled = False
    TabOrder = 11
    OnClick = BinsertClick
  end
  object ADOConnExcel: TADOConnection
    KeepConnection = False
    LoginPrompt = False
    Left = 128
    Top = 136
  end
  object QueryExcel: TADOQuery
    Connection = ADOConnExcel
    Parameters = <>
    Left = 208
    Top = 136
  end
  object dsexcel: TDataSource
    AutoEdit = False
    DataSet = QueryExcel
    Left = 280
    Top = 136
  end
  object ADOConnAccess: TADOConnection
    KeepConnection = False
    LoginPrompt = False
    Mode = cmReadWrite
    Provider = 'Microsoft.ACE.OLEDB.12.0'
    Left = 944
    Top = 312
  end
  object QueryAccess: TADOQuery
    Connection = ADOConnAccess
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      
        'SELECT FIRST(Descrizione) as Descrizione,FIRST(Extra3 )as Extra3' +
        ' FROM mytable   Group By Extra3 '
      '')
    Left = 1032
    Top = 312
  end
  object ComAccess: TADOCommand
    Connection = ADOConnAccess
    Parameters = <
      item
        Name = 'descrizione'
        DataType = ftString
        Size = -1
        Value = Null
      end
      item
        Name = 'note'
        DataType = ftString
        Size = -1
        Value = Null
      end>
    Left = 720
    Top = 240
  end
  object dsaccess: TDataSource
    DataSet = QueryAccess
    Left = 1104
    Top = 312
  end
  object OpenDialog1: TOpenDialog
    Left = 592
    Top = 16
  end
  object ImportExcel1: TImportExcel
    Left = 720
    Top = 120
  end
  object MainMenu1: TMainMenu
    Left = 536
    Top = 16
    object ApriExcel1: TMenuItem
      Caption = 'Apri Excel'
      OnClick = ApriExcel1Click
    end
    object Nuovo1: TMenuItem
      Caption = 'Nuova Importazione'
      OnClick = Nuovo1Click
    end
    object Info1: TMenuItem
      Caption = 'Configurazione'
      OnClick = Info1Click
    end
  end
  object ComAccessdelete: TADOCommand
    Parameters = <>
    Left = 719
    Top = 192
  end
end
