object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Form1'
  ClientHeight = 534
  ClientWidth = 1079
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  OnCreate = FormCreate
  TextHeight = 15
  object pnlPrincipal: TPanel
    Left = 0
    Top = 0
    Width = 1079
    Height = 534
    Align = alClient
    TabOrder = 0
    ExplicitWidth = 1077
    ExplicitHeight = 526
    object pnlTop: TPanel
      Left = 1
      Top = 1
      Width = 1077
      Height = 168
      Align = alTop
      TabOrder = 0
      ExplicitWidth = 1075
      object edtDestino: TEdit
        Left = 16
        Top = 120
        Width = 569
        Height = 23
        Enabled = False
        TabOrder = 0
      end
      object btnSelecionarDestino: TButton
        Left = 608
        Top = 119
        Width = 75
        Height = 25
        Caption = 'Selecionar'
        TabOrder = 1
        OnClick = btnSelecionarDestinoClick
      end
      object btnCriarExcel: TBitBtn
        Left = 720
        Top = 119
        Width = 75
        Height = 25
        Caption = 'Criar Excel'
        TabOrder = 2
        OnClick = btnCriarExcelClick
      end
      object edtOrigem: TEdit
        Left = 16
        Top = 82
        Width = 569
        Height = 23
        Enabled = False
        TabOrder = 3
      end
      object btnSelecionarOrigem: TButton
        Left = 608
        Top = 81
        Width = 75
        Height = 25
        Caption = 'Selecionar'
        TabOrder = 4
        OnClick = btnSelecionarOrigemClick
      end
      object btnProcessar: TBitBtn
        Left = 720
        Top = 81
        Width = 75
        Height = 25
        Caption = 'Processar'
        TabOrder = 5
        OnClick = btnProcessarClick
      end
    end
    object pnlGrid: TPanel
      Left = 1
      Top = 169
      Width = 1077
      Height = 364
      Align = alClient
      TabOrder = 1
      ExplicitWidth = 1075
      ExplicitHeight = 356
      object dbgView: TDBGrid
        Left = 1
        Top = 1
        Width = 1075
        Height = 362
        Align = alClient
        DataSource = dsExcel
        TabOrder = 0
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -12
        TitleFont.Name = 'Segoe UI'
        TitleFont.Style = []
        Columns = <
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'CODIGO'
            Title.Alignment = taCenter
            Title.Caption = 'C'#243'digo'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'CLIENTE'
            Title.Alignment = taCenter
            Title.Caption = 'Cliente'
            Width = 250
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'EMPRESA'
            Title.Alignment = taCenter
            Title.Caption = 'Empresa'
            Width = 250
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'CATEGORIA'
            Title.Alignment = taCenter
            Title.Caption = 'Categoria'
            Width = 150
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'TIPO_PAGAMENTO'
            Title.Alignment = taCenter
            Title.Caption = 'Tipo de pagamento'
            Width = 150
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'VALOR'
            Title.Alignment = taCenter
            Title.Caption = 'Total'
            Width = 80
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'VALOR_PROF'
            Title.Alignment = taCenter
            Title.Caption = 'Profissional'
            Width = 80
            Visible = True
          end>
      end
      object pnlProcessamento: TPanel
        Left = 328
        Top = 144
        Width = 409
        Height = 64
        TabOrder = 1
        object lblStatus: TLabel
          Left = 9
          Top = 14
          Width = 45
          Height = 15
          Alignment = taCenter
          Caption = 'lblStatus'
        end
        object pbProcessamento: TProgressBar
          Left = 2
          Top = 37
          Width = 402
          Height = 21
          TabOrder = 0
        end
      end
    end
  end
  object cdsExecel: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 969
    Top = 257
    object cdsExecelCODIGO: TIntegerField
      FieldName = 'CODIGO'
    end
    object cdsExecelCLIENTE: TStringField
      FieldName = 'CLIENTE'
      Size = 150
    end
    object cdsExecelEMPRESA: TStringField
      FieldName = 'EMPRESA'
      Size = 150
    end
    object cdsExecelCATEGORIA: TStringField
      FieldName = 'CATEGORIA'
      Size = 100
    end
    object cdsExecelTIPO_PAGAMENTO: TStringField
      FieldName = 'TIPO_PAGAMENTO'
      Size = 100
    end
    object cdsExecelVALOR: TCurrencyField
      FieldName = 'VALOR'
    end
    object cdsExecelVALOR_PROF: TCurrencyField
      FieldName = 'VALOR_PROF'
    end
  end
  object dsExcel: TDataSource
    DataSet = cdsExecel
    Left = 864
    Top = 256
  end
  object opdArquivo: TOpenDialog
    Left = 961
    Top = 73
  end
  object sdlArquivo: TSaveDialog
    Left = 961
    Top = 22
  end
end
