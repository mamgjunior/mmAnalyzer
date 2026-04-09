object dm: Tdm
  Height = 750
  Width = 1000
  PixelsPerInch = 120
  object dsExcel: TDataSource
    DataSet = cdsExecel
    Left = 69
    Top = 69
  end
  object cdsExecel: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 156
    Top = 68
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
end
