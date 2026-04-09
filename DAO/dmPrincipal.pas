unit dmDAO;

interface

uses
  System.SysUtils, System.Classes, Data.DB, Datasnap.DBClient;

type
  Tdm = class(TDataModule)
    dsExcel: TDataSource;
    cdsExecel: TClientDataSet;
    cdsExecelCODIGO: TIntegerField;
    cdsExecelCLIENTE: TStringField;
    cdsExecelEMPRESA: TStringField;
    cdsExecelCATEGORIA: TStringField;
    cdsExecelTIPO_PAGAMENTO: TStringField;
    cdsExecelVALOR: TCurrencyField;
    cdsExecelVALOR_PROF: TCurrencyField;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  dm: Tdm;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

end.
