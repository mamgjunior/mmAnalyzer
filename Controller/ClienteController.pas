unit ClienteController;

interface

uses
  Cliente;

type
  TClienteController = class
    private
      procedure Inserir(const pCliente: TCliente);
      procedure Editar(const pCliente: TCliente);
      function LocalizarRegistroAgrupado(const pCliente: TCliente): Boolean;
    public
      procedure Salvar(const pCliente: TCliente);
      procedure PrepararClientDataSet;
  end;

implementation

uses
  System.SysUtils, System.Variants, Constantes, dmDAO;

{ TClienteController }

procedure TClienteController.Editar(const pCliente: TCliente);
begin
  try
    dm.cdsExecel.Edit;
    dm.cdsExecelVALOR.AsCurrency :=
      dm.cdsExecelVALOR.AsCurrency + pCliente.valor;
    dm.cdsExecelVALOR_PROF.AsCurrency :=
      dm.cdsExecelVALOR_PROF.AsCurrency + pCliente.valorprof;
    dm.cdsExecel.Post;
  except
    raise Exception.Create(MSG_ERRO_EDITAR);
  end;
end;

procedure TClienteController.Inserir(const pCliente: TCliente);
begin
  try
    dm.cdsExecel.Append;
    dm.cdsExecelCODIGO.AsInteger := pCliente.codigo;
    dm.cdsExecelCLIENTE.AsString := pCliente.nome;
    dm.cdsExecelEMPRESA.AsString := pCliente.empresa;
    dm.cdsExecelCATEGORIA.AsString := pCliente.categoria;
    dm.cdsExecelTIPO_PAGAMENTO.AsString := pCliente.tipoPagamento;
    dm.cdsExecelVALOR.AsCurrency := pCliente.valor;
    dm.cdsExecelVALOR_PROF.AsCurrency := pCliente.valorprof;
    dm.cdsExecel.Post;
  except
     raise Exception.Create(MSG_ERRO_INSERIR);
  end;
end;

function TClienteController.LocalizarRegistroAgrupado(
  const pCliente: TCliente): Boolean;
begin
    Result := dm.cdsExecel.Locate(IDX_AGRUPAMENTO,
    VarArrayOf([pCliente.codigo, pCliente.nome, pCliente.categoria, pCliente.tipoPagamento]),
    []
  );
end;

procedure TClienteController.PrepararClientDataSet;
begin
  if dm.cdsExecel.Active then
    dm.cdsExecel.Close;

  dm.cdsExecel.CreateDataSet;
  dm.cdsExecel.Open;

  dm.cdsExecel.IndexFieldNames := IDX_AGRUPAMENTO;
end;

procedure TClienteController.Salvar(const pCliente: TCliente);
begin
  if LocalizarRegistroAgrupado(pCliente) then
    Editar(pCliente)
  else
    Inserir(pCliente);
end;

end.
