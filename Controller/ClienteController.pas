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
  principal, System.SysUtils, System.Variants, Constantes;

{ TClienteController }

procedure TClienteController.Editar(const pCliente: TCliente);
begin
  try
    frmPrincipal.cdsExecel.Edit;
    frmPrincipal.cdsExecelVALOR.AsCurrency :=
      frmPrincipal.cdsExecelVALOR.AsCurrency + pCliente.valor;
    frmPrincipal.cdsExecelVALOR_PROF.AsCurrency :=
      frmPrincipal.cdsExecelVALOR_PROF.AsCurrency + pCliente.valorprof;
    frmPrincipal.cdsExecel.Post;
  except
    raise Exception.Create(MSG_ERRO_EDITAR);
  end;
end;

procedure TClienteController.Inserir(const pCliente: TCliente);
begin
  try
    frmPrincipal.cdsExecel.Append;
    frmPrincipal.cdsExecelCODIGO.AsInteger := pCliente.codigo;
    frmPrincipal.cdsExecelCLIENTE.AsString := pCliente.nome;
    frmPrincipal.cdsExecelEMPRESA.AsString := pCliente.empresa;
    frmPrincipal.cdsExecelCATEGORIA.AsString := pCliente.categoria;
    frmPrincipal.cdsExecelTIPO_PAGAMENTO.AsString := pCliente.tipoPagamento;
    frmPrincipal.cdsExecelVALOR.AsCurrency := pCliente.valor;
    frmPrincipal.cdsExecelVALOR_PROF.AsCurrency := pCliente.valorprof;
    frmPrincipal.cdsExecel.Post;
  except
     raise Exception.Create(MSG_ERRO_INSERIR);
  end;
end;

function TClienteController.LocalizarRegistroAgrupado(
  const pCliente: TCliente): Boolean;
begin
    Result := frmPrincipal.cdsExecel.Locate(IDX_AGRUPAMENTO,
    VarArrayOf([pCliente.codigo, pCliente.nome, pCliente.categoria, pCliente.tipoPagamento]),
    []
  );
end;

procedure TClienteController.PrepararClientDataSet;
begin
  if frmPrincipal.cdsExecel.Active then
    frmPrincipal.cdsExecel.Close;

  frmPrincipal.cdsExecel.CreateDataSet;
  frmPrincipal.cdsExecel.Open;

  frmPrincipal.cdsExecel.IndexFieldNames := IDX_AGRUPAMENTO;
end;

procedure TClienteController.Salvar(const pCliente: TCliente);
begin
  if LocalizarRegistroAgrupado(pCliente) then
    Editar(pCliente)
  else
    Inserir(pCliente);
end;

end.
