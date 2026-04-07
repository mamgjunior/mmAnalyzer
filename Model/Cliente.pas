unit Cliente;

interface

type
  TCliente = class
    private
      Fcodigo: integer;
      Fvalorprof: Currency;
      Fvalor: Currency;
      FtipoPagamento: string;
      Fcategoria: string;
      Fempresa: string;
    Fnome: string;
      procedure Setcodigo(const Value: integer);
      procedure Setcategoria(const Value: string);
      procedure Setempresa(const Value: string);
      procedure SettipoPagamento(const Value: string);
      procedure Setvalor(const Value: Currency);
      procedure Setvalorprof(const Value: Currency);
    procedure Setnome(const Value: string);
    public
      property codigo: integer read Fcodigo write Setcodigo;
      property nome: string read Fnome write Setnome;
      property empresa: string read Fempresa write Setempresa;
      property categoria: string read Fcategoria write Setcategoria;
      property tipoPagamento: string read FtipoPagamento write SettipoPagamento;
      property valor: Currency read Fvalor write Setvalor;
      property valorprof: Currency read Fvalorprof write Setvalorprof;
  end;

implementation

{ TCliente }

procedure TCliente.Setcategoria(const Value: string);
begin
  Fcategoria := Value;
end;

procedure TCliente.Setcodigo(const Value: integer);
begin
  Fcodigo := Value;
end;

procedure TCliente.Setempresa(const Value: string);
begin
  Fempresa := Value;
end;

procedure TCliente.Setnome(const Value: string);
begin
  Fnome := Value;
end;

procedure TCliente.SettipoPagamento(const Value: string);
begin
  FtipoPagamento := Value;
end;

procedure TCliente.Setvalor(const Value: Currency);
begin
  Fvalor := Value;
end;

procedure TCliente.Setvalorprof(const Value: Currency);
begin
  Fvalorprof := Value;
end;

end.
