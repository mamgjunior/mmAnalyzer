unit principal;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.Grids, Data.DB,
  Vcl.DBGrids, Datasnap.DBClient, Vcl.StdCtrls, Vcl.Buttons, ComObj, System.Generics.Collections,
  Vcl.ComCtrls;

type
  TForm1 = class(TForm)
    pnlPrincipal: TPanel;
    pnlTop: TPanel;
    pnlGrid: TPanel;
    cdsExecel: TClientDataSet;
    dsExcel: TDataSource;
    dbgView: TDBGrid;
    edtDestino: TEdit;
    btnSelecionarDestino: TButton;
    opdArquivo: TOpenDialog;
    btnCriarExcel: TBitBtn;
    cdsExecelCODIGO: TIntegerField;
    cdsExecelCLIENTE: TStringField;
    cdsExecelEMPRESA: TStringField;
    cdsExecelCATEGORIA: TStringField;
    cdsExecelTIPO_PAGAMENTO: TStringField;
    cdsExecelVALOR: TCurrencyField;
    cdsExecelVALOR_PROF: TCurrencyField;
    pnlProcessamento: TPanel;
    lblStatus: TLabel;
    pbProcessamento: TProgressBar;
    edtOrigem: TEdit;
    btnSelecionarOrigem: TButton;
    btnProcessar: TBitBtn;
    sdlArquivo: TSaveDialog;
    procedure FormCreate(Sender: TObject);
    procedure btnSelecionarOrigemClick(Sender: TObject);
    procedure btnProcessarClick(Sender: TObject);
    procedure btnSelecionarDestinoClick(Sender: TObject);
    procedure btnCriarExcelClick(Sender: TObject);
  private
    FArquivoSelecionado: string;
    FDestinoArquivo: string;
    procedure PrepararClientDataSet;
    procedure CarregarExcelNoGrid(const ArquivoExcel: string);
    procedure ExportarClientDataSetParaExcel(const ArquivoDestino: string);
    function ValorParaFloat(const V: Variant): Double;
    function LocalizarRegistroAgrupado(const ACodigo, ACliente, ACategoria, ATipoPagamento: string): Boolean;
  public
    { Public declarations }
  end;

  TResumoItem = record
    CodigoCliente: string;
    NomeCliente: string;
    Empresa: string;
    Categoria: string;
    TipoPagamento: string;
    TotalValor: Double;
    TotalValorProf: Double;
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.btnCriarExcelClick(Sender: TObject);
begin
  if Trim(FArquivoSelecionado) = '' then
  begin
    ShowMessage('Selecione o arquivo de origem.');
    Exit;
  end;

  if Trim(FDestinoArquivo) = '' then
  begin
    ShowMessage('Escolha o local onde o novo arquivo será salvo.');
    Exit;
  end;

  btnSelecionarOrigem.Enabled := False;
  btnSelecionarDestino.Enabled := False;
  btnCriarExcel.Enabled := False;

  try
    ExportarClientDataSetParaExcel(FDestinoArquivo);

    ShowMessage('Arquivo gerado com sucesso em:'#13#10 + FDestinoArquivo);
  finally
    btnSelecionarOrigem.Enabled := True;
    btnSelecionarDestino.Enabled := True;
    btnCriarExcel.Enabled := True;
  end;
end;

procedure TForm1.btnProcessarClick(Sender: TObject);
begin
    if Trim(FArquivoSelecionado) = '' then
  begin
    ShowMessage('Selecione um arquivo primeiro.');
    Exit;
  end;

  CarregarExcelNoGrid(FArquivoSelecionado);
  ShowMessage('Concluído. ' + IntToStr(cdsExecel.RecordCount) + ' registros agrupados.');
end;

procedure TForm1.btnSelecionarDestinoClick(Sender: TObject);
begin
  sdlArquivo.Title := 'Escolha onde salvar o arquivo agrupado';
  sdlArquivo.Filter := 'Arquivo Excel|*.xlsx';
  sdlArquivo.DefaultExt := 'xlsx';
  sdlArquivo.FileName := 'Resumo_Agrupado.xlsx';

  if sdlArquivo.Execute then
  begin
    FDestinoArquivo := sdlArquivo.FileName;
    edtDestino.Text := FDestinoArquivo;
  end;
end;

procedure TForm1.btnSelecionarOrigemClick(Sender: TObject);
begin
  opdArquivo.Title := 'Selecione a planilha';
  opdArquivo.Filter := 'Arquivos Excel|*.xlsx;*.xls';
  opdArquivo.DefaultExt := 'xlsx';

  if opdArquivo.Execute then
  begin
    FArquivoSelecionado := opdArquivo.FileName;
    edtOrigem.Text := FArquivoSelecionado;
  end;
end;

procedure TForm1.CarregarExcelNoGrid(const ArquivoExcel: string);
var
  ExcelApp, Workbook, Worksheet: OleVariant;
  UltimaLinha, UltimaColuna: Integer;
  Coluna, Linha: Integer;
  Cabecalho: string;

  ColCodigoCliente: Integer;
  ColNomeCliente: Integer;
  ColEmpresaCliente: Integer;
  ColCategoria: Integer;
  ColValor: Integer;
  ColValorProf: Integer;
  ColTipoPagamento: Integer;

  CodigoCliente: string;
  NomeCliente: string;
  EmpresaCliente: string;
  Categoria: string;
  TipoPagamento: string;
  Valor: Double;
  ValorProf: Double;

  UltCodigoCliente: string;
  UltNomeCliente: string;
  UltEmpresaCliente: string;
  UltCategoria: string;
  UltTipoPagamento: string;
begin
  PrepararClientDataSet;

  pnlProcessamento.Visible := True;
  pbProcessamento.Min := 0;
  pbProcessamento.Position := 0;
  lblStatus.Caption := 'Iniciando processamento...';
  Application.ProcessMessages;

  ExcelApp := CreateOleObject('Excel.Application');
  try
    ExcelApp.Visible := False;
    ExcelApp.DisplayAlerts := False;

    Workbook := ExcelApp.Workbooks.Open(ArquivoExcel);
    Worksheet := Workbook.WorkSheets[1];

    UltimaLinha := Worksheet.UsedRange.Rows.Count;
    UltimaColuna := Worksheet.UsedRange.Columns.Count;

    pbProcessamento.Max := UltimaLinha - 1;
    pbProcessamento.Position := 0;

    ColCodigoCliente := 0;
    ColNomeCliente := 0;
    ColEmpresaCliente := 0;
    ColCategoria := 0;
    ColValor := 0;
    ColValorProf := 0;
    ColTipoPagamento := 0;

    for Coluna := 1 to UltimaColuna do
    begin
      Cabecalho := Trim(LowerCase(VarToStr(Worksheet.Cells[1, Coluna].Value)));

      if (Cabecalho = 'cód. cliente') or (Cabecalho = 'cod. cliente') or (Cabecalho = 'código cliente') or (Cabecalho = 'codigo cliente') then
        ColCodigoCliente := Coluna
      else if Cabecalho = 'nome cliente' then
        ColNomeCliente := Coluna
      else if Cabecalho = 'empresa' then
        ColEmpresaCliente := Coluna
      else if Cabecalho = 'categoria' then
        ColCategoria := Coluna
      else if Cabecalho = 'valor' then
        ColValor := Coluna
      else if Pos('valor prof', Cabecalho) > 0 then
        ColValorProf := Coluna
      else if Pos('tipo de pagamento', Cabecalho) > 0 then
        ColTipoPagamento := Coluna;
    end;

    if ColCodigoCliente = 0 then
      raise Exception.Create('Coluna "Código cliente" não encontrada.');
    if ColNomeCliente = 0 then
      raise Exception.Create('Coluna "Nome cliente" não encontrada.');
    if ColEmpresaCliente = 0 then
      raise Exception.Create('Coluna "Empresa" não encontrada.');
    if ColCategoria = 0 then
      raise Exception.Create('Coluna "Categoria" não encontrada.');
    if ColValor = 0 then
      raise Exception.Create('Coluna "Valor" não encontrada.');
    if ColValorProf = 0 then
      raise Exception.Create('Coluna "Valor prof." não encontrada.');
    if ColTipoPagamento = 0 then
      raise Exception.Create('Coluna "Tipo de Pagamento" não encontrada.');

    UltCodigoCliente := '';
    UltNomeCliente := '';
    UltEmpresaCliente := '';
    UltCategoria := '';
    UltTipoPagamento := '';

    cdsExecel.DisableControls;
    try
      for Linha := 2 to UltimaLinha do
      begin
        CodigoCliente := Trim(VarToStr(Worksheet.Cells[Linha, ColCodigoCliente].Value));
        NomeCliente := Trim(VarToStr(Worksheet.Cells[Linha, ColNomeCliente].Value));
        EmpresaCliente := Trim(VarToStr(Worksheet.Cells[Linha, ColEmpresaCliente].Value));
        Categoria := Trim(VarToStr(Worksheet.Cells[Linha, ColCategoria].Value));
        TipoPagamento := Trim(VarToStr(Worksheet.Cells[Linha, ColTipoPagamento].Value));

        // reaproveita último valor não vazio
        if CodigoCliente = '' then
          CodigoCliente := UltCodigoCliente
        else
          UltCodigoCliente := CodigoCliente;

        if NomeCliente = '' then
          NomeCliente := UltNomeCliente
        else
          UltNomeCliente := NomeCliente;

        if EmpresaCliente = '' then
          EmpresaCliente := UltEmpresaCliente
        else
          UltEmpresaCliente := EmpresaCliente;

        if Categoria = '' then
          Categoria := UltCategoria
        else
          UltCategoria := Categoria;

        if TipoPagamento = '' then
          TipoPagamento := UltTipoPagamento
        else
          UltTipoPagamento := TipoPagamento;

        if (CodigoCliente = '') and (NomeCliente = '') then
          Continue;

        Valor := ValorParaFloat(Worksheet.Cells[Linha, ColValor].Value);
        ValorProf := ValorParaFloat(Worksheet.Cells[Linha, ColValorProf].Value);

        if LocalizarRegistroAgrupado(CodigoCliente, NomeCliente, Categoria, TipoPagamento) then
        begin
          cdsExecel.Edit;
          cdsExecelVALOR.AsCurrency := cdsExecelVALOR.AsCurrency + Valor;
          cdsExecelVALOR_PROF.AsCurrency := cdsExecelVALOR_PROF.AsCurrency + ValorProf;
          cdsExecel.Post;
        end
        else
        begin
          cdsExecel.Append;
          cdsExecelCODIGO.AsInteger := StrToIntDef(CodigoCliente, 0);
          cdsExecelCLIENTE.AsString := NomeCliente;
          cdsExecelEMPRESA.AsString := EmpresaCliente;
          cdsExecelCATEGORIA.AsString := Categoria;
          cdsExecelTIPO_PAGAMENTO.AsString := TipoPagamento;
          cdsExecelVALOR.AsCurrency := Valor;
          cdsExecelVALOR_PROF.AsCurrency := ValorProf;
          cdsExecel.Post;
        end;

//        lblTeste.Caption := 'Processando ... ' + CodigoCliente + ' | ' +
//          NomeCliente + ' | ' + EmpresaCliente + ' | ' + TipoPagamento;

        pbProcessamento.Position := Linha - 1;
        lblStatus.Caption := 'Processando linha ' + IntToStr(Linha) + ' de ' + IntToStr(UltimaLinha);

        if (Linha mod 20 = 0) then
          Application.ProcessMessages;

      end;
    finally
      cdsExecel.EnableControls;
    end;

    Workbook.Close(False);
    ExcelApp.Quit;

  finally
    Worksheet := Unassigned;
    Workbook := Unassigned;
    ExcelApp := Unassigned;

    pnlProcessamento.Visible := false;
    pbProcessamento.Min := 0;
    pbProcessamento.Position := 0;
    lblStatus.Caption := '';
  end;
end;

procedure TForm1.ExportarClientDataSetParaExcel(const ArquivoDestino: string);
var
  ExcelApp, Workbook, Worksheet: OleVariant;
  LinhaExcel: Integer;
begin
  if not cdsExecel.Active then
    raise Exception.Create('Não há dados para exportar.');

  if cdsExecel.IsEmpty then
    raise Exception.Create('O conjunto de dados está vazio.');

  pnlProcessamento.Visible := True;
  pbProcessamento.Min := 0;
  pbProcessamento.Position := 0;
  pbProcessamento.Max := cdsExecel.RecordCount;
  lblStatus.Caption := 'Gerando arquivo Excel...';
  Application.ProcessMessages;

  ExcelApp := CreateOleObject('Excel.Application');
  try
    ExcelApp.Visible := False;
    ExcelApp.DisplayAlerts := False;

    Workbook := ExcelApp.Workbooks.Add;
    Worksheet := Workbook.WorkSheets[1];
    Worksheet.Name := 'Resumo Agrupado';

    // Cabeçalhos
    Worksheet.Cells[1, 1].Value := 'Código';
    Worksheet.Cells[1, 2].Value := 'Cliente';
    Worksheet.Cells[1, 3].Value := 'Empresa';
    Worksheet.Cells[1, 4].Value := 'Categoria';
    Worksheet.Cells[1, 5].Value := 'Tipo de Pagamento';
    Worksheet.Cells[1, 6].Value := 'Valor';
    Worksheet.Cells[1, 7].Value := 'Valor Prof.';

    // Dados
    LinhaExcel := 2;
    cdsExecel.First;
    while not cdsExecel.Eof do
    begin
      Worksheet.Cells[LinhaExcel, 1].Value := IntToStr(cdsExecelCODIGO.AsInteger);
      Worksheet.Cells[LinhaExcel, 2].Value := cdsExecelCLIENTE.AsString;
      Worksheet.Cells[LinhaExcel, 3].Value := cdsExecelEMPRESA.AsString;
      Worksheet.Cells[LinhaExcel, 4].Value := cdsExecelCATEGORIA.AsString;
      Worksheet.Cells[LinhaExcel, 5].Value := cdsExecelTIPO_PAGAMENTO.AsString;
      Worksheet.Cells[LinhaExcel, 6].Value := cdsExecelVALOR.AsCurrency;
      Worksheet.Cells[LinhaExcel, 7].Value := cdsExecelVALOR_PROF.AsCurrency;

      pbProcessamento.Position := LinhaExcel - 1;
      Application.ProcessMessages;

      Inc(LinhaExcel);
      cdsExecel.Next;
    end;

    // Formatação
    Worksheet.Columns[1].AutoFit;
    Worksheet.Columns[2].AutoFit;
    Worksheet.Columns[3].AutoFit;
    Worksheet.Columns[4].AutoFit;
    Worksheet.Columns[5].AutoFit;
    Worksheet.Columns[6].NumberFormat := '#,##0.00';
    Worksheet.Columns[7].NumberFormat := '#,##0.00';
    Worksheet.Columns[6].AutoFit;
    Worksheet.Columns[7].AutoFit;

    // Cabeçalho em negrito
    Worksheet.Range['A1:G1'].Font.Bold := True;

    Workbook.SaveAs(ArquivoDestino);
    Workbook.Close(False);
    ExcelApp.Quit;

    lblStatus.Caption := 'Arquivo Excel criado com sucesso.';
    Application.ProcessMessages;

    ShowMessage('Arquivo Excel criado com sucesso.');
  finally
    Worksheet := Unassigned;
    Workbook := Unassigned;
    ExcelApp := Unassigned;

    pnlProcessamento.Visible := false;
    pbProcessamento.Min := 0;
    pbProcessamento.Position := 0;
    lblStatus.Caption := '';
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  pnlProcessamento.Visible := false;
  pbProcessamento.Min := 0;
  pbProcessamento.Position := 0;
  lblStatus.Caption := '';
end;

function TForm1.LocalizarRegistroAgrupado(const ACodigo, ACliente, ACategoria,
  ATipoPagamento: string): Boolean;
begin
  Result := cdsExecel.Locate(
    'CODIGO;CLIENTE;CATEGORIA;TIPO_PAGAMENTO',
    VarArrayOf([ACodigo, ACliente, ACategoria, ATipoPagamento]),
    []
  );
end;

procedure TForm1.PrepararClientDataSet;
begin
  if cdsExecel.Active then
    cdsExecel.Close;

  cdsExecel.CreateDataSet;
  cdsExecel.Open;

  cdsExecel.IndexFieldNames := 'CODIGO;CLIENTE;CATEGORIA;TIPO_PAGAMENTO';
end;

function TForm1.ValorParaFloat(const V: Variant): Double;
begin
  Result := 0;

  if VarIsNull(V) or VarIsEmpty(V) then
    Exit;

  try
    if VarIsNumeric(V) then
      Result := V
    else
      Result := StrToFloat(StringReplace(VarToStr(V), '.', ',', [rfReplaceAll]));
  except
    Result := 0;
  end;
end;

end.
