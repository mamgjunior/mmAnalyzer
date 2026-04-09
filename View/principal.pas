unit principal;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.Grids, Data.DB,
  Vcl.DBGrids, Datasnap.DBClient, Vcl.StdCtrls, Vcl.Buttons, ComObj, System.Generics.Collections,
  Vcl.ComCtrls;

type
  TfrmPrincipal = class(TForm)
    pnlPrincipal: TPanel;
    pnlTop: TPanel;
    pnlGrid: TPanel;
    dbgView: TDBGrid;
    edtDestino: TEdit;
    btnSelecionarDestino: TButton;
    opdArquivo: TOpenDialog;
    btnCriarExcel: TBitBtn;
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
    procedure CarregarExcelNoGrid(const ArquivoExcel: string);
    procedure ExportarClientDataSetParaExcel(const ArquivoDestino: string);
    procedure EstadoInicialDoPainelDeStatus;
    procedure SelecionarOrigem;
    procedure CriarExcel;
    procedure Processar;
    procedure SelecionarDestino;
  public
    { Public declarations }
  end;

var
  frmPrincipal: TfrmPrincipal;

implementation

uses
  Cliente, ClienteController, Util, Constantes, dmDAO;


{$R *.dfm}

procedure TfrmPrincipal.btnCriarExcelClick(Sender: TObject);
begin
  CriarExcel;
end;

procedure TfrmPrincipal.btnProcessarClick(Sender: TObject);
begin
  Processar;
end;

procedure TfrmPrincipal.btnSelecionarDestinoClick(Sender: TObject);
begin
  SelecionarDestino;
end;

procedure TfrmPrincipal.btnSelecionarOrigemClick(Sender: TObject);
begin
  SelecionarOrigem;
end;

procedure TfrmPrincipal.CarregarExcelNoGrid(const ArquivoExcel: string);
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

  UltCodigoCliente: string;
  UltNomeCliente: string;
  UltCategoria: string;
  UltTipoPagamento: string;

  lCliente: TCliente;
  lClienteController: TClienteController;
begin
  lClienteController := TClienteController.Create;
  lClienteController.PrepararClientDataSet;

  pnlProcessamento.Visible := True;
  pbProcessamento.Min := ZERO;
  pbProcessamento.Position := ZERO;
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
    pbProcessamento.Position := ZERO;

    ColCodigoCliente := ZERO;
    ColNomeCliente := ZERO;
    ColEmpresaCliente := ZERO;
    ColCategoria := ZERO;
    ColValor := ZERO;
    ColValorProf := ZERO;
    ColTipoPagamento := ZERO;

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

    if ColCodigoCliente = ZERO then
      raise Exception.Create('Coluna "Código cliente" não encontrada.');
    if ColNomeCliente = ZERO then
      raise Exception.Create('Coluna "Nome cliente" não encontrada.');
    if ColEmpresaCliente = ZERO then
      raise Exception.Create('Coluna "Empresa" não encontrada.');
    if ColCategoria = ZERO then
      raise Exception.Create('Coluna "Categoria" não encontrada.');
    if ColValor = ZERO then
      raise Exception.Create('Coluna "Valor" não encontrada.');
    if ColValorProf = ZERO then
      raise Exception.Create('Coluna "Valor prof." não encontrada.');
    if ColTipoPagamento = ZERO then
      raise Exception.Create('Coluna "Tipo de Pagamento" não encontrada.');

    UltCodigoCliente := VAZIO;
    UltNomeCliente := VAZIO;
    UltCategoria := VAZIO;
    UltTipoPagamento := VAZIO;

    dm.cdsExecel.DisableControls;
    try
      for Linha := 2 to UltimaLinha do
      begin
        lCliente := TCliente.Create;

        lCliente.codigo := StrToIntDef(Trim(VarToStr(Worksheet.Cells[Linha, ColCodigoCliente].Value)),0);
        lCliente.nome := Trim(VarToStr(Worksheet.Cells[Linha, ColNomeCliente].Value));
        lCliente.empresa := Trim(VarToStr(Worksheet.Cells[Linha, ColEmpresaCliente].Value));
        lCliente.categoria := Trim(VarToStr(Worksheet.Cells[Linha, ColCategoria].Value));
        lCliente.tipoPagamento := Trim(VarToStr(Worksheet.Cells[Linha, ColTipoPagamento].Value));

        // reaproveita último valor não vazio
        if lCliente.codigo = ZERO then
          lCliente.codigo := StrToInt(UltCodigoCliente)
        else
          UltCodigoCliente := IntToStr(lCliente.codigo);

        if lCliente.nome = VAZIO then
          lCliente.nome := UltNomeCliente
        else
          UltNomeCliente := lCliente.nome;

        if lCliente.empresa = VAZIO then
          lCliente.empresa := '-';

        if lCliente.categoria = VAZIO then
          lCliente.categoria := UltCategoria
        else
          UltCategoria := lCliente.categoria;

        if lCliente.tipoPagamento = VAZIO then
          lCliente.tipoPagamento := UltTipoPagamento
        else
          UltTipoPagamento := lCliente.tipoPagamento;

        if (lCliente.codigo = ZERO) and (lCliente.nome = VAZIO) then
          Continue;

        lCliente.valor := ValorParaFloat(Worksheet.Cells[Linha, ColValor].Value);
        lCliente.valorprof := ValorParaFloat(Worksheet.Cells[Linha, ColValorProf].Value);

        lClienteController.Salvar(lCliente);
        FreeAndNil(lCliente);

//        lblTeste.Caption := 'Processando ... ' + CodigoCliente + ' | ' +
//          NomeCliente + ' | ' + EmpresaCliente + ' | ' + TipoPagamento;

        pbProcessamento.Position := Linha - 1;
        lblStatus.Caption := 'Processando linha ' + IntToStr(Linha) + ' de ' + IntToStr(UltimaLinha);

        if (Linha mod 20 = 0) then
          Application.ProcessMessages;

      end;
    finally
      dm.cdsExecel.EnableControls;
    end;

    Workbook.Close(False);
    ExcelApp.Quit;

  finally
    Worksheet := Unassigned;
    Workbook := Unassigned;
    ExcelApp := Unassigned;

    EstadoInicialDoPainelDeStatus;
    FreeAndNil(lClienteController);
  end;
end;

procedure TfrmPrincipal.CriarExcel;
begin
  if Trim(FArquivoSelecionado) = VAZIO then
  begin
    ShowMessage(MSG_SELECIONE_ORIGEM);
    Exit;
  end;

  if Trim(FDestinoArquivo) = VAZIO then
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

procedure TfrmPrincipal.EstadoInicialDoPainelDeStatus;
begin
  pnlProcessamento.Visible := False;
  pbProcessamento.Min := ZERO;
  pbProcessamento.Position := ZERO;
  lblStatus.Caption := VAZIO;
end;

procedure TfrmPrincipal.ExportarClientDataSetParaExcel(const ArquivoDestino: string);
var
  ExcelApp, Workbook, Worksheet: OleVariant;
  LinhaExcel: Integer;
begin
  if not dm.cdsExecel.Active then
    raise Exception.Create(MSG_DADOS_VAZIOS);

  if dm.cdsExecel.IsEmpty then
    raise Exception.Create(MSG_DATASET_VAZIO);

  pnlProcessamento.Visible := True;
  pbProcessamento.Min := ZERO;
  pbProcessamento.Position := ZERO;
  pbProcessamento.Max := dm.cdsExecel.RecordCount;
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
    dm.cdsExecel.First;
    while not dm.cdsExecel.Eof do
    begin
      Worksheet.Cells[LinhaExcel, 1].Value := IntToStr(dm.cdsExecelCODIGO.AsInteger);
      Worksheet.Cells[LinhaExcel, 2].Value := dm.cdsExecelCLIENTE.AsString;
      Worksheet.Cells[LinhaExcel, 3].Value := dm.cdsExecelEMPRESA.AsString;
      Worksheet.Cells[LinhaExcel, 4].Value := dm.cdsExecelCATEGORIA.AsString;
      Worksheet.Cells[LinhaExcel, 5].Value := dm.cdsExecelTIPO_PAGAMENTO.AsString;
      Worksheet.Cells[LinhaExcel, 6].Value := dm.cdsExecelVALOR.AsCurrency;
      Worksheet.Cells[LinhaExcel, 7].Value := dm.cdsExecelVALOR_PROF.AsCurrency;

      pbProcessamento.Position := LinhaExcel - 1;
      Application.ProcessMessages;

      Inc(LinhaExcel);
      dm.cdsExecel.Next;
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

    EstadoInicialDoPainelDeStatus;
  end;
end;

procedure TfrmPrincipal.FormCreate(Sender: TObject);
begin
  EstadoInicialDoPainelDeStatus;
end;

procedure TfrmPrincipal.Processar;
begin
  if Trim(FArquivoSelecionado) = '' then
  begin
    ShowMessage('Selecione um arquivo primeiro.');
    Exit;
  end;

  CarregarExcelNoGrid(FArquivoSelecionado);
  ShowMessage('Concluído. ' + IntToStr(dm.cdsExecel.RecordCount) + ' registros agrupados.');
end;

procedure TfrmPrincipal.SelecionarDestino;
begin
  sdlArquivo.Title := MSG_ESCOLHA_DESTINO;
  sdlArquivo.Filter := 'Arquivo Excel|*.xlsx';
  sdlArquivo.DefaultExt := 'xlsx';
  sdlArquivo.FileName := 'Resumo_Agrupado.xlsx';

  if sdlArquivo.Execute then
  begin
    FDestinoArquivo := sdlArquivo.FileName;
    edtDestino.Text := FDestinoArquivo;
  end;
end;

procedure TfrmPrincipal.SelecionarOrigem;
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

end.
