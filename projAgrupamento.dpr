program projAgrupamento;

uses
  Vcl.Forms,
  principal in 'View\principal.pas' {frmPrincipal},
  Cliente in 'Model\Cliente.pas',
  ClienteController in 'Controller\ClienteController.pas',
  Util in 'Util.pas',
  Constantes in 'Constantes.pas',
  dmDAO in 'DAO\dmDAO.pas' {dm: TDataModule};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmPrincipal, frmPrincipal);
  Application.CreateForm(Tdm, dm);
  Application.Run;
end.
