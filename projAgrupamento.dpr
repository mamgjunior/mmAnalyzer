program projAgrupamento;

uses
  Vcl.Forms,
  principal in 'View\principal.pas' {frmPrincipal},
  Cliente in 'Model\Cliente.pas',
  ClienteController in 'Controller\ClienteController.pas',
  Util in 'Util.pas',
  Constantes in 'Constantes.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmPrincipal, frmPrincipal);
  Application.Run;
end.
