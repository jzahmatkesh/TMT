program TMT;

uses
  Forms,
  MainUn in 'MainUn.pas' {FmMain},
  TablesUn in 'TablesUn.pas' {DmTables: TDataModule},
  Config in 'Config.pas',
  LoginUs in 'LoginUs.pas' {FmLogin};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFmMain, FmMain);
  Application.CreateForm(TDmTables, DmTables);
  Application.CreateForm(TFmLogin, FmLogin);
  if FmLogin.ShowModal = 6 then
    Application.Run
  else
  begin
    Application.Terminate;
  end;
end.
