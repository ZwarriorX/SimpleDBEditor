program MsAccess_clone;

uses
  Vcl.Forms,
  MainUnit_u in 'MainUnit_u.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
