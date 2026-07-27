unit Unit1;

interface

uses
  Winapi.Windows,
  System.SysUtils,
  System.Classes,
  Vcl.Graphics,
  Vcl.Controls,
  Vcl.Forms,
  Vcl.Dialogs,
  Vcl.StdCtrls,
  Vcl.ExtCtrls,
  Vcl.DBGrids,
  Vcl.DBCtrls,
  Data.DB,
  ADODB;

type
  TMainForm = class(TForm)
  public
    constructor Create(AOwner: TComponent); override;
  private
    TopPanel: TPanel;
    BtnOpenDB: TButton;
    BtnRefresh: TButton;
    BtnCloseDB: TButton;
    CboTables: TComboBox;
    LblFile: TLabel;
    Grid: TDBGrid;
    Nav: TDBNavigator;
    OpenDlg: TOpenDialog;

    Conn: TADOConnection;
    Table: TADOTable;
    DS: TDataSource;

    procedure BuildUI;
    procedure BuildData;
    procedure OpenDatabaseClick(Sender: TObject);
    procedure RefreshTablesClick(Sender: TObject);
    procedure CloseDatabaseClick(Sender: TObject);
    procedure TableChanged(Sender: TObject);
    procedure OpenSelectedTable;
    procedure FillTableList;
    function BuildConnectionString(const FileName: string): string;
    procedure UpdateCaption;
  end;

var
  MainForm: TMainForm;

implementation
{$R-}
constructor TMainForm.Create(AOwner: TComponent);
begin
  inherited CreateNew(AOwner);

  Name := 'MainForm';
  Caption := 'Simple Database Editor';
  Position := poScreenCenter;
  Width := 1100;
  Height := 700;

  BuildUI;
  BuildData;
end;

procedure TMainForm.BuildUI;
begin
  Color := clBtnFace;

  TopPanel := TPanel.Create(Self);
  TopPanel.Parent := Self;
  TopPanel.Align := alTop;
  TopPanel.Height := 56;
  TopPanel.BevelOuter := bvNone;
  TopPanel.Padding.Left := 8;
  TopPanel.Padding.Top := 8;
  TopPanel.Padding.Right := 8;
  TopPanel.Padding.Bottom := 8;

  BtnOpenDB := TButton.Create(Self);
  BtnOpenDB.Parent := TopPanel;
  BtnOpenDB.Left := 8;
  BtnOpenDB.Top := 12;
  BtnOpenDB.Width := 120;
  BtnOpenDB.Caption := 'Open Database';
  BtnOpenDB.OnClick := OpenDatabaseClick;

  BtnRefresh := TButton.Create(Self);
  BtnRefresh.Parent := TopPanel;
  BtnRefresh.Left := 136;
  BtnRefresh.Top := 12;
  BtnRefresh.Width := 90;
  BtnRefresh.Caption := 'Refresh';
  BtnRefresh.OnClick := RefreshTablesClick;

  BtnCloseDB := TButton.Create(Self);
  BtnCloseDB.Parent := TopPanel;
  BtnCloseDB.Left := 234;
  BtnCloseDB.Top := 12;
  BtnCloseDB.Width := 90;
  BtnCloseDB.Caption := 'Close DB';
  BtnCloseDB.OnClick := CloseDatabaseClick;

  CboTables := TComboBox.Create(Self);
  CboTables.Parent := TopPanel;
  CboTables.Left := 338;
  CboTables.Top := 13;
  CboTables.Width := 280;
  CboTables.Style := csDropDownList;
  CboTables.OnChange := TableChanged;

  LblFile := TLabel.Create(Self);
  LblFile.Parent := TopPanel;
  LblFile.Left := 636;
  LblFile.Top := 17;
  LblFile.Caption := 'No database opened';

  Grid := TDBGrid.Create(Self);
  Grid.Parent := Self;
  Grid.Align := alClient;
  Grid.Options := Grid.Options + [dgEditing, dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines];
  Grid.ReadOnly := False;

  Nav := TDBNavigator.Create(Self);
  Nav.Parent := Self;
  Nav.Align := alBottom;
  Nav.Height := 44;
  Nav.VisibleButtons := [nbFirst, nbPrior, nbNext, nbLast, nbInsert, nbDelete, nbEdit, nbPost, nbCancel, nbRefresh];

  OpenDlg := TOpenDialog.Create(Self);
  OpenDlg.Options := [ofFileMustExist, ofPathMustExist];
OpenDlg.Filter := 'Access Database (*.mdb)|*.mdb|All Files (*.*)|*.*';
end;

procedure TMainForm.BuildData;
begin
  Conn := TADOConnection.Create(Self);
  Conn.LoginPrompt := False;
  Conn.KeepConnection := False;

  Table := TADOTable.Create(Self);
  Table.Connection := Conn;
  Table.CursorType := ctStatic;
  Table.LockType := ltOptimistic;

  DS := TDataSource.Create(Self);
  DS.DataSet := Table;

  Grid.DataSource := DS;
  Nav.DataSource := DS;
end;

function TMainForm.BuildConnectionString(const FileName: string): string;
begin
  if SameText(ExtractFileExt(FileName), '.mdb') then
    Result :=
      'Provider=Microsoft.Jet.OLEDB.4.0;' +
      'Data Source=' + FileName + ';' +
      'Persist Security Info=False;'
  else
    Result :=
      'Provider=Microsoft.ACE.OLEDB.12.0;' +
      'Data Source=' + FileName + ';' +
      'Persist Security Info=False;';
end;

procedure TMainForm.UpdateCaption;
begin
  if Conn.Connected then
    Caption := 'Simple Database Editor - Connected'
  else
    Caption := 'Simple Database Editor';
end;

procedure TMainForm.OpenDatabaseClick(Sender: TObject);
begin
  if not OpenDlg.Execute then
    Exit;

  try
    if Conn.Connected then
      Conn.Close;

    Conn.ConnectionString := BuildConnectionString(OpenDlg.FileName);
    Conn.Open;

    LblFile.Caption := ExtractFileName(OpenDlg.FileName);
    FillTableList;

    if CboTables.Items.Count > 0 then
    begin
      CboTables.ItemIndex := 0;
      OpenSelectedTable;
    end;

    UpdateCaption;
  except
    on E: Exception do
    begin
      Conn.Close;
      MessageDlg('Could not open database:'#13#10 + E.Message, mtError, [mbOK], 0);
      UpdateCaption;
    end;
  end;
end;

procedure TMainForm.FillTableList;
begin
  CboTables.Items.BeginUpdate;
  try
    CboTables.Items.Clear;
    if Conn.Connected then
      Conn.GetTableNames(CboTables.Items, False);
  finally
    CboTables.Items.EndUpdate;
  end;
end;

procedure TMainForm.OpenSelectedTable;
var
  I: Integer;
  ColWidth: Integer;
begin
  if not Conn.Connected then
    Exit;

  if CboTables.ItemIndex < 0 then
    Exit;

  try
    Table.Close;
    Table.TableName := CboTables.Items[CboTables.ItemIndex];
    Table.Open;

    // Automatically size all columns to fit the grid
    if Grid.Columns.Count > 0 then
    begin
      ColWidth := (Grid.ClientWidth - 20) div Grid.Columns.Count;

      for I := 0 to Grid.Columns.Count - 1 do
        Grid.Columns[I].Width := ColWidth;
    end;

  except
    on E: Exception do
      MessageDlg('Could not open table:'#13#10 + E.Message, mtError, [mbOK], 0);
  end;
end;

procedure TMainForm.TableChanged(Sender: TObject);
begin
  OpenSelectedTable;
end;

procedure TMainForm.RefreshTablesClick(Sender: TObject);
begin
  if not Conn.Connected then
  begin
    MessageDlg('Open a database first.', mtInformation, [mbOK], 0);
    Exit;
  end;

  FillTableList;
  if CboTables.Items.Count > 0 then
  begin
    if CboTables.ItemIndex < 0 then
      CboTables.ItemIndex := 0;
    OpenSelectedTable;
  end;
end;

procedure TMainForm.CloseDatabaseClick(Sender: TObject);
begin
  try
    Table.Close;
    Conn.Close;
    CboTables.Items.Clear;
    CboTables.ItemIndex := -1;
    LblFile.Caption := 'No database opened';
    UpdateCaption;
  except
    on E: Exception do
      MessageDlg(E.Message, mtError, [mbOK], 0);
  end;
end;

end.
