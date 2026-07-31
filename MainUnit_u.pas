unit MainUnit_u;

interface

uses
  Winapi.Windows,
  System.SysUtils,
  System.Classes,
  System.Variants,
  System.Win.ComObj,
  Vcl.Graphics,
  Vcl.Controls,
  Vcl.Forms,
  Vcl.Dialogs,
  Vcl.StdCtrls,
  Vcl.ExtCtrls,
  Vcl.ComCtrls,
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
    BtnNewDB: TButton;
    BtnOpenDB: TButton;
    BtnRefresh: TButton;
    BtnCloseDB: TButton;
    BtnConvert: TButton;
    BtnAddColumn: TButton;
    BtnDeleteColumn: TButton;
    BtnToggleView: TButton;
    CboTables: TComboBox;
    LblFile: TLabel;
    Grid: TDBGrid;
    DesignView: TListView;
    Nav: TDBNavigator;
    OpenDlg: TOpenDialog;

    Conn: TADOConnection;
    Table: TADOTable;
    DS: TDataSource;

    ShowingDesign: Boolean;

    procedure BuildUI;
    procedure BuildData;
    procedure NewDatabaseClick(Sender: TObject);
    procedure OpenDatabaseClick(Sender: TObject);
    procedure RefreshTablesClick(Sender: TObject);
    procedure CloseDatabaseClick(Sender: TObject);
    procedure ConvertDatabaseClick(Sender: TObject);
    procedure AddColumnClick(Sender: TObject);
    procedure DeleteColumnClick(Sender: TObject);
    procedure ToggleViewClick(Sender: TObject);
    procedure DesignViewDblClick(Sender: TObject);
    procedure TableChanged(Sender: TObject);
    procedure OpenSelectedTable;
    procedure FillTableList;
    procedure AutoSizeGridColumns;
    procedure PopulateDesignView;
    procedure SwitchToDatasheetView;
    procedure UpdateUIState;
    function OpenConnection(const ConnStr, DisplayName: string): Boolean;
    function BuildConnectionString(const FileName: string): string;
    procedure UpdateCaption;
  end;

var
  MainForm: TMainForm;

implementation

function QuoteIdent(const S: string): string;
begin
  Result := '[' + S + ']';
end;

// Display name for a field's type, in the style Access uses in its
// table designer (Short Text, Number, Date/Time, etc).
function AccessTypeName(F: TField): string;
begin
  case F.DataType of
    ftString, ftWideString:
      Result := 'Short Text';
    ftMemo, ftWideMemo, ftFmtMemo:
      Result := 'Long Text (Memo)';
    ftAutoInc:
      Result := 'AutoNumber';
    ftSmallint, ftInteger, ftWord, ftLargeint:
      Result := 'Number';
    ftFloat, ftBCD, ftFMTBcd, ftSingle:
      Result := 'Number';
    ftCurrency:
      Result := 'Currency';
    ftDateTime, ftDate, ftTime:
      Result := 'Date/Time';
    ftBoolean:
      Result := 'Yes/No';
    ftBlob, ftGraphic, ftOraBlob:
      Result := 'OLE Object';
    ftGuid:
      Result := 'Replication ID';
  else
    Result := 'Unknown';
  end;
end;

// Index into the editable type list used by the Add Column / Edit Field
// dialogs: 0=Short Text 1=Memo 2=Integer 3=Double 4=Currency 5=DateTime
// 6=Yes/No. Returns -1 for types that can't be reassigned through that
// list (AutoNumber, OLE Object, GUID, ...).
function TypeIndexForField(F: TField): Integer;
begin
  case F.DataType of
    ftString, ftWideString: Result := 0;
    ftMemo, ftWideMemo, ftFmtMemo: Result := 1;
    ftSmallint, ftInteger, ftWord, ftLargeint: Result := 2;
    ftFloat, ftBCD, ftFMTBcd, ftSingle: Result := 3;
    ftCurrency: Result := 4;
    ftDateTime, ftDate, ftTime: Result := 5;
    ftBoolean: Result := 6;
  else
    Result := -1;
  end;
end;

function TypeSQLForIndex(Index: Integer): string;
begin
  case Index of
    0: Result := 'TEXT(255)';
    1: Result := 'MEMO';
    2: Result := 'INTEGER';
    3: Result := 'DOUBLE';
    4: Result := 'CURRENCY';
    5: Result := 'DATETIME';
    6: Result := 'YESNO';
  else
    Result := 'TEXT(255)';
  end;
end;

// SQL type keyword used when reconstructing a table's structure during
// mdb <-> accdb conversion.
function FieldSQLType(F: TField): string;
var
  Sz: Integer;
begin
  case F.DataType of
    ftString, ftWideString:
      begin
        Sz := F.Size;
        if (Sz <= 0) or (Sz > 255) then
          Sz := 255;
        Result := Format('TEXT(%d)', [Sz]);
      end;
    ftMemo, ftWideMemo, ftFmtMemo: Result := 'MEMO';
    ftAutoInc: Result := 'COUNTER';
    ftSmallint, ftWord: Result := 'SMALLINT';
    ftInteger, ftLargeint: Result := 'INTEGER';
    ftFloat, ftBCD, ftFMTBcd, ftSingle: Result := 'DOUBLE';
    ftCurrency: Result := 'CURRENCY';
    ftDateTime, ftDate, ftTime: Result := 'DATETIME';
    ftBoolean: Result := 'YESNO';
    ftBlob, ftGraphic, ftOraBlob: Result := 'OLEOBJECT';
    ftGuid: Result := 'GUID';
  else
    Result := 'TEXT(255)';
  end;
end;

// Renaming a column isn't supported by Jet/ACE's ALTER TABLE syntax, but
// setting an ADOX Column's Name property does rename it in place.
procedure RenameColumnADOX(Conn: TADOConnection; const TableName, OldName, NewName: string);
var
  Catalog: OleVariant;
begin
  Catalog := CreateOleObject('ADOX.Catalog');
  Catalog.ActiveConnection := Conn.ConnectionObject;
  Catalog.Tables[TableName].Columns[OldName].Name := NewName;
  Catalog := Unassigned;
end;

function PromptNewColumn(out AName, AType: string): Boolean;
var
  Dlg: TForm;
  EdName: TEdit;
  CboType: TComboBox;
  LblName, LblType: TLabel;
  BtnOK, BtnCancel: TButton;
begin
  Result := False;
  Dlg := TForm.CreateNew(Application);
  try
    Dlg.Caption := 'Add Column';
    Dlg.Position := poScreenCenter;
    Dlg.BorderStyle := bsDialog;
    Dlg.BorderIcons := [biSystemMenu];
    Dlg.Width := 320;
    Dlg.Height := 190;

    LblName := TLabel.Create(Dlg);
    LblName.Parent := Dlg;
    LblName.Left := 16;
    LblName.Top := 16;
    LblName.Caption := 'Column name:';

    EdName := TEdit.Create(Dlg);
    EdName.Parent := Dlg;
    EdName.Left := 16;
    EdName.Top := 36;
    EdName.Width := 270;

    LblType := TLabel.Create(Dlg);
    LblType.Parent := Dlg;
    LblType.Left := 16;
    LblType.Top := 68;
    LblType.Caption := 'Data type:';

    CboType := TComboBox.Create(Dlg);
    CboType.Parent := Dlg;
    CboType.Left := 16;
    CboType.Top := 88;
    CboType.Width := 270;
    CboType.Style := csDropDownList;
    CboType.Items.Add('Short Text');
    CboType.Items.Add('Long Text (Memo)');
    CboType.Items.Add('Integer Number');
    CboType.Items.Add('Double Number');
    CboType.Items.Add('Currency');
    CboType.Items.Add('Date/Time');
    CboType.Items.Add('Yes/No');
    CboType.ItemIndex := 0;

    BtnOK := TButton.Create(Dlg);
    BtnOK.Parent := Dlg;
    BtnOK.Caption := 'OK';
    BtnOK.ModalResult := mrOk;
    BtnOK.Default := True;
    BtnOK.Left := 128;
    BtnOK.Top := 130;
    BtnOK.Width := 75;

    BtnCancel := TButton.Create(Dlg);
    BtnCancel.Parent := Dlg;
    BtnCancel.Caption := 'Cancel';
    BtnCancel.ModalResult := mrCancel;
    BtnCancel.Cancel := True;
    BtnCancel.Left := 211;
    BtnCancel.Top := 130;
    BtnCancel.Width := 75;

    if Dlg.ShowModal = mrOk then
    begin
      AName := Trim(EdName.Text);
      if AName = '' then
      begin
        MessageDlg('Please enter a column name.', mtWarning, [mbOK], 0);
        Exit(False);
      end;

      AType := TypeSQLForIndex(CboType.ItemIndex);
      Result := True;
    end;
  finally
    Dlg.Free;
  end;
end;

function PromptSelectColumn(Items: TStrings; const ACaption, APrompt: string;
  out ASelected: string): Boolean;
var
  Dlg: TForm;
  Lbl: TLabel;
  Cbo: TComboBox;
  BtnOK, BtnCancel: TButton;
begin
  Result := False;
  Dlg := TForm.CreateNew(Application);
  try
    Dlg.Caption := ACaption;
    Dlg.Position := poScreenCenter;
    Dlg.BorderStyle := bsDialog;
    Dlg.BorderIcons := [biSystemMenu];
    Dlg.Width := 320;
    Dlg.Height := 150;

    Lbl := TLabel.Create(Dlg);
    Lbl.Parent := Dlg;
    Lbl.Left := 16;
    Lbl.Top := 16;
    Lbl.Caption := APrompt;

    Cbo := TComboBox.Create(Dlg);
    Cbo.Parent := Dlg;
    Cbo.Left := 16;
    Cbo.Top := 36;
    Cbo.Width := 270;
    Cbo.Style := csDropDownList;
    Cbo.Items.Assign(Items);
    if Cbo.Items.Count > 0 then
      Cbo.ItemIndex := 0;

    BtnOK := TButton.Create(Dlg);
    BtnOK.Parent := Dlg;
    BtnOK.Caption := 'OK';
    BtnOK.ModalResult := mrOk;
    BtnOK.Default := True;
    BtnOK.Left := 128;
    BtnOK.Top := 90;
    BtnOK.Width := 75;

    BtnCancel := TButton.Create(Dlg);
    BtnCancel.Parent := Dlg;
    BtnCancel.Caption := 'Cancel';
    BtnCancel.ModalResult := mrCancel;
    BtnCancel.Cancel := True;
    BtnCancel.Left := 211;
    BtnCancel.Top := 90;
    BtnCancel.Width := 75;

    if Cbo.Items.Count = 0 then
    begin
      MessageDlg('There are no columns to choose from.', mtInformation, [mbOK], 0);
      Exit(False);
    end;

    if Dlg.ShowModal = mrOk then
    begin
      if Cbo.ItemIndex < 0 then
        Exit(False);
      ASelected := Cbo.Items[Cbo.ItemIndex];
      Result := True;
    end;
  finally
    Dlg.Free;
  end;
end;

function PromptChooseFormat(out AIsAccdb: Boolean): Boolean;
var
  Dlg: TForm;
  RbMdb, RbAccdb: TRadioButton;
  BtnOK, BtnCancel: TButton;
begin
  Result := False;
  Dlg := TForm.CreateNew(Application);
  try
    Dlg.Caption := 'New Database';
    Dlg.Position := poScreenCenter;
    Dlg.BorderStyle := bsDialog;
    Dlg.BorderIcons := [biSystemMenu];
    Dlg.Width := 340;
    Dlg.Height := 170;

    RbMdb := TRadioButton.Create(Dlg);
    RbMdb.Parent := Dlg;
    RbMdb.Left := 16;
    RbMdb.Top := 16;
    RbMdb.Width := 290;
    RbMdb.Caption := 'Access 97-2003 Database (*.mdb)';
    RbMdb.Checked := True;

    RbAccdb := TRadioButton.Create(Dlg);
    RbAccdb.Parent := Dlg;
    RbAccdb.Left := 16;
    RbAccdb.Top := 44;
    RbAccdb.Width := 290;
    RbAccdb.Caption := 'Access 2007-2016 Database (*.accdb)';

    BtnOK := TButton.Create(Dlg);
    BtnOK.Parent := Dlg;
    BtnOK.Caption := 'OK';
    BtnOK.ModalResult := mrOk;
    BtnOK.Default := True;
    BtnOK.Left := 148;
    BtnOK.Top := 96;
    BtnOK.Width := 75;

    BtnCancel := TButton.Create(Dlg);
    BtnCancel.Parent := Dlg;
    BtnCancel.Caption := 'Cancel';
    BtnCancel.ModalResult := mrCancel;
    BtnCancel.Cancel := True;
    BtnCancel.Left := 231;
    BtnCancel.Top := 96;
    BtnCancel.Width := 75;

    if Dlg.ShowModal = mrOk then
    begin
      AIsAccdb := RbAccdb.Checked;
      Result := True;
    end;
  finally
    Dlg.Free;
  end;
end;

// Edit dialog for Design View: lets the user rename a field and, if the
// current type is one we know how to reassign, change its data type too.
function PromptEditColumn(const CurrentName, CurrentTypeLabel: string;
  CurrentTypeIndex: Integer; AllowTypeChange: Boolean;
  out NewName: string; out NewTypeIndex: Integer): Boolean;
var
  Dlg: TForm;
  EdName: TEdit;
  CboType: TComboBox;
  LblName, LblType, LblFixed: TLabel;
  BtnOK, BtnCancel: TButton;
begin
  Result := False;
  NewTypeIndex := -1;
  CboType := nil;
  Dlg := TForm.CreateNew(Application);
  try
    Dlg.Caption := 'Edit Field';
    Dlg.Position := poScreenCenter;
    Dlg.BorderStyle := bsDialog;
    Dlg.BorderIcons := [biSystemMenu];
    Dlg.Width := 320;
    Dlg.Height := 190;

    LblName := TLabel.Create(Dlg);
    LblName.Parent := Dlg;
    LblName.Left := 16;
    LblName.Top := 16;
    LblName.Caption := 'Field name:';

    EdName := TEdit.Create(Dlg);
    EdName.Parent := Dlg;
    EdName.Left := 16;
    EdName.Top := 36;
    EdName.Width := 270;
    EdName.Text := CurrentName;

    LblType := TLabel.Create(Dlg);
    LblType.Parent := Dlg;
    LblType.Left := 16;
    LblType.Top := 68;
    LblType.Caption := 'Data type:';

    if AllowTypeChange then
    begin
      CboType := TComboBox.Create(Dlg);
      CboType.Parent := Dlg;
      CboType.Left := 16;
      CboType.Top := 88;
      CboType.Width := 270;
      CboType.Style := csDropDownList;
      CboType.Items.Add('Short Text');
      CboType.Items.Add('Long Text (Memo)');
      CboType.Items.Add('Integer Number');
      CboType.Items.Add('Double Number');
      CboType.Items.Add('Currency');
      CboType.Items.Add('Date/Time');
      CboType.Items.Add('Yes/No');
      if (CurrentTypeIndex >= 0) and (CurrentTypeIndex <= 6) then
        CboType.ItemIndex := CurrentTypeIndex
      else
        CboType.ItemIndex := 0;
    end
    else
    begin
      LblFixed := TLabel.Create(Dlg);
      LblFixed.Parent := Dlg;
      LblFixed.Left := 16;
      LblFixed.Top := 90;
      LblFixed.Caption := CurrentTypeLabel + ' (type cannot be changed)';
      LblFixed.Font.Style := [fsItalic];
    end;

    BtnOK := TButton.Create(Dlg);
    BtnOK.Parent := Dlg;
    BtnOK.Caption := 'OK';
    BtnOK.ModalResult := mrOk;
    BtnOK.Default := True;
    BtnOK.Left := 128;
    BtnOK.Top := 130;
    BtnOK.Width := 75;

    BtnCancel := TButton.Create(Dlg);
    BtnCancel.Parent := Dlg;
    BtnCancel.Caption := 'Cancel';
    BtnCancel.ModalResult := mrCancel;
    BtnCancel.Cancel := True;
    BtnCancel.Left := 211;
    BtnCancel.Top := 130;
    BtnCancel.Width := 75;

    if Dlg.ShowModal = mrOk then
    begin
      NewName := Trim(EdName.Text);
      if NewName = '' then
      begin
        MessageDlg('Please enter a field name.', mtWarning, [mbOK], 0);
        Exit(False);
      end;

      if AllowTypeChange and Assigned(CboType) then
        NewTypeIndex := CboType.ItemIndex
      else
        NewTypeIndex := -1;

      Result := True;
    end;
  finally
    Dlg.Free;
  end;
end;

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
  UpdateUIState;
end;

procedure TMainForm.BuildUI;
begin
  Color := clBtnFace;

  TopPanel := TPanel.Create(Self);
  TopPanel.Parent := Self;
  TopPanel.Align := alTop;
  TopPanel.Height := 80;
  TopPanel.BevelOuter := bvNone;
  TopPanel.Padding.Left := 8;
  TopPanel.Padding.Top := 8;
  TopPanel.Padding.Right := 8;
  TopPanel.Padding.Bottom := 8;

  // -- Row 1: database-level actions --
  BtnNewDB := TButton.Create(Self);
  BtnNewDB.Parent := TopPanel;
  BtnNewDB.Left := 8;
  BtnNewDB.Top := 8;
  BtnNewDB.Width := 110;
  BtnNewDB.Caption := 'New Database';
  BtnNewDB.OnClick := NewDatabaseClick;

  BtnOpenDB := TButton.Create(Self);
  BtnOpenDB.Parent := TopPanel;
  BtnOpenDB.Left := 126;
  BtnOpenDB.Top := 8;
  BtnOpenDB.Width := 120;
  BtnOpenDB.Caption := 'Open Database';
  BtnOpenDB.OnClick := OpenDatabaseClick;

  BtnRefresh := TButton.Create(Self);
  BtnRefresh.Parent := TopPanel;
  BtnRefresh.Left := 254;
  BtnRefresh.Top := 8;
  BtnRefresh.Width := 90;
  BtnRefresh.Caption := 'Refresh';
  BtnRefresh.OnClick := RefreshTablesClick;

  BtnCloseDB := TButton.Create(Self);
  BtnCloseDB.Parent := TopPanel;
  BtnCloseDB.Left := 352;
  BtnCloseDB.Top := 8;
  BtnCloseDB.Width := 90;
  BtnCloseDB.Caption := 'Close DB';
  BtnCloseDB.OnClick := CloseDatabaseClick;

  BtnConvert := TButton.Create(Self);
  BtnConvert.Parent := TopPanel;
  BtnConvert.Left := 450;
  BtnConvert.Top := 8;
  BtnConvert.Width := 150;
  BtnConvert.Caption := 'Convert to mdb/accdb';
  BtnConvert.OnClick := ConvertDatabaseClick;

  // -- Row 2: table / column / view actions --
  BtnAddColumn := TButton.Create(Self);
  BtnAddColumn.Parent := TopPanel;
  BtnAddColumn.Left := 8;
  BtnAddColumn.Top := 40;
  BtnAddColumn.Width := 110;
  BtnAddColumn.Caption := 'Add Column';
  BtnAddColumn.OnClick := AddColumnClick;

  BtnDeleteColumn := TButton.Create(Self);
  BtnDeleteColumn.Parent := TopPanel;
  BtnDeleteColumn.Left := 126;
  BtnDeleteColumn.Top := 40;
  BtnDeleteColumn.Width := 110;
  BtnDeleteColumn.Caption := 'Delete Column';
  BtnDeleteColumn.OnClick := DeleteColumnClick;

  BtnToggleView := TButton.Create(Self);
  BtnToggleView.Parent := TopPanel;
  BtnToggleView.Left := 244;
  BtnToggleView.Top := 40;
  BtnToggleView.Width := 110;
  BtnToggleView.Caption := 'Design View';
  BtnToggleView.OnClick := ToggleViewClick;

  CboTables := TComboBox.Create(Self);
  CboTables.Parent := TopPanel;
  CboTables.Left := 362;
  CboTables.Top := 41;
  CboTables.Width := 220;
  CboTables.Style := csDropDownList;
  CboTables.OnChange := TableChanged;

  LblFile := TLabel.Create(Self);
  LblFile.Parent := TopPanel;
  LblFile.Left := 590;
  LblFile.Top := 45;
  LblFile.Caption := 'No database opened';

  // -- Datasheet (normal grid) view --
  Grid := TDBGrid.Create(Self);
  Grid.Parent := Self;
  Grid.Align := alClient;
  Grid.Options := Grid.Options + [dgEditing, dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines];
  Grid.ReadOnly := False;

  // -- Design view (editable field name / data type / size, like MS Access) --
  DesignView := TListView.Create(Self);
  DesignView.Parent := Self;
  DesignView.Align := alClient;
  DesignView.ViewStyle := vsReport;
  DesignView.RowSelect := True;
  DesignView.GridLines := True;
  DesignView.ReadOnly := True; // native inline editing disabled; we use our own dialog
  DesignView.Visible := False;
  DesignView.ShowHint := True;
  DesignView.Hint := 'Double-click a field to rename it or change its data type.';
  DesignView.OnDblClick := DesignViewDblClick;
  DesignView.Columns.Add.Caption := 'Field Name';
  DesignView.Columns[0].Width := 240;
  DesignView.Columns.Add.Caption := 'Data Type';
  DesignView.Columns[1].Width := 160;
  DesignView.Columns.Add.Caption := 'Size';
  DesignView.Columns[2].Width := 80;

  Nav := TDBNavigator.Create(Self);
  Nav.Parent := Self;
  Nav.Align := alBottom;
  Nav.Height := 44;
  Nav.VisibleButtons := [nbFirst, nbPrior, nbNext, nbLast, nbInsert, nbDelete, nbEdit, nbPost, nbCancel, nbRefresh];

  OpenDlg := TOpenDialog.Create(Self);
  OpenDlg.Options := [ofFileMustExist, ofPathMustExist];
  OpenDlg.Filter := 'Access Database (*.mdb;*.accdb)|*.mdb;*.accdb|All Files (*.*)|*.*';
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

  ShowingDesign := False;
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

procedure TMainForm.UpdateUIState;
var
  Connected, HasTable: Boolean;
begin
  Connected := Conn.Connected;
  HasTable := Connected and Table.Active;

  BtnRefresh.Enabled := Connected;
  BtnCloseDB.Enabled := Connected;
  BtnConvert.Enabled := Connected;
  CboTables.Enabled := Connected;
  BtnAddColumn.Enabled := HasTable;
  BtnDeleteColumn.Enabled := HasTable;
  BtnToggleView.Enabled := HasTable;
end;

// Shared by "New Database" / "Open Database" / a completed conversion:
// opens Conn with the given connection string and refreshes the UI.
function TMainForm.OpenConnection(const ConnStr, DisplayName: string): Boolean;
begin
  Result := False;
  try
    if Conn.Connected then
    begin
      Table.Close;
      Conn.Close;
    end;

    Conn.ConnectionString := ConnStr;
    Conn.Open;

    LblFile.Caption := DisplayName;
    FillTableList;
    CboTables.ItemIndex := -1;
    SwitchToDatasheetView;
    UpdateCaption;
    Result := True;
  except
    on E: Exception do
    begin
      Conn.Close;
      MessageDlg('Could not open database:'#13#10 + E.Message, mtError, [mbOK], 0);
      UpdateCaption;
    end;
  end;
end;

procedure TMainForm.NewDatabaseClick(Sender: TObject);
var
  IsAccdb: Boolean;
  SaveDlg: TSaveDialog;
  ConnStr: string;
  Catalog: OleVariant;
begin
  if not PromptChooseFormat(IsAccdb) then
    Exit;

  SaveDlg := TSaveDialog.Create(nil);
  try
    if IsAccdb then
    begin
      SaveDlg.Filter := 'Access 2007-2016 Database (*.accdb)|*.accdb';
      SaveDlg.DefaultExt := 'accdb';
    end
    else
    begin
      SaveDlg.Filter := 'Access 97-2003 Database (*.mdb)|*.mdb';
      SaveDlg.DefaultExt := 'mdb';
    end;
    SaveDlg.Options := [ofOverwritePrompt, ofPathMustExist];

    if not SaveDlg.Execute then
      Exit;

    if IsAccdb then
      ConnStr := 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + SaveDlg.FileName + ';'
    else
      ConnStr := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + SaveDlg.FileName + ';';

    Screen.Cursor := crHourGlass;
    try
      try
        if Conn.Connected then
        begin
          Table.Close;
          Conn.Close;
        end;

        // ADOX creates the physical database file; late-bound so no
        // extra unit/type-library import is required.
        Catalog := CreateOleObject('ADOX.Catalog');
        Catalog.Create(ConnStr);
        Catalog := Unassigned;

        OpenConnection(ConnStr, ExtractFileName(SaveDlg.FileName));
      except
        on E: Exception do
          MessageDlg('Could not create database:'#13#10 + E.Message, mtError, [mbOK], 0);
      end;
    finally
      Screen.Cursor := crDefault;
      UpdateUIState;
    end;
  finally
    SaveDlg.Free;
  end;
end;

procedure TMainForm.OpenDatabaseClick(Sender: TObject);
begin
  if not OpenDlg.Execute then
    Exit;

  Screen.Cursor := crHourGlass;
  try
    OpenConnection(BuildConnectionString(OpenDlg.FileName), ExtractFileName(OpenDlg.FileName));
  finally
    Screen.Cursor := crDefault;
    UpdateUIState;
  end;
end;

// Converts the currently-open database to the other Access format
// (.mdb <-> .accdb) by creating a new file and copying each table's
// structure and data across. Queries, relationships, and indexes are
// not carried over - only tables and their data.
procedure TMainForm.ConvertDatabaseClick(Sender: TObject);
var
  TargetIsAccdb: Boolean;
  SaveDlg: TSaveDialog;
  DestConnStr: string;
  Catalog: OleVariant;
  DestConn: TADOConnection;
  SrcNames: TStringList;
  I, J: Integer;
  SrcTable, DestTable: TADOTable;
  SQL, ColDefs, TableName, FailedTable: string;
  F: TField;
begin
  if not Conn.Connected then
  begin
    MessageDlg('Open a database first.', mtInformation, [mbOK], 0);
    Exit;
  end;

  if Pos('ACE.OLEDB', UpperCase(Conn.ConnectionString)) > 0 then
    TargetIsAccdb := False   // currently .accdb -> convert to .mdb
  else
    TargetIsAccdb := True;   // currently .mdb -> convert to .accdb

  SaveDlg := TSaveDialog.Create(nil);
  try
    if TargetIsAccdb then
    begin
      SaveDlg.Filter := 'Access 2007-2016 Database (*.accdb)|*.accdb';
      SaveDlg.DefaultExt := 'accdb';
      SaveDlg.Title := 'Save As Access 2007-2016 Database';
    end
    else
    begin
      SaveDlg.Filter := 'Access 97-2003 Database (*.mdb)|*.mdb';
      SaveDlg.DefaultExt := 'mdb';
      SaveDlg.Title := 'Save As Access 97-2003 Database';
    end;
    SaveDlg.Options := [ofOverwritePrompt, ofPathMustExist];

    if not SaveDlg.Execute then
      Exit;

    if TargetIsAccdb then
      DestConnStr := 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + SaveDlg.FileName + ';'
    else
      DestConnStr := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + SaveDlg.FileName + ';';

    Screen.Cursor := crHourGlass;
    Application.ProcessMessages;

    SrcNames := TStringList.Create;
    DestConn := TADOConnection.Create(nil);
    SrcTable := TADOTable.Create(nil);
    DestTable := TADOTable.Create(nil);
    FailedTable := '';
    try
      try
        Catalog := CreateOleObject('ADOX.Catalog');
        Catalog.Create(DestConnStr);
        Catalog := Unassigned;

        DestConn.LoginPrompt := False;
        DestConn.ConnectionString := DestConnStr;
        DestConn.Open;

        Conn.GetTableNames(SrcNames, False);

        SrcTable.Connection := Conn;
        DestTable.Connection := DestConn;

        for I := 0 to SrcNames.Count - 1 do
        begin
          TableName := SrcNames[I];
          FailedTable := TableName;

          SrcTable.Close;
          SrcTable.TableName := TableName;
          SrcTable.Open;

          ColDefs := '';
          for J := 0 to SrcTable.FieldCount - 1 do
          begin
            F := SrcTable.Fields[J];
            if ColDefs <> '' then
              ColDefs := ColDefs + ', ';
            ColDefs := ColDefs + QuoteIdent(F.FieldName) + ' ' + FieldSQLType(F);
          end;

          SQL := 'CREATE TABLE ' + QuoteIdent(TableName) + ' (' + ColDefs + ')';
          DestConn.Execute(SQL);

          DestTable.Close;
          DestTable.TableName := TableName;
          DestTable.Open;

          SrcTable.First;
          while not SrcTable.Eof do
          begin
            DestTable.Append;
            for J := 0 to SrcTable.FieldCount - 1 do
            begin
              F := SrcTable.Fields[J];
              if F.DataType <> ftAutoInc then
                DestTable.FieldByName(F.FieldName).Value := F.Value;
            end;
            DestTable.Post;
            SrcTable.Next;
          end;

          SrcTable.Close;
          DestTable.Close;
        end;

        MessageDlg(Format('Converted %d table(s) to:'#13#10 + '%s',
          [SrcNames.Count, SaveDlg.FileName]), mtInformation, [mbOK], 0);

        if MessageDlg('Open the converted database now?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          if DestConn.Connected then
            DestConn.Close;
          OpenConnection(DestConnStr, ExtractFileName(SaveDlg.FileName));
        end;
      except
        on E: Exception do
          MessageDlg('Conversion failed on table "' + FailedTable + '":'#13#10 + E.Message,
            mtError, [mbOK], 0);
      end;
    finally
      SrcTable.Free;
      DestTable.Free;
      if DestConn.Connected then
        DestConn.Close;
      DestConn.Free;
      SrcNames.Free;
      Screen.Cursor := crDefault;
      UpdateUIState;
    end;
  finally
    SaveDlg.Free;
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

procedure TMainForm.AutoSizeGridColumns;
var
  I, ColWidth: Integer;
begin
  if Grid.Columns.Count = 0 then
    Exit;

  ColWidth := (Grid.ClientWidth - 20) div Grid.Columns.Count;
  if ColWidth < 40 then
    ColWidth := 40;

  for I := 0 to Grid.Columns.Count - 1 do
    Grid.Columns[I].Width := ColWidth;
end;

procedure TMainForm.PopulateDesignView;
var
  I: Integer;
  Item: TListItem;
  F: TField;
begin
  DesignView.Items.BeginUpdate;
  try
    DesignView.Items.Clear;
    if not Table.Active then
      Exit;

    for I := 0 to Table.FieldCount - 1 do
    begin
      F := Table.Fields[I];
      Item := DesignView.Items.Add;
      Item.Caption := F.FieldName;
      Item.SubItems.Add(AccessTypeName(F));
      if F.DataType in [ftString, ftWideString] then
        Item.SubItems.Add(IntToStr(F.Size))
      else
        Item.SubItems.Add('');
    end;
  finally
    DesignView.Items.EndUpdate;
  end;
end;

procedure TMainForm.SwitchToDatasheetView;
begin
  ShowingDesign := False;
  DesignView.Visible := False;
  Grid.Visible := True;
  Nav.Visible := True;
  BtnToggleView.Caption := 'Design View';
end;

procedure TMainForm.ToggleViewClick(Sender: TObject);
begin
  if not Table.Active then
  begin
    MessageDlg('Open a database and select a table first.', mtInformation, [mbOK], 0);
    Exit;
  end;

  ShowingDesign := not ShowingDesign;

  if ShowingDesign then
  begin
    PopulateDesignView;
    Grid.Visible := False;
    Nav.Visible := False;
    DesignView.Visible := True;
    BtnToggleView.Caption := 'Datasheet View';
  end
  else
    SwitchToDatasheetView;
end;

// Double-clicking a row in Design View lets the user rename the field
// and/or change its data type (where that's supported).
procedure TMainForm.DesignViewDblClick(Sender: TObject);
var
  Item: TListItem;
  F: TField;
  OldName, NewName, TargetColName, SQL: string;
  TIdx, NewTypeIdx: Integer;
  AllowType, NameChanged, TypeChanged: Boolean;
begin
  if not Table.Active then
    Exit;

  Item := DesignView.Selected;
  if Item = nil then
    Exit;

  OldName := Item.Caption;
  F := Table.FindField(OldName);
  if F = nil then
    Exit;

  TIdx := TypeIndexForField(F);
  AllowType := TIdx >= 0;

  if not PromptEditColumn(OldName, AccessTypeName(F), TIdx, AllowType, NewName, NewTypeIdx) then
    Exit;

  NameChanged := not SameText(NewName, OldName);
  TypeChanged := AllowType and (NewTypeIdx >= 0) and (NewTypeIdx <> TIdx);

  if not NameChanged and not TypeChanged then
    Exit;

  if NameChanged and (Table.FindField(NewName) <> nil) then
  begin
    MessageDlg('A column named "' + NewName + '" already exists.', mtWarning, [mbOK], 0);
    Exit;
  end;

  Screen.Cursor := crHourGlass;
  try
    try
      Table.Close;

      if NameChanged then
        RenameColumnADOX(Conn, Table.TableName, OldName, NewName);

      if NameChanged then
        TargetColName := NewName
      else
        TargetColName := OldName;

      if TypeChanged then
      begin
        SQL := Format('ALTER TABLE %s ALTER COLUMN %s %s',
          [QuoteIdent(Table.TableName), QuoteIdent(TargetColName), TypeSQLForIndex(NewTypeIdx)]);
        Conn.Execute(SQL);
      end;

      Table.Open;
      AutoSizeGridColumns;
      PopulateDesignView;
    except
      on E: Exception do
      begin
        MessageDlg('Could not update field:'#13#10 + E.Message, mtError, [mbOK], 0);
        if not Table.Active then
          Table.Open;
      end;
    end;
  finally
    Screen.Cursor := crDefault;
    UpdateUIState;
  end;
end;

procedure TMainForm.OpenSelectedTable;
begin
  if not Conn.Connected then
    Exit;

  if CboTables.ItemIndex < 0 then
    Exit;

  try
    Table.Close;
    Table.TableName := CboTables.Items[CboTables.ItemIndex];
    Table.Open;
    AutoSizeGridColumns;
    SwitchToDatasheetView;
    if ShowingDesign then
      PopulateDesignView;
  except
    on E: Exception do
      MessageDlg('Could not open table:'#13#10 + E.Message, mtError, [mbOK], 0);
  end;

  UpdateUIState;
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
    SwitchToDatasheetView;
    DesignView.Items.Clear;
    UpdateCaption;
  except
    on E: Exception do
      MessageDlg(E.Message, mtError, [mbOK], 0);
  end;
  UpdateUIState;
end;

procedure TMainForm.AddColumnClick(Sender: TObject);
var
  ColName, ColType, SQL: string;
begin
  if not (Conn.Connected and Table.Active) then
  begin
    MessageDlg('Open a database and select a table first.', mtInformation, [mbOK], 0);
    Exit;
  end;

  if not PromptNewColumn(ColName, ColType) then
    Exit;

  if Table.FindField(ColName) <> nil then
  begin
    MessageDlg('A column named "' + ColName + '" already exists.', mtWarning, [mbOK], 0);
    Exit;
  end;

  SQL := Format('ALTER TABLE %s ADD COLUMN %s %s',
    [QuoteIdent(Table.TableName), QuoteIdent(ColName), ColType]);

  Screen.Cursor := crHourGlass;
  try
    try
      Table.Close;
      Conn.Execute(SQL);
      Table.Open;
      AutoSizeGridColumns;
      if ShowingDesign then
        PopulateDesignView;
    except
      on E: Exception do
      begin
        MessageDlg('Could not add column:'#13#10 + E.Message, mtError, [mbOK], 0);
        if not Table.Active then
          Table.Open;
      end;
    end;
  finally
    Screen.Cursor := crDefault;
    UpdateUIState;
  end;
end;

procedure TMainForm.DeleteColumnClick(Sender: TObject);
var
  ColNames: TStringList;
  I: Integer;
  ColName, SQL: string;
begin
  if not (Conn.Connected and Table.Active) then
  begin
    MessageDlg('Open a database and select a table first.', mtInformation, [mbOK], 0);
    Exit;
  end;

  ColNames := TStringList.Create;
  try
    for I := 0 to Table.FieldCount - 1 do
      ColNames.Add(Table.Fields[I].FieldName);

    if not PromptSelectColumn(ColNames, 'Delete Column', 'Column to delete:', ColName) then
      Exit;
  finally
    ColNames.Free;
  end;

  if MessageDlg(Format('Delete column "%s"?'#13#10 +
    'All data in this column will be permanently lost.', [ColName]),
    mtWarning, [mbYes, mbNo], 0) <> mrYes then
    Exit;

  SQL := Format('ALTER TABLE %s DROP COLUMN %s',
    [QuoteIdent(Table.TableName), QuoteIdent(ColName)]);

  Screen.Cursor := crHourGlass;
  try
    try
      Table.Close;
      Conn.Execute(SQL);
      Table.Open;
      AutoSizeGridColumns;
      if ShowingDesign then
        PopulateDesignView;
    except
      on E: Exception do
      begin
        MessageDlg('Could not delete column:'#13#10 + E.Message, mtError, [mbOK], 0);
        if not Table.Active then
          Table.Open;
      end;
    end;
  finally
    Screen.Cursor := crDefault;
    UpdateUIState;
  end;
end;

end.
