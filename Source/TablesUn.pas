unit TablesUn;

interface

uses
  SysUtils, Classes, DB, ADODB;


type
  TUser = class
      id: Integer;
      family: String;
      mobile: String;
      UsrMng: Boolean;
      Blk: Boolean;
      Saham: Boolean;
      MPY: Boolean;
      SMV: Boolean;
      Rep1: Boolean;
      Rep2: Boolean;
      Rep3: Boolean;
      Rep4: Boolean;
      Rep5: Boolean;
      Rep6: Boolean;
      Rep7: Boolean;
      Rep8: Boolean;
      Rep9: Boolean;
      Rep10: Boolean;
  end;
type
  TDmTables = class(TDataModule)
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    AdqSahamPic: TADOQuery;
    ADOExcelConnection: TADOConnection;
    AdqExcel: TADOQuery;
    DsExcel: TDataSource;
    AdqBlockPic: TADOQuery;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    CurrentUser: TUser;
    Function Connect: Boolean;
    Procedure RunQuery(Sql: String);
    Procedure ExecQuery(Sql: String);
    Function CheckPrivilages(Name: String): Boolean;
  end;

var
  DmTables: TDmTables;

implementation

uses Config;

{$R *.dfm}

function TDmTables.Connect: Boolean;
var
  txt: TStrings;
begin
  try
    txt := TStringList.Create;
    txt.LoadFromFile('Server.config');
    ADOConnection1.Connected := False;
//    Provider=;Password=Sanyar@jz@ss;Persist Security Info=True;User ID=sa;Initial Catalog=TMT;Data Source=10.211.55.2
    ADOConnection1.ConnectionString := 'Provider=SQLOLEDB.1;Persist Security Info=True;User ID=sa;password=Sanyar@jz@ss;Initial Catalog=TMT;Data Source='+Trim(txt[0]);
//    ADOConnection1.ConnectionString := 'Provider=SQLNCLI11.1;Persist Security Info=False;User ID=sa;password=P@$$w0rd;Initial Catalog=TMT;Data Source='+Trim(txt[0])+';Initial File Name="";Server SPN="";Option=2';
    ADOConnection1.Connected := True;
    Result := true;
  Except
    On E:Exception do
    begin
      ShowError('خطا در اتصال به سرور. لطفا با واحد پشتیبانی تماس حاصل نمایید'+E.ToString);
      Result := False;
    end;
  end;
End;

procedure TDmTables.DataModuleCreate(Sender: TObject);
begin
  CurrentUser := TUser.Create;
end;

Procedure TDmTables.RunQuery(Sql: String);
Begin
  ADOQuery1.Close;
  ADOQuery1.SQL.Text := Sql;
  ADOQuery1.Open;
End;

Procedure TDmTables.ExecQuery(Sql: String);
Begin
  ADOQuery1.Close;
  ADOQuery1.SQL.Text := Sql;
  ADOQuery1.ExecSQL;
End;

Function TDmTables.CheckPrivilages(Name: String): Boolean;
Begin
  with ADOQuery1 do
  begin
    Close;
    SQL.Clear;
    SQL.Add('Exec PrcCheckPrivilage '+IntToStr(CurrentUser.id)+','+QuotedStr(Name));
    Open;
    Result := Fields[0].AsString = 'Success';
  end;
End;

end.
