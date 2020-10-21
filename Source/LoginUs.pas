unit LoginUs;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, pngimage, ExtCtrls, DBCtrls;

type
  TFmLogin = class(TForm)
    Label1: TLabel;
    Shape1: TShape;
    Shape3: TShape;
    Image1: TImage;
    Label2: TLabel;
    Label3: TLabel;
    EDUserName: TEdit;
    EDPass: TEdit;
    Button1: TButton;
    Button2: TButton;
    Shape2: TShape;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Label1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FmLogin: TFmLogin;

implementation

uses Config, TablesUn;

{$R *.dfm}

procedure TFmLogin.Button1Click(Sender: TObject);
begin
  if Trim(EDUserName.Text) = '' then
    ShowError('نام کاربری مشخص نشده است')
  Else if Trim(EDPass.Text) = '' then
    ShowError('کلمه عبور مشخص نشده است')
  else With DmTables.ADOQuery1 do
    if DmTables.Connect then
    Try
      Close;
      SQL.Clear;
      SQL.Add('Exec PrcAuthenticate '+QuotedStr(EDUserName.Text)+','+QuotedStr(EncryptStr(EDPass.Text)));
      Open;
      if IsEmpty then
        ShowError('نام کاربری یا کلمه عبور اشتباه می باشد')
      else
      begin
        DmTables.CurrentUser.ID := DmTables.AdoQuery1.FindField('ID').AsInteger;
        DmTables.CurrentUser.family := DmTables.AdoQuery1.FindField('Family').AsString;
        DmTables.CurrentUser.mobile := DmTables.AdoQuery1.FindField('Mobile').AsString;
        DmTables.CurrentUser.UsrMng := DmTables.AdoQuery1.FindField('UsrMng').AsInteger=1;
        DmTables.CurrentUser.Blk := DmTables.AdoQuery1.FindField('Blk').AsInteger=1;
        DmTables.CurrentUser.Saham := DmTables.AdoQuery1.FindField('Saham').AsInteger=1;
        DmTables.CurrentUser.MPY := DmTables.AdoQuery1.FindField('MPY').AsInteger=1;
        DmTables.CurrentUser.SMV := DmTables.AdoQuery1.FindField('SMV').AsInteger=1;
        DmTables.CurrentUser.Rep1 := DmTables.AdoQuery1.FindField('Rep1').AsInteger=1;
        DmTables.CurrentUser.Rep2 := DmTables.AdoQuery1.FindField('Rep2').AsInteger=1;
        DmTables.CurrentUser.Rep3 := DmTables.AdoQuery1.FindField('Rep3').AsInteger=1;
        DmTables.CurrentUser.Rep4 := DmTables.AdoQuery1.FindField('Rep4').AsInteger=1;
        DmTables.CurrentUser.Rep5 := DmTables.AdoQuery1.FindField('Rep5').AsInteger=1;
        DmTables.CurrentUser.Rep6 := DmTables.AdoQuery1.FindField('Rep6').AsInteger=1;
        DmTables.CurrentUser.Rep7 := DmTables.AdoQuery1.FindField('Rep7').AsInteger=1;
        DmTables.CurrentUser.Rep8 := DmTables.AdoQuery1.FindField('Rep8').AsInteger=1;
        DmTables.CurrentUser.Rep9 := DmTables.AdoQuery1.FindField('Rep9').AsInteger=1;
        DmTables.CurrentUser.Rep10 := DmTables.AdoQuery1.FindField('Rep10').AsInteger=1;
        ModalResult := mrYes;
      end;
    Except
      ShowError('خطا در شناسایی مشخصات کاربر');
    end;
end;

procedure TFmLogin.Button2Click(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TFmLogin.Label1MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  ReleaseCapture;
  SendMessage(Self.Handle, WM_SYSCOMMAND, 61458, 0) ;
end;

end.
