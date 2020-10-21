unit MainUn;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons, Grids, DBGrids, DB, ADODB, DBClient,
  AdvObj, BaseGrid, AdvGrid, DBAdvGrid, ToolWin, ActnMan, ActnCtrls, ActnMenus,
  Menus, ImgList, DBCtrls, Mask, ComCtrls, Tabs, TabNotBk, ExtDlgs, jpeg,
  SolarCalendarPackage, PNGImage, frxClass, frxDBSet, frxDesgn, frxDMPExport,
  frxPreview;

type
  TFmMain = class(TForm)
    Label1: TLabel;
    Shape1: TShape;
    Shape2: TShape;
    Shape3: TShape;
    Panel1: TPanel;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    Panel2: TPanel;
    AdsSaham: TADOQuery;
    DsSaham: TDataSource;
    AdsSahamID: TIntegerField;
    AdsSahamRadif: TSmallintField;
    AdsSahamSerial: TIntegerField;
    AdsSahamUserID: TSmallintField;
    AdsSahamSerial1: TIntegerField;
    AdsSahamSerial2: TIntegerField;
    AdsSahamSerial3: TIntegerField;
    AdsSahamName: TWideStringField;
    AdsSahamFamily: TWideStringField;
    AdsSahamFather: TWideStringField;
    AdsSahamNationalID: TWideStringField;
    AdsSahamStatus: TWordField;
    AdsSahamSS: TWideStringField;
    AdsSahamSPlace: TWideStringField;
    AdsSahamBirthPlace: TWideStringField;
    AdsSahamBirthDate: TStringField;
    AdsSahamSSerial: TWideStringField;
    AdsSahamNationalSerial: TWideStringField;
    AdsSahamJob: TWideStringField;
    AdsSahamTel: TWideStringField;
    AdsSahamTel2: TWideStringField;
    AdsSahamMobile: TWideStringField;
    AdsSahamAddress: TWideStringField;
    AdsSahamJAddress: TWideStringField;
    AdsSahamLawyer: TWideStringField;
    AdsSahamLFamily: TWideStringField;
    AdsSahamLNationalID: TWideStringField;
    AdsSahamWife: TWideStringField;
    AdsSahamWFamily: TWideStringField;
    AdsSahamWFather: TWideStringField;
    AdsSahamWSS: TWideStringField;
    AdsSahamWNationalID: TWideStringField;
    AdsSahamWSPlace: TWideStringField;
    AdsSahamWBirthPlace: TWideStringField;
    AdsSahamWBirthDate: TStringField;
    AdsSahamWSSerial: TWideStringField;
    AdsSahamWNationalSerial: TWideStringField;
    AdsSahamWJob: TWideStringField;
    AdsSahamWTel: TWideStringField;
    AdsSahamWMobile: TWideStringField;
    AdsSahamWAddress: TWideStringField;
    AdsSahamWJAddress: TWideStringField;
    AdsSahamBLockID: TSmallintField;
    AdsSahamTabaghe: TWordField;
    AdsSahamVahed: TWordField;
    AdsSahamTabagheCount: TWordField;
    AdsSahamMetraj: TSmallintField;
    AdsSahamTip: TWideStringField;
    AdsSahamPoint1: TIntegerField;
    AdsSahamPoint2: TIntegerField;
    AdsSahamAdamMarghobiat: TIntegerField;
    AdsSahamHologramSerial: TIntegerField;
    AdsSahamSanadSerial: TIntegerField;
    AdsSahamPayment: TBCDField;
    AdsSahamBedehi: TBCDField;
    AdsSahamArzeshAfzodeh: TBCDField;
    AdsSahamTaminEjtemaei: TBCDField;
    AdsSahamMaliatTakilifi: TBCDField;
    AdsSahamTasviyePeymankaran: TBCDField;
    AdsSahamHazineSodorSanad: TBCDField;
    PnGridSetting: TPanel;
    Label2: TLabel;
    Shape4: TShape;
    Shape5: TShape;
    Shape6: TShape;
    SpeedButton4: TSpeedButton;
    AdsGridSetting: TClientDataSet;
    DsGridSetting: TDataSource;
    AdsGridSettingName: TStringField;
    AdsGridSettingEngName: TStringField;
    AdsGridSettingActive: TIntegerField;
    Panel3: TPanel;
    DBAdvGrid1: TDBAdvGrid;
    Panel4: TPanel;
    BnUp: TSpeedButton;
    BtnDown: TSpeedButton;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    Notebook1: TNotebook;
    DBGUsers: TDBGrid;
    DsUser: TDataSource;
    AdsUsers: TADOStoredProc;
    AdsUsersID: TSmallintField;
    AdsUsersFamily: TWideStringField;
    AdsUsersMobile: TWideStringField;
    AdsUsersUserName: TWideStringField;
    AdsUsersEmail: TWideStringField;
    AdsUsersSemat: TWideStringField;
    AdsUsersDate: TStringField;
    PNPass: TPanel;
    Label3: TLabel;
    Label4: TLabel;
    Shape7: TShape;
    Shape8: TShape;
    Shape9: TShape;
    Label5: TLabel;
    EDPass1: TEdit;
    EDPass2: TEdit;
    SpeedButton3: TSpeedButton;
    SpeedButton6: TSpeedButton;
    Shape10: TShape;
    N4: TMenuItem;
    ImageList1: TImageList;
    Panel5: TPanel;
    Splitter1: TSplitter;
    ADTGroup: TADOTable;
    ADTGroupID: TSmallintField;
    ADTGroupName: TWideStringField;
    ADTGroupUsrMng: TWordField;
    ADTGroupBlk: TWordField;
    ADTGroupSaham: TWordField;
    ADTGroupMPY: TWordField;
    ADTGroupSMV: TWordField;
    ADTGroupRep1: TWordField;
    ADTGroupRep2: TWordField;
    ADTGroupRep3: TWordField;
    ADTGroupRep4: TWordField;
    ADTGroupRep5: TWordField;
    ADTGroupRep6: TWordField;
    ADTGroupRep7: TWordField;
    ADTGroupRep8: TWordField;
    ADTGroupRep9: TWordField;
    ADTGroupRep10: TWordField;
    DsGroup: TDataSource;
    DSGrpUser: TDataSource;
    DBAdvGrid2: TDBAdvGrid;
    ADTGroupNote: TStringField;
    AdsGrpUser: TClientDataSet;
    AdsGrpUserChk: TIntegerField;
    AdsGrpUserID: TIntegerField;
    AdsGrpUserFamily: TStringField;
    AdsGrpUserMobile: TStringField;
    Panel6: TPanel;
    DBGGroup: TDBGrid;
    Panel7: TPanel;
    DBCheckBox1: TDBCheckBox;
    DBCheckBox2: TDBCheckBox;
    DBCheckBox3: TDBCheckBox;
    DBCheckBox4: TDBCheckBox;
    DBCheckBox5: TDBCheckBox;
    SpeedButton7: TSpeedButton;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Panel8: TPanel;
    BGBlock: TDBGrid;
    AdqBlock: TADOQuery;
    DsBlock: TDataSource;
    AdqBlockID: TSmallintField;
    AdqBlockX: TIntegerField;
    AdqBlockY: TIntegerField;
    AdqBlockTabaghat: TWordField;
    AdqBlockKind: TWideStringField;
    AdqBlockMasahat: TIntegerField;
    AdqBlockParvane: TWideStringField;
    AdqBlockNamayande: TWideStringField;
    AdqBlockNTel: TWideStringField;
    AdqBlockSazande: TWideStringField;
    AdqBlockSTel: TWideStringField;
    AdqBlockNote: TWideStringField;
    PNSahamTool: TPanel;
    BtnSahamSetting: TSpeedButton;
    PNUserTool: TPanel;
    BtnSetPass: TSpeedButton;
    SpeedButton8: TSpeedButton;
    SpeedButton11: TSpeedButton;
    Panel11: TPanel;
    DbgSaham: TDBGrid;
    Splitter2: TSplitter;
    AdsChilds: TADOQuery;
    DsChilds: TDataSource;
    AdsChildsRadif: TSmallintField;
    AdsChildsName: TWideStringField;
    AdsChildsFamily: TWideStringField;
    AdsChildsNesbat: TWideStringField;
    AdsChildsBirthDate: TStringField;
    AdsChildsEducation: TWideStringField;
    AdsChildsTel: TWideStringField;
    AdsChildsPic: TVarBytesField;
    AdsChildsUserID: TSmallintField;
    AdsChildsEDate: TDateTimeField;
    AdsChildsUserFamily: TWideStringField;
    PNChilds: TPanel;
    DbgChild: TDBGrid;
    Panel13: TPanel;
    SpeedButton12: TSpeedButton;
    PnSaham: TPanel;
    LBSahamTitle: TLabel;
    Shape14: TShape;
    Shape15: TShape;
    Shape16: TShape;
    Panel12: TPanel;
    NtbSaham: TNotebook;
    Panel14: TPanel;
    BtnNtbSaham: TSpeedButton;
    SpeedButton14: TSpeedButton;
    SpeedButton15: TSpeedButton;
    SpeedButton16: TSpeedButton;
    Shape17: TShape;
    Shape18: TShape;
    Label20: TLabel;
    Shape19: TShape;
    Shape20: TShape;
    Label24: TLabel;
    Label25: TLabel;
    Label26: TLabel;
    Label27: TLabel;
    Label28: TLabel;
    Label29: TLabel;
    Label30: TLabel;
    Label31: TLabel;
    Label32: TLabel;
    OpenPictureDialog1: TOpenPictureDialog;
    Panel15: TPanel;
    ImgSaham: TImage;
    AdsSahamEDate: TStringField;
    Label33: TLabel;
    Label34: TLabel;
    Label35: TLabel;
    Label36: TLabel;
    Label37: TLabel;
    Label40: TLabel;
    Label41: TLabel;
    Label42: TLabel;
    Label43: TLabel;
    Label44: TLabel;
    Label45: TLabel;
    Label46: TLabel;
    Label47: TLabel;
    Label48: TLabel;
    Label49: TLabel;
    Label50: TLabel;
    Label51: TLabel;
    Label52: TLabel;
    Label53: TLabel;
    Label54: TLabel;
    Label55: TLabel;
    Label21: TLabel;
    Label56: TLabel;
    Label57: TLabel;
    Label58: TLabel;
    Label62: TLabel;
    Label59: TLabel;
    Label60: TLabel;
    Label61: TLabel;
    Label63: TLabel;
    Label64: TLabel;
    Label65: TLabel;
    Label66: TLabel;
    Label67: TLabel;
    Label69: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    AdsSahamFaz: TWideStringField;
    AdsSahamHesabNo: TWideStringField;
    AdsSahamBank: TWideStringField;
    AdsSahamShaba: TWideStringField;
    AdsSahamMostajer: TWideStringField;
    AdsSahamMosTel: TWideStringField;
    AdsSahamMostajerEdate: TStringField;
    Label38: TLabel;
    Label39: TLabel;
    Label68: TLabel;
    Label77: TLabel;
    Label78: TLabel;
    Label79: TLabel;
    Label80: TLabel;
    DBEdit8: TDBEdit;
    DBEdit9: TDBEdit;
    DBEdit10: TDBEdit;
    DBEdit11: TDBEdit;
    DBEdit12: TDBEdit;
    DBEdit13: TDBEdit;
    DBEdit15: TDBEdit;
    DBEdit16: TDBEdit;
    DBEdit17: TDBEdit;
    DBEdit18: TDBEdit;
    DBEdit19: TDBEdit;
    DBEdit20: TDBEdit;
    DBEdit21: TDBEdit;
    DBEdit22: TDBEdit;
    DBEdit23: TDBEdit;
    DBEdit24: TDBEdit;
    DBEdit27: TDBEdit;
    DBEdit28: TDBEdit;
    DBEdit29: TDBEdit;
    DBEdit30: TDBEdit;
    DBEdit31: TDBEdit;
    EDMosEDate: TSolarDatePicker;
    EDBirthDate: TSolarDatePicker;
    DBEdit25: TDBEdit;
    DBEdit26: TDBEdit;
    DBEdit32: TDBEdit;
    DBEdit33: TDBEdit;
    DBEdit35: TDBEdit;
    DBEdit36: TDBEdit;
    DBEdit37: TDBEdit;
    EDWBirthDate: TSolarDatePicker;
    DBEdit34: TDBEdit;
    DBEdit38: TDBEdit;
    DBEdit39: TDBEdit;
    DBEdit40: TDBEdit;
    DBEdit41: TDBEdit;
    DBEdit42: TDBEdit;
    DBEdit43: TDBEdit;
    DBEdit44: TDBEdit;
    DBEdit45: TDBEdit;
    DBEdit47: TDBEdit;
    DBEdit46: TDBEdit;
    DBEdit48: TDBEdit;
    DBEdit49: TDBEdit;
    DBEdit50: TDBEdit;
    DBEdit51: TDBEdit;
    DBEdit52: TDBEdit;
    DBEdit53: TDBEdit;
    DBEdit54: TDBEdit;
    DBEdit55: TDBEdit;
    DBEdit56: TDBEdit;
    DBEdit57: TDBEdit;
    DBEdit58: TDBEdit;
    SpeedButton18: TSpeedButton;
    SpeedButton19: TSpeedButton;
    SpeedButton5: TSpeedButton;
    EdCsID1: TEdit;
    EDCsID2: TEdit;
    EDCsName: TEdit;
    EDCsFamily: TEdit;
    EDCsNationalID: TEdit;
    EDCsBlock: TEdit;
    EDCsVahed: TEdit;
    EDCsTabaghe: TEdit;
    AdsUsersPic: TBlobField;
    PnPayment: TPanel;
    Label11: TLabel;
    Shape11: TShape;
    Shape12: TShape;
    Shape13: TShape;
    SpeedButton9: TSpeedButton;
    Panel10: TPanel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    SpeedButton10: TSpeedButton;
    DBText1: TDBText;
    DBText2: TDBText;
    DBText3: TDBText;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    DBEdit3: TDBEdit;
    DBEdit4: TDBEdit;
    DBEdit5: TDBEdit;
    DBEdit6: TDBEdit;
    DBEdit7: TDBEdit;
    AdsSahamShenasePardakht: TWideStringField;
    AdsSahamEducation: TWideStringField;
    AdsSahamNoGharardad: TWordField;
    AdsSahamVam: TWideStringField;
    AdsSahamdateGharardad: TStringField;
    AdsSahamInfoSh: TWideStringField;
    AdsSahamMvDate: TStringField;
    AdsSahamMvFrm: TIntegerField;
    AdsSahamMvDaftarID: TIntegerField;
    AdsSahamMvPageID: TIntegerField;
    AdsSahamMvHologramSerial: TIntegerField;
    AdsSahamMVSanadSerial: TIntegerField;
    AdsSahamMvSahamSerial: TWideStringField;
    SpeedButton17: TSpeedButton;
    FileOpenDialog1: TFileOpenDialog;
    PNExcel: TPanel;
    Label70: TLabel;
    Shape21: TShape;
    Shape22: TShape;
    Shape23: TShape;
    Panel9: TPanel;
    SpeedButton20: TSpeedButton;
    DBGrid1: TDBGrid;
    Panel16: TPanel;
    Label71: TLabel;
    CMBExelID: TComboBox;
    CMBExcelMyField: TComboBox;
    CMBExcelOthField: TComboBox;
    BtnDoExcelExport: TSpeedButton;
    CMBExcelMyFieldEng: TComboBox;
    PNExcelWaiting: TPanel;
    Label72: TLabel;
    Label73: TLabel;
    AdsSahamParentID: TIntegerField;
    Label74: TLabel;
    Label75: TLabel;
    Label76: TLabel;
    DBEdit59: TDBEdit;
    DBEdit60: TDBEdit;
    DBEdit61: TDBEdit;
    Label81: TLabel;
    DBEdit62: TDBEdit;
    DBEdit63: TDBEdit;
    Label82: TLabel;
    Label84: TLabel;
    BtnMoveMentSec: TSpeedButton;
    Label83: TLabel;
    DBEdit65: TDBEdit;
    Label85: TLabel;
    DBEdit66: TDBEdit;
    Label86: TLabel;
    DBEdit67: TDBEdit;
    Label87: TLabel;
    DBEdit68: TDBEdit;
    Label88: TLabel;
    DBEdit69: TDBEdit;
    Label89: TLabel;
    DBEdit70: TDBEdit;
    SpeedButton21: TSpeedButton;
    DBComboBox1: TDBComboBox;
    EDSahamMvDate: TSolarDatePicker;
    frxReport1: TfrxReport;
    AdsChildsSahamID: TIntegerField;
    N5: TMenuItem;
    NRep: TMenuItem;
    frxDBDataset1: TfrxDBDataset;
    N7: TMenuItem;
    N8: TMenuItem;
    N9: TMenuItem;
    N10: TMenuItem;
    AdsSahamX: TIntegerField;
    AdsSahamY: TIntegerField;
    AdsSahamTabaghat: TWordField;
    AdsSahamKind: TWideStringField;
    AdsSahamMasahat: TIntegerField;
    AdsSahamParvane: TWideStringField;
    AdsSahamNamayande: TWideStringField;
    AdsSahamNTel: TWideStringField;
    AdsSahamSazande: TWideStringField;
    AdsSahamSTel: TWideStringField;
    AdsSahamNote: TWideStringField;
    AdsSahamMBaravord: TBCDField;
    AdsSahamMHazineAmadesazi: TBCDField;
    AdsSahamMWater: TBCDField;
    AdsSahamMBargh: TBCDField;
    AdsSahamMFazelab: TBCDField;
    AdsSahamMKharidSahamEzafe: TBCDField;
    AdsSahamMBimeGozaresh: TBCDField;
    AdsSahamMMaliatTaklifiGozaresh: TBCDField;
    AdsSahamMArzeshAfzodeGozaresh: TBCDField;
    AdsSahamMBerozresani: TBCDField;
    AdsSahamMParvaneSakht: TBCDField;
    AdsSahamMKharidZamin: TBCDField;
    AdsSahamMMablaghVamMehr: TBCDField;
    AdsSahamMMablaghPardakhtiGhabli: TBCDField;
    AdsSahamMMablaghPardakhtiBadi: TBCDField;
    AdsSahamMPardakhtNahae: TBCDField;
    AdsSahamMMablaghtahatorVahed: TBCDField;
    AdsSahamMPeymankarTahator: TWideStringField;
    AdsSahamMVarizi1: TBCDField;
    AdsSahamMVarizi2: TBCDField;
    AdsSahamMVarizi3: TBCDField;
    AdsSahamMVarizi4: TBCDField;
    AdsSahamMVarizi5: TBCDField;
    AdsSahamMVarizi6: TBCDField;
    AdsSahamJamKolBedehi: TFloatField;
    AdsSahamJamKolPardakhti: TFloatField;
    SpeedButton13: TSpeedButton;
    Label90: TLabel;
    Label91: TLabel;
    Label92: TLabel;
    Label93: TLabel;
    Label94: TLabel;
    Label95: TLabel;
    Label96: TLabel;
    Label97: TLabel;
    Label98: TLabel;
    Label99: TLabel;
    Label100: TLabel;
    Label101: TLabel;
    Label102: TLabel;
    Label103: TLabel;
    Label104: TLabel;
    Label105: TLabel;
    Label106: TLabel;
    Label107: TLabel;
    Label108: TLabel;
    Label109: TLabel;
    Label110: TLabel;
    Label111: TLabel;
    Label112: TLabel;
    Label113: TLabel;
    Label114: TLabel;
    Label115: TLabel;
    DBEdit14: TDBEdit;
    DBEdit64: TDBEdit;
    DBEdit71: TDBEdit;
    DBEdit72: TDBEdit;
    DBEdit73: TDBEdit;
    DBEdit74: TDBEdit;
    DBEdit75: TDBEdit;
    DBEdit76: TDBEdit;
    DBEdit77: TDBEdit;
    DBEdit78: TDBEdit;
    DBEdit79: TDBEdit;
    DBEdit80: TDBEdit;
    DBEdit81: TDBEdit;
    DBEdit82: TDBEdit;
    DBEdit83: TDBEdit;
    DBEdit84: TDBEdit;
    DBEdit85: TDBEdit;
    DBEdit86: TDBEdit;
    DBEdit87: TDBEdit;
    DBEdit88: TDBEdit;
    DBEdit89: TDBEdit;
    DBEdit90: TDBEdit;
    DBEdit91: TDBEdit;
    DBEdit92: TDBEdit;
    DBText4: TDBText;
    DBText5: TDBText;
    SpeedButton22: TSpeedButton;
    SaveDialog1: TSaveDialog;
    PnBlockTool: TPanel;
    SpeedButton23: TSpeedButton;
    AdqBlockPic: TBlobField;
    PnGetSahamID: TPanel;
    Label116: TLabel;
    Edit1: TEdit;
    Shape24: TShape;
    EDCsMobile: TEdit;
    EDCsWife: TEdit;
    SpeedButton24: TSpeedButton;
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure AdsSahamStatusGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure SpeedButton4Click(Sender: TObject);
    procedure BtnSahamSettingClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure AdsGridSettingBeforePost(DataSet: TDataSet);
    procedure BnUpClick(Sender: TObject);
    procedure Label1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Label2MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure EDSearchChange(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure Notebook1PageChanged(Sender: TObject);
    procedure AdsUsersBeforePost(DataSet: TDataSet);
    procedure AdsUsersPassGetText(Sender: TField; var Text: string;
      DisplayText: Boolean);
    procedure SpeedButton6Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure AdsUsersBeforeEdit(DataSet: TDataSet);
    procedure AdsSahamBeforeEdit(DataSet: TDataSet);
    procedure BtnSetPassClick(Sender: TObject);
    procedure ADTGroupBeforeInsert(DataSet: TDataSet);
    procedure AdqGrpUserBeforePost(DataSet: TDataSet);
    procedure ADTGroupAfterScroll(DataSet: TDataSet);
    procedure SpeedButton7Click(Sender: TObject);
    procedure AdqBlockBeforeInsert(DataSet: TDataSet);
    procedure AdqBlockBeforePost(DataSet: TDataSet);
    procedure SpeedButton9Click(Sender: TObject);
    procedure SpeedButton8Click(Sender: TObject);
    procedure SpeedButton10Click(Sender: TObject);
    procedure SpeedButton11Click(Sender: TObject);
    procedure AdsSahamAfterScroll(DataSet: TDataSet);
    procedure AdsChildsBeforePost(DataSet: TDataSet);
    procedure SpeedButton12Click(Sender: TObject);
    procedure BtnNtbSahamClick(Sender: TObject);
    procedure DbgSahamDblClick(Sender: TObject);
    procedure DBImage1DblClick(Sender: TObject);
    procedure SpeedButton18Click(Sender: TObject);
    procedure SpeedButton19Click(Sender: TObject);
    procedure EDMosEDateChange(Sender: TObject);
    procedure LBSahamTitleMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure EdCsID1Change(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SpeedButton17Click(Sender: TObject);
    procedure SpeedButton20Click(Sender: TObject);
    procedure Label70MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure BtnDoExcelExportClick(Sender: TObject);
    procedure Label11MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure SpeedButton21Click(Sender: TObject);
    procedure AdsSahamStatusSetText(Sender: TField; const Text: string);
    procedure DbgSahamTitleClick(Column: TColumn);
    procedure AdsSahamBeforePost(DataSet: TDataSet);
    procedure NRepClick(Sender: TObject);
    procedure Label1DblClick(Sender: TObject);
    procedure AdsSahamCalcFields(DataSet: TDataSet);
    procedure SpeedButton22Click(Sender: TObject);
    procedure frxReport1BeforePrint(Sender: TfrxReportComponent);
    procedure SpeedButton23Click(Sender: TObject);
    procedure Edit1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure SpeedButton24Click(Sender: TObject);
  private
    { Private declarations }
    DesignRep: Boolean;
    procedure PrcLoadGridSetting(z: Integer);
    function GetHeader(const AFile: string; const AByteCount: integer): string;
    function GetImageFromBlob(const ABlobField: TBlobField): TGraphic;
  public
    { Public declarations }
  end;

var
  FmMain: TFmMain;

implementation

uses Config, TablesUn;

{$R *.dfm}

function TFmMain.GetHeader(const AFile: string; const AByteCount: integer): string;
const
  HEADER_STR = '%s_HEADER: array [0 .. %d] of byte = (%s)';
var
  _HeaderStream: TMemoryStream;
  _FileStream: TMemoryStream;
  _Buf: integer;
  _Ext: string;
  _FullByteStrArr: string;
  _ByteStr: string;
  i: integer;
begin
  Result := '';
  if not FileExists(AFile) then
    Exit;

  _HeaderStream := TMemoryStream.Create;
  _FileStream := TMemoryStream.Create;
  try
    _FileStream.LoadFromFile(AFile);
    _FileStream.Position := 0;
    _HeaderStream.CopyFrom(_FileStream, 5);
    if _HeaderStream.Size > 4 then
    begin
      _HeaderStream.Position := 0;
      _ByteStr := '';
      _FullByteStrArr := '';
      for i := 0 to AByteCount do
      begin
        _HeaderStream.Read(_Buf, 1);
        _ByteStr := IntToHex(_Buf, 2);
        _FullByteStrArr := _FullByteStrArr + ', $' +
          Copy(_ByteStr, Length(_ByteStr) - 1, 2);
      end;
      _FullByteStrArr := Copy(_FullByteStrArr, 3, Length(_FullByteStrArr));

      _Ext := UpperCase(ExtractFileExt(AFile));
      _Ext := Copy(_Ext, 2, Length(_Ext));
      Result := Format(HEADER_STR, [_Ext, AByteCount, _FullByteStrArr]);
    end;
  finally
    FreeAndNil(_FileStream);
    FreeAndNil(_HeaderStream);
  end;
end;

function TFmMain.GetImageFromBlob(const ABlobField: TBlobField): TGraphic;
CONST
  JPG_HEADER: array [0 .. 2] of byte = ($FF, $D8, $FF);
  GIF_HEADER: array [0 .. 2] of byte = ($47, $49, $46);
  BMP_HEADER: array [0 .. 1] of byte = ($42, $4D);
  PNG_HEADER: array [0 .. 3] of byte = ($89, $50, $4E, $47);
  TIF_HEADER: array [0 .. 2] of byte = ($49, $49, $2A);
  TIF_HEADER2: array [0 .. 2] of byte = (77, 77, 00);
  PCX_HEADER: array [0 .. 2] of byte = (10, 5, 1);

var
  _HeaderStream: TMemoryStream;
  _ImgStream: TMemoryStream;
  _GraphicClassName: string;
  _GraphicClass: TGraphicClass;
begin
  Result := nil;

  _HeaderStream := TMemoryStream.Create;
  _ImgStream := TMemoryStream.Create;
  try
    ABlobField.SaveToStream(_ImgStream);
    _ImgStream.Position := 0;
    _HeaderStream.CopyFrom(_ImgStream, 5);
    if _HeaderStream.Size > 4 then
    begin
      if CompareMem(_HeaderStream.Memory, @JPG_HEADER, SizeOf(JPG_HEADER)) then
        _GraphicClassName := 'TJPEGImage'
      else if CompareMem(_HeaderStream.Memory, @GIF_HEADER, SizeOf(GIF_HEADER))
      then
        _GraphicClassName := 'TGIFImage'
      else if CompareMem(_HeaderStream.Memory, @PNG_HEADER, SizeOf(PNG_HEADER))
      then
        _GraphicClassName := 'TPNGImage'
      else if CompareMem(_HeaderStream.Memory, @BMP_HEADER, SizeOf(BMP_HEADER))
      then
        _GraphicClassName := 'TBitmap'
      else if CompareMem(_HeaderStream.Memory, @TIF_HEADER, SizeOf(TIF_HEADER))
      then
        _GraphicClassName := 'TWICImage'
      else if CompareMem(_HeaderStream.Memory, @TIF_HEADER2, SizeOf(TIF_HEADER2))
      then
        _GraphicClassName := 'TWICImage'
      else if CompareMem(_HeaderStream.Memory, @PCX_HEADER, SizeOf(PCX_HEADER))
      then
        _GraphicClassName := 'PCXImage';

      RegisterClasses([TIcon, TMetafile, TBitmap, TJPEGImage, TPngImage,
        TWICImage]);
      _GraphicClass := TGraphicClass(FindClass(_GraphicClassName));
      if (_GraphicClass <> nil) then
      begin
        Result := _GraphicClass.Create; // Create appropriate graphic class
        _ImgStream.Position := 0;
        Result.LoadFromStream(_ImgStream);
      end;
    end;
  finally
    FreeAndNil(_ImgStream);
    FreeAndNil(_HeaderStream);
  end;
end;

procedure TFmMain.PrcLoadGridSetting(z: Integer);
var
  I, RNo: Integer;
begin
  try
    AdsGridSetting.DisableControls;
    if Not AdsGridSetting.Active then
    begin
      AdsGridSetting.CreateDataSet;
      RNo := 1;
    end
    else
    begin
      RNo := AdsGridSetting.RecNo;
      AdsGridSetting.EmptyDataSet;
    end;
    for I := 0 to DbgSaham.Columns.Count - 1 do
    begin
      AdsGridSetting.Append;
      if DbgSaham.Columns[i].Visible then
        AdsGridSettingActive.AsInteger := 1
      Else
        AdsGridSettingActive.AsInteger := 0;
      AdsGridSettingName.AsString := DbgSaham.Columns[i].Field.DisplayLabel;
      AdsGridSettingEngName.AsString := DbgSaham.Columns[i].Field.FieldName;
      AdsGridSetting.Post;
    end;
    if (AdsGridSetting.RecordCount > RNo+z) And (Rno+z >= 0) then
      if RNo + Z = 0 then
        AdsGridSetting.First
      else
        AdsGridSetting.RecNo := RNo+z;
  finally
    AdsGridSetting.EnableControls;
  end;
end;

procedure TFmMain.AdqBlockBeforeInsert(DataSet: TDataSet);
begin
  if Not DmTables.CurrentUser.Blk then
  begin
    ShowError('مجاز به تعریف/ویرایش اطلاعات بلوک نمی باشید');
    Abort;
  end;
end;

procedure TFmMain.AdqBlockBeforePost(DataSet: TDataSet);
begin
  if AdqBlockID.AsInteger = 0 then
  begin
    ShowError('کد بلاک مشخص نشده است');
    Abort;
  end
end;

procedure TFmMain.AdqGrpUserBeforePost(DataSet: TDataSet);
begin
  if DmTables.CurrentUser.UsrMng then
    try
      DmTables.RunQuery('Exec PrcSetUserToGroup '+ADTGroupID.AsString+','+AdsGrpUserID.AsString);
      ADTGroupAfterScroll(ADTGroup);
    Except
      On E:Exception do
      begin
        ShowError('خطا در تخصیص دسترسی');
        Abort;
      end;
    end
    else
    begin
      ShowError('مجاز به تخصیص کاربر به گروه نمی باشید');
      AdsGrpUser.Cancel;
      Abort;
    end;
end;

procedure TFmMain.AdsChildsBeforePost(DataSet: TDataSet);
begin
  if Trim(AdsChildsName.AsString) = '' then
  begin
    ShowError('نام فرزند مشخص نشده است');
    Abort;
  end
  else if Trim(AdsChildsFamily.AsString) = '' then
  begin
    ShowError('نام خانوادگی فرزند مشخص نشده است');
    Abort;
  end
  else if Trim(AdsChildsNesbat.AsString) = '' then
  begin
    ShowError('نسبت مشخص نشده است');
    Abort;
  end
  else
  begin
    DmTables.RunQuery('Select Max(Radif) From TBSahamChild Where SahamID = '+IntToStr(AdsSahamRadif.AsInteger));
    if DmTables.ADOQuery1.RecordCount > 0 then
      AdsChildsRadif.AsInteger := DmTables.ADOQuery1.Fields[0].AsInteger+1;
    AdsChildsSahamID.AsInteger := AdsSahamRadif.AsInteger;
    AdsChildsUserID.AsInteger := DmTables.CurrentUser.id;
  end;
end;

procedure TFmMain.AdsGridSettingBeforePost(DataSet: TDataSet);
var
  I: Integer;
begin
  for I := 0 to DbgSaham.Columns.Count - 1 do
    if DbgSaham.Columns[i].Field.FieldName = AdsGridSettingEngName.AsString  then
      DbgSaham.Columns[i].Visible := AdsGridSettingActive.AsInteger = 1;
end;

procedure TFmMain.AdsSahamAfterScroll(DataSet: TDataSet);
begin
  if PNChilds.Visible then
  begin
    AdsChilds.Close;
    AdsChilds.Parameters[0].Value := AdsSahamRadif.AsInteger;
    AdsChilds.Open;
  end;
end;

procedure TFmMain.AdsSahamBeforeEdit(DataSet: TDataSet);
begin
  if Not DmTables.CurrentUser.Saham then
  begin
    ShowError('مجاز به ویرایش اطلاعات سهام نمی باشید');
    Abort;
  end;
end;

procedure TFmMain.AdsSahamBeforePost(DataSet: TDataSet);
Var
  Sql: TStrings;
  I: Integer;
  comma: String;
begin
  AdsSahamUserID.AsInteger := DmTables.CurrentUser.id;
  try
    Sql := TStringList.Create;
    for I := 0 to 107 do
    begin
      if Sql.Count > 0 then
        comma := ', ';
      if (AdsSaham.Fields[i] <> AdsSahamRadif) And (AdsSaham.Fields[i] <> AdsSahamEDate) And (AdsSaham.Fields[i].FieldKind = fkData) then
        if ((AdsSaham.Fields[i] <> AdsSahamBLockID) or (AdsSahamBLockID.AsInteger > 0)) And ((AdsSaham.Fields[i] <> AdsSahamParentID) or (AdsSahamParentID.AsInteger > 0)) then
          if (AdsSaham.Fields[i].DataType = ftString) or (AdsSaham.Fields[i].DataType = ftWideString) then
            Sql.Add(comma+AdsSaham.Fields[i].FieldName+' = '+QuotedStr(AdsSaham.Fields[i].AsString))
          Else
            Sql.Add(comma+AdsSaham.Fields[i].FieldName+' = '+FloatToStr(AdsSaham.Fields[i].AsFloat));
    end;

    Sql.Text := 'Update TBSaham Set '+Sql.Text+' Where Radif = '+IntToStr(AdsSahamRadif.AsInteger);
    DmTables.ExecQuery(Sql.Text);
    AdsSaham.Cancel;
    Tag := AdsSahamRadif.AsInteger;
    AdsSaham.Requery();
    AdsSaham.Locate('Radif', Tag, [loCaseInsensitive]);
    TForm(PnSaham.Parent).Close;
    Abort;
  Except
    On E:Exception do
    begin
      if Trim(LowerCase(e.Message)) <> 'operation aborted'  then
        ShowError('خطا در ذخیره اطلاعات'+e.Message);
      Abort;
    end;
  end;
end;

procedure TFmMain.AdsSahamCalcFields(DataSet: TDataSet);
begin
  AdsSahamJamKolBedehi.AsFloat := AdsSahamMBaravord.AsFloat+AdsSahamMHazineAmadesazi.AsFloat+AdsSahamMWater.AsFloat+AdsSahamMBargh.AsFloat+AdsSahamMFazelab.AsFloat+AdsSahamMKharidSahamEzafe.AsFloat+AdsSahamMBimeGozaresh.AsFloat+AdsSahamMMaliatTaklifiGozaresh.AsFloat+AdsSahamMArzeshAfzodeGozaresh.AsFloat+AdsSahamMBerozresani.AsFloat+AdsSahamMParvaneSakht.AsFloat+AdsSahamMKharidZamin.AsFloat;
  AdsSahamJamKolPardakhti.AsFloat := AdsSahamMMablaghVamMehr.AsFloat+AdsSahamMMablaghPardakhtiGhabli.AsFloat+AdsSahamMMablaghPardakhtiBadi.AsFloat+AdsSahamMPardakhtNahae.AsFloat;
end;

procedure TFmMain.AdsSahamStatusGetText(Sender: TField; var Text: string;
  DisplayText: Boolean);
begin
  if Sender.AsInteger=1 then
    Text := 'آپارتمان'
  Else if Sender.AsInteger=2 then
    Text := 'در حال ساخت'
  Else if Sender.AsInteger=3 then
    Text := 'انتخاب بلوک'
  Else if Sender.AsInteger=4 then
    Text := 'بدون زمین'
  Else if Sender.AsInteger=5 then
    Text := 'باطل شده'
  Else
    Text := '';
end;

procedure TFmMain.AdsSahamStatusSetText(Sender: TField; const Text: string);
begin
  if Text = 'آپارتمان' then
    Sender.AsInteger := 1
  Else if Text = 'در حال ساخت' then
    Sender.AsInteger := 2
  Else if Text = 'انتخاب بلوک' then
    Sender.AsInteger := 3
  Else if Text = 'بدون زمین' then
    Sender.AsInteger := 4
  Else if Text = 'باطل شده' then
    Sender.AsInteger := 5
  Else
    Sender.AsInteger := 0;
end;

procedure TFmMain.AdsUsersBeforeEdit(DataSet: TDataSet);
begin
  if Not DmTables.CurrentUser.UsrMng then
  begin
    ShowError('مجاز به ویرایش اطلاعات کاربران نمی باشید');
    Abort;
  end;
end;

procedure TFmMain.AdsUsersBeforePost(DataSet: TDataSet);
begin
  if AdsUsersID.AsInteger = 0 then
  Begin
    ShowError('کد کاربر مشخص نشده است');
    Abort
  End
  else if Trim(AdsUsersFamily.AsString) = '' then
  Begin
    ShowError('نام خانوادگی مشخص نشده است');
    Abort
  End
  else if Trim(AdsUsersUserName.AsString) = '' then
  Begin
    ShowError('نام کاربری مشخص نشده است');
    Abort
  End;
end;

procedure TFmMain.AdsUsersPassGetText(Sender: TField; var Text: string;
  DisplayText: Boolean);
begin
  Text := '********';
end;

procedure TFmMain.ADTGroupAfterScroll(DataSet: TDataSet);
begin
  try
    try
      AdsGrpUser.BeforePost := nil;
      if AdsGrpUser.Active then
        AdsGrpUser.EmptyDataSet
      else
        AdsGrpUser.CreateDataSet;
      DmTables.RunQuery('Select Case When B.UserID Is Not Null then 1 Else 0 End Chk,A.ID, A.Family, A.Mobile From  TBUsers A Left Outer Join TBUserGroup B On A.ID = B.UserID And B.GrpID = '+IntToStr(ADTGroupID.AsInteger));
      while Not DmTables.ADOQuery1.Eof do
      begin
        AdsGrpUser.Append;
        AdsGrpUserChk.AsInteger := DmTables.ADOQuery1.FindField('Chk').AsInteger;
        AdsGrpUserID.AsString := DmTables.ADOQuery1.FindField('ID').AsString;
        AdsGrpUserFamily.AsString := DmTables.ADOQuery1.FindField('Family').AsString;
        AdsGrpUserMobile.AsString := DmTables.ADOQuery1.FindField('Mobile').AsString;
        AdsGrpUser.Post;
        DmTables.ADOQuery1.Next;
      end;
    finally
      AdsGrpUser.BeforePost := AdqGrpUserBeforePost;
    end;
  Except
    ShowError('خطا در بارگذاری کاربران گروه');
  end;
end;

procedure TFmMain.ADTGroupBeforeInsert(DataSet: TDataSet);
begin
  if Not DmTables.CurrentUser.UsrMng  then
  begin
    ShowError('مجاز به ویرایش اطلاعات گروه های کاربری نمی باشید');
    Abort;
  end;
end;

procedure TFmMain.BnUpClick(Sender: TObject);
var
  i, j: Integer;
  F: TField;
begin
  try
    J := TSpeedButton(Sender).Tag;
    for I := 0 to DbgSaham.Columns.Count - 1 do
      if DbgSaham.Columns[i].Field.FieldName = AdsGridSettingEngName.AsString then
        if (i+j >= 0) And (DbgSaham.Columns.Count > i+j) then
        begin
          F := DbgSaham.Columns[i].Field;
          DbgSaham.Columns[i].Field := DbgSaham.Columns[i+j].Field;
          DbgSaham.Columns[i+j].Field := F;
          Abort;
        end;
  finally
    PrcLoadGridSetting(j);
  end;
end;

procedure TFmMain.EdCsID1Change(Sender: TObject);
var
  S: String;
begin
  if (Trim(EdCsID1.Text) <> '') And (Trim(EdCsID2.Text) <> '') then
    S := 'ID >= '+EdCsID1.Text+' And ID <= '+EdCsID2.Text
  Else if (Trim(EdCsID1.Text) <> '') And (Trim(EdCsID2.Text) = '') then
    S := 'ID = '+EdCsID1.Text
  Else if (Trim(EdCsID1.Text) = '') And (Trim(EdCsID2.Text) <> '') then
    S := 'ID = '+EdCsID2.Text;

  if EDCsName.Text <> '' then
    if Trim(S) <> '' then
      S := S + ' And Name Like '+QuotedStr('%'+EDCsName.Text+'%')
    Else
      S := 'Name Like '+QuotedStr('%'+EDCsName.Text+'%');

  if EDCsFamily.Text <> '' then
    if Trim(S) <> '' then
      S := S + ' And Family Like '+QuotedStr('%'+EDCsFamily.Text+'%')
    Else
      S := 'Family Like '+QuotedStr('%'+EDCsFamily.Text+'%');

  if EDCsWife.Text <> '' then
    if Trim(S) <> '' then
      S := S + ' And (Wife Like '+QuotedStr('%'+EDCsWife.Text+'%')+' or WFamily Like '+QuotedStr('%'+EDCsWife.Text+'%')+')'
    Else
      S := '(Wife Like '+QuotedStr('%'+EDCsWife.Text+'%')+' or WFamily Like '+QuotedStr('%'+EDCsWife.Text+'%')+')';

  if EDCsNationalID.Text <> '' then
    if Trim(S) <> '' then
      S := S + ' And NationalID Like '+QuotedStr('%'+EDCsNationalID.Text+'%')
    Else
      S := 'NationalID Like '+QuotedStr('%'+EDCsNationalID.Text+'%');

  if EDCsBlock.Text <> '' then
    if Trim(S) <> '' then
      S := S + ' And BLockID = '+EDCsBlock.Text
    Else
      S := 'BLockID = '+EDCsBlock.Text;

  if EDCsVahed.Text <> '' then
    if Trim(S) <> '' then
      S := S + ' And Vahed = '+EDCsVahed.Text
    Else
      S := 'Vahed = '+EDCsVahed.Text;

  if EDCsTabaghe.Text <> '' then
    if Trim(S) <> '' then
      S := S + ' And Tabaghe = '+EDCsTabaghe.Text
    Else
      S := 'Tabaghe = '+EDCsTabaghe.Text;

  if EDCsMobile.Text <> '' then
    if Trim(S) <> '' then
      S := S + ' And Mobile Like '+QuotedStr('%'+EDCsMobile.Text+'%')
    Else
      S := 'Mobile Like '+QuotedStr('%'+EDCsMobile.Text+'%');

  AdsSaham.Filter := S;
  AdsSaham.Filtered := Trim(S) <> '';
end;

procedure TFmMain.Edit1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  ID, Status, blockid, tabaghe, vahed, TabagheCount, Metraj, AdamMarghobiat: Integer;
  Mostajer, MosTel, mosdate, tip, faz: String;
begin
  if key=13 then
    if Not isnumeric(TEdit(Sender).Text) then
      ShowMessage('کد سهام وارد شده صحیح نمی باشد')
    Else
      if AdsSaham.Locate('ID', TEdit(Sender).Text, [loCaseInsensitive]) then
        try
          Tag := AdsSahamRadif.AsInteger;
          ID := AdsSahamID.AsInteger;
          Status := AdsSahamStatus.AsInteger;
          blockid := AdsSahamBLockID.AsInteger;
          tabaghe := AdsSahamTabaghe.AsInteger;
          vahed := AdsSahamVahed.AsInteger;
          TabagheCount := AdsSahamTabagheCount.AsInteger;
          Metraj := AdsSahamMetraj.AsInteger;
          AdamMarghobiat := AdsSahamAdamMarghobiat.AsInteger;
          Mostajer := AdsSahamMostajer.AsString;
          MosTel := AdsSahamMosTel.AsString;
          mosdate := AdsSahamMostajerEdate.AsString;
          tip := AdsSahamTip.AsString;
          faz := AdsSahamFaz.AsString;

          LBSahamTitle.Caption := AdsSahamID.AsString+' - '+AdsSahamName.AsString+' '+AdsSahamFamily.AsString+' - بلوک '+IntToStr(AdsSahamBLockID.AsInteger)+' طبقه '+IntToStr(AdsSahamTabaghe.AsInteger)+' واحد '+IntToStr(AdsSahamVahed.AsInteger)+' - '+AdsSahamStatus.AsString;

          DmTables.RunQuery('Select Max(Radif) From TBSaham');

          AdsSaham.Append;

          AdsSahamRadif.AsInteger := DmTables.ADOQuery1.Fields[0].AsInteger+1;

          AdsSahamParentID.AsInteger := Tag;
          AdsSahamID.AsInteger := ID;
          AdsSahamStatus.AsInteger := Status;
          AdsSahamBLockID.AsInteger := blockid;
          AdsSahamTabaghe.AsInteger := tabaghe;
          AdsSahamVahed.AsInteger := vahed;
          AdsSahamTabagheCount.AsInteger := TabagheCount;
          AdsSahamMetraj.AsInteger := Metraj;
          AdsSahamAdamMarghobiat.AsInteger := AdamMarghobiat;
          AdsSahamMostajer.AsString := Mostajer;
          AdsSahamMosTel.AsString := MosTel;
          AdsSahamMostajerEdate.AsString := mosdate;
          AdsSahamTip.AsString := tip;
          AdsSahamFaz.AsString := faz;

          EDBirthDate.Text   := '';
          EDMosEDate.Text    := '';
          EDWBirthDate.Text  := '';
          EDSahamMvDate.Text := GetDate();

          BtnNtbSaham.Down := True;
          NtbSaham.PageIndex := 0;
          BtnMoveMentSec.Visible := AdsSahamParentID.AsInteger > 0;
          SepratePanel(Self, PNSaham, '', PNSaham.Width+5, alClient);
          TForm(PnGetSahamID.Parent).Close;
        finally

        end;
end;

procedure TFmMain.EDMosEDateChange(Sender: TObject);
begin
  if Not(AdsSaham.State in [dsEdit,dsInsert]) then
    AdsSaham.Edit;
end;

procedure TFmMain.EDSearchChange(Sender: TObject);
//var
//  Str: String;
//  I: Integer;
begin
//  if Trim(EDSearch.Text) <> '' then
//    for I := 0 to DbgSaham.Columns.Count - 1 do
//      if DbgSaham.Columns[i].Visible then
//        if (DbgSaham.Columns[i].Field.DataType = ftString) or (DbgSaham.Columns[i].Field.DataType = ftWideString) then
//          if Trim(Str) = '' then
//            Str := DbgSaham.Columns[i].Field.FieldName+' like '+QuotedStr('%'+StringReplace(EDSearch.Text, ' ', '%', [rfReplaceAll])+'%')
//          Else
//            Str := Str + ' OR ' + DbgSaham.Columns[i].Field.FieldName+' like '+QuotedStr('%'+StringReplace(EDSearch.Text, ' ', '%', [rfReplaceAll])+'%')
//        Else if IsNumeric(EDSearch.Text) And ((DbgSaham.Columns[i].Field.DataType = ftFloat) or (DbgSaham.Columns[i].Field.DataType = ftBCD)) then
//          if Trim(Str) = '' then
//            Str := DbgSaham.Columns[i].Field.FieldName+' = '+EDSearch.Text
//          Else
//            Str := Str + ' OR ' + DbgSaham.Columns[i].Field.FieldName+' = '+EDSearch.Text;
//  AdsSaham.Filter := Str;
//  AdsSaham.Filtered := Trim(Str) <> '';
end;

procedure TFmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  DbgSaham.Columns.SaveToFile('GridFields'+IntToStr(DmTables.CurrentUser.id)+'.Config');
end;

procedure TFmMain.FormShow(Sender: TObject);
begin
  Notebook1.PageIndex := 0;
  AdsSaham.Open;
  if FileExists('GridFields'+IntToStr(DmTables.CurrentUser.id)+'.Config') then
    DbgSaham.Columns.LoadFromFile('GridFields'+IntToStr(DmTables.CurrentUser.id)+'.Config');
end;

procedure TFmMain.frxReport1BeforePrint(Sender: TfrxReportComponent);
begin
  DmTables.RunQuery('Select Pic From TBBlock Where ID = '+IntToStr(AdsSahamBLockID.AsInteger)+' And Pic Is Not Null');
  if frxReport1.FindObject('ImgBlock') <> Nil then
  try
    if Not DmTables.ADOQuery1.IsEmpty then
    TfrxPictureView(frxReport1.FindObject('ImgBlock')).Picture.Assign(GetImageFromBlob(TBlobField(DmTables.ADOQuery1.Fields[0])));
  except
  end;
//    TfrxPictureView(frxReport1.FindObject('ImgBlock')).Picture.LoadFromFile(IncludeTrailingBackslash(ExtractFilePath(Application.ExeName))+'Chart.Bmp');
end;

procedure TFmMain.Label11MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  ReleaseCapture;
  SendMessage(TForm(PnPayment.Parent).Handle, WM_SYSCOMMAND, 61458, 0) ;
end;

procedure TFmMain.LBSahamTitleMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  ReleaseCapture;
  SendMessage(TForm(PNSaham.Parent).Handle, WM_SYSCOMMAND, 61458, 0) ;
end;

procedure TFmMain.Label1DblClick(Sender: TObject);
begin
  if DmTables.CurrentUser.id=1 then
    DesignRep := True;
end;

procedure TFmMain.Label1MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  ReleaseCapture;
  SendMessage(Self.Handle, WM_SYSCOMMAND, 61458, 0) ;
end;

procedure TFmMain.Label2MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  ReleaseCapture;
  SendMessage(Tform(PnGridSetting.Parent).Handle, WM_SYSCOMMAND, 61458, 0) ;
end;

procedure TFmMain.Label70MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  ReleaseCapture;
  SendMessage(Tform(PNExcel.Parent).Handle, WM_SYSCOMMAND, 61458, 0) ;
end;

procedure TFmMain.N3Click(Sender: TObject);
begin
  Notebook1.PageIndex := TMenuItem(Sender).Tag;
end;

procedure TFmMain.NRepClick(Sender: TObject);
begin
  frxReport1.LoadFromFile('report'+IntToStr(TMenuItem(Sender).Tag)+'.fr3');
  if DesignRep then
    frxReport1.DesignReport
  Else
  try
    AdsSaham.DisableControls;
    frxReport1.PrepareReport();
    frxReport1.ShowReport();
  finally
    AdsSaham.EnableControls;
    AdsSaham.First;
  end;
  DesignRep := False;
end;

procedure TFmMain.Notebook1PageChanged(Sender: TObject);
begin
  PNSahamTool.Visible := Notebook1.PageIndex = 0;
  PnBlockTool.Visible := Notebook1.PageIndex = 1;
  PNUserTool.Visible  := Notebook1.PageIndex = 2;
  if Notebook1.PageIndex = 0 then
    if Not AdsSaham.Active then
      AdsSaham.Open
    else
      AdsSaham.Requery()
  Else if Notebook1.PageIndex = 2 then
    if Not AdsUsers.Active then
      AdsUsers.Open
    else
      AdsUsers.Requery()
  Else if Notebook1.PageIndex=3 then
    if Not ADTGroup.Active then
      ADTGroup.Open
    else
      ADTGroup.Requery()
  Else if Notebook1.PageIndex=1 then
    if Not AdqBlock.Active then
      AdqBlock.Open
    else
      AdqBlock.Requery()
end;

procedure TFmMain.SpeedButton10Click(Sender: TObject);
begin
  if AdsSaham.State in [dsEdit,dsInsert] then
    AdsSaham.Post;
  TForm(PnPayment.Parent).Close;
end;

procedure TFmMain.SpeedButton11Click(Sender: TObject);
begin
  PNChilds.Visible := Not PNChilds.Visible;
  if PNChilds.Visible then
    AdsSahamAfterScroll(AdsSaham);
end;

procedure TFmMain.SpeedButton12Click(Sender: TObject);
begin
  if not AdsChilds.IsEmpty then
    if MessageDlg('آیا مایل به حذف '+AdsChildsName.AsString+' '+AdsChildsFamily.AsString+' می باشید؟', mtConfirmation, [mbYes,mbNo], 0) = mrYes then
    try
      DmTables.RunQuery('Delete TBSahamChild OutPut ''Success'' Where SID = '+IntToStr(AdsSahamID.AsInteger)+' And SRadif = '+IntToStr(AdsSahamRadif.AsInteger)+' And Radif = '+IntToStr(AdsChildsRadif.AsInteger));
      AdsChilds.Requery();
    Except
      ShowError('خطا در حذف فرزند. لطفا مجددا سعی نمایید');
    end;
end;

procedure TFmMain.BtnNtbSahamClick(Sender: TObject);
begin
  NtbSaham.PageIndex := TSpeedButton(Sender).Tag;
end;

procedure TFmMain.SpeedButton17Click(Sender: TObject);
var
  strConn: widestring;
  I: Integer;
begin
  if DmTables.CurrentUser.ID <> 1 then
    ShowError('مجاز به دسترسی به این ابزار نمی باشید')
  Else if FileOpenDialog1.Execute then
    begin
      strConn := 'Provider=Microsoft.ACE.OLEDB.12.0;' + 'Data Source=' + FileOpenDialog1.FileName + ';' + 'Extended Properties=Excel 12.0;';
      try
        DmTables.AdoExcelConnection.Connected := false;
        DmTables.AdoExcelConnection.ConnectionString := strConn;
        try
          DmTables.AdoExcelConnection.Open;
          DmTables.AdqExcel.Sql.Text := 'Select * From [Sheet1$]';
          DmTables.AdqExcel.Open;

          CMBExelID.Items.Clear;
          CMBExcelOthField.Items.Clear;
          CMBExcelMyField.Items.Clear;
          CMBExcelMyFieldEng.Items.Clear;
          for I := 0 to DmTables.AdqExcel.Fields.Count - 1 do
          begin
            CMBExelID.Items.Add(DmTables.AdqExcel.Fields[i].FieldName);
            CMBExcelOthField.Items.Add(DmTables.AdqExcel.Fields[i].FieldName);
          end;
          for I := 0 to AdsSaham.Fields.Count - 1 do
            if (AdsSaham.Fields[i] <> AdsSahamRadif) And (AdsSaham.Fields[i] <> AdsSahamUserID) And (AdsSaham.Fields[i] <> AdsSahamEDate) then
            begin
              CMBExcelMyField.Items.Add(AdsSaham.Fields[i].DisplayLabel);
              CMBExcelMyFieldEng.Items.Add(AdsSaham.Fields[i].FieldName);
            end;
          CMBExelID.Text := '';
          CMBExcelOthField.Text := '';
          CMBExcelMyField.Text := '';
          CMBExcelMyFieldEng.Text := '';
          SepratePanel(Self, PnExcel, '', PnExcel.Width + 3, alClient);
        except
          On E: Exception do
            ShowError('خطا در اتصال به فایل اکسل');
        end;
      finally
        if DmTables.ADOExcelConnection.Connected then
          DmTables.ADOExcelConnection.Close;
      end;
    end;
end;

procedure TFmMain.SpeedButton18Click(Sender: TObject);
begin
  if AdsSaham.State in [dsEdit,dsInsert] then
    AdsSaham.Cancel;
  TForm(PNSaham.Parent).Close;
end;

procedure TFmMain.SpeedButton19Click(Sender: TObject);
begin
  if AdsSahamID.AsInteger = 0  then
  begin
    ShowError('شماره سهام مشخص نشده است');
    Abort;
  end;
  if (Trim(AdsSahamName.AsString) = '') or (Trim(AdsSahamFamily.AsString) = '') then
  begin
    ShowError('نام و نام خانوادگی سهامدار مشخص نشده است');
    Abort;
  end;
  if AdsSahamStatus.AsInteger = 0  then
  begin
    ShowError('وضعیت سهام مشخص نشده است');
    Abort;
  end;
  if AdsSahamParentID.AsInteger > 0 then
    if AdsSahamMvFrm.AsInteger = 0 then
    begin
      ShowError('شماره فرم نقل و انتقال مشخص نشده است');
      NtbSaham.PageIndex := 4;
      Abort;
    end
    else if Trim(EDSahamMvDate.Text) = '' then
    begin
      ShowError('تاریخ نقل و انتقال مشخص نشده است');
      NtbSaham.PageIndex := 4;
      Abort;
    end;

  if AdsSaham.State in [dsEdit,dsInsert] then
  begin
    AdsSahamMvDate.AsString := EDSahamMvDate.Text;
    AdsSahamBirthDate.AsString := EDBirthDate.Text;
    AdsSahamMostajerEdate.AsString := EDMosEDate.Text;
    AdsSahamWBirthDate.AsString := EDWBirthDate.Text;
    AdsSaham.Post;
  end;
  TForm(PNSaham.Parent).Close;
end;

procedure TFmMain.SpeedButton1Click(Sender: TObject);
begin
  if Notebook1.PageIndex > 0 then
    Notebook1.PageIndex := 0
  else
    Application.Terminate;
end;

procedure TFmMain.SpeedButton20Click(Sender: TObject);
begin
  TForm(PNExcel.Parent).Close;
end;

procedure TFmMain.SpeedButton21Click(Sender: TObject);
begin
  SepratePanel(Self, PnGetSahamID, '', PnGetSahamID.Width+5, alClient)
end;

procedure TFmMain.SpeedButton22Click(Sender: TObject);
begin
  if SaveDialog1.Execute then
    ExportRecordsetToMSExcel(SaveDialog1.FileName, AdsSaham.Recordset);
end;

procedure TFmMain.SpeedButton23Click(Sender: TObject);
var
  jp:TJpegimage;
begin
  if OpenPictureDialog1.Execute then
  try
    jp:=TJpegimage.Create;
    jp.LoadFromFile(OpenPictureDialog1.FileName);
    DmTables.AdqBlockPic.Close;
    DmTables.AdqBlockPic.Parameters[0].Assign(jp);
    DmTables.AdqBlockPic.Parameters[1].Value := AdqBlockID.AsInteger;
    DmTables.AdqBlockPic.ExecSQL;
  finally
    jp.Free;
  end;
end;

procedure TFmMain.SpeedButton24Click(Sender: TObject);
begin
  Self.WindowState := wsMinimized
end;

procedure TFmMain.SpeedButton2Click(Sender: TObject);
begin
  if Self.WindowState = wsMaximized then
    Self.WindowState := wsNormal
  else
    Self.WindowState := wsMaximized
end;

 procedure TFmMain.SpeedButton3Click(Sender: TObject);
begin
  if Trim(EDPass1.Text) <> Trim(EDPass1.Text) then
    ShowError('تکرار رمز عبور صحیح نمی باشد')
  Else if Length(Trim(EDPass1.Text)) < 6 then
    ShowError('طول رمز عبور می بایست حداقل 6 کاراکتر باشد')
  Else
  try
    DmTables.RunQuery('Exec PrcSetPass '+IntToStr(DmTables.CurrentUser.id)+','+AdsUsersID.AsString+','+QuotedStr(EncryptStr(EDPass1.Text)));
    if DmTables.ADOQuery1.Fields[0].AsString <> 'Success' then
      ShowError(DmTables.ADOQuery1.Fields[0].AsString)
    Else
      TForm(PNPass.Parent).Close;
  Except
    On E:Exception do
      ShowError('خطا در ذخیره رمز عبور');
  end;
end;

procedure TFmMain.BtnDoExcelExportClick(Sender: TObject);
var
  Sql: TStrings;
begin
  if DmTables.AdqExcel.FindField(CMBExelID.Text) = nil then
    ShowMessage('فیلد شماره سهام فایل اکسل قابل شناسایی نمی باشد')
  else if DmTables.AdqExcel.FindField(CMBExcelOthField.Text) = nil then
    ShowMessage('فیلد فایل اکسل قابل شناسایی نمی باشد')
  Else if MessageDlg('آیا از ادامه عملیات اطمینان دارید؟ (بروز رسانی غیرقابل بازگشت می باشید)', mtConfirmation, [mbYes,mbNo], 0) = mrYes then
  try
    PNExcelWaiting.Visible := True;
    Application.ProcessMessages;
    Sql := TStringList.Create;
    CMBExcelMyFieldEng.ItemIndex := CMBExcelMyField.ItemIndex;
    with DmTables do
    try
      ExecQuery('Backup DataBase Saham To Disk = ''Saham('+GetDate('', '-')+').Bak''');
      DmTables.ADOConnection1.BeginTrans;
      AdqExcel.First;
      while Not AdqExcel.Eof do
        if isnumeric(AdqExcel.FindField(CMBExelID.Text).AsString) then
        begin
          Sql.Clear;
          Sql.Add('Declare @Radif Int');
          Sql.Add('Select @Radif = A.Radif');
          Sql.Add('From TBSaham A');
          Sql.Add(' Left Outer Join TBSaham B On A.Radif = B.ParentID');
          Sql.Add('Where B.Radif Is Null And A.ID = '+AdqExcel.FindField(CMBExelID.Text).AsString);
          Sql.Add('Update TBSaham');
          Sql.Add('Set '+CMBExcelMyFieldEng.Text+' = '+QuotedStr(AdqExcel.FindField(CMBExcelOthField.Text).AsString));
          Sql.Add('Where Radif = @Radif');
          ExecQuery(Sql.Text);
          AdqExcel.Next;
        end
        else
        begin
          ShowMessage('شماره سهام عددی نمی باشد '+AdqExcel.FindField(CMBExelID.Text).AsString);
          Abort;
        end;
      if DmTables.ADOConnection1.InTransaction then
        DmTables.ADOConnection1.CommitTrans;
      AdsSaham.Requery();
      ShowMessage('بارگذاری اطلاعات از فایل اکسل با موفقیت انجام گردید');
    Except
      On E:Exception do
      begin
        if DmTables.ADOConnection1.InTransaction then
          DmTables.ADOConnection1.RollbackTrans;
        ShowError('خطا در بروز رسانی بانک اطلاعاتی '+#13#10+e.ToString);
      end;
    end;
  finally
    PNExcelWaiting.Visible := False;
    if DmTables.ADOConnection1.InTransaction then
      DmTables.ADOConnection1.CommitTrans;
  end;
end;

procedure TFmMain.BtnSahamSettingClick(Sender: TObject);
var
  I: Integer;
begin
  PrcLoadGridSetting(0);
  SepratePanel(Self, PnGridSetting, '', PnGridSetting.Width+10, alClient);
  DbgSaham.Columns.SaveToFile('GridFields'+IntToStr(DmTables.CurrentUser.id)+'.Config');
end;

procedure TFmMain.BtnSetPassClick(Sender: TObject);
begin
  if (DmTables.CurrentUser.id <> AdsUsersID.AsInteger) And Not DmTables.CurrentUser.UsrMng then
  begin
    ShowError('مجاز به ویرایش اطلاعات کاربران نمی باشید');
    Abort;
  end;
  SepratePanel(Self, PNPass, '', PNPass.Width+5, alClient);
end;

procedure TFmMain.DbgSahamDblClick(Sender: TObject);
//var
//  g:TGraphic;
//  b: TBitmap;
begin
  if Not DmTables.CurrentUser.Saham then
  begin
    ShowError('مجاز به ویرایش اطلاعات سهام نمی باشید');
    Abort;
  end;
  DmTables.RunQuery('Select Pic From TBSaham Where ID = '+IntToStr(AdsSahamID.AsInteger)+' And Radif = '+IntToStr(AdsSahamRadif.AsInteger)+' And Pic Is Not Null');
  if Not DmTables.ADOQuery1.IsEmpty then
    try
      try
    //    g:=TJpegimage.Create;
    //    b := TBitmap.Create;
    //    b.Assign(DmTables.ADOQuery1.Fields[0]);
    //    g.Assign(DmTables.ADOQuery1.Fields[0]);
        ImgSaham.Picture.Assign(GetImageFromBlob(TBlobField(DmTables.ADOQuery1.Fields[0])));
      finally
    //    b.Free;
      end
    Except
      ImgSaham.Picture.Assign(Nil);
    end
  else
    ImgSaham.Picture.Assign(Nil);
  EDBirthDate.Text   := AdsSahamBirthDate.AsString;
  EDMosEDate.Text    := AdsSahamMostajerEdate.AsString;
  EDWBirthDate.Text  := AdsSahamWBirthDate.AsString;
  EDSahamMvDate.Text := AdsSahamMvDate.AsString;
  BtnNtbSaham.Down := True;
  NtbSaham.PageIndex := 0;
  BtnMoveMentSec.Visible := AdsSahamParentID.AsInteger > 0;
  LBSahamTitle.Caption := 'اطلاعات سهام';
  SepratePanel(Self, PNSaham, '', PnSaham.Width+5, alClient)
end;

procedure TFmMain.DbgSahamTitleClick(Column: TColumn);
begin
  if AdsSaham.Sort = Column.FieldName then
    AdsSaham.Sort := Column.FieldName+' DESC'
  Else
    AdsSaham.Sort := Column.FieldName;
end;

procedure TFmMain.DBImage1DblClick(Sender: TObject);
var
  img: TImage;
  S : TMemoryStream;
  jp:TJpegimage;
begin
  if OpenPictureDialog1.Execute then
  try
    jp:=TJpegimage.Create;
    jp.LoadFromFile(OpenPictureDialog1.FileName);
    DmTables.AdqSahamPic.Close;
    DmTables.AdqSahamPic.Parameters[0].Assign(jp);
    DmTables.AdqSahamPic.Parameters[1].Value := AdsSahamID.AsInteger;
    DmTables.AdqSahamPic.Parameters[2].Value := AdsSahamRadif.AsInteger;
    DmTables.AdqSahamPic.ExecSQL;
    ImgSaham.Picture.LoadFromFile(OpenPictureDialog1.FileName);
//    img := TImage.Create(self);
//    img.Picture.LoadFromFile(OpenPictureDialog1.FileName);
//    S := TMemoryStream.Create;
//    Img.Picture.Bitmap.SaveToStream(S);
//    S.Position := 0;
//    AdsSaham.Edit;
//    AdsSahamPic.LoadFromStream(S);
//    AdsSaham.Post
  finally
    jp.Free;
  end;
end;

procedure TFmMain.SpeedButton4Click(Sender: TObject);
begin
  TForm(PnGridSetting.Parent).Close;
end;

procedure TFmMain.SpeedButton5Click(Sender: TObject);
begin
  PopupMenu1.Popup(Mouse.CursorPos.X, Mouse.CursorPos.Y);
end;

procedure TFmMain.SpeedButton6Click(Sender: TObject);
begin
  TForm(PNPass.Parent).Close;
end;

procedure TFmMain.SpeedButton7Click(Sender: TObject);
begin
  if ADTGroup.State in [dsEdit,dsInsert] then
    ADTGroup.Post;
end;

procedure TFmMain.SpeedButton8Click(Sender: TObject);
begin
  if Not DmTables.CurrentUser.MPY then
    ShowError('مجاز به ثبت پرداختی سهامداران نمی باشید')
  Else
    SepratePanel(Self, PnPayment, '', PnPayment.Width+5, alClient)
end;

procedure TFmMain.SpeedButton9Click(Sender: TObject);
begin
  if AdsSaham.State in [dsEdit,dsInsert] then
    AdsSaham.Cancel;
  TForm(PnPayment.Parent).Close;
end;



end.
