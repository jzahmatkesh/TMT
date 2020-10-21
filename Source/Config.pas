unit Config;

interface

uses Dialogs, SysUtils, ExtCtrls, Forms, Controls, Classes, ADODB, ComObj;

const CKEY1 = 53761;
      CKEY2 = 32618;

type
  TSepratorHolder = class(TWinControl)
  private
    MainForm: TObject;
    ParentHolder: TWinControl;
    Panel: TObject;
    Align: TAlign;
    Width: Integer;
    Height: Integer;
    Visible: Boolean;
  public
  end;

  TEventHandlers = class(TWinControl)
  Public
    Procedure FormClose(Sender: TObject; var Action: TCloseAction);
  end;

Procedure ShowError(Error: String);
function EncryptStr(const S :WideString): String;
function DecryptStr(const S: String): String;
Function SepratePanel(Self: TObject; Sender: TWinControl; cptn: string; Wdth: Integer; Algn: TAlign)  : TModalResult;
function isnumeric(const S: string): Boolean;
Function GetDate(dt: string = ''; DtSeprator: string = '/'): String;
procedure ExportRecordsetToMSExcel(DestName: string; Data: _Recordset);


implementation

var
  OldParent: TWinControl;

Procedure TEventHandlers.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if TSepratorHolder(TForm(Sender).Components[0]).Action <> nil then
    TSepratorHolder(TForm(Sender).Components[0]).Action.Execute;
  TWinControl(TSepratorHolder(TForm(Sender).Components[0]).Panel).Align :=
    TSepratorHolder(TForm(Sender).Components[0]).Align;
  TWinControl(TSepratorHolder(TForm(Sender).Components[0]).Panel).parent :=
    TSepratorHolder(TForm(Sender).Components[0]).ParentHolder;
  TWinControl(TSepratorHolder(TForm(Sender).Components[0]).Panel).Width :=
    TSepratorHolder(TForm(Sender).Components[0]).Width;
  TWinControl(TSepratorHolder(TForm(Sender).Components[0]).Panel).Height :=
    TSepratorHolder(TForm(Sender).Components[0]).Height;
  TWinControl(TSepratorHolder(TForm(Sender).Components[0]).Panel).Visible :=
    TSepratorHolder(TForm(Sender).Components[0]).Visible;
  Action := caFree;
end;

Procedure ShowError(Error: String);
Begin
  MessageDlg(Error, mtError, [mbok], 0);
End;

function EncryptStr(const S :WideString): String;
var   i          :Integer;
      RStr       :RawByteString;
      RStrB      :TBytes Absolute RStr;
      Key: Word;
begin
  Key := 223;
  Result:= '';
  RStr:= UTF8Encode(S);
  for i := 0 to Length(RStr)-1 do begin
    RStrB[i] := RStrB[i] xor (Key shr 8);
    Key := (RStrB[i] + Key) * CKEY1 + CKEY2;
  end;
  for i := 0 to Length(RStr)-1 do begin
    Result:= Result + IntToHex(RStrB[i], 2);
  end;
end;

function DecryptStr(const S: String): String;
var   i, tmpKey  :Integer;
      RStr       :RawByteString;
      RStrB      :TBytes Absolute RStr;
      tmpStr     :string;
      Key: Word;
begin
  key := 223;
  tmpStr:= UpperCase(S);
  SetLength(RStr, Length(tmpStr) div 2);
  i:= 1;
  try
    while (i < Length(tmpStr)) do begin
      RStrB[i div 2]:= StrToInt('$' + tmpStr[i] + tmpStr[i+1]);
      Inc(i, 2);
    end;
  except
    Result:= '';
    Exit;
  end;
  for i := 0 to Length(RStr)-1 do begin
    tmpKey:= RStrB[i];
    RStrB[i] := RStrB[i] xor (Key shr 8);
    Key := (tmpKey + Key) * CKEY1 + CKEY2;
  end;
  Result:= UTF8Decode(RStr);
end;

Function SepratePanel(Self: TObject; Sender: TWinControl; cptn: string; Wdth: Integer; Algn: TAlign)  : TModalResult;
var
  WinForm: TForm;
  Holder: TSepratorHolder;
  ev: TEventHandlers;
  I: Integer;
begin
  if Sender.parent.ClassType = TForm then
    TForm(Sender.parent).Close
  else
  begin
    WinForm := TForm.Create(TComponent(Self));
    with WinForm do
    begin
      BiDiMode := bdRightToLeft;
      Position := poScreenCenter;
      KeyPreview := True;
      BorderStyle := bsNone;
      caption := cptn;
      Font := TForm(Self).Font;
      Height := Sender.Height + 30;
      Width := Wdth;
      Tag := 500;
      // FormStyle   := fsStayOnTop;
      OnClose := ev.FormClose;
    end;
    Holder := TSepratorHolder.Create(WinForm);
    Holder.parent := WinForm;
    // if Application.FindComponent(TWinControl(Self).Name) <> nil then
    if Self <> nil then
    begin
      Holder.MainForm := Self;
      Holder.ParentHolder := Sender.parent;
      Holder.Panel := Sender;
      Holder.Align := Sender.Align;
      Holder.Width := Sender.Width;
      Holder.Height := Sender.Height;
      Holder.Visible := Sender.Visible;
      Sender.parent := WinForm;
      Sender.Visible := True;
      with Sender do
      begin
        Align := Algn;
        Left := 0;
        Top := 0;
      end;
      TForm(Self).AlphaBlendValue := 200;
      TForm(Self).AlphaBlend := True;
      Result := WinForm.ShowModal;
      TForm(Self).AlphaBlendValue := 255;
      TForm(Self).AlphaBlend := False;
      FreeAndNil(WinForm);
    end;
  end;
end;

function isnumeric(const S: string): Boolean;
var
  P: PChar;
  I: SmallInt;
begin
  I := 0;
  if trim(S) = '' then
  begin
    Result := False;
    Exit;
  end;
  P := PChar(S);
  Result := False;
  while P^ <> #0 do
  begin
    if not((P^ in ['0' .. '9']) or (P^ = '.') or ((I = 0) and (P^ = '-')))
      then
      Exit; // or (P^ = '-')
    I := I + 1;
    Inc(P);
  end;
  Result := True;
end;

Function GetDate(dt: string = ''; DtSeprator: string = '/'): String;
const
  count_days: array [1 .. 12] of Byte =
    (31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);
var
  I: Byte;
  st: String;
  day_year: Integer;
  Year, Month, Day: Word;
begin
  if dt = '' then
    dt := DateToStr(Now);
  DecodeDate(StrToDate(dt), Year, Month, Day);
  day_year := 0;
  for I := 1 to Month - 1 do
    day_year := day_year + count_days[I];
  day_year := day_year + Day;

  if IsLeapYear(Year) and (Month > 2) then
    Inc(day_year);

  if (day_year <= 79) then
  begin
    if ((Year - 1) mod 4 = 0) then
      day_year := day_year + 11
    else
      day_year := day_year + 10;

    Year := Year - 622;

    if (day_year mod 30 = 0) then
    begin
      Month := (day_year div 30) + 9;
      Day := 30;
    end
    else
    begin
      Month := (day_year div 30) + 10;
      Day := day_year mod 30;
    end;
  end
  else
  begin
    Year := Year - 621;

    day_year := day_year - 79;
    if (day_year <= 186) then
    begin
      if (day_year mod 31 = 0) then
      begin
        Month := (day_year div 31);
        Day := 31;
      end
      else
      begin
        Month := (day_year div 31) + 1;
        Day := day_year mod 31;
      end;
    end
    else
    begin
      day_year := day_year - 186;
      if (day_year mod 30 = 0) then
      begin
        Month := (day_year div 30) + 6;
        Day := 30;
      end
      else
      begin
        Month := (day_year div 30) + 7;
        Day := day_year mod 30;
      end;
    end;
  end; // else  .

  st := IntToStr(Year) + DtSeprator;
  if (Month < 10) then
    st := st + '0';
  st := st + IntToStr(Month) + DtSeprator;
  if (Day < 10) then
    st := st + '0';
  st := st + IntToStr(Day);

  Result := st;
end;

procedure ExportRecordsetToMSExcel(DestName: string; Data: _Recordset);
var
  ovExcelApp: OleVariant;
  ovExcelWorkbook: OleVariant;
  ovWS: OleVariant;
  ovRange: OleVariant;
begin
  ovExcelApp := CreateOleObject('Excel.Application'); //If Excel isnt installed will raise an exception
  try
    ovExcelWorkbook   := ovExcelApp.WorkBooks.Add;
    ovWS := ovExcelWorkbook.Worksheets.Item[1]; // go to first worksheet
    ovWS.Activate;
    ovWS.Select;
    ovRange := ovWS.Range['A1', 'A1']; //go to first cell
    ovRange.Resize[Data.RecordCount, Data.Fields.Count];
    ovRange.CopyFromRecordset(Data, Data.RecordCount, Data.Fields.Count); //this copy the entire recordset to the selected range in excel
    ovWS.SaveAs(DestName, 1, '', '', False, False);
  finally
    ovExcelWorkbook.Close();
//    ovWS := ;
//    ovExcelWorkbook := Unassigned;
//    ovExcelApp := Unassigned;
  end;
end;

end.
