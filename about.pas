unit about;
{ About Window }


interface

uses Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls,
  Buttons, ExtCtrls, ShellAPI, StrUtils;

type
  TAboutBox = class(TForm)
    Panel1: TPanel;
    Timer1: TTimer;
    Panel3: TPanel;
    IconImage: TImage;
    Label1: TLabel;
    Label2: TLabel;
    VersionInfo: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    OKButton: TButton;
    Panel2: TPanel;
    Image1: TImage;
    TextField: TMemo;
    procedure Label5Click(Sender: TObject);
    procedure Label5MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure FormCreate(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure OKButtonClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Panel3MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
  private
    // FFileName, VerNum : string;
    // VersionStrings : TStringList;
    opening : boolean;
    function Sto_GetFmtFileVersion(const FileName: String;
       const Fmt: String = 'Version: %d.%d  (Build %d)'): String;
    { Private declarations }
  public
    { Public declarations }
  protected
    procedure CreateParams(var Params: TCreateParams); override; // For shadow
  end;

var
  AboutBox: TAboutBox;


{ ---------------------------------------------------------------------------- }


implementation
{$R *.DFM}

uses
  main;


{ ---------------------------------------------------------------------------- }


procedure TAboutBox.Label5Click(Sender: TObject);
{ Open my webpage }
begin
    ShellExecute(Handle, 'open', 'http://www.alexschreyer.de/projects/xloptim/', nil, nil,
    SW_SHOWNORMAL);
end;


procedure TAboutBox.Label5MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
{ Underline hyperlink when mouse hovers over }
begin
  Label5.Font.Style := [fsUnderline];
end;


procedure TAboutBox.Panel3MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
{ Don't underline hyperlink when mouse somewhere else }
begin
  Label5.Font.Style := [];
end;


{ ---------------------------------------------------------------------------- }


procedure TAboutBox.FormCreate(Sender: TObject);
{ Initialize form and show version info }
begin
  TextField.Lines.LoadFromFile(ExtractFilePath(Application.ExeName) + 'README.TXT');
  VersionInfo.Caption := Sto_GetFmtFileVersion(Application.ExeName);
  AlphaBlendValue := 0;
  opening := true;
end;


procedure TAboutBox.Timer1Timer(Sender: TObject);
{  Fade in the about form }
begin
  if opening then begin
    while AlphaBlendValue < 255 do
      AlphaBlendValue :=  AlphaBlendValue + 5;
    AlphaBlendValue := 255;
  end
  else begin
    while AlphaBlendValue > 0 do
      AlphaBlendValue :=  AlphaBlendValue - 5;
    AlphaBlendValue := 0;
  end;
end;


procedure TAboutBox.OKButtonClick(Sender: TObject);
{ Prepare for closing fade }
begin
  opening := false;
end;


procedure TAboutBox.FormClose(Sender: TObject; var Action: TCloseAction);
{ Fade out } 
begin
  Timer1Timer(OKButton);
end;


procedure TAboutBox.CreateParams(var Params: TCreateParams);
{ Show a little drop shadow }
const
 CS_DROPSHADOW = $00020000;
function IsWinXP: Boolean;
  begin
   Result := (Win32Platform = VER_PLATFORM_WIN32_NT) and
     (Win32MajorVersion >= 5) and (Win32MinorVersion >= 1);
  end;
begin
inherited;
 if IsWinXP then
 Params.WindowClass.Style := Params.WindowClass.Style or CS_DROPSHADOW
 else
end;


{ ---------------------------------------------------------------------------- }


////////////////////////////////////////////////////////////////
// this function reads the file ressource of "FileName" and returns the version
// number as formatted text. in "Fmt" you can use at most four integer values.
// if "Fmt" is invalid, the function may raise an EConvertError exception.
// examples for "Fmt" with versioninfo 4.3.128.0:
// '%d.%d.%d.%d' => '4.13.128.0'
// '%.2d-%.2d-%.2d' => '04-13-128'
//
function TAboutBox.Sto_GetFmtFileVersion(const FileName: String;
  const Fmt: String = 'Version: %d.%d  (Build %d)'): String;    // Removed: .%d
var
  iBufferSize: DWORD;
  iDummy: DWORD;
  pBuffer: Pointer;
  pFileInfo: Pointer;
  iVer: Array[1..4] of Word;
begin
  // set default value
  Result := '';
  // get size of version info (0 if no version info exists)
  iBufferSize := GetFileVersionInfoSize(PChar(FileName), iDummy);
  if (iBufferSize > 0) then
  begin
    GetMem(pBuffer, iBufferSize);
    try
    // get fixed file info
    GetFileVersionInfo(PChar(FileName), 0, iBufferSize, pBuffer);
    VerQueryValue(pBuffer, '\', pFileInfo, iDummy);
    // read version blocks
    iVer[1] := HiWord(PVSFixedFileInfo(pFileInfo)^.dwFileVersionMS);
    iVer[2] := LoWord(PVSFixedFileInfo(pFileInfo)^.dwFileVersionMS);
    iVer[3] := HiWord(PVSFixedFileInfo(pFileInfo)^.dwFileVersionLS);
    iVer[4] := LoWord(PVSFixedFileInfo(pFileInfo)^.dwFileVersionLS);
    finally
      FreeMem(pBuffer);
    end;
    // format result string
    Result := Format(Fmt, [iVer[1], iVer[2], iVer[4]]);  // Removed:  iVer[3],
  end;
end;


{ ---------------------------------------------------------------------------- }


end.

