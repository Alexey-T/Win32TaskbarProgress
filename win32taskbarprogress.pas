{*******************************************************************************

Windows 7 TaskBar Progress Unit File

originally from:
https://github.com/tarampampam
https://stackoverflow.com/questions/5814765/how-do-i-show-progress-in-status-task-bar-button-using-delphi-7

adapted to Lazarus by Alexey Torgashin:
https://github.com/Alexey-T/Win32TaskbarProgress
License: MIT

*******************************************************************************}

unit win32taskbarprogress;

{$mode objfpc}{$H+}

interface

uses
  Windows, SysUtils, ActiveX;

type
  ITaskbarList = interface(IUnknown)
    ['{56FDF342-FD6D-11D0-958A-006097C9A090}']
    function HrInit: HRESULT; stdcall;
    function AddTab(hwnd: LongWord): HRESULT; stdcall;
    function DeleteTab(hwnd: LongWord): HRESULT; stdcall;
    function ActivateTab(hwnd: LongWord): HRESULT; stdcall;
    function SetActiveAlt(hwnd: LongWord): HRESULT; stdcall;
  end;

  ITaskbarList2 = interface(ITaskbarList)
    ['{602D4995-B13A-429B-A66E-1935E44F4317}']
    function MarkFullscreenWindow(hwnd: LongWord;
      fFullscreen: LongBool): HRESULT; stdcall;
  end;

  ITaskbarList3 = interface(ITaskbarList2)
    ['{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}']
    procedure SetProgressValue(hwnd: LongWord; ullCompleted: UInt64; ullTotal: UInt64); stdcall;
    procedure SetProgressState(hwnd: LongWord; tbpFlags: Integer); stdcall;
  end;

type
  TTaskBarProgressStyle = (tbpsNone, tbpsIndeterminate, tbpsNormal, tbpsError, tbpsPaused);

  TWin7TaskProgressBar = class
  private
    glHandle: THandle;
    glMin,
    glMax,
    glValue: Integer;
    glStyle: TTaskBarProgressStyle;
    glVisible,
    glMarquee: Boolean;
    glTaskBarInterface: ITaskbarList3;
  private
    procedure SetProgress(const Value: Integer);
    procedure SetMax(const Value: Integer);
    procedure SetStyle(const Style: TTaskBarProgressStyle);
    procedure SetVisible(const IsVisible: Boolean);
    procedure SetMarquee(const IsMarquee: Boolean);
  public
    constructor Create(const Handle: THandle);
    property Max: Integer read glMax write SetMax default 100;
    property Min: Integer read glMin default 0;
    property Progress: Integer read glValue write SetProgress default 0;
    property Marquee: Boolean read glMarquee write SetMarquee default False;
    property Style: TTaskBarProgressStyle read glStyle write SetStyle default tbpsNone;
    property Visible: Boolean read glVisible write SetVisible default False;
    destructor Destroy; override;
  end;

var
  Win7TaskbarProgress: TWin7TaskProgressBar = nil;

implementation

procedure TWin7TaskProgressBar.SetMax(const Value: Integer);
begin
  glMax := Value;
  SetProgress(glValue);
end;

procedure TWin7TaskProgressBar.SetProgress(const Value: Integer);
begin
  if (glTaskBarInterface <> nil) and (glHandle <> 0) then begin
    glValue := Value;
    if not glMarquee then
      glTaskBarInterface.SetProgressValue(glHandle, UInt64(glValue), UInt64(glMax));
  end;
end;

procedure TWin7TaskProgressBar.SetStyle(const Style: TTaskBarProgressStyle);
const
  Flags: array[TTaskBarProgressStyle] of Cardinal = (0, 1, 2, 4, 8);
begin
  if (glTaskBarInterface <> nil) and (glHandle <> 0) then
    glTaskBarInterface.SetProgressState(glHandle, Flags[Style]);

  glStyle := Style;
end;

procedure TWin7TaskProgressBar.SetVisible(const IsVisible: Boolean);
begin
  if IsVisible then begin
    if (glStyle <> tbpsNormal) then
      SetStyle(tbpsNormal)
  end else
    SetStyle(tbpsNone);

  glVisible := IsVisible;
end;

procedure TWin7TaskProgressBar.SetMarquee(const IsMarquee: Boolean);
begin
  if IsMarquee then
    SetStyle(tbpsIndeterminate)
  else begin
    SetStyle(tbpsNone);
    if glVisible then begin
      SetProgress(glValue);
      SetStyle(tbpsNormal);
    end;
  end;

  glMarquee := IsMarquee;
end;

constructor TWin7TaskProgressBar.Create(const Handle: THandle);
const
  CLSID_TaskbarList: TGUID = '{56FDF344-FD6D-11d0-958A-006097C9A090}';
var
  OSVersionInfo : TOSVersionInfo;
begin
  OSVersionInfo.dwOSVersionInfoSize := sizeof(TOSVersionInfo);
  if (Handle <> 0) and GetVersionEx(OSVersionInfo) then
    if OSVersionInfo.dwMajorVersion >= 6 then try
      glHandle := Handle;
      CoCreateInstance(CLSID_TaskbarList, nil, CLSCTX_INPROC, ITaskbarList3, glTaskBarInterface);

      if (glTaskBarInterface <> nil) then
        glTaskBarInterface.SetProgressState(glHandle, 0);

      glMin := 0;
      glMax := 100;
      glValue := 10;
      glStyle := tbpsNormal;

      SetStyle(glStyle);
      SetVisible(glVisible);
    except
      glTaskBarInterface := nil;
    end;
end;


destructor TWin7TaskProgressBar.Destroy;
begin
  if (glTaskBarInterface <> nil) then begin
    glTaskBarInterface.SetProgressState(glHandle, 0);
    glTaskBarInterface := nil;
  end;
end;

end.
