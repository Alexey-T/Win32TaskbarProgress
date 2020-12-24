{*******************************************************************************

Windows TaskBar Progress Unit File

originally from:
https://stackoverflow.com/questions/5814765/how-do-i-show-progress-in-status-task-bar-button-using-delphi-7

refactored, adapted to Lazarus by Alexey Torgashin:
https://github.com/Alexey-T/Win32TaskbarProgress
License: MIT

*******************************************************************************}

unit win32taskbarprogress;

{$mode objfpc}{$H+}

interface

uses
  SysUtils, ShlObj, ComObj;

type
  TTaskBarProgressStyle = (tbpsNone, tbpsIndeterminate, tbpsNormal, tbpsError, tbpsPaused);

  TWin7TaskProgressBar = class
  private
    FHandle: THandle;
    FMin: Integer;
    FMax: Integer;
    FValue: Integer;
    FStyle: TTaskBarProgressStyle;
    FVisible: Boolean;
    FMarquee: Boolean;
    FIntf: ITaskbarList3;
    procedure SetProgress(const AValue: Integer);
    procedure SetMax(const AValue: Integer);
    procedure SetStyle(const AValue: TTaskBarProgressStyle);
    procedure SetVisible(const AValue: Boolean);
    procedure SetMarquee(const AValue: Boolean);
  public
    // AHandle param should be Application.Handle, not form's Handle
    constructor Create(const AHandle: THandle);
    destructor Destroy; override;
    property Max: Integer read FMax write SetMax;
    property Min: Integer read FMin;
    property Style: TTaskBarProgressStyle read FStyle write SetStyle;
    property Progress: Integer read FValue write SetProgress;
    // Marquee is simplified prop to set Style: Indeterminate or None/Normal
    property Marquee: Boolean read FMarquee write SetMarquee;
    // Visible is simplified prop to set Style: None or Normal
    property Visible: Boolean read FVisible write SetVisible;
  end;

var
  Win7TaskbarProgress: TWin7TaskProgressBar = nil;

implementation

procedure TWin7TaskProgressBar.SetMax(const AValue: Integer);
begin
  FMax := AValue;
  SetProgress(FValue);
end;

procedure TWin7TaskProgressBar.SetProgress(const AValue: Integer);
begin
  if (FIntf <> nil) and (FHandle <> 0) then
  begin
    FValue := AValue;
    if not FMarquee then
      FIntf.SetProgressValue(FHandle, UInt64(FValue), UInt64(FMax));
  end;
end;

procedure TWin7TaskProgressBar.SetStyle(const AValue: TTaskBarProgressStyle);
const
  Flags: array[TTaskBarProgressStyle] of Cardinal = (0, 1, 2, 4, 8);
begin
  if (FIntf <> nil) and (FHandle <> 0) then
    FIntf.SetProgressState(FHandle, Flags[AValue]);

  FStyle := AValue;
end;

procedure TWin7TaskProgressBar.SetVisible(const AValue: Boolean);
begin
  if AValue then
  begin
    if (FStyle <> tbpsNormal) then
      SetStyle(tbpsNormal)
  end
  else
    SetStyle(tbpsNone);

  FVisible := AValue;
end;

procedure TWin7TaskProgressBar.SetMarquee(const AValue: Boolean);
begin
  if AValue then
    SetStyle(tbpsIndeterminate)
  else
  begin
    SetStyle(tbpsNone);
    if FVisible then
    begin
      SetProgress(FValue);
      SetStyle(tbpsNormal);
    end;
  end;

  FMarquee := AValue;
end;

constructor TWin7TaskProgressBar.Create(const AHandle: THandle);
begin
  if (AHandle = 0) then exit;

  if Win32MajorVersion >= 6 then
    try
      FHandle := AHandle;
      FIntf:= CreateComObject(CLSID_TaskbarList) as ITaskbarList3;

      if (FIntf <> nil) then
        FIntf.SetProgressState(FHandle, 0);

      FMin := 0;
      FMax := 100;
      FValue := 10;
      FStyle := tbpsNormal;

      SetStyle(FStyle);
      SetVisible(FVisible);
    except
      FIntf := nil;
    end;
end;


destructor TWin7TaskProgressBar.Destroy;
begin
  if (FIntf <> nil) then
  begin
    FIntf.SetProgressState(FHandle, 0);
    FIntf := nil;
  end;
end;

end.
