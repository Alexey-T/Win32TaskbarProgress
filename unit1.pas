unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, Spin, StdCtrls;

type

  { TForm1 }

  TForm1 = class(TForm)
    ComboBoxStyle: TComboBox;
    ed1: TSpinEdit;
    Label1: TLabel;
    Label2: TLabel;
    procedure ComboBoxStyleChange(Sender: TObject);
    procedure ed1Change(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private

  public

  end;

var
  Form1: TForm1;

implementation

uses
  win32taskbarprogress;

{$R *.lfm}

{ TForm1 }

procedure TForm1.ed1Change(Sender: TObject);
begin
  GlobalTaskbarProgress.Progress:= ed1.Value;
end;

procedure TForm1.FormShow(Sender: TObject);
begin
  GlobalTaskbarProgress:= TWin7TaskProgressBar.Create;
end;

procedure TForm1.ComboBoxStyleChange(Sender: TObject);
begin
  GlobalTaskbarProgress.Style:= TTaskBarProgressStyle(ComboBoxStyle.ItemIndex);
end;

end.

