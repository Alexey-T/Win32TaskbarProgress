unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, Spin, StdCtrls;

type

  { TForm1 }

  TForm1 = class(TForm)
    ComboBox1: TComboBox;
    ed1: TSpinEdit;
    Label1: TLabel;
    Label2: TLabel;
    procedure ComboBox1Change(Sender: TObject);
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
  Win7TaskbarProgress.Progress:= ed1.Value;
end;

procedure TForm1.FormShow(Sender: TObject);
begin
  Win7TaskbarProgress:= TWin7TaskProgressBar.Create(Application.{%H-}Handle);
  Win7TaskbarProgress.Visible:= true;
  ComboBox1.ItemIndex:= Ord(tbpsNormal);
  ComboBox1Change(Self);
end;

procedure TForm1.ComboBox1Change(Sender: TObject);
begin
  Win7TaskbarProgress.Style:= TTaskBarProgressStyle(ComboBox1.ItemIndex);
end;

end.

