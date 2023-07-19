unit uPrincipal;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, Buttons, ACBrNFe,
  ACBrDANFCeFortesFr, ACBrNFeDANFeRLClass, LCLIntf, StdCtrls, ComCtrls,
  ExtCtrls;

type

  { TfrmPrincipal }

  TfrmPrincipal = class(TForm)
    ACBrNFe1: TACBrNFe;
    ACBrNFeDANFeRL1: TACBrNFeDANFeRL;
    btConvert: TBitBtn;
    Image1: TImage;
    mmXML: TMemo;
    OpenDialog1: TOpenDialog;
    Panel1: TPanel;
    Panel2: TPanel;
    StatusBar1: TStatusBar;
    procedure btConvertClick(Sender: TObject);
    procedure Image1Click(Sender: TObject);
  private

  public

  end;

var
  frmPrincipal: TfrmPrincipal;

implementation

{$R *.lfm}

{ TfrmPrincipal }

procedure TfrmPrincipal.btConvertClick(Sender: TObject);
var pathStr,fileStr: string;

begin
  if OpenDialog1.Execute then
  begin
    try
      ACBrNFe1.NotasFiscais.Clear;
      ACBrNFe1.NotasFiscais.LoadFromFile(OpenDialog1.FileName);
      ACBrNFe1.Configuracoes.Arquivos.PathSalvar:= '/';
      ACBrNFe1.NotasFiscais.ImprimirPDF;
      //
      pathStr:= '';
      fileStr:= '';
      pathStr:= ExtractFileDir(Paramstr(0));
      fileStr:= ACBrNFeDANFeRL1.ArquivoPDF;
      fileStr:= Copy(fileStr, 2, Length(fileStr));
    finally
      if fileStr <> '' then
      begin
        mmXML.Lines.Add(pathStr+fileStr);
        OpenDocument(pathStr+fileStr);
      end
      else
      begin
        Beep;
        ShowMessage('O arquivo escolhido não é um XML de NF-e 4.01 válido!')
      end;
    end;
  end;
end;

procedure TfrmPrincipal.Image1Click(Sender: TObject);
begin
  OpenURL('https://connectivaredes.com/');
end;

end.

