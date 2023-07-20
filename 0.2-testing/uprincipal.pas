unit uPrincipal;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, Buttons, ACBrNFe,
  ACBrDANFCeFortesFr, ACBrNFeDANFeRLClass, LCLIntf, StdCtrls, ComCtrls,
  ExtCtrls, ActnList, pcnConversao;

type

  { TfrmPrincipal }

  TfrmPrincipal = class(TForm)
    ACBrNFe1: TACBrNFe;
    ACBrNFeDANFeRL1: TACBrNFeDANFeRL;
    ReadINIFile: TAction;
    NewINIFile: TAction;
    ActionList1: TActionList;
    btConsultaXML: TBitBtn;
    btCleanLog: TBitBtn;
    btConvert: TBitBtn;
    Image1: TImage;
    mmXML: TMemo;
    OpenDialog1: TOpenDialog;
    Panel1: TPanel;
    Panel2: TPanel;
    StatusBar1: TStatusBar;
    procedure ACBrNFe1StatusChange(Sender: TObject);
    procedure btConsultaXMLClick(Sender: TObject);
    procedure btCleanLogClick(Sender: TObject);
    procedure btConvertClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Image1Click(Sender: TObject);
    procedure NewINIFileExecute(Sender: TObject);
    procedure ReadINIFileExecute(Sender: TObject);
  private

  public

  end;

var
  frmPrincipal: TfrmPrincipal;

  NFeTXTVendas_, NFeTXTVendasSTQLog_: string;
  NFeXMLVendas_, NFePathSchemas_,
  DANFEPath_: string;
  //
  Cert_SSLLib_, Cert_CryptLib_, Cert_HttpLib_, Cert_XmlSignLib_: integer;
  //URLPFX_: string;
  Cert_Caminho_, Cert_Senha_: string;
  //NumSerie_: string;
  //
  WS_UF_: string;
  WS_TipoAmb_: Integer;
  WS_Visualizar_, WS_SalvarSOAP_, WS_AjustarAut_: Boolean;
  WS_Aguardar_, WS_Tentativas_, WS_Intervalo_: String;
  WS_TimeOut_, WS_SSLType_: integer;
  //
  WS_TipoAmb_Current: TpcnTipoAmbiente = taHomologacao;

implementation

uses pcnConversaoNFe, ACBrDFeSSL, {synacode,} blcksock,
     Frm_Status, IniFiles;

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
        ShowMessage('O arquivo selecionado não é um XML de NF-e 4.01 válido!')
      end;
    end;
  end;
end;

procedure TfrmPrincipal.FormCreate(Sender: TObject);
begin
  try
    if not FileExists('.\danfev.ini') then
    begin
      Beep;
      ShowMessage('Arquivo danfev.ini não encontrado!'+#13+
                  'Um arquivo INI com a configurações padrão será criado.'+#13+#13+
                  'Antes de continuar, configure o caminho e a senha do certificado digital A1 no arquivo danfev.ini.');
      NewINIFile.Execute;
    end;
  finally
    ReadINIFile.Execute;
  end;
end;

procedure TfrmPrincipal.btConsultaXMLClick(Sender: TObject);
var chaveStr, retornoWS, statusWS: string;
    Arq: TextFile;
    p,q: Integer;

begin
  OpenDialog1.Title := 'Escolha o XML da Nota Fiscal Eletrônica';
  OpenDialog1.DefaultExt := '*.xml';
  OpenDialog1.Filter := 'Arquivos NFe (*.xml)|*.xml';

  if OpenDialog1.Execute then
  begin
    try
    AssignFile(Arq, OpenDialog1.FileName);
    Reset(Arq);
    Read(Arq, chaveStr);

    p:= Pos('Id=', chaveStr);
    if p > 0 then
    begin
      // Adicione a posição do início e o comprimento da tag ao resultado para encontrar o fim do texto desejado
      p:= p + Length('Id=');
      chaveStr:= Copy(chaveStr, p+4, 44);
      CloseFile(Arq);

      //ShowMessage(chaveStr);
      ACBrNFe1.NotasFiscais.Clear;
      ACBrNFe1.WebServices.Consulta.NFeChave:= chaveStr;
      ACBrNFe1.WebServices.Consulta.Executar;

      retornoWS:= ACBrNFe1.WebServices.Consulta.RetornoWS;
      mmXML.Lines.Add('===== Inicio da resposta da SEFAZ =====');
      mmXML.Lines.Add(retornoWS);
      mmXML.Lines.Add('=====  Fim da resposta da SEFAZ  =====');

      p:= Pos('<xMotivo>', retornoWS);
      if p > 0 then
      begin
        // Adicione a posição do início e o comprimento da tag ao resultado para encontrar o fim do texto desejado
        p:= p + Length('<xMotivo>');
        q:= Pos('</xMotivo>', retornoWS);
        if q > 0 then
        begin
          // Copie o texto desejado
          statusWS:= Copy(retornoWS, p, q - p);
          ShowMessage(statusWS);
        end;
      end;
      //ShowMessage(ACBrNFe1.WebServices.Consulta.Protocolo);
    end;
    except
      ShowMessage('Não foi possível ler o XML ou o arquivo está inválido!');
    end;
  end;
end;

procedure TfrmPrincipal.btCleanLogClick(Sender: TObject);
begin
  mmXML.Lines.Clear;
end;

procedure TfrmPrincipal.ACBrNFe1StatusChange(Sender: TObject);
begin
  case ACBrNFe1.Status of
    stIdle:
      begin
        if ( frmStatus <> nil ) then
          frmStatus.Hide;
      end;

    stNFeStatusServico:
      begin
        if ( frmStatus = nil ) then
          frmStatus := TfrmStatus.Create(Application);

        frmStatus.lblStatus.Caption := 'Verificando Status do servico...';
        //=frmStatus.Show;
        frmStatus.Show;
        frmStatus.BringToFront;
      end;

    stNFeRecepcao:
      begin
        if ( frmStatus = nil ) then
          frmStatus := TfrmStatus.Create(Application);

        frmStatus.lblStatus.Caption := 'Enviando dados da NFe...';
        frmStatus.Show;
        frmStatus.BringToFront;
      end;

    stNfeRetRecepcao:
      begin
        if ( frmStatus = nil ) then
          frmStatus := TfrmStatus.Create(Application);

        frmStatus.lblStatus.Caption := 'Recebendo dados da NFe...';
        frmStatus.Show;
        frmStatus.BringToFront;
      end;

    stNfeConsulta:
      begin
        if ( frmStatus = nil ) then
          frmStatus := TfrmStatus.Create(Application);

        frmStatus.lblStatus.Caption := 'Consultando NFe...';
        frmStatus.Show;
        frmStatus.BringToFront;
      end;

    stNfeCancelamento:
      begin
        if ( frmStatus = nil ) then
          frmStatus := TfrmStatus.Create(Application);

        frmStatus.lblStatus.Caption := 'Enviando cancelamento de NFe...';
        frmStatus.Show;
        frmStatus.BringToFront;
      end;

    stNfeInutilizacao:
      begin
        if ( frmStatus = nil ) then
          frmStatus := TfrmStatus.Create(Application);

        frmStatus.lblStatus.Caption := 'Enviando pedido de Inutilização...';
        frmStatus.Show;
        frmStatus.BringToFront;
      end;

    stNFeRecibo:
      begin
        if ( frmStatus = nil ) then
          frmStatus := TfrmStatus.Create(Application);

        frmStatus.lblStatus.Caption := 'Consultando Recibo de Lote...';
        frmStatus.Show;
        frmStatus.BringToFront;
      end;

    stNFeCadastro:
      begin
        if ( frmStatus = nil ) then
          frmStatus := TfrmStatus.Create(Application);

        frmStatus.lblStatus.Caption := 'Consultando Cadastro...';
        frmStatus.Show;
        frmStatus.BringToFront;
      end;

    stNFeEmail:
      begin
        if ( frmStatus = nil ) then
          frmStatus := TfrmStatus.Create(Application);

        frmStatus.lblStatus.Caption := 'Enviando Email...';
        frmStatus.Show;
        frmStatus.BringToFront;
      end;

    stNFeCCe:
      begin
        if ( frmStatus = nil ) then
          frmStatus := TfrmStatus.Create(Application);

        frmStatus.lblStatus.Caption := 'Enviando Carta de Correção...';
        frmStatus.Show;
        frmStatus.BringToFront;
      end;

    stNFeEvento:
      begin
        if ( frmStatus = nil ) then
          frmStatus := TfrmStatus.Create(Application);

        frmStatus.lblStatus.Caption := 'Enviando Evento...';
        frmStatus.Show;
        frmStatus.BringToFront;
      end;
  end;

  Application.ProcessMessages;
end;

procedure TfrmPrincipal.Image1Click(Sender: TObject);
begin
  OpenURL('https://connectivaredes.com/');
end;

procedure TfrmPrincipal.NewINIFileExecute(Sender: TObject);
var arq: TextFile;
begin
  try
    AssignFile(arq, '.\danfev.ini');
    Rewrite(arq);
    Writeln(Arq, '[CERTIFICADO]');
    Writeln(Arq, 'SSLLib=4');
    Writeln(Arq, 'CryptLib=3');
    Writeln(Arq, 'HttpLib=2');
    Writeln(Arq, 'XmlSignLib=4');
    Writeln(Arq, 'URL=');
    Writeln(Arq, 'Caminho=.\certificado.pfx');
    Writeln(Arq, 'Senha=123456');
    Writeln(Arq, 'NumSerie=');
    Writeln(Arq);
    Writeln(Arq, '[WEBSERVICE]');
    Writeln(Arq, 'UF=MG');
    Writeln(Arq, 'Ambiente=1');
    Writeln(Arq, 'Visualizar=0');
    Writeln(Arq, 'SalvarSOAP=0');
    Writeln(Arq, 'AjustarAut=0');
    Writeln(Arq, 'Aguardar=0');
    Writeln(Arq, 'Tentativas=5');
    Writeln(Arq, 'Intervalo=0');
    Writeln(Arq, 'TimeOut=30000');
    Writeln(Arq, 'SSLType=5');
    Writeln(Arq);
    Writeln(Arq, '[CAMINHOS]');
    Writeln(Arq, 'NFePathSchemas=.\Schemas\NFe\');
    Writeln(Arq, 'DANFEPath=.\Docs\');
    CloseFile(Arq);
  except
    ShowMessage('Falha ao criar arquivo danfe.ini. Verifique as permissões do diretório ou se já existe um arquivo com o mesmo nome e que esteja sendo usado por outro processo.');
  end;
  Application.Terminate;
end;

procedure TfrmPrincipal.ReadINIFileExecute(Sender: TObject);
var ArqIni: TIniFile;
begin
  //READ INI FILE - danfev.ini
  ArqIni:=            TIniFile.Create('.\danfev.ini');

  //CAMINHOS
  NFePathSchemas_:=     ArqIni.ReadString('CAMINHOS', 'NFePathSchemas', '.\Schemas\NFe\');
  DANFEPath_:=          ArqIni.ReadString('CAMINHOS', 'DANFEPath', '.\Docs\');

  //CERTIFICADO
  Cert_SSLLib_:=        ArqIni.ReadInteger('CERTIFICADO', 'SSLLib',     4);
  Cert_CryptLib_:=      ArqIni.ReadInteger('CERTIFICADO', 'CryptLib',   3);
  Cert_HttpLib_:=       ArqIni.ReadInteger('CERTIFICADO', 'HttpLib',    2);
  Cert_XmlSignLib_:=    ArqIni.ReadInteger('CERTIFICADO', 'XmlSignLib', 4);
  //URLPFX.Text:=         Ini.ReadString( 'CERTIFICADO', 'URL',        '');
  Cert_Caminho_:=       ArqIni.ReadString( 'CERTIFICADO', 'Caminho',    '.\certificado.pfx');
  Cert_Senha_:=         ArqIni.ReadString( 'CERTIFICADO', 'Senha',      '123456');
  //NumSerie:=          Ini.ReadString( 'CERTIFICADO', 'NumSerie',   '');

  //WEBSERVICE
  WS_UF_:=              ArqIni.ReadString('WEBSERVICE', 'UF',   'MG');
  //informa  se as NF-e serão emitidas em Produção ou Homologação
  WS_TipoAmb_:=         ArqIni.ReadInteger('WEBSERVICE', 'Ambiente',   1);

  if WS_TipoAmb_ = 1 then //1 = Producao / 2 = Homologacao
    WS_TipoAmb_Current:= taProducao
    else
      WS_TipoAmb_Current:= taHomologacao;

  WS_Visualizar_:=      ArqIni.ReadBool(   'WEBSERVICE', 'Visualizar', False);
  WS_SalvarSOAP_:=      ArqIni.ReadBool(   'WEBSERVICE', 'SalvarSOAP', False);
  WS_AjustarAut_:=      ArqIni.ReadBool(   'WEBSERVICE', 'AjustarAut', False);
  WS_Aguardar_:=        ArqIni.ReadString( 'WEBSERVICE', 'Aguardar',   '0');
  WS_Tentativas_:=      ArqIni.ReadString( 'WEBSERVICE', 'Tentativas', '5');
  WS_Intervalo_:=       ArqIni.ReadString( 'WEBSERVICE', 'Intervalo',  '0');
  WS_TimeOut_:=         ArqIni.ReadInteger('WEBSERVICE', 'TimeOut',    30000);
  WS_SSLType_:=         ArqIni.ReadInteger('WEBSERVICE', 'SSLType',    5);
  //
  ArqIni.Destroy;
  //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

  //CARREGA INFORMAÇÕES DO CERTIFICADO
  ACBrNFe1.Configuracoes.Geral.SSLLib             := TSSLLib(Cert_SSLLib_);
  ACBrNFe1.Configuracoes.Geral.SSLCryptLib        := TSSLCryptLib(Cert_CryptLib_);
  ACBrNFe1.Configuracoes.Geral.SSLHttpLib         := TSSLHttpLib(Cert_HttpLib_);
  ACBrNFe1.Configuracoes.Geral.SSLXmlSignLib      := TSSLXmlSignLib(Cert_XmlSignLib_);
  //  ACBrNFe1.SSL.SSLType                            := LT_TLSv1_2;
  //ACBrNFe1.Configuracoes.Certificados.URLPFX      := uPrincipal.Cert_URLPFX_;
  ACBrNFe1.Configuracoes.Certificados.ArquivoPFX  := Cert_Caminho_;
  ACBrNFe1.Configuracoes.Certificados.Senha       := Cert_Senha_;
  //ACBrNFe1.Configuracoes.Certificados.NumeroSerie := Cert_NumSerie_;

  //CONFIGURA CAMINHOS
  ACBrNFe1.Configuracoes.Arquivos.PathSchemas     := NFePathSchemas_;
  ACBrNFeDANFeRL1.PathPDF                         := DANFEPath_;
  if not DirectoryExists(DANFEPath_) then
    CreateDir(DANFEPath_);

  //CARREGA INFORMAÇÕES DO WEBSERVICE
  with ACBrNFe1.Configuracoes.WebServices do
  begin
    UF:=         WS_UF_;
    Ambiente:=   WS_TipoAmb_Current;
    Visualizar:= WS_Visualizar_;
    Salvar:=     WS_SalvarSOAP_;

    AjustaAguardaConsultaRet:= WS_AjustarAut_;
    TimeOut:=    WS_TimeOut_;
    //ProxyHost:= //TODO
    //ProxyPort:= //TODO
    //ProxyUser:= //TODO
    //ProxyPass:= //TODO
  end;
  //CONFIGURA SSL
  ACBrNFe1.SSL.SSLType := TSSLType(WS_SSLType_);
end;

end.

