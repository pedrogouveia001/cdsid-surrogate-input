unit UntEntrada;

interface

uses
  Classes, SysUtils, IWAppForm, IWApplication, IWColor, IWTypes, IWCompButton,
  IWCompListbox, IWCompEdit, IWVCLBaseControl, IWBaseControl, IWBaseHTMLControl,
  IWControl, IWCompLabel, Controls, Forms, IWVCLBaseContainer, IWContainer,
  IWHTMLContainer, IWHTML40Container, IWRegion, IWGrids, ComObj, Dialogs,
  Variants, Math, UserSessionUnit, IWExtCtrls, jpeg, windows, ActiveX,
  IWWebGrid, IWAdvWebGrid, IWCompTabControl, IWAdvWebGridExcel, IWCompText,
  IWCompCheckbox, IWBaseComponent, IWBaseHTMLComponent, IWBaseHTML40Component,
  chart, series, pngimage, Grids,
  IWCompProgressBar, ExtCtrls, IWHTMLControls, TeeGDIPlus, TeEngine, TeeProcs,
  IWTMSPopup, IWHelpTip, Menus,

  BubbleCh, TeePolar,
  IWCompRadioButton, IWTMSCheckList, TeeLisB,
  TeeSeriesTextEd, Buttons, CheckLst,
  ShellApi, TeeLegendScrollBar, Spin, ToolWin, StdActns,
  TeeURL, TeeXML, AppEvnts, IWCompDynamicChart, IWVCLComponent, TeeProCo,
  Messages, Graphics, ComCtrls, StdCtrls, ErrorBar,
  IWProgressIndicator, IWDBGrids,
  IWColorComboBox, IWTMSImgCtrls, IWTMSCtrls,
  IWTMSMenus, RpDefine, RpRender, RpRenderText,
  IWAdvCheckGroup, IWDBAdvCheckGroup, GanttCh, IWCompRectangle,
  TeeDoubleHorizBar;

type
  TIWForm1 = class(TIWAppForm)
    rgnmatcons: TIWRegion;
    IWImage31: TIWImage;
    IWImage32: TIWImage;
    IWLabel20: TIWLabel;
    IWRegion18: TIWRegion;
    IWRegion30: TIWRegion;
    IWRegion1: TIWRegion;
    IWLabel3: TIWLabel;
    IWLabel1: TIWLabel;
    IWLabel2: TIWLabel;
    btnupdate: TIWButton;
    rgnmatrizdeconsequencias: TIWRegion;
    IWLabel4: TIWLabel;
    btnshowresults: TIWButton;
    grdmatrizconseq: TTIWAdvWebGrid;
    cmbbxauxiliar: TIWComboBox;
    rgnresultado: TIWRegion;
    IWImage1: TIWImage;
    IWImage2: TIWImage;
    IWLabel7: TIWLabel;
    IWRegion3: TIWRegion;
    IWRegion4: TIWRegion;
    chrtResultado: TChart;
    IWRegion5: TIWRegion;
    grdresultados: TIWGrid;
    IWRegion6: TIWRegion;
    grdcases: TIWGrid;
    IWLabel8: TIWLabel;
    IWLabel9: TIWLabel;
    Chart1: TChart;
    Series2: TBarSeries;
    edtnumalt: TIWEdit;
    grdImport: TTIWAdvWebGrid;
    TIWAdvWebGridExcelIO1: TTIWAdvWebGridExcelIO;
    IWLabel10: TIWLabel;
    FileArquivo: TIWFile;
    btnimport: TIWButton;
    IWLink1: TIWLink;
    rgnfiltro: TIWRegion;
    IWImage3: TIWImage;
    IWImage4: TIWImage;
    IWLabel6: TIWLabel;
    IWLink3: TIWLink;
    IWRegion7: TIWRegion;
    IWRegion8: TIWRegion;
    IWRegion9: TIWRegion;
    grdfiltro: TIWGrid;
    IWLabel15: TIWLabel;
    cmbbxfiltro: TIWComboBox;
    rgnfiltro2: TIWRegion;
    IWImage5: TIWImage;
    IWImage6: TIWImage;
    IWLabel16: TIWLabel;
    IWLink4: TIWLink;
    IWLabel22: TIWLabel;
    IWRegion10: TIWRegion;
    IWRegion11: TIWRegion;
    IWRegion12: TIWRegion;
    grdprimeiro: TIWGrid;
    IWRegion13: TIWRegion;
    grdsegundo: TIWGrid;
    IWLabel23: TIWLabel;
    rgnAnaliseEstatistica: TIWRegion;
    IWImage7: TIWImage;
    IWImage8: TIWImage;
    IWLabel24: TIWLabel;
    IWLink5: TIWLink;
    IWRegion14: TIWRegion;
    IWRegion15: TIWRegion;
    IWRegion16: TIWRegion;
    grdestatistica: TIWGrid;
    IWText1: TIWText;
    IWText2: TIWText;
    IWLink9: TIWLink;
    IWRegion2: TIWRegion;
    grdmaxmin: TIWGrid;
    IWText5: TIWText;
    IWRegion25: TIWRegion;
    imgResultado: TIWImage;
    Series1: TBarSeries;
    rgnAnaliseEstatisticaPromethee: TIWRegion;
    IWImage11: TIWImage;
    IWImage12: TIWImage;
    IWLabel32: TIWLabel;
    IWLink10: TIWLink;
    IWText3: TIWText;
    IWText4: TIWText;
    IWText6: TIWText;
    IWRegion23: TIWRegion;
    IWRegion24: TIWRegion;
    IWRegion26: TIWRegion;
    grdestatisticaPromethee: TIWGrid;
    IWRegion27: TIWRegion;
    grdmaxminPromethee: TIWGrid;
    IWRegion28: TIWRegion;
    IWLabel35: TIWLabel;
    rdgVisualizacao: TIWRadioGroup;
    btnupdatevisualizacao: TIWButton;
    rgncomparacao: TIWRegion;
    IWImage13: TIWImage;
    IWImage14: TIWImage;
    IWLabel37: TIWLabel;
    IWLabel38: TIWLabel;
    IWLabel39: TIWLabel;
    IWLink11: TIWLink;
    IWRegion32: TIWRegion;
    IWRegion33: TIWRegion;
    IWRegion34: TIWRegion;
    grdcomparacao: TIWGrid;
    IWRegion35: TIWRegion;
    grdcasescomparacao: TIWGrid;
    IWRegion36: TIWRegion;
    imgcomparacao: TIWImage;
    chrtsimilaridade: TChart;
    BarSeries3: TBarSeries;
    btncomparacao: TIWButton;
    grdexportacao: TTIWAdvWebGrid;
    lnkExportar: TIWLink;
    lnkexportacaopromethee: TIWLink;
    cmbbxcriteria: TIWComboBox;
    rgnregradedecisao: TIWRegion;
    IWImage15: TIWImage;
    IWImage16: TIWImage;
    IWLabel42: TIWLabel;
    IWLink12: TIWLink;
    IWRegion37: TIWRegion;
    IWRegion38: TIWRegion;
    IWRegion39: TIWRegion;
    grdregradedecisao: TIWGrid;
    IWLink13: TIWLink;
    txtregradedecisao: TIWText;
    rgndescricao: TIWRegion;
    IWText7: TIWText;
    IWLink15: TIWLink;
    IWLabel48: TIWLabel;
    lblsimilaridade: TIWLabel;
    btnforcacrit: TIWButton;
    btnestatistica: TIWButton;
    rgnresultadopromethee: TIWRegion;
    IWImage19: TIWImage;
    IWImage20: TIWImage;
    IWLabel49: TIWLabel;
    IWLabel50: TIWLabel;
    IWLabel51: TIWLabel;
    IWLink2: TIWLink;
    IWLink16: TIWLink;
    btncomparacaopromethee: TIWButton;
    btnforcacritpromethee: TIWButton;
    btnestatisticapromethee: TIWButton;
    IWRegion45: TIWRegion;
    IWRegion46: TIWRegion;
    IWRegion47: TIWRegion;
    grdresultadospromethee: TIWGrid;
    IWRegion48: TIWRegion;
    grdcasespromethee: TIWGrid;
    IWRegion49: TIWRegion;
    imgResultadoPromethee: TIWImage;
    IWRegion50: TIWRegion;
    IWLabel54: TIWLabel;
    rdgVisualizacaoPromethee: TIWRadioGroup;
    btnupdatevisualizacaoPromethee: TIWButton;
    chrtresultadopromethee: TChart;
    BarSeries2: TBarSeries;
    grdrecomendacao: TIWGrid;
    IWLabel27: TIWLabel;
    IWLabel28: TIWLabel;
    IWLabel29: TIWLabel;
    lnkdescricaolimiar: TIWLink;
    rgnlimiares: TIWRegion;
    grdlimiares: TTIWAdvWebGrid;
    btnatualizarlimiares: TIWButton;
    txtdescricaoregra: TIWText;
    IWLabel30: TIWLabel;
    rgnregradedecisaoPromethee: TIWRegion;
    IWImage9: TIWImage;
    IWImage10: TIWImage;
    IWLabel31: TIWLabel;
    IWLink6: TIWLink;
    txtregradedecisaopromethee: TIWText;
    grdrecomendacaopromethee: TIWGrid;
    IWLabel46: TIWLabel;
    IWLabel47: TIWLabel;
    lnkdescricaolimiarPromethee: TIWLink;
    IWRegion19: TIWRegion;
    IWRegion20: TIWRegion;
    IWRegion21: TIWRegion;
    grdregradedecisaopromethee: TIWGrid;
    rgndescricaoPromethee: TIWRegion;
    IWText9: TIWText;
    IWLabel55: TIWLabel;
    rgnlimiaresPromethee: TIWRegion;
    btnatualizarlimiaresPromethee: TIWButton;
    txtdescricaoregraPromethee: TIWText;
    IWLabel56: TIWLabel;
    grdlimiarespromethee: TTIWAdvWebGrid;
    lnkregrasdedecisao: TIWLink;
    grdperdamediaemaxima: TIWGrid;
    lblp1: TIWLabel;
    lblp2: TIWLabel;
    lblp3: TIWLabel;
    lblp4: TIWLabel;
    cmbbxp1: TIWComboBox;
    cmbbxp2: TIWComboBox;
    cmbbxp3: TIWComboBox;
    cmbbxp4: TIWComboBox;
    cmbbxp1promethee: TIWComboBox;
    lblp1promethee: TIWLabel;
    cmbbxp2promethee: TIWComboBox;
    cmbbxp3promethee: TIWComboBox;
    cmbbxp4promethee: TIWComboBox;
    lblp2promethee: TIWLabel;
    lblp3promethee: TIWLabel;
    lblp4promethee: TIWLabel;
    lnkregrasdedecisaoPromethee: TIWLink;
    grdperdamediaemaximapromethee: TIWGrid;
    IWLabel57: TIWLabel;
    IWLabel58: TIWLabel;
    IWLabel59: TIWLabel;
    IWLabel60: TIWLabel;
    IWLabel61: TIWLabel;
    IWLabel62: TIWLabel;
    IWLabel63: TIWLabel;
    IWLabel64: TIWLabel;
    IWLabel65: TIWLabel;
    IWLabel66: TIWLabel;
    rgnintro: TIWRegion;
    lblWarning: TIWLabel;
    btnLogin: TIWButton;
    btnRegister: TIWButton;
    IWImage17: TIWImage;
    IWLabel70: TIWLabel;
    IWImage18: TIWImage;
    IWImage27: TIWImage;
    txtwarnning: TIWText;
    iwrgnCaixaDialogoRU: TIWRegion;
    lblWg: TIWLabel;
    IWRegionru: TIWRegion;
    txtCaixaDialogoRU: TIWText;
    btnBackRU: TIWButton;
    btncontinueRU: TIWButton;
    IWRegion31: TIWRegion;
    IWRegion17: TIWRegion;
    rgnEnterLogin: TIWRegion;
    lbluserid: TIWLabel;
    edtemail: TIWEdit;
    btnEnter: TIWButton;
    lblSenha: TIWLabel;
    edtSenha: TIWEdit;
    lblpassword: TIWLabel;
    lblemail: TIWLabel;
    btnrecover: TIWButton;
    IWImage21: TIWImage;
    IWImage33: TIWImage;
    lnkesqueceusenha: TIWLink;
    lnkbacklogin: TIWLink;
    IWLabel76: TIWLabel;
    IWRegion22: TIWRegion;
    IWRegion40: TIWRegion;
    iwrgnCaixaDialogoRU2: TIWRegion;
    lblWg2: TIWLabel;
    IWRegionru2: TIWRegion;
    txtCaixaDialogoRU2: TIWText;
    btnBackRU2: TIWButton;
    btncontinueRU2: TIWButton;
    IWRegion41: TIWRegion;
    rgnOpcaoInput: TIWRegion;
    lblpleasechoose: TIWLabel;
    btnRegisternew: TIWButton;
    btnContinueregist: TIWButton;
    btnLogoutRgnOpcaoInput: TIWButton;
    IWImage37: TIWImage;
    IWImage38: TIWImage;
    IWLink7: TIWLink;
    IWImage47: TIWImage;
    IWImage48: TIWImage;
    IWLabel77: TIWLabel;
    IWRegion42: TIWRegion;
    IWRegion43: TIWRegion;
    IWRegion44: TIWRegion;
    rgnLoginproject: TIWRegion;
    lblChooseproj: TIWLabel;
    lstbxProjects: TIWListbox;
    btnReset3: TIWButton;
    btnLogoutRgnLogin: TIWButton;
    lblMexplanation: TIWLabel;
    btnauxchoose: TIWButton;
    IWImage25: TIWImage;
    IWImage34: TIWImage;
    IWLabel81: TIWLabel;
    IWRegion51: TIWRegion;
    IWRegion52: TIWRegion;
    IWRegion53: TIWRegion;
    IWRegion54: TIWRegion;
    IWLabel84: TIWLabel;
    edtnameproblem: TIWEdit;
    IWLink8: TIWLink;
    IWLink14: TIWLink;
    IWTimer1: TIWTimer;
    IWImage22: TIWImage;
    IWImage23: TIWImage;
    IWImage24: TIWImage;
    IWImage26: TIWImage;
    IWImage28: TIWImage;
    IWImage29: TIWImage;
    IWImage30: TIWImage;
    IWImage35: TIWImage;
    IWImage36: TIWImage;
    IWImage39: TIWImage;
    IWImage40: TIWImage;
    rgnelicitacaocrit: TIWRegion;
    IWImage41: TIWImage;
    IWImage42: TIWImage;
    IWLabel5: TIWLabel;
    IWLabel12: TIWLabel;
    IWImage43: TIWImage;
    IWRegion55: TIWRegion;
    IWRegion56: TIWRegion;
    IWRegion57: TIWRegion;
    grdcomparacaocrit: TIWGrid;
    IWLink18: TIWLink;
    IWText8: TIWText;
    IWImage44: TIWImage;
    IWImage45: TIWImage;
    IWImage46: TIWImage;
    rgnelicitacaoalt: TIWRegion;
    IWImage49: TIWImage;
    IWImage50: TIWImage;
    IWLabel11: TIWLabel;
    IWLabel13: TIWLabel;
    IWImage51: TIWImage;
    IWLink17: TIWLink;
    IWText10: TIWText;
    IWRegion58: TIWRegion;
    IWRegion59: TIWRegion;
    IWRegion60: TIWRegion;
    grdcomparacaoalt: TIWGrid;
    IWLink19: TIWLink;
    btnaltxalt: TIWButton;
    btncritxcrit: TIWButton;
    rgncritxcrit: TIWRegion;
    IWImage52: TIWImage;
    IWImage53: TIWImage;
    IWLabel14: TIWLabel;
    IWLabel17: TIWLabel;
    IWImage54: TIWImage;
    IWLink20: TIWLink;
    IWText11: TIWText;
    IWRegion61: TIWRegion;
    IWRegion62: TIWRegion;
    IWRegion63: TIWRegion;
    grdcritxcrit: TIWGrid;
    rgnaltxalt: TIWRegion;
    IWImage55: TIWImage;
    IWImage56: TIWImage;
    IWLabel18: TIWLabel;
    IWLabel19: TIWLabel;
    IWImage57: TIWImage;
    IWLink22: TIWLink;
    IWText12: TIWText;
    IWRegion64: TIWRegion;
    IWRegion65: TIWRegion;
    IWRegion66: TIWRegion;
    grdaltxalt: TIWGrid;
    procedure rgnmatconsCreate(Sender: TObject);
    procedure btnupdateClick(Sender: TObject);
    procedure cmbbxauxiliarAsyncChange(Sender: TObject;
      EventParams: TStringList);
    procedure cmbbxauxiliarChange(Sender: TObject);
    procedure btnshowresultsClick(Sender: TObject);
    procedure GerarCases;
    procedure Surrogate;
    procedure rgnresultadoCreate(Sender: TObject);
    procedure SetarCores;
    procedure btnimportClick(Sender: TObject);
    procedure IWLink1Click(Sender: TObject);
    procedure cmbbxfiltroChange(Sender: TObject);
    procedure IWLink2Click(Sender: TObject);
    procedure rgnfiltroCreate(Sender: TObject);
    procedure IWLink3Click(Sender: TObject);
    procedure rgnfiltro2Create(Sender: TObject);
    procedure IWLink4Click(Sender: TObject);
    procedure LimpaVariaveis;
    procedure lnkestatisticaClick(Sender: TObject);
    procedure rgnAnaliseEstatisticaCreate(Sender: TObject);
    procedure IWLink5Click(Sender: TObject);
    procedure ComparacaoParAPar;
    procedure Promethee;
    procedure ResultadosPromethee;
    procedure CalculoFluxos;
    procedure IWLink9Click(Sender: TObject);
    procedure lnkregrasdedecisaoPrometheeClick(Sender: TObject);
    procedure rgnresultadoPrometheeCreate(Sender: TObject);
    procedure IWLink6Click(Sender: TObject);
    procedure IWLink10Click(Sender: TObject);
    procedure lnkestatisticaPrometheeClick(Sender: TObject);
    procedure rgnAnaliseEstatisticaPrometheeCreate(Sender: TObject);
    procedure btnupdatevisualizacaoClick(Sender: TObject);
    procedure ComparacaoMetodos;
    procedure rgncomparacaoCreate(Sender: TObject);
    procedure btncomparacaoClick(Sender: TObject);
    procedure IWButton2Click(Sender: TObject);
    procedure lnkExportarClick(Sender: TObject);
    procedure lnkexportacaoprometheeClick(Sender: TObject);
    function TrocaPV(Num: String): Double;
    function TrocaVP(Num: Double): String;
    function trocavp1(Num: String): Double;
    procedure rgnregradedecisaoCreate(Sender: TObject);
    procedure IWLink13Click(Sender: TObject);
    procedure IWLink12Click(Sender: TObject);
    procedure IWLink11Click(Sender: TObject);
    procedure rgnregradedecisaoprometheeCreate(Sender: TObject);
    procedure IWLink15Click(Sender: TObject);
    procedure IWLink14Click(Sender: TObject);
    procedure btnforcacritClick(Sender: TObject);
    procedure btnestatisticaClick(Sender: TObject);
    procedure btncomparacaoprometheeClick(Sender: TObject);
    procedure btnestatisticaprometheeClick(Sender: TObject);
    procedure btnupdatevisualizacaoPrometheeClick(Sender: TObject);
    procedure IWLink16Click(Sender: TObject);
    procedure lnkdescricaolimiarClick(Sender: TObject);
    procedure btnatualizarlimiaresClick(Sender: TObject);
    procedure lnkdescricaolimiarPrometheeClick(Sender: TObject);
    procedure btnatualizarlimiaresPrometheeClick(Sender: TObject);
    procedure lnkregrasdedecisaoClick(Sender: TObject);
    procedure cmbbxp1Change(Sender: TObject);
    procedure cmbbxp2Change(Sender: TObject);
    procedure cmbbxp3Change(Sender: TObject);
    procedure cmbbxp4Change(Sender: TObject);
    procedure cmbbxp1prometheeChange(Sender: TObject);
    procedure cmbbxp2prometheeChange(Sender: TObject);
    procedure cmbbxp3prometheeChange(Sender: TObject);
    procedure cmbbxp4prometheeChange(Sender: TObject);
    procedure grdlimiaresComboChange(Sender: TObject;
      RowIndex, ColumnIndex: Integer; AValue: string);
    procedure grdlimiaresEdit(Sender: TObject);
    procedure DecrementaValorLimiar(j, k: Integer);
    procedure DecrementaValorLimiarPromethee(j, k: Integer);
    procedure EscalaRazao;
    procedure rgnintroCreate(Sender: TObject);
    procedure btnRegisterClick(Sender: TObject);
    procedure btncontinueRUClick(Sender: TObject);
    procedure btnBackRUClick(Sender: TObject);
    procedure btnBackRU2Click(Sender: TObject);
    procedure lnkesqueceusenhaClick(Sender: TObject);
    procedure btncontinueRU2Click(Sender: TObject);
    procedure btnEnterClick(Sender: TObject);
    procedure rgnEnterLoginCreate(Sender: TObject);
    procedure rgnOpcaoInputCreate(Sender: TObject);
    procedure rgnLoginprojectCreate(Sender: TObject);
    procedure IWLink7Click(Sender: TObject);
    procedure lnkbackloginClick(Sender: TObject);
    procedure btnLoginClick(Sender: TObject);
    procedure btnRegisternewClick(Sender: TObject);
    procedure btnContinueregistClick(Sender: TObject);
    procedure IWLink8Click(Sender: TObject);
    procedure btnauxchooseClick(Sender: TObject);
    procedure SaveProblem;
    procedure rgnmatconsRender(Sender: TObject);
    procedure edtnumaltSubmit(Sender: TObject);
    procedure edtnumaltAsyncExit(Sender: TObject; EventParams: TStringList);
    procedure AnaliseParaElicitacao;
    procedure SelecionarPergunta;
    procedure AtualizarPopulacao;
    procedure rgnelicitacaocritCreate(Sender: TObject);
    procedure IWLink18Click(Sender: TObject);
    procedure IWLink17Click(Sender: TObject);
    procedure IWLink19Click(Sender: TObject);
    procedure rgnelicitacaoaltCreate(Sender: TObject);
    procedure rgncritxcritCreate(Sender: TObject);
    procedure rgnaltxaltCreate(Sender: TObject);
    procedure btncritxcritClick(Sender: TObject);
    procedure IWLink20Click(Sender: TObject);
    procedure btnaltxaltClick(Sender: TObject);
    procedure IWLink22Click(Sender: TObject);
  public
  end;

implementation

uses ServerController;
{$R *.dfm}

Function TIWForm1.TrocaPV(Num: String): Double;
var
  S: string;
  i: Integer;
begin
  S := Num;
  for i := 2 to length(S) - 1 do
    if S[i] = '.' then
      S[i] := ',';

  Result := StrToFloat(S);
end;

Function TIWForm1.TrocaVP(Num: Double): String;
var
  S: string;
  i: Integer;
begin
  S := FloatToStr(Num);
  for i := 2 to length(S) - 1 do
    if S[i] = ',' then
      S[i] := '.';

  Result := S;
end;

Function TIWForm1.trocavp1(Num: String): Double;
var
  S: string;
  i: Integer;
begin
  S := Num;
  for i := 2 to length(S) - 1 do
    if S[i] = ',' then
      S[i] := '.';

  Result := StrToFloat(S);
end;

procedure TIWForm1.ComparacaoParAPar;
var
  xcrit, l, c: Integer;
begin

  SetLength(TIWUserSession(WebApplication.Data).matrizParaPar, TIWUserSession
      (WebApplication.Data).Numcrit, TIWUserSession(WebApplication.Data)
      .Numalt, TIWUserSession(WebApplication.Data).Numalt);
  xcrit := 0;
  for xcrit := 1 to TIWUserSession(WebApplication.Data).Numcrit do
    for l := 1 to TIWUserSession(WebApplication.Data).Numalt - 1 do
      for c := 1 to TIWUserSession(WebApplication.Data).Numalt - 1 do
        TIWUserSession(WebApplication.Data)
          .matrizParaPar[xcrit - 1, l - 1, c - 1] := 0;

  for xcrit := 1 to TIWUserSession(WebApplication.Data).Numcrit do
  begin
    if (TIWUserSession(WebApplication.Data).tipocrit[xcrit - 1] = 1) or
      (TIWUserSession(WebApplication.Data).tipocrit[xcrit - 1] = 3) or
      (TIWUserSession(WebApplication.Data).tipocrit[xcrit - 1] = 5) then
    begin
      for l := 1 to TIWUserSession(WebApplication.Data).Numalt do
      begin
        for c := 1 to TIWUserSession(WebApplication.Data).Numalt do
        begin
          TIWUserSession(WebApplication.Data).comp1 := TIWUserSession
            (WebApplication.Data).MatrizConseq[c - 1, xcrit - 1];
          { Alternativa a(j) }
          TIWUserSession(WebApplication.Data).comp2 := TIWUserSession
            (WebApplication.Data).MatrizConseq[l - 1, xcrit - 1];
          { Alternativa a(i) }
          if not(c = l) then
            if (TIWUserSession(WebApplication.Data).comp2 - TIWUserSession
                (WebApplication.Data).comp1 > 0) then
              TIWUserSession(WebApplication.Data)
                .matrizParaPar[xcrit - 1, l - 1, c - 1] := 1
              { preenche a matriz Par a Par do critério }
            else
              TIWUserSession(WebApplication.Data)
                .matrizParaPar[xcrit - 1, l - 1, c - 1] := 0;
          { preenche a matriz Par a Par do critério }
        end;
      end;
    end;
    if (TIWUserSession(WebApplication.Data).tipocrit[xcrit - 1] = 0) or
      (TIWUserSession(WebApplication.Data).tipocrit[xcrit - 1] = 2) or
      (TIWUserSession(WebApplication.Data).tipocrit[xcrit - 1] = 4) then
    begin

      for l := 1 to TIWUserSession(WebApplication.Data).Numalt do
      begin
        for c := 1 to TIWUserSession(WebApplication.Data).Numalt do
        begin
          TIWUserSession(WebApplication.Data).comp1 := TIWUserSession
            (WebApplication.Data).MatrizConseq[c - 1, xcrit - 1];
          { Alternativa a(j) }
          TIWUserSession(WebApplication.Data).comp2 := TIWUserSession
            (WebApplication.Data).MatrizConseq[l - 1, xcrit - 1];
          { Alternativa a(i) }
          if not(c = l) then
            if (TIWUserSession(WebApplication.Data).comp1 - TIWUserSession
                (WebApplication.Data).comp2 > 0) then
              TIWUserSession(WebApplication.Data)
                .matrizParaPar[xcrit - 1, l - 1, c - 1] := 1
              { preenche a matriz Par a Par do critério }
            else
              TIWUserSession(WebApplication.Data)
                .matrizParaPar[xcrit - 1, l - 1, c - 1] := 0;
          { preenche a matriz Par a Par do critério }
        end;
      end;
    end;
  end;
end;

procedure TIWForm1.CalculoFluxos;
var
  i, j, l, c, xcrit: Integer;
  soma, Min: real;

begin

  // matriz grau de sobreclassificação

  SetLength(TIWUserSession(WebApplication.Data).SobClassMatrix, TIWUserSession
      (WebApplication.Data).Numalt, TIWUserSession(WebApplication.Data).Numalt);

  for i := 0 to TIWUserSession(WebApplication.Data).Numalt - 1 do
    for j := 0 to TIWUserSession(WebApplication.Data).Numalt - 1 do
      TIWUserSession(WebApplication.Data).SobClassMatrix[i, j] := 0;

  for i := 0 to TIWUserSession(WebApplication.Data).Numalt - 1 do
  begin
    for j := 0 to TIWUserSession(WebApplication.Data).Numalt - 1 do
    begin
      for xcrit := 0 to TIWUserSession(WebApplication.Data).Numcrit - 1 do
      begin
        TIWUserSession(WebApplication.Data).SobClassMatrix[i, j] :=
          TIWUserSession(WebApplication.Data).SobClassMatrix[i, j]
          + TIWUserSession(WebApplication.Data).PesoCrit[xcrit] * TIWUserSession
          (WebApplication.Data).matrizParaPar[xcrit, i, j];
      end;
    end;
  end;

  // fluxo positivo proposto por Brans e Mareschal (2002)

  SetLength(TIWUserSession(WebApplication.Data).PositiveFlow, TIWUserSession
      (WebApplication.Data).Numalt);

  for l := 0 to TIWUserSession(WebApplication.Data).Numalt - 1 do
  { calcular o TOTAL dos produtos PesoDoCriterio (i) x AgregValor Critério (i) }
  begin
    soma := 0;
    for c := 0 to TIWUserSession(WebApplication.Data).Numalt - 1 do
    begin
      soma := soma + TIWUserSession(WebApplication.Data).SobClassMatrix[l, c];
    end;
    TIWUserSession(WebApplication.Data).PositiveFlow[l] := soma /
      (TIWUserSession(WebApplication.Data).Numalt - 1);
    { FLUXO POSITIVO - preenchimento }
  end;

  // fluxo negativo

  SetLength(TIWUserSession(WebApplication.Data).NegativeFlow, TIWUserSession
      (WebApplication.Data).Numalt);

  for c := 0 to TIWUserSession(WebApplication.Data).Numalt - 1 do
  { calcular o TOTAL dos produtos PesoDoCriterio (i) x AgregValor Critério (i) }
  begin
    soma := 0;
    for l := 0 to TIWUserSession(WebApplication.Data).Numalt - 1 do
    begin
      soma := soma + TIWUserSession(WebApplication.Data).SobClassMatrix[l, c];
    end;
    TIWUserSession(WebApplication.Data).NegativeFlow[c] := soma /
      (TIWUserSession(WebApplication.Data).Numalt - 1);
  end;

  // fluxo líquido

  SetLength(TIWUserSession(WebApplication.Data).NetFlow, TIWUserSession
      (WebApplication.Data).Numalt);

  for c := 0 to TIWUserSession(WebApplication.Data).Numalt - 1 do
  { calcular o TOTAL dos produtos PesoDoCriterio (i) x AgregValor Critério (i) }
  begin
    TIWUserSession(WebApplication.Data).NetFlow[c] := TIWUserSession
      (WebApplication.Data).PositiveFlow[c] - TIWUserSession
      (WebApplication.Data).NegativeFlow[c];
  end;

end;

procedure TIWForm1.ResultadosPromethee;
var
  i, j, k, qtdpareto: Integer;
  auxregra, aux1, aux2, aux3: real;
  bar: array of TBarSeries;
  SolRegra1, SolRegra2, SolRegra3: Integer;
  totalacumulado: real;
  epsilon, omega, faixaprob: array of real;
  recomendacao: boolean;
  Px: real;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    chrtresultadopromethee.RemoveAllSeries;
    chrtresultadopromethee.Legend.Items.Clear;
    // chrtresultadoPromethee.ClearChart;
    chrtresultadopromethee.BottomAxis.Items.Clear;
    chrtresultadopromethee.View3D := false;
    SetLength(bar, Numalt);
    i := 0;
    Total := length(CasesOrdemcrit);
    totalacumulado := 0;
    for j := 0 to Numalt - 1 do
    begin
      if totalacumulado > 0.8 * Total then
      begin
        bar[i] := TBarSeries.Create(self);
        chrtresultadopromethee.AddSeries(bar[i]);
        bar[i].AddXY(0, Total - totalacumulado, formatfloat
            ('0', Total - totalacumulado), coresalt[i]);
        // chrtResultadopromethee.BottomAxis.Items.Add(i, 'S' + inttostr(i + 1));

        chrtresultadopromethee.Legend.Item[i].Text := 'Other Alternatives';
        bar[i].Color := coresalt[i];
        break;
      end;
      if ResultAltPromethee[OrdemaltPromethee[j]] <> 0 then
      begin
        bar[i] := TBarSeries.Create(self);
        chrtresultadopromethee.AddSeries(bar[i]);
        bar[i].AddXY(0, ResultAltPromethee[OrdemaltPromethee[j]], formatfloat
            ('0', ResultAltPromethee[OrdemaltPromethee[j]]), coresalt[i]);
        // chrtResultadoPromethee.BottomAxis.Items.Add(i, 'S' + inttostr(i + 1));
        chrtresultadopromethee.Legend.Item[i].Text := nomealt
          [OrdemaltPromethee[j]];
        bar[i].Color := coresalt[i];
        totalacumulado := totalacumulado + ResultAltPromethee
          [OrdemaltPromethee[j]];
        inc(i);
      end;
    end;
    Total := 0;
    bar[1] := TBarSeries.Create(self);
    chrtresultadopromethee.AddSeries(bar[1]);
    for i := 0 to length(ResultSolPromethee) - 1 do
      Total := Total + ResultSolPromethee[i];
    bar[1].AddXY(0, Total, '', clWhite);
    bar[1].Visible := false;
    // chrtResultadoPromethee.LeftAxis.Increment := 0.2; // ?
    // chrtResultadoPromethee.Legend.Visible := false;
    chrtresultadopromethee.BottomAxis.LabelsAlternate;
    chrtresultadopromethee.BottomAxis.Items.Automatic := false;
    chrtresultadopromethee.LeftAxis.Maximum := Total;
    chrtresultadopromethee.LeftAxis.AutomaticMaximum := false;
    /// chrtResultadoPromethee.LeftAxis.MaximumOffset := 16;
    chrtresultadopromethee.LeftAxis.Minimum := 0;
    // chrtresultadoPromethee.leftaxis.Items.Automatic := false;
    chrtresultadopromethee.LeftAxis.Increment := Total;

    imgResultadoPromethee.Picture.Bitmap.Assign
      (chrtresultadopromethee.TeeCreateBitmap(clWebPURPLE, rect(0, 0,
          imgResultadoPromethee.Width, imgResultadoPromethee.Height)));

    grdresultadospromethee.Clear;
    NaltselPromethee := 0;
    for i := 0 to length(ResultAltPromethee) - 1 do
    begin
      if ResultAltPromethee[i] <> 0 then
        inc(NaltselPromethee);
    end;
    grdresultadospromethee.rowcount := 2 + NaltselPromethee;
    grdresultadospromethee.ColumnCount := 3;
    { if grdresultados.ColumnCount * 80 > 832 then
      grdresultados.Width := grdresultados.ColumnCount * 80
      else
      grdresultados.Width := 832; }
    grdresultadospromethee.Cell[0, 0].Text := 'Alternative';
    grdresultadospromethee.Cell[0, 1].Text := 'Frequency';
    grdresultadospromethee.Cell[0, 2].Text := '%';
    i := 0;
    for j := 0 to Numalt - 1 do
    begin
      if ResultAltPromethee[OrdemaltPromethee[j]] <> 0 then
      begin
        grdresultadospromethee.Cell[i + 1, 0].Text := nomealt
          [OrdemaltPromethee[j]];
        grdresultadospromethee.Cell[i + 1, 1].Text := FloatToStr
          (ResultAltPromethee[OrdemaltPromethee[j]]);
        grdresultadospromethee.Cell[i + 1, 2].Text := formatfloat
          ('0', 100 * ResultAltPromethee[OrdemaltPromethee[j]] / Total) + ' %';
        inc(i);
      end;
    end;
    // grdresultados.Cell[0, Numalt+1].Text := 'Cases';
    // grdresultados.Cell[0, Naltsel + 1].Text := 'Total';
    { for i := 0 to length(ResultSol) - 1 do
      grdresultados.Cell[i + 1, 0].Text := 'Sol ' + inttostr(i + 1) + ':'; }
    grdresultadospromethee.Cell[1 + i, 0].Text := 'Total:';
    grdresultadospromethee.Cell[1 + i, 1].Text := FloatToStr(Total);
    grdresultadospromethee.Cell[1 + i, 2].Text := '100%';
    { for i := 0 to length(ResultSol) - 1 do
      begin
      k := 0;
      for j := 0 to Numalt - 1 do
      begin
      if ResultAlt[j] <> 0 then
      begin
      if MatrizSol[OrdemSol[i], j] = 1 then
      grdresultados.Cell[i + 1, k + 1].Text := 'X'; // 'μ|Max|σ';
      inc(k);
      end;
      end;
      end; }

    {
      for i := 0 to length(ResultSol) - 1 do
      grdresultados.Cell[i + 1, Naltsel + 1].Text := floattostr
      (ResultSol[OrdemSol[i]]);
      }
    for j := 0 to grdresultadospromethee.ColumnCount - 1 do
      for i := 0 to grdresultadospromethee.rowcount - 1 do
      begin
        grdresultadospromethee.Cell[i, j].Alignment := taCenter;
        grdresultadospromethee.Cell[i, j].wrap := True;
        // grdresultados.Cell[0,j].Font.Style := [fsbold];
      end;

    for j := 0 to grdresultadospromethee.ColumnCount - 1 do
    begin
      grdresultadospromethee.Cell[0, j].Font.Style := [fsbold];
      grdresultadospromethee.Cell[1 + NaltselPromethee, j].Font.Color := clwebgray;
      grdresultadospromethee.Cell[1 + NaltselPromethee, j].BGColor := clWebWhiteSmoke;
    end;
    grdresultadospromethee.Cell[1 + NaltselPromethee, 0].Font.Style := [fsbold];

    { for I := 0 to length(ResultSol) - 1 do
      begin
      grdresultados.Cell[i+1, Numalt+1].Font.Size := 8;
      for j := 0 to Round(Total) - 1 do
      begin
      if Casesol[j] = i+1 then
      begin
      if grdresultados.Cell[i+1, Numalt+1].Text = '' then
      grdresultados.Cell[i+1, Numalt+1].Text := inttostr(j+1)
      else
      grdresultados.Cell[i+1, Numalt+1].Text := grdresultados.Cell[i+1, Numalt+1].Text + ', ' + inttostr(j+1);
      end;
      end;
      end; }

    grdcasespromethee.rowcount := Round(Total) + 1;
    grdcasespromethee.ColumnCount := 2 + Numcrit;

    for j := 0 to grdcasespromethee.ColumnCount - 1 do
      for i := 0 to grdcasespromethee.rowcount - 1 do
      begin
        grdcasespromethee.Cell[i, j].Alignment := taCenter;
        grdcasespromethee.Cell[i, j].wrap := True;
        // grdresultados.Cell[0,j].Font.Style := [fsbold];
      end;
    for j := 0 to grdcasespromethee.ColumnCount - 1 do
      grdcasespromethee.Cell[0, j].Font.Style := [fsbold];

    grdcasespromethee.Cell[0, 0].Text := 'Surrogate';
    for i := 0 to Numcrit - 1 do
      grdcasespromethee.Cell[0, i + 1].Text := Nomecrit[i];

    grdcasespromethee.Cell[0, Numcrit + 1].Text := 'Solution';

    for j := 0 to Round(Total) - 1 do
    begin
      grdcasespromethee.Cell[j + 1, 0].Text := 'S' + inttostr(j + 1);
      for i := 0 to Numalt - 1 do
      begin
        if matrizpoaPromethee[j, i] = 1 then
        begin
          grdcasespromethee.Cell[j + 1, Numcrit + 1].Text := nomealt[i] + ' ';
        end;
      end;
      for i := 0 to Numcrit - 1 do
        grdcasespromethee.Cell[j + 1, i + 1].Text := inttostr
          (CasesOrdemcrit[j, i]) + 'º'
    end;

    Total := 0;
    for i := 0 to length(ResultSolPromethee) - 1 do
      Total := Total + ResultSolPromethee[i];

    grdestatisticaPromethee.Clear;
    grdestatisticaPromethee.rowcount := 3 + length(MediaDifSolPromethee);
    grdestatisticaPromethee.ColumnCount := 4 + length(MediaDifSolPromethee);

    for i := 0 to length(MediaDifSolPromethee) - 1 do
    begin
      grdestatisticaPromethee.Cell[0, i + 4].Alignment := taCenter;
      // grdestatisticaPromethee.Cell[0, i+1].wrap := True;
      grdestatisticaPromethee.Cell[0, i + 4].Font.Style := [fsbold];
      grdestatisticaPromethee.Cell[1, i + 4].Alignment := taCenter;
      // grdestatisticaPromethee.Cell[1, i+1].wrap := True;
      grdestatisticaPromethee.Cell[1, i + 4].Font.Style := [fsbold];
      grdestatisticaPromethee.Cell[0, i + 4].Text := 'Sol ' + inttostr(i + 1);
      grdestatisticaPromethee.Cell[0, i + 4].Font.Color := clwebgray;
      grdestatisticaPromethee.Cell[1, i + 4].Text := ' μ | σ ';

      grdestatisticaPromethee.Cell[i + 2, 3].Alignment := taCenter;
      grdestatisticaPromethee.Cell[i + 2, 3].wrap := True;
      grdestatisticaPromethee.Cell[i + 2, 3].Font.Style := [fsbold];
      grdestatisticaPromethee.Cell[i + 2, 3].Font.Color := clwebgray;
      for j := 0 to Numalt - 1 do
      begin
        if matrizsolPromethee[ordemsolPromethee[i], j] = 1 then
        begin
          if grdestatisticaPromethee.Cell[i + 2, 3].Text = '' then
            grdestatisticaPromethee.Cell[i + 2, 3].Text := nomealt[j]
          else
            grdestatisticaPromethee.Cell[i + 2, 3].Text :=
              grdestatisticaPromethee.Cell[i + 2, 3].Text + ' + ' + nomealt[j];
        end;
      end;
      grdestatisticaPromethee.Cell[0, 4 + i].Text :=
        grdestatisticaPromethee.Cell[i + 2, 3].Text;

      grdestatisticaPromethee.Cell[i + 2, 1].Alignment := taCenter;
      grdestatisticaPromethee.Cell[i + 2, 1].wrap := True;
      // grdestatisticaPromethee.Cell[i + 2, 1].Font.Style := [fsbold];
      grdestatisticaPromethee.Cell[i + 2, 1].Font.Color := clWebBlue;
      grdestatisticaPromethee.Cell[i + 2, 1].Text := formatfloat
        ('0.00', MediaGeralPromethee[ordemsolPromethee[i]]);

      grdestatisticaPromethee.Cell[i + 2, 2].Alignment := taCenter;
      grdestatisticaPromethee.Cell[i + 2, 2].wrap := True;
      // grdestatisticaPromethee.Cell[i + 2, 1].Font.Style := [fsbold];
      grdestatisticaPromethee.Cell[i + 2, 2].Font.Color := clWebBlue;
      grdestatisticaPromethee.Cell[i + 2, 2].Text := formatfloat
        ('0.00', DesvioPadraoGeralPromethee[ordemsolPromethee[i]]);

      grdestatisticaPromethee.Cell[i + 2, 0].Alignment := taCenter;
      grdestatisticaPromethee.Cell[i + 2, 0].wrap := True;
      // grdestatisticaPromethee.Cell[i + 2, 1].Font.Style := [fsbold];
      grdestatisticaPromethee.Cell[i + 2, 0].Font.Color := clWebBlue;
      grdestatisticaPromethee.Cell[i + 2, 0].Text := formatfloat
        ('0.00', (ResultSolPromethee[ordemsolPromethee[i]] / Total) * 100)
        + '%';
    end;

    grdestatisticaPromethee.Cell[1, 1].Alignment := taCenter;
    grdestatisticaPromethee.Cell[1, 1].wrap := True;
    grdestatisticaPromethee.Cell[1, 1].Font.Style := [fsbold];
    grdestatisticaPromethee.Cell[1, 1].Font.Color := clWebBlue;
    grdestatisticaPromethee.Cell[1, 1].Text := 'μ';

    grdestatisticaPromethee.Cell[1, 2].Alignment := taCenter;
    grdestatisticaPromethee.Cell[1, 2].wrap := True;
    grdestatisticaPromethee.Cell[1, 2].Font.Style := [fsbold];
    grdestatisticaPromethee.Cell[1, 2].Font.Color := clWebBlue;
    grdestatisticaPromethee.Cell[1, 2].Text := 'σ';

    grdestatisticaPromethee.Cell[1, 0].Alignment := taCenter;
    grdestatisticaPromethee.Cell[1, 0].wrap := True;
    grdestatisticaPromethee.Cell[1, 0].Font.Style := [fsbold];
    grdestatisticaPromethee.Cell[1, 0].Font.Color := clWebBlue;
    grdestatisticaPromethee.Cell[1, 0].Text := '%';

    grdestatisticaPromethee.Cell[2 + length(MediaDifSolPromethee), 3]
      .Alignment := taCenter;
    grdestatisticaPromethee.Cell[2 + length(MediaDifSolPromethee), 3].wrap :=
      True;
    grdestatisticaPromethee.Cell[2 + length(MediaDifSolPromethee), 3]
      .Font.Style := [fsbold];
    grdestatisticaPromethee.Cell[2 + length(MediaDifSolPromethee), 3]
      .Font.Color := clWebBlue;
    grdestatisticaPromethee.Cell[2 + length(MediaDifSolPromethee), 3].Text :=
      'μ:';

    { grdestatisticaPromethee.Cell[3+length(MediaDifSolPromethee), 3].Alignment := taCenter;
      grdestatisticaPromethee.Cell[3+length(MediaDifSolPromethee), 3].wrap := True;
      grdestatisticaPromethee.Cell[3+length(MediaDifSolPromethee), 3].Font.Style := [fsbold];
      grdestatisticaPromethee.Cell[3+length(MediaDifSolPromethee), 3].Font.Color := clWebBlue;
      grdestatisticaPromethee.Cell[3+length(MediaDifSolPromethee), 3].Text := 'σ:'; }

    for i := 0 to length(MediaDifSolPromethee) - 1 do
    begin
      grdestatisticaPromethee.Cell[2 + length(MediaDifSolPromethee), 4 + i]
        .Text := formatfloat('0.00',
        MediaGeralNaoSolPromethee[ordemsolPromethee[i]]);
      grdestatisticaPromethee.Cell[2 + length(MediaDifSolPromethee), 4 + i]
        .Alignment := taCenter;
      grdestatisticaPromethee.Cell[2 + length(MediaDifSolPromethee), 4 + i]
        .wrap := True;
      grdestatisticaPromethee.Cell[2 + length(MediaDifSolPromethee), 4 + i]
        .Font.Color := clWebBlue;

      { grdestatisticaPromethee.Cell[3+length(MediaDifSolPromethee), 4+i].Text := formatfloat('0.00', DesvioPadraoGeralNaoSolPromethee[OrdemSolPromethee[i]]);
        grdestatisticaPromethee.Cell[3+length(MediaDifSolPromethee), 4+i].Alignment := taCenter;
        grdestatisticaPromethee.Cell[3+length(MediaDifSolPromethee), 4+i].wrap := True;
        grdestatisticaPromethee.Cell[3+length(MediaDifSolPromethee), 4+i].Font.Color := clWebBlue; }
    end;

    for i := 0 to length(MediaDifSolPromethee) - 1 do
      for j := 0 to length(MediaDifSolPromethee) - 1 do
      begin
        grdestatisticaPromethee.Cell[i + 2, j + 4].Alignment := taCenter;
        grdestatisticaPromethee.Cell[i + 2, j + 4].wrap := True;
        if i = j then
          grdestatisticaPromethee.Cell[i + 2, j + 4].Text := ''
        else
        begin
          grdestatisticaPromethee.Cell[i + 2, j + 4].Text := formatfloat
            ('0.00', MediaDifSolPromethee[ordemsolPromethee[i],
            ordemsolPromethee[j]]) + '|' + formatfloat('0.00',
            DesvioPadraoDifSolPromethee[ordemsolPromethee[i],
            ordemsolPromethee[j]]);
        end;
      end;

    grdmaxminPromethee.Clear;
    grdmaxminPromethee.rowcount := 7 + length(MediaDifSolPromethee);
    grdmaxminPromethee.ColumnCount := 5 + length(MediaDifSolPromethee);

    for i := 0 to length(MediaDifSolPromethee) - 1 do
    begin
      grdmaxminPromethee.Cell[0, i + 5].Alignment := taCenter;
      // grdmaxminPromethee.Cell[0, i+1].wrap := True;
      grdmaxminPromethee.Cell[0, i + 5].Font.Style := [fsbold];
      grdmaxminPromethee.Cell[1, i + 5].Alignment := taCenter;
      // grdmaxminPromethee.Cell[1, i+1].wrap := True;
      grdmaxminPromethee.Cell[1, i + 5].Font.Style := [fsbold];
      // grdmaxminPromethee.Cell[0, i + 5].Text := 'Sol ' + inttostr(i + 1);
      grdmaxminPromethee.Cell[0, i + 5].Font.Color := clwebgray;
      grdmaxminPromethee.Cell[1, i + 5].Text := '  Max | Min ';

      grdmaxminPromethee.Cell[i + 2, 4].Alignment := taCenter;
      grdmaxminPromethee.Cell[i + 2, 4].wrap := True;
      grdmaxminPromethee.Cell[i + 2, 4].Font.Style := [fsbold];
      grdmaxminPromethee.Cell[i + 2, 4].Font.Color := clwebgray;
      for j := 0 to Numalt - 1 do
      begin
        if matrizsolPromethee[ordemsolPromethee[i], j] = 1 then
        begin
          if grdmaxminPromethee.Cell[i + 2, 4].Text = '' then
            grdmaxminPromethee.Cell[i + 2, 4].Text := nomealt[j]
          else
            grdmaxminPromethee.Cell[i + 2, 4].Text := grdmaxminPromethee.Cell
              [i + 2, 4].Text + ' + ' + nomealt[j];
        end;
      end;
      grdmaxminPromethee.Cell[0, 5 + i].Text := grdmaxminPromethee.Cell
        [i + 2, 4].Text;

      grdmaxminPromethee.Cell[i + 2, 1].Alignment := taCenter;
      grdmaxminPromethee.Cell[i + 2, 1].wrap := True;
      // grdestatisticaPromethee.Cell[i + 2, 1].Font.Style := [fsbold];
      grdmaxminPromethee.Cell[i + 2, 1].Font.Color := clWebBlue;
      grdmaxminPromethee.Cell[i + 2, 1].Text := formatfloat
        ('0.00', MinimoGeralPromethee[ordemsolPromethee[i]]);

      grdmaxminPromethee.Cell[i + 2, 2].Alignment := taCenter;
      grdmaxminPromethee.Cell[i + 2, 2].wrap := True;
      // grdestatisticaPromethee.Cell[i + 2, 1].Font.Style := [fsbold];
      grdmaxminPromethee.Cell[i + 2, 2].Font.Color := clWebBlue;
      grdmaxminPromethee.Cell[i + 2, 2].Text := formatfloat
        ('0.00', MaxMinPromethee[ordemsolPromethee[i]]);

      grdmaxminPromethee.Cell[i + 2, 3].Alignment := taCenter;
      grdmaxminPromethee.Cell[i + 2, 3].wrap := True;
      // grdestatisticaPromethee.Cell[i + 2, 1].Font.Style := [fsbold];
      grdmaxminPromethee.Cell[i + 2, 3].Font.Color := clWebBlue;
      grdmaxminPromethee.Cell[i + 2, 3].Text := formatfloat
        ('0.00', MinMaxPromethee[ordemsolPromethee[i]]);

      grdmaxminPromethee.Cell[i + 2, 0].Alignment := taCenter;
      grdmaxminPromethee.Cell[i + 2, 0].wrap := True;
      // grdestatisticaPromethee.Cell[i + 2, 1].Font.Style := [fsbold];
      grdmaxminPromethee.Cell[i + 2, 0].Font.Color := clWebBlue;
      grdmaxminPromethee.Cell[i + 2, 0].Text := formatfloat
        ('0.00', MaximoGeralPromethee[ordemsolPromethee[i]]);

      grdmaxminPromethee.Cell[1, 1].Alignment := taCenter;
      grdmaxminPromethee.Cell[1, 1].wrap := True;
      grdmaxminPromethee.Cell[1, 1].Font.Style := [fsbold];
      grdmaxminPromethee.Cell[1, 1].Font.Color := clWebBlue;
      grdmaxminPromethee.Cell[1, 1].Text := 'Min-Min';

      grdmaxminPromethee.Cell[1, 2].Alignment := taCenter;
      grdmaxminPromethee.Cell[1, 2].wrap := True;
      grdmaxminPromethee.Cell[1, 2].Font.Style := [fsbold];
      grdmaxminPromethee.Cell[1, 2].Font.Color := clWebBlue;
      grdmaxminPromethee.Cell[1, 2].Text := 'Max-Min';

      grdmaxminPromethee.Cell[1, 3].Alignment := taCenter;
      grdmaxminPromethee.Cell[1, 3].wrap := True;
      grdmaxminPromethee.Cell[1, 3].Font.Style := [fsbold];
      grdmaxminPromethee.Cell[1, 3].Font.Color := clWebBlue;
      grdmaxminPromethee.Cell[1, 3].Text := 'Min-Max';

      grdmaxminPromethee.Cell[1, 0].Alignment := taCenter;
      grdmaxminPromethee.Cell[1, 0].wrap := True;
      grdmaxminPromethee.Cell[1, 0].Font.Style := [fsbold];
      grdmaxminPromethee.Cell[1, 0].Font.Color := clWebBlue;
      grdmaxminPromethee.Cell[1, 0].Text := 'Max-Max';

    end;

    grdmaxminPromethee.Cell[2 + length(MediaDifSolPromethee), 4].Alignment :=
      taCenter;
    grdmaxminPromethee.Cell[2 + length(MediaDifSolPromethee), 4].wrap := True;
    grdmaxminPromethee.Cell[2 + length(MediaDifSolPromethee), 4].Font.Style :=
      [fsbold];
    grdmaxminPromethee.Cell[2 + length(MediaDifSolPromethee), 4].Font.Color :=
      clWebBlue;
    grdmaxminPromethee.Cell[2 + length(MediaDifSolPromethee), 4].Text :=
      'Max-Max:';

    grdmaxminPromethee.Cell[3 + length(MediaDifSolPromethee), 4].Alignment :=
      taCenter;
    grdmaxminPromethee.Cell[3 + length(MediaDifSolPromethee), 4].wrap := True;
    grdmaxminPromethee.Cell[3 + length(MediaDifSolPromethee), 4].Font.Style :=
      [fsbold];
    grdmaxminPromethee.Cell[3 + length(MediaDifSolPromethee), 4].Font.Color :=
      clWebBlue;
    grdmaxminPromethee.Cell[3 + length(MediaDifSolPromethee), 4].Text :=
      'Min-Min:';

    grdmaxminPromethee.Cell[4 + length(MediaDifSolPromethee), 4].Alignment :=
      taCenter;
    grdmaxminPromethee.Cell[4 + length(MediaDifSolPromethee), 4].wrap := True;
    grdmaxminPromethee.Cell[4 + length(MediaDifSolPromethee), 4].Font.Style :=
      [fsbold];
    grdmaxminPromethee.Cell[4 + length(MediaDifSolPromethee), 4].Font.Color :=
      clWebBlue;
    grdmaxminPromethee.Cell[4 + length(MediaDifSolPromethee), 4].Text :=
      'Max-Min:';

    grdmaxminPromethee.Cell[5 + length(MediaDifSolPromethee), 4].Alignment :=
      taCenter;
    grdmaxminPromethee.Cell[5 + length(MediaDifSolPromethee), 4].wrap := True;
    grdmaxminPromethee.Cell[5 + length(MediaDifSolPromethee), 4].Font.Style :=
      [fsbold];
    grdmaxminPromethee.Cell[5 + length(MediaDifSolPromethee), 4].Font.Color :=
      clWebBlue;
    grdmaxminPromethee.Cell[5 + length(MediaDifSolPromethee), 4].Text :=
      'Min-Max:';

    grdmaxminPromethee.Cell[6 + length(MediaDifSolPromethee), 4].Alignment :=
      taCenter;
    grdmaxminPromethee.Cell[6 + length(MediaDifSolPromethee), 4].wrap := True;
    grdmaxminPromethee.Cell[6 + length(MediaDifSolPromethee), 4].Font.Style :=
      [fsbold];
    grdmaxminPromethee.Cell[6 + length(MediaDifSolPromethee), 4].Font.Color :=
      clWebBlue;
    grdmaxminPromethee.Cell[6 + length(MediaDifSolPromethee), 4].Text :=
      'Min-Max-Max:';

    minmaxmaxnaosolPromethee := 99999999999;
    for i := 0 to length(MediaDifSolPromethee) - 1 do
    begin
      if minmaxmaxnaosolPromethee > MaximoGeralNaosolPromethee
        [ordemsolPromethee[i]] then
      begin
        TIWUserSession(WebApplication.Data).minmaxmaxnaosolPromethee :=
          MaximoGeralNaosolPromethee[ordemsolPromethee[i]];
        j := i;
      end;
    end;
    grdmaxminPromethee.Cell[6 + length(MediaDifSolPromethee), 5 + j].Text :=
      formatfloat('0.00', minmaxmaxnaosolPromethee);
    grdmaxminPromethee.Cell[6 + length(MediaDifSolPromethee), 5 + j]
      .Alignment := taCenter;
    grdmaxminPromethee.Cell[6 + length(MediaDifSolPromethee), 5 + j].wrap :=
      True;
    grdmaxminPromethee.Cell[6 + length(MediaDifSolPromethee), 5 + j]
      .Font.Color := clWebBlue;

    for i := 0 to length(MediaDifSolPromethee) - 1 do
    begin
      grdmaxminPromethee.Cell[2 + length(MediaDifSolPromethee), 5 + i].Text :=
        formatfloat('0.00', MaximoGeralNaosolPromethee[ordemsolPromethee[i]]);
      grdmaxminPromethee.Cell[2 + length(MediaDifSolPromethee), 5 + i]
        .Alignment := taCenter;
      grdmaxminPromethee.Cell[2 + length(MediaDifSolPromethee), 5 + i].wrap :=
        True;
      grdmaxminPromethee.Cell[2 + length(MediaDifSolPromethee), 5 + i]
        .Font.Color := clWebBlue;

      grdmaxminPromethee.Cell[3 + length(MediaDifSolPromethee), 5 + i].Text :=
        formatfloat('0.00', MinimoGeralNaoSolPromethee[ordemsolPromethee[i]]);
      grdmaxminPromethee.Cell[3 + length(MediaDifSolPromethee), 5 + i]
        .Alignment := taCenter;
      grdmaxminPromethee.Cell[3 + length(MediaDifSolPromethee), 5 + i].wrap :=
        True;
      grdmaxminPromethee.Cell[3 + length(MediaDifSolPromethee), 5 + i]
        .Font.Color := clWebBlue;

      grdmaxminPromethee.Cell[4 + length(MediaDifSolPromethee), 5 + i].Text :=
        formatfloat('0.00', MaxMinNaoSolPromethee[ordemsolPromethee[i]]);
      grdmaxminPromethee.Cell[4 + length(MediaDifSolPromethee), 5 + i]
        .Alignment := taCenter;
      grdmaxminPromethee.Cell[4 + length(MediaDifSolPromethee), 5 + i].wrap :=
        True;
      grdmaxminPromethee.Cell[4 + length(MediaDifSolPromethee), 5 + i]
        .Font.Color := clWebBlue;

      grdmaxminPromethee.Cell[5 + length(MediaDifSolPromethee), 5 + i].Text :=
        formatfloat('0.00', MinMaxNaoSolPromethee[ordemsolPromethee[i]]);
      grdmaxminPromethee.Cell[5 + length(MediaDifSolPromethee), 5 + i]
        .Alignment := taCenter;
      grdmaxminPromethee.Cell[5 + length(MediaDifSolPromethee), 5 + i].wrap :=
        True;
      grdmaxminPromethee.Cell[5 + length(MediaDifSolPromethee), 5 + i]
        .Font.Color := clWebBlue;

      { grdmaxminPromethee.Cell[3+length(MediaDifSolPromethee), 4+i].Text := formatfloat('0.00', DesvioPadraoGeralNaoSolPromethee[OrdemSolPromethee[i]]);
        grdmaxminPromethee.Cell[3+length(MediaDifSolPromethee), 4+i].Alignment := taCenter;
        grdmaxminPromethee.Cell[3+length(MediaDifSolPromethee), 4+i].wrap := True;
        grdmaxminPromethee.Cell[3+length(MediaDifSolPromethee), 4+i].Font.Color := clWebBlue; }
    end;

    for i := 0 to length(MediaDifSolPromethee) - 1 do
      for j := 0 to length(MediaDifSolPromethee) - 1 do
      begin
        grdmaxminPromethee.Cell[i + 2, j + 5].Alignment := taCenter;
        grdmaxminPromethee.Cell[i + 2, j + 5].wrap := True;
        if i = j then
          grdmaxminPromethee.Cell[i + 2, j + 5].Text := ''
        else
        begin
          grdmaxminPromethee.Cell[i + 2, j + 5].Text := formatfloat
            ('0.00', MaxDifSolPromethee[ordemsolPromethee[i],
            ordemsolPromethee[j]]) + '|' + formatfloat
            ('0.00', MinDifSolPromethee[ordemsolPromethee[i],
            ordemsolPromethee[j]]);
        end;
      end;

    auxregra := 0;
    qtdpareto := 0;
    while auxregra < 0.8 * Total do
    begin
      auxregra := auxregra + ResultSolPromethee[ordemsolPromethee[qtdpareto]];
      inc(qtdpareto);
    end;

    grdperdamediaemaximapromethee.Clear;
    grdperdamediaemaximapromethee.ColumnCount := 4;
    grdperdamediaemaximapromethee.rowcount := 1 + qtdpareto;
    for i := 0 to grdperdamediaemaximapromethee.ColumnCount - 1 do
    begin
      for j := 0 to grdperdamediaemaximapromethee.rowcount - 1 do
      begin
        grdperdamediaemaximapromethee.Cell[j, i].Alignment := taCenter;
        grdperdamediaemaximapromethee.Cell[j, i].wrap := True;
        if j = 0 then
        begin
          grdperdamediaemaximapromethee.Cell[j, i].BGColor := clWebWhiteSmoke;
          grdperdamediaemaximapromethee.Cell[j, i].Font.Style := [fsbold];
        end;
      end;
    end;
    grdperdamediaemaximapromethee.Cell[0, 0].Text := 'Alternatives';
    grdperdamediaemaximapromethee.Cell[0, 1].Text := 'Probability[P(Ai)]';
    grdperdamediaemaximapromethee.Cell[0, 2].Text := 'Average Loss[μ(Aj,Ai)]:';
    grdperdamediaemaximapromethee.Cell[0, 3].Text :=
      'Maximum Loss[Dmax(Aj, Ai)]:';
    for j := 0 to qtdpareto - 1 do
    begin
      for i := 0 to Numalt - 1 do
      begin
        if matrizsolPromethee[ordemsolPromethee[j], i] = 1 then
        begin
          if grdperdamediaemaximapromethee.Cell[1 + j, 0].Text = '' then
            grdperdamediaemaximapromethee.Cell[1 + j, 0].Text := nomealt[i]
          else
            grdperdamediaemaximapromethee.Cell[1 + j, 0].Text :=
              grdperdamediaemaximapromethee.Cell[1 + j, 0]
              .Text + ' + ' + nomealt[i];
        end;
        grdperdamediaemaximapromethee.Cell[j + 1, 1].Text := formatfloat
          ('0', (ResultSolPromethee[ordemsolPromethee[j]] / Total) * 100) + '%';
        grdperdamediaemaximapromethee.Cell[j + 1, 2].Text := formatfloat
          ('0.00', MediaGeralNaoSolPromethee[ordemsolPromethee[j]]);
        grdperdamediaemaximapromethee.Cell[j + 1, 3].Text := formatfloat
          ('0.00', MaximoGeralNaosolPromethee[ordemsolPromethee[j]]);
      end;
    end;

    grdregradedecisaopromethee.Clear;
    grdregradedecisaopromethee.ColumnCount := 1 + qtdpareto;
    grdregradedecisaopromethee.rowcount := 11;
    for i := 0 to grdregradedecisaopromethee.ColumnCount - 1 do
    begin
      if i >= 1 + qtdpareto then
      begin
        grdregradedecisaopromethee.Cell[0, i].Font.Color := ClWebRed;
        grdregradedecisaopromethee.Cell[0, i].Font.Style := [fsbold];
      end;
      for j := 0 to grdregradedecisaopromethee.rowcount - 1 do
      begin
        grdregradedecisaopromethee.Cell[j, i].Alignment := taCenter;
        grdregradedecisaopromethee.Cell[j, i].wrap := True;
        if i = 0 then
          grdregradedecisaopromethee.Cell[j, 0].Font.Style := [fsbold];

        if j = 0 then
          grdregradedecisaopromethee.Cell[j, i].BGColor := clWebWhiteSmoke;

        if j >= 8 then
        begin
          grdregradedecisaopromethee.Cell[j, i].Font.Color := ClWebRed;
          grdregradedecisaopromethee.Cell[j, i].BGColor := clWebWhiteSmoke;
        end;
      end;

    end;

    grdregradedecisaopromethee.Cell[0, 0].Text := 'Alternative:';
    grdregradedecisaopromethee.Cell[1, 0].Text := 'Probability [P(Ai)]:';
    grdregradedecisaopromethee.Cell[2, 0].Text := 'Average Loss[μ(Aj,Ai)]:';
    grdregradedecisaopromethee.Cell[3, 0].Text := 'Average Gain[μ(Ai,Aj)]:';
    grdregradedecisaopromethee.Cell[4, 0].Text :=
      'Expected Loss[[1-P(Ai)]*μ(Aj,Ai)]:';
    grdregradedecisaopromethee.Cell[5, 0].Text :=
      'Expected Gain[P(Ai)* μ(Aj,Ai)]:';
    grdregradedecisaopromethee.Cell[6, 0].Text := 'Maximum Loss[Dmax(Aj, Ai)]:';
    grdregradedecisaopromethee.Cell[7, 0].Text := 'Global Performance[G(Ai)]:';

    grdregradedecisaopromethee.Cell[8, 0].Text := 'Minimum Avarage Loss:';
    grdregradedecisaopromethee.Cell[9, 0].Text := 'Minimum Maximum Loss:';
    grdregradedecisaopromethee.Cell[10, 0].Text :=
      'Maximum Global Performance:';

    aux1 := 99999999999;
    aux2 := 99999999999;
    aux3 := -9999999999;

    for i := 0 to qtdpareto - 1 do
    begin

      for j := 0 to Numalt - 1 do
      begin
        if matrizsolPromethee[ordemsolPromethee[i], j] = 1 then
        begin
          if grdregradedecisaopromethee.Cell[0, 1 + i].Text = '' then
            grdregradedecisaopromethee.Cell[0, 1 + i].Text := nomealt[j]
          else
            grdregradedecisaopromethee.Cell[0, 1 + i].Text :=
              grdregradedecisaopromethee.Cell[0, 1 + i].Text + ' + ' + nomealt
              [j];
        end;
      end;

      // grdregradedecisao.Cell[0,1 + i].Text := 'Sol ' + inttostr(i + 1);
      grdregradedecisaopromethee.Cell[1, 1 + i].Text := formatfloat
        ('0', (ResultSolPromethee[ordemsolPromethee[i]] / Total) * 100) + '%';
      grdregradedecisaopromethee.Cell[2, 1 + i].Text := formatfloat
        ('0.00', MediaGeralNaoSolPromethee[ordemsolPromethee[i]]);
      grdregradedecisaopromethee.Cell[3, 1 + i].Text := formatfloat
        ('0.00', MediaGeralPromethee[ordemsolPromethee[i]]);
      grdregradedecisaopromethee.Cell[4, 1 + i].Text := formatfloat
        ('0.00', MediaGeralNaoSolPromethee[ordemsolPromethee[i]] *
          (1 - (ResultSolPromethee[ordemsolPromethee[i]] / Total)));
      grdregradedecisaopromethee.Cell[5, 1 + i].Text := formatfloat
        ('0.00', MediaGeralPromethee[ordemsolPromethee[i]] *
          (ResultSolPromethee[ordemsolPromethee[i]] / Total));
      grdregradedecisaopromethee.Cell[6, 1 + i].Text := formatfloat
        ('0.00', MaximoGeralNaosolPromethee[ordemsolPromethee[i]]);
      auxregra := 0;
      auxregra := MediaGeralPromethee[ordemsolPromethee[i]] *
        (ResultSolPromethee[ordemsolPromethee[i]] / Total)
        - MediaGeralNaoSolPromethee[ordemsolPromethee[i]] *
        ((Total - ResultSolPromethee[ordemsolPromethee[i]]) / Total);
      grdregradedecisaopromethee.Cell[7, 1 + i].Text := formatfloat
        ('0.00', auxregra);

      if MediaGeralNaoSolPromethee[ordemsolPromethee[i]] < aux1 then
      begin

        aux1 := MediaGeralNaoSolPromethee[ordemsolPromethee[i]];
        SolRegra1 := i;
      end;
      if MaximoGeralNaosolPromethee[ordemsolPromethee[i]] < aux2 then
      begin
        aux2 := MaximoGeralNaosolPromethee[ordemsolPromethee[i]];
        SolRegra2 := i;
      end;
      if (MediaGeralPromethee[ordemsolPromethee[i]] *
          (ResultSolPromethee[ordemsolPromethee[i]] / Total)
          - MediaGeralNaoSolPromethee[ordemsolPromethee[i]] *
          ((Total - ResultSolPromethee[ordemsolPromethee[i]]) / Total))
        > aux3 then
      begin
        aux3 := (MediaGeralPromethee[ordemsolPromethee[i]] *
            (ResultSolPromethee[ordemsolPromethee[i]] / Total)
            - MediaGeralNaoSolPromethee[ordemsolPromethee[i]] *
            ((Total - ResultSolPromethee[ordemsolPromethee[i]]) / Total));
        SolRegra3 := i;
      end;

    end;

    grdregradedecisaopromethee.Cell[8, 1 + SolRegra1].Text := formatfloat
      ('0.00', MediaGeralNaoSolPromethee[ordemsolPromethee[SolRegra1]]);

    grdregradedecisaopromethee.Cell[9, 1 + SolRegra2].Text := formatfloat
      ('0.00', MaximoGeralNaosolPromethee[ordemsolPromethee[SolRegra2]]);

    grdregradedecisaopromethee.Cell[10, 1 + SolRegra3].Text := formatfloat
      ('0.00', (MediaGeralPromethee[ordemsolPromethee[SolRegra3]] *
          (ResultSolPromethee[ordemsolPromethee[SolRegra3]] / Total)
          - MediaGeralNaoSolPromethee[ordemsolPromethee[SolRegra3]] *
          ((Total - ResultSolPromethee[ordemsolPromethee[SolRegra3]]) / Total))
      );

    SetLength(epsilon, 4);
    SetLength(omega, 4);
    SetLength(faixaprob, 4);

    epsilon[0] := 1;
    epsilon[1] := 0.4;
    epsilon[2] := 0.3;
    epsilon[3] := 0.2;

    omega[0] := 1;
    omega[1] := 0.2;
    omega[2] := 0.15;
    omega[3] := 0.10;

    faixaprob[0] := 80;
    faixaprob[1] := 70;
    faixaprob[2] := 60;
    faixaprob[3] := 50;

    { grdlimiares.Cells[0, 0] := '80% < P(x)';
      grdlimiares.Cells[0, 1] := '70% < P(x) < 80%';
      grdlimiares.Cells[0, 2] := '60% < P(x) < 70%';
      grdlimiares.Cells[0, 3] := '50% < P(x) < 60%';
      }
    grdlimiarespromethee.DynEdits[1, 0] := '0.50';
    grdlimiarespromethee.DynEdits[1, 1] := '0.40';
    grdlimiarespromethee.DynEdits[1, 2] := '0.30';
    grdlimiarespromethee.DynEdits[1, 3] := '0.20';

    grdlimiarespromethee.DynEdits[2, 0] := '0.25';
    grdlimiarespromethee.DynEdits[2, 1] := '0.20';
    grdlimiarespromethee.DynEdits[2, 2] := '0.15';
    grdlimiarespromethee.DynEdits[2, 3] := '0.10';

    recomendacao := false;

    // Para alt1
    Px := ((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
          [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
        / Total) * 100;

    recomendacao := false;

    grdrecomendacaopromethee.Clear;
    grdrecomendacaopromethee.rowcount := 4;
    grdrecomendacaopromethee.ColumnCount := 2;
    for i := 0 to grdrecomendacaopromethee.ColumnCount - 1 do
    begin

      grdrecomendacaopromethee.Cell[0, i].Font.Color := ClWebRed;
      grdrecomendacaopromethee.Cell[0, i].Font.Style := [fsbold];
      grdrecomendacaopromethee.Cell[0, i].BGColor := clWebWhiteSmoke;
      for j := 0 to grdrecomendacaopromethee.rowcount - 1 do
      begin
        grdrecomendacaopromethee.Cell[j, i].Alignment := taCenter;
        grdrecomendacaopromethee.Cell[j, i].wrap := True;
      end;
    end;

    grdrecomendacaopromethee.Cell[0, 0].Text := 'Solution';
    grdrecomendacaopromethee.Cell[0, 1].Text := 'Possibles Recommendations';
    grdrecomendacaopromethee.Cell[1, 1].Text := 'Best Alternative';
    grdrecomendacaopromethee.Cell[2, 1].Text :=
      'Two/Three Alternatives are Competitive';
    grdrecomendacaopromethee.Cell[3, 1].Text := 'Unable to make it';
    // txtdescricaoregrapromethee.Lines[2] := 'P(x) = ' + 'P(' + grdperdamediaemaximapromethee.Cell[1, 0].Text + ')' + ' + ' + 'P(' + grdperdamediaemaximapromethee.Cell[2, 0]
    // .Text + ')' + ' + ' + 'P(' + grdperdamediaemaximapromethee.Cell[3, 0].Text + ')';
    // Para alt1

    Px := ((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
          [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
        / Total) * 100;
    if (TIWUserSession(WebApplication.Data).ResultSolPromethee
        [ordemsolPromethee[0]] / Total) >= faixaprob[0] then // Se a alt[1] tiver uma probabilidade >= 80%, ela é recomendada
    begin
      Px := (TIWUserSession(WebApplication.Data).ResultSolPromethee
          [ordemsolPromethee[0]] / Total) * 100;
      if (TIWUserSession(WebApplication.Data).MaximoGeralNaosolPromethee
          [ordemsolPromethee[0]] < epsilon[0]) and
        (TIWUserSession(WebApplication.Data).MediaGeralNaoSolPromethee
          [ordemsolPromethee[0]] < omega[0]) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[1, 0].Text = '' then
              grdrecomendacaopromethee.Cell[1, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[1, 0].Text :=
                grdrecomendacaopromethee.Cell[1, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[1, 0].Text :=
          grdrecomendacaopromethee.Cell[1, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
        recomendacao := false;
    end;
    if ((ResultSolPromethee[ordemsolPromethee[0]] / Total) < faixaprob[0]) and
      ((ResultSolPromethee[ordemsolPromethee[0]] / Total) >= faixaprob[1]) and
      (recomendacao = false) then
    begin
      Px := (TIWUserSession(WebApplication.Data).ResultSolPromethee
          [ordemsolPromethee[0]] / Total) * 100;
      if (MaximoGeralNaosolPromethee[ordemsolPromethee[0]] < epsilon[1]) and
        (MediaGeralNaoSolPromethee[ordemsolPromethee[0]] < omega[1]) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[1, 0].Text = '' then
              grdrecomendacaopromethee.Cell[1, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[1, 0].Text :=
                grdrecomendacaopromethee.Cell[1, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[1, 0].Text :=
          grdrecomendacaopromethee.Cell[1, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if ((ResultSolPromethee[ordemsolPromethee[0]] / Total) < faixaprob[1]) and
      ((ResultSolPromethee[ordemsolPromethee[0]] / Total) >= faixaprob[2]) and
      (recomendacao = false) then
    begin
      Px := (TIWUserSession(WebApplication.Data).ResultSolPromethee
          [ordemsolPromethee[0]] / Total) * 100;
      if (MaximoGeralNaosolPromethee[ordemsolPromethee[0]] < epsilon[2]) and
        (MediaGeralNaoSolPromethee[ordemsolPromethee[0]] < omega[2]) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[1, 0].Text = '' then
              grdrecomendacaopromethee.Cell[1, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[1, 0].Text :=
                grdrecomendacaopromethee.Cell[1, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[1, 0].Text :=
          grdrecomendacaopromethee.Cell[1, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if ((ResultSolPromethee[ordemsolPromethee[0]] / Total) < faixaprob[2]) and
      ((ResultSolPromethee[ordemsolPromethee[0]] / Total) >= faixaprob[3]) and
      (recomendacao = false) then
    begin
      Px := (TIWUserSession(WebApplication.Data).ResultSolPromethee
          [ordemsolPromethee[0]] / Total) * 100;
      if (MaximoGeralNaosolPromethee[ordemsolPromethee[0]] < epsilon[3]) and
        (MediaGeralNaoSolPromethee[ordemsolPromethee[0]] < omega[3]) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[1, 0].Text = '' then
              grdrecomendacaopromethee.Cell[1, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[1, 0].Text :=
                grdrecomendacaopromethee.Cell[1, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[1, 0].Text :=
          grdrecomendacaopromethee.Cell[1, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;

    // Para alt1 e alt2
    if (((TIWUserSession(WebApplication.Data).ResultSolPromethee
            [ordemsolPromethee[0]] + ResultSolPromethee[ordemsolPromethee[1]])
          / Total) >= faixaprob[0]) and (recomendacao = false) then
    // Se a alt[1] tiver uma probabilidade >= 80%, ela é recomendada
    begin
      Px := ((TIWUserSession(WebApplication.Data).ResultSolPromethee
            [ordemsolPromethee[0]] + ResultSolPromethee[ordemsolPromethee[1]])
          / Total) * 100;
      if ((MaximoGeralNaosolPromethee[ordemsolPromethee[0]] < epsilon[0]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[0]] < omega[0])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[1]] < epsilon[0]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[1]] < omega[0])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if TIWUserSession(WebApplication.Data).matrizsolPromethee
            [ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;

        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[1], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[2, 0].Text :=
          grdrecomendacaopromethee.Cell[2, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[2, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
        recomendacao := false;
    end;
    if (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]]) / Total) < faixaprob[0]) and
      (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]]) / Total) >= faixaprob[1]) and
      (recomendacao = false) then
    begin
      Px := ((TIWUserSession(WebApplication.Data).ResultSolPromethee
            [ordemsolPromethee[0]] + ResultSolPromethee[ordemsolPromethee[1]])
          / Total) * 100;
      if ((MaximoGeralNaosolPromethee[ordemsolPromethee[0]] < epsilon[1]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[0]] < omega[1])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[1]] < epsilon[1]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[1]] < omega[1])) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;

        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[1], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[2, 0].Text :=
          grdrecomendacaopromethee.Cell[2, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[2, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]]) / Total) < faixaprob[1]) and
      (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]]) / Total) >= faixaprob[2]) and
      (recomendacao = false) then
    begin
      Px := ((TIWUserSession(WebApplication.Data).ResultSolPromethee
            [ordemsolPromethee[0]] + ResultSolPromethee[ordemsolPromethee[1]])
          / Total) * 100;
      if ((MaximoGeralNaosolPromethee[ordemsolPromethee[0]] < epsilon[2]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[0]] < omega[2])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[1]] < epsilon[2]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[1]] < omega[2])) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;

        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[1], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[2, 0].Text :=
          grdrecomendacaopromethee.Cell[2, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[2, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]]) / Total) < faixaprob[2]) and
      (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]]) / Total) >= faixaprob[3]) and
      (recomendacao = false) then
    begin
      Px := ((TIWUserSession(WebApplication.Data).ResultSolPromethee
            [ordemsolPromethee[0]] + ResultSolPromethee[ordemsolPromethee[1]])
          / Total) * 100;
      if ((MaximoGeralNaosolPromethee[ordemsolPromethee[0]] < epsilon[3]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[0]] < omega[3])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[1]] < epsilon[3]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[1]] < omega[3])) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;

        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[1], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[2, 0].Text :=
          grdrecomendacaopromethee.Cell[2, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[2, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;

    // Para alt1, alt2 e alt3
    if ((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
          [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
        / Total) >= faixaprob[0] then
    // Se a alt[1] tiver uma probabilidade >= 80%, ela é recomendada
    begin
      Px := ((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
          / Total) * 100;
      if ((MaximoGeralNaosolPromethee[ordemsolPromethee[0]] < epsilon[0]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[0]] < omega[0])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[1]] < epsilon[0]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[1]] < omega[0])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[2]] < epsilon[0]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[2]] < omega[0])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[1], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[2], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[2, 0].Text :=
          grdrecomendacaopromethee.Cell[2, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[2, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[3, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
        recomendacao := false;
    end;
    if (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
          / Total) < faixaprob[0]) and
      (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
          / Total) >= faixaprob[1]) and (recomendacao = false) then
    begin
      Px := ((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
          / Total) * 100;
      if ((MaximoGeralNaosolPromethee[ordemsolPromethee[0]] < epsilon[1]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[0]] < omega[1])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[1]] < epsilon[1]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[1]] < omega[1])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[2]] < epsilon[1]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[2]] < omega[1])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[1], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[2], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[2, 0].Text :=
          grdrecomendacaopromethee.Cell[2, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[2, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[3, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
          / Total) < faixaprob[1]) and
      (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
          / Total) >= faixaprob[2]) and (recomendacao = false) then
    begin
      Px := ((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
          / Total) * 100;
      if ((MaximoGeralNaosolPromethee[ordemsolPromethee[0]] < epsilon[2]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[0]] < omega[2])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[1]] < epsilon[2]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[1]] < omega[2])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[2]] < epsilon[2]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[2]] < omega[2])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[1], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[2], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[2, 0].Text :=
          grdrecomendacaopromethee.Cell[2, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[2, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[3, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
          / Total) < faixaprob[2]) and
      (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
          / Total) >= faixaprob[3]) and (recomendacao = false) then
    begin
      Px := ((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
          / Total) * 100;
      if ((MaximoGeralNaosolPromethee[ordemsolPromethee[0]] < epsilon[3]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[0]] < omega[3])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[1]] < epsilon[3]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[1]] < omega[3])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[2]] < epsilon[3]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[2]] < omega[3])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[1], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[2], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[2, 0].Text :=
          grdrecomendacaopromethee.Cell[2, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[2, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[3, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;

    // caso não passe em nenhuma regra
    if recomendacao = false then
      grdrecomendacaopromethee.Cell[3, 0].Text := 'Not Possible' + '(P(x) = ' +
        formatfloat('0', Px) + '%)';
  end;
end;

procedure TIWForm1.Promethee;
var
  i, j, l, c, xcrit, k, aux, t: Integer;
  soma, Min, auxresultado: real;
  NovaSol: boolean;
  pular: array of boolean;
  pular2: boolean;
  Fatorial: real;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    SetLength(TIWUserSession(WebApplication.Data).matrizsolPromethee, 1,
      Numalt);
    SetLength(TIWUserSession(WebApplication.Data).ResultAltPromethee, Numalt);
    SetLength(TIWUserSession(WebApplication.Data).ResultSolPromethee, 1);
    for j := 0 to Numalt - 1 do
      matrizsolPromethee[0, j] := matrizpoaPromethee[0, j];
    ResultSolPromethee[0] := 1;

    casesolPromethee[0] := 1;

    for i := 1 to length(matrizpoaPromethee) - 1 do
    begin

      k := 0;
      NovaSol := True;
      while k <= length(matrizsolPromethee) - 1 do
      begin
        aux := 0;
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[k, j] = matrizpoaPromethee[i, j] then
            inc(aux);
        end;
        if aux = Numalt then
        begin
          ResultSolPromethee[k] := ResultSolPromethee[k] + 1;
          casesolPromethee[i] := k + 1;
          k := length(matrizsolPromethee);
          NovaSol := false;
        end
        else
        begin
          inc(k);
        end;
      end; // k
      if NovaSol = True then
      begin
        SetLength(TIWUserSession(WebApplication.Data).matrizsolPromethee, length
            (matrizsolPromethee) + 1, Numalt);
        SetLength(TIWUserSession(WebApplication.Data).ResultSolPromethee, length
            (matrizsolPromethee));
        for j := 0 to Numalt - 1 do
          matrizsolPromethee[length(matrizsolPromethee) - 1, j] :=
            matrizpoaPromethee[i, j];
        ResultSolPromethee[length(matrizsolPromethee) - 1] := 1;
        TIWUserSession(WebApplication.Data).casesolPromethee[i] := length
          (matrizsolPromethee);
      end;

    end; // i
    SetLength(TIWUserSession(WebApplication.Data).ResultAltPromethee, Numalt);
    for j := 0 to Numalt - 1 do
    begin
      ResultAltPromethee[j] := 0;
      for i := 0 to length(matrizpoaPromethee) - 1 do
      begin
        ResultAltPromethee[j] := ResultAltPromethee[j] + matrizpoaPromethee
          [i, j];
      end;
    end;

    SetLength(ordemsolPromethee, length(matrizsolPromethee));
    SetLength(pular, length(matrizsolPromethee));

    for i := 0 to length(matrizsolPromethee) - 1 do
      pular[i] := false;

    for i := 0 to length(matrizsolPromethee) - 1 do
    begin
      auxresultado := 0;
      for j := 0 to length(matrizsolPromethee) - 1 do
      begin
        if pular[j] = True then
          continue;
        if TIWUserSession(WebApplication.Data).ResultSolPromethee[j]
          > auxresultado then
          auxresultado := ResultSolPromethee[j];
      end;
      pular2 := false;
      for j := 0 to length(matrizsolPromethee) - 1 do
      begin
        if (pular2 = True) or (pular[j] = True) then
          continue;
        if auxresultado = ResultSolPromethee[j] then
        begin
          pular[j] := True;
          TIWUserSession(WebApplication.Data).ordemsolPromethee[i] := j;
          pular2 := True;
        end;

      end;
    end;

    SetLength(OrdemaltPromethee, Numalt);
    SetLength(pular, Numalt);
    for i := 0 to Numalt - 1 do
      pular[i] := false;

    for i := 0 to Numalt - 1 do
    begin
      auxresultado := 0;
      for j := 0 to Numalt - 1 do
      begin
        if pular[j] = True then
          continue;
        if TIWUserSession(WebApplication.Data).ResultAltPromethee[j]
          > auxresultado then
          auxresultado := ResultAltPromethee[j];
      end;
      pular2 := false;
      for j := 0 to Numalt - 1 do
      begin
        if (pular2 = True) or (pular[j] = True) then
          continue;
        if auxresultado = ResultAltPromethee[j] then
        begin
          pular[j] := True;
          TIWUserSession(WebApplication.Data).OrdemaltPromethee[i] := j;
          pular2 := True;
        end;

      end;
    end;

    SetLength(MediaDifSolPromethee, length(matrizsolPromethee), length
        (matrizsolPromethee));
    SetLength(MaxDifSolPromethee, length(matrizsolPromethee), length
        (matrizsolPromethee));
    SetLength(MinDifSolPromethee, length(matrizsolPromethee), length
        (matrizsolPromethee));
    SetLength(DesvioPadraoDifSolPromethee, length(matrizsolPromethee), length
        (matrizsolPromethee));
    SetLength(MediaGeralPromethee, length(matrizsolPromethee));
    SetLength(DesvioPadraoGeralPromethee, length(matrizsolPromethee));
    SetLength(MaximoGeralPromethee, length(matrizsolPromethee));
    SetLength(MinimoGeralPromethee, length(matrizsolPromethee));
    SetLength(MinMaxPromethee, length(matrizsolPromethee));
    SetLength(MaxMinPromethee, length(matrizsolPromethee));
    SetLength(MediaGeralNaoSolPromethee, length(matrizsolPromethee));
    SetLength(DesvioPadraoGeralNaosolPromethee, length(matrizsolPromethee));
    SetLength(MaximoGeralNaosolPromethee, length(matrizsolPromethee));
    SetLength(MinimoGeralNaoSolPromethee, length(matrizsolPromethee));
    SetLength(MinMaxNaoSolPromethee, length(matrizsolPromethee));
    SetLength(MaxMinNaoSolPromethee, length(matrizsolPromethee));
    for i := 0 to length(matrizsolPromethee) - 1 do
    begin
      MediaGeralPromethee[i] := 0;
      MaximoGeralPromethee[i] := 0;
      MinimoGeralPromethee[i] := 999999999999;
      MinMaxPromethee[i] := 9999999999999999;
      MaxMinPromethee[i] := 0;
      MediaGeralNaoSolPromethee[i] := 0;
      MaximoGeralNaosolPromethee[i] := 0;
      MinimoGeralNaoSolPromethee[i] := 999999999999;
      MinMaxNaoSolPromethee[i] := 9999999999999999;
      MaxMinNaoSolPromethee[i] := 0;
      for j := 0 to length(matrizsolPromethee) - 1 do
      begin
        if i = j then
          continue;
        MaxDifSolPromethee[i, j] := 0;
        MinDifSolPromethee[i, j] := 9999999999;
        for k := 0 to length(casesolPromethee) - 1 do
        begin
          if casesolPromethee[k] <> i + 1 then
          begin
            for t := 0 to Numalt - 1 do
            begin
              TIWUserSession(WebApplication.Data).MediaGeralNaoSolPromethee[i]
                := MediaGeralNaoSolPromethee[i] + TIWUserSession
                (WebApplication.Data).matrizsolPromethee[i, t] * TIWUserSession
                (WebApplication.Data).MatrizDifVGPromethee[k, t];

            end; // t
          end;
          if casesolPromethee[k] = i + 1 then
          begin
            for t := 0 to Numalt - 1 do
            begin
              MediaDifSolPromethee[i, j] := MediaDifSolPromethee[i, j]
                + TIWUserSession(WebApplication.Data).matrizsolPromethee[j, t]
                * TIWUserSession(WebApplication.Data).MatrizDifVGPromethee
                [k, t];

              MediaGeralPromethee[i] := MediaGeralPromethee[i] + TIWUserSession
                (WebApplication.Data).matrizsolPromethee[j, t] * TIWUserSession
                (WebApplication.Data).MatrizDifVGPromethee[k, t];

              if MaxDifSolPromethee[i, j] <
                (matrizsolPromethee[j, t] * TIWUserSession(WebApplication.Data)
                  .MatrizDifVGPromethee[k, t]) then
              begin
                TIWUserSession(WebApplication.Data).MaxDifSolPromethee[i, j] :=
                  matrizsolPromethee[j, t] * TIWUserSession(WebApplication.Data)
                  .MatrizDifVGPromethee[k, t];

                { if MaxDifSol[i,j] < MinMax[i] then
                  TIWUserSession(WebApplication.Data).MinMax[i] := MaxDifSol[i,j];
                  } end;

              if (matrizsolPromethee[j, t] = 1) and
                (MinDifSolPromethee[i, j] >
                  (matrizsolPromethee[j, t] * TIWUserSession
                    (WebApplication.Data).MatrizDifVGPromethee[k, t])) then
              begin
                TIWUserSession(WebApplication.Data).MinDifSolPromethee[i, j] :=
                  matrizsolPromethee[j, t] * TIWUserSession(WebApplication.Data)
                  .MatrizDifVGPromethee[k, t];

                { if MinDifSol[i,j] > MaxMin[i] then
                  TIWUserSession(WebApplication.Data).MaxMin[i] := MinDifSol[i,j];
                  } end;

              { if MaximoGeral[i] < (MatrizSol[j, t] * TIWUserSession
                (WebApplication.Data).MatrizDifVG[k, t]) then
                TIWUserSession(WebApplication.Data).MaximoGeral[i] := MatrizSol
                [j, t] * TIWUserSession(WebApplication.Data)
                .MatrizDifVG[k, t];

                if (MatrizSol[j, t] = 1) and
                (MinimoGeral[i] > (MatrizSol[j, t] * TIWUserSession
                (WebApplication.Data).MatrizDifVG[k, t])) then
                TIWUserSession(WebApplication.Data).MinimoGeral[i] := MatrizSol
                [j, t] * TIWUserSession(WebApplication.Data)
                .MatrizDifVG[k, t]; }

            end;
          end; // t
        end; // k
        auxresultado := 0;
        for t := 0 to Numalt - 1 do
          auxresultado := auxresultado + matrizsolPromethee[j, t];

        MediaDifSolPromethee[i, j] := TIWUserSession(WebApplication.Data)
          .MediaDifSolPromethee[i, j] / (auxresultado * TIWUserSession
            (WebApplication.Data).ResultSolPromethee[i]);
      end; // j
      if (auxresultado > 0) and (ResultSolPromethee[i] <> 0) then
      TIWUserSession(WebApplication.Data).MediaGeralPromethee[i] :=
        MediaGeralPromethee[i] / (auxresultado * TIWUserSession
          (WebApplication.Data).ResultSolPromethee[i] *
          (length(TIWUserSession(WebApplication.Data).matrizsolPromethee)));

      MaxMinPromethee[i] := 0;
      MinMaxPromethee[i] := 9999999999;
      for j := 0 to length(matrizsolPromethee) - 1 do
      begin
        if i = j then
          continue;
        if MaxDifSolPromethee[i, j] < MinMaxPromethee[i] then
          MinMaxPromethee[i] := MaxDifSolPromethee[i, j];
        if MinDifSolPromethee[i, j] > MaxMinPromethee[i] then
          MaxMinPromethee[i] := MinDifSolPromethee[i, j];
        if MaxDifSolPromethee[i, j] > MaximoGeralPromethee[i] then
          TIWUserSession(WebApplication.Data).MaximoGeralPromethee[i] :=
            MaxDifSolPromethee[i, j];
        if MinDifSolPromethee[i, j] < MinimoGeralPromethee[i] then
          TIWUserSession(WebApplication.Data).MinimoGeralPromethee[i] :=
            MinDifSolPromethee[i, j];
      end;
    end; // i
    for i := 0 to length(matrizsolPromethee) - 1 do
    begin
      Fatorial := 1;
      for j := 1 to Numcrit do
        Fatorial := Fatorial * j;

      Fatorial := Fatorial + 1;
      if (auxresultado <> 0) and ((Fatorial - TIWUserSession(WebApplication.Data)
            .ResultSolPromethee[i]) * (length(TIWUserSession(WebApplication.Data).matrizsolPromethee)) <> 0) then

      TIWUserSession(WebApplication.Data).MediaGeralNaoSolPromethee[i] :=
        MediaGeralNaoSolPromethee[i] /
        (auxresultado * (Fatorial - TIWUserSession(WebApplication.Data)
            .ResultSolPromethee[i]) *
          (length(TIWUserSession(WebApplication.Data).matrizsolPromethee)));

      for j := 0 to length(matrizsolPromethee) - 1 do
      begin
        if i = j then
          continue;
        if MaxDifSolPromethee[i, j] < MinMaxNaoSolPromethee[j] then
          TIWUserSession(WebApplication.Data).MinMaxNaoSolPromethee[j] :=
            MaxDifSolPromethee[i, j];
        if MinDifSolPromethee[i, j] > MaxMinNaoSolPromethee[j] then
          TIWUserSession(WebApplication.Data).MaxMinNaoSolPromethee[j] :=
            MinDifSolPromethee[i, j];

        if MaxDifSolPromethee[i, j] > MaximoGeralNaosolPromethee[j] then
          TIWUserSession(WebApplication.Data).MaximoGeralNaosolPromethee[j] :=
            MaxDifSolPromethee[i, j];
        if MinDifSolPromethee[i, j] < MinimoGeralNaoSolPromethee[j] then
          TIWUserSession(WebApplication.Data).MinimoGeralNaoSolPromethee[j] :=
            TIWUserSession(WebApplication.Data).MinDifSolPromethee[i, j];
      end;
    end;

    for i := 0 to length(matrizsolPromethee) - 1 do
    begin
      DesvioPadraoGeralPromethee[i] := 0;
      for j := 0 to length(matrizsolPromethee) - 1 do
      begin
        if i = j then
          continue;

        DesvioPadraoDifSolPromethee[i, j] := 0;
        for k := 0 to length(casesolPromethee) - 1 do
        begin
          if casesolPromethee[k] <> i + 1 then
          begin
            // DesvioPadraoGeralNaosolPromethee[i] := DesvioPadraoGeralNaosolPromethee[i] + sqr(abs(MediaDifSolPromethee[i, j] - TIWUserSession(WebApplication.Data).MatrizDifVGPromethee[k, t]));
            { DesvioPadraoGeralNaosol[i] := DesvioPadraoGeralNaosol[i] + TIWUserSession
              (WebApplication.Data).MatrizSol[i, t] * sqr
              (abs(MediaGeral[i] - TIWUserSession(WebApplication.Data)
              .MatrizDifVG[k, t])); }
          end;

          if casesolPromethee[k] = i + 1 then
          begin
            for t := 0 to Numalt - 1 do
            begin
              DesvioPadraoDifSolPromethee[i, j] := DesvioPadraoDifSolPromethee
                [i, j] + TIWUserSession(WebApplication.Data)
                .matrizsolPromethee[j, t] * sqr
                (abs(MediaDifSolPromethee[i, j] - TIWUserSession
                    (WebApplication.Data).MatrizDifVGPromethee[k, t]));

              DesvioPadraoGeralPromethee[i] := DesvioPadraoGeralPromethee[i]
                + TIWUserSession(WebApplication.Data).matrizsolPromethee[j, t]
                * sqr(abs(MediaGeralPromethee[i] - TIWUserSession
                    (WebApplication.Data).MatrizDifVGPromethee[k, t]));
            end;
          end;

        end; // k
        auxresultado := 0;
        for t := 0 to Numalt - 1 do
          auxresultado := auxresultado + matrizsolPromethee[j, t];

        TIWUserSession(WebApplication.Data).DesvioPadraoDifSolPromethee[i, j]
          := sqrt(DesvioPadraoDifSolPromethee[i, j] /
            (auxresultado * TIWUserSession(WebApplication.Data)
              .ResultSolPromethee[i]));
      end; // j
      if (auxresultado * TIWUserSession
            (WebApplication.Data).ResultSolPromethee[i] *
            (length(matrizsolPromethee))) <> 0 then
      TIWUserSession(WebApplication.Data).DesvioPadraoGeralPromethee[i] := sqrt
        (DesvioPadraoGeralPromethee[i] / (auxresultado * TIWUserSession
            (WebApplication.Data).ResultSolPromethee[i] *
            (length(matrizsolPromethee))));
    end; // i
    for i := 0 to length(matrizsolPromethee) - 1 do
    begin
      if (auxresultado * (Fatorial - TIWUserSession(WebApplication.Data)
              .ResultSolPromethee[i]) *
            (length(TIWUserSession(WebApplication.Data).matrizsolPromethee))) <> 0 then

      TIWUserSession(WebApplication.Data).DesvioPadraoGeralNaosolPromethee[i]
        := sqrt(DesvioPadraoGeralNaosolPromethee[i] /
          (auxresultado * (Fatorial - TIWUserSession(WebApplication.Data)
              .ResultSolPromethee[i]) *
            (length(TIWUserSession(WebApplication.Data).matrizsolPromethee))));
    end;

  end;
end;

procedure TIWForm1.Surrogate;
var
  i, j, k, t, S, q, aux, qtdpareto: Integer;
  auxpeso, soma, auxresultado, vg1, vg2, auxregra: real;
  NovaSol, pular2: boolean;
  pular: array of boolean;
  Fatorial: real;
  auxsol: Integer;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    SetLength(PesoCrit, Numcrit);
    SetLength(ResultadoROC, length(CasesOrdemcrit), Numalt);
    SetLength(ResultadoPromethee, length(CasesOrdemcrit), Numalt);
    SetLength(MatrizDifVG, length(CasesOrdemcrit), Numalt);
    SetLength(MatrizPOA, length(CasesOrdemcrit), Numalt);
    SetLength(MatrizDifVGPromethee, length(CasesOrdemcrit), Numalt);
    SetLength(matrizpoaPromethee, length(CasesOrdemcrit), Numalt);
    SetLength(casesol, length(CasesOrdemcrit));
    SetLength(casesolPromethee, length(CasesOrdemcrit));
    SetLength(crit10, Numcrit);
    SetLength(crit12, Numcrit);
    SetLength(crit13, Numcrit);
    SetLength(crit14, Numcrit);
    SetLength(crit20, Numcrit);
    SetLength(crit21, Numcrit);
    SetLength(crit23, Numcrit);
    SetLength(crit24, Numcrit);
    SetLength(MaxAltVG, Numalt);
    SetLength(MaxAltSolVG, Numalt);
    SetLength(MinAltVG, Numalt);
    SetLength(MinAltSolVG, Numalt);

    for k := 0 to length(CasesOrdemcrit) - 2 do
    begin
      for i := 0 to length(CasesOrdemcrit[0]) - 1 do
      begin
        soma := 0;
        for j := i + 1 to TIWUserSession(WebApplication.Data).Numcrit do
        begin
          soma := soma + 1 / j;
        end;
        auxpeso := (1 / TIWUserSession(WebApplication.Data).Numcrit) * soma;
        TIWUserSession(WebApplication.Data).PesoCrit[CasesOrdemcrit[k, i] - 1]
          := auxpeso;
        // simpeso[indexroc[i]] := auxpeso
      end;
      for j := 0 to Numalt - 1 do
        for i := 0 to Numcrit - 1 do
        begin
          TIWUserSession(WebApplication.Data).ResultadoROC[k, j] := ResultadoROC
            [k, j] + TIWUserSession(WebApplication.Data).MatrizConseqNorm[j, i]
            * PesoCrit[i];
        end;
      auxresultado := 0;
      for j := 0 to Numalt - 1 do
      begin
        if ResultadoROC[k, j] > auxresultado then
          auxresultado := ResultadoROC[k, j];
      end;
      for j := 0 to Numalt - 1 do
      begin
        TIWUserSession(WebApplication.Data).MatrizDifVG[k, j] :=
          auxresultado - ResultadoROC[k, j];
        if auxresultado = ResultadoROC[k, j] then  //Pedro - 17/12/2025
          TIWUserSession(WebApplication.Data).MatrizPOA[k, j] := 1;
      end;

      // Promethee
      ComparacaoParAPar;
      CalculoFluxos;

      for j := 0 to Numalt - 1 do
      begin
        ResultadoPromethee[k, j] := NetFlow[j];
      end;
      auxresultado := 0;
      for j := 0 to Numalt - 1 do
      begin
        if ResultadoPromethee[k, j] > auxresultado then
          auxresultado := ResultadoPromethee[k, j];
      end;
      for j := 0 to Numalt - 1 do
      begin
        TIWUserSession(WebApplication.Data).MatrizDifVGPromethee[k, j] :=
          auxresultado - ResultadoPromethee[k, j];
        if auxresultado = ResultadoPromethee[k, j] then
          TIWUserSession(WebApplication.Data).matrizpoaPromethee[k, j] := 1;
      end;
      // Fim Promethee
    end; // k

    for j := 0 to Numalt - 1 do
      for i := 0 to Numcrit - 1 do
      begin
        TIWUserSession(WebApplication.Data)
          .ResultadoROC[length(CasesOrdemcrit) - 1, j] := ResultadoROC[k, j]
          + MatrizConseqNorm[j, i] * (1 / Numcrit);

      end;
    k := length(CasesOrdemcrit) - 1;
    auxresultado := 0;
    for j := 0 to Numalt - 1 do
    begin
      if ResultadoROC[k, j] > auxresultado then
        auxresultado := ResultadoROC[k, j];
    end;
    for j := 0 to Numalt - 1 do
    begin
      TIWUserSession(WebApplication.Data).MatrizDifVG[k, j] :=
        auxresultado - ResultadoROC[k, j];
      if auxresultado = ResultadoROC[k, j] then
        TIWUserSession(WebApplication.Data).MatrizPOA[k, j] := 1;
    end;

    // Promethee - Caso do Surrogate
    for i := 0 to Numcrit - 1 do
      PesoCrit[i] := 1 / Numcrit;

    ComparacaoParAPar;
    CalculoFluxos;

    for j := 0 to Numalt - 1 do
    begin
      TIWUserSession(WebApplication.Data).ResultadoPromethee
        [length(CasesOrdemcrit) - 1, j] := NetFlow[j];
    end;
    k := length(CasesOrdemcrit) - 1;
    auxresultado := 0;
    for j := 0 to Numalt - 1 do
    begin
      if ResultadoPromethee[k, j] > auxresultado then
        auxresultado := ResultadoPromethee[k, j];
    end;
    for j := 0 to Numalt - 1 do
    begin
      TIWUserSession(WebApplication.Data).MatrizDifVGPromethee[k, j] :=
        auxresultado - ResultadoPromethee[k, j];
      if auxresultado = ResultadoPromethee[k, j] then
        TIWUserSession(WebApplication.Data).matrizpoaPromethee[k, j] := 1;
    end;
    // Fim do Promethee - Caso do Surrogate
    Promethee; // Armazena as soluções para o Promethee

    SetLength(MatrizSol, 1, Numalt);
    SetLength(ResultAlt, Numalt);
    SetLength(ResultSol, 1);
    for j := 0 to Numalt - 1 do
      MatrizSol[0, j] := MatrizPOA[0, j];
    ResultSol[0] := 1;

    casesol[0] := 1;

    for i := 1 to length(MatrizPOA) - 1 do
    begin

      k := 0;
      NovaSol := True;
      while k <= length(MatrizSol) - 1 do
      begin
        aux := 0;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[k, j] = MatrizPOA[i, j] then
            inc(aux);
        end;
        if aux = Numalt then
        begin
          ResultSol[k] := ResultSol[k] + 1;
          casesol[i] := k + 1;
          k := length(MatrizSol);
          NovaSol := false;
        end
        else
        begin
          inc(k);
        end;
      end; // k
      if NovaSol = True then
      begin
        SetLength(TIWUserSession(WebApplication.Data).MatrizSol, length
            (MatrizSol) + 1, Numalt);
        SetLength(TIWUserSession(WebApplication.Data).ResultSol, length
            (MatrizSol));
        for j := 0 to Numalt - 1 do
          MatrizSol[length(MatrizSol) - 1, j] := MatrizPOA[i, j];
        ResultSol[length(MatrizSol) - 1] := 1;
        TIWUserSession(WebApplication.Data).casesol[i] := length(MatrizSol);
      end;

    end; // i
    SetLength(TIWUserSession(WebApplication.Data).ResultAlt, Numalt);
    for j := 0 to Numalt - 1 do
    begin
      ResultAlt[j] := 0;
      for i := 0 to length(MatrizPOA) - 1 do
      begin
        ResultAlt[j] := ResultAlt[j] + MatrizPOA[i, j];
      end;
    end;
    SetLength(OrdemSol, length(MatrizSol));
    SetLength(pular, length(MatrizSol));
    for i := 0 to length(MatrizSol) - 1 do
      pular[i] := false;

    for i := 0 to length(MatrizSol) - 1 do
    begin
      auxresultado := 0;
      for j := 0 to length(MatrizSol) - 1 do
      begin
        if pular[j] = True then
          continue;
        if TIWUserSession(WebApplication.Data).ResultSol[j] > auxresultado then
          auxresultado := ResultSol[j];
      end;
      pular2 := false;
      for j := 0 to length(MatrizSol) - 1 do
      begin
        if (pular2 = True) or (pular[j] = True) then
          continue;
        if auxresultado = ResultSol[j] then
        begin
          pular[j] := True;
          TIWUserSession(WebApplication.Data).OrdemSol[i] := j;
          pular2 := True;
        end;

      end;
    end;

    SetLength(Ordemalt, Numalt);
    SetLength(pular, Numalt);
    for i := 0 to Numalt - 1 do
      pular[i] := false;

    for i := 0 to Numalt - 1 do
    begin
      auxresultado := 0;
      for j := 0 to Numalt - 1 do
      begin
        if pular[j] = True then
          continue;
        if TIWUserSession(WebApplication.Data).ResultAlt[j] > auxresultado then
          auxresultado := ResultAlt[j];
      end;
      pular2 := false;
      for j := 0 to Numalt - 1 do
      begin
        if (pular2 = True) or (pular[j] = True) then
          continue;
        if auxresultado = ResultAlt[j] then
        begin
          pular[j] := True;
          TIWUserSession(WebApplication.Data).Ordemalt[i] := j;
          pular2 := True;
        end;

      end;
    end;
    for j := 0 to numalt - 1 do
    begin
    TIWUserSession(WebApplication.Data).maxaltsolVG[j] := 0;
    TIWUserSession(WebApplication.Data).minaltsolVG[j] := 99999999;
    TIWUserSession(WebApplication.Data).maxaltVG[j] := 0;
    TIWUserSession(WebApplication.Data).minaltVG[j] := 99999999;
      for i := 0 to length(MatrizSol) - 1 do
      begin
        if MatrizPOA[i, j] = 0 then
        begin
        if maxaltVG[j] < ResultadoROC[i,j] then
        maxaltVG[j] := ResultadoROC[i,j];

        if minaltVG[j] > ResultadoROC[i,j] then
        minaltVG[j] := ResultadoROC[i,j];
        end;

        if MatrizPOA[i, j] = 1 then
        begin
        if maxaltsolVG[j] < ResultadoROC[i,j] then
        maxaltsolVG[j] := ResultadoROC[i,j];

        if minaltsolVG[j] > ResultadoROC[i,j] then
        minaltsolVG[j] := ResultadoROC[i,j];

        end;


      end;
      if minaltVG[j] = 99999999 then minaltVG[j] := 0;
      if minaltSolVG[j] = 99999999 then minaltSolVG[j] := 0;
    end;
    SetLength(MediaDifSol, length(MatrizSol), length(MatrizSol));
    SetLength(MaxDifSol, length(MatrizSol), length(MatrizSol));
    SetLength(MinDifSol, length(MatrizSol), length(MatrizSol));
    SetLength(DesvioPadraoDifSol, length(MatrizSol), length(MatrizSol));
    SetLength(MediaGeral, length(MatrizSol));
    SetLength(DesvioPadraoGeral, length(MatrizSol));
    SetLength(MaximoGeral, length(MatrizSol));
    SetLength(MinimoGeral, length(MatrizSol));
    SetLength(MinMax, length(MatrizSol));
    SetLength(MaxMin, length(MatrizSol));
    SetLength(MediaGeralNaosol, length(MatrizSol));
    SetLength(DesvioPadraoGeralNaosol, length(MatrizSol));
    SetLength(MaximoGeralNaosol, length(MatrizSol));
    SetLength(MinimoGeralNaosol, length(MatrizSol));
    SetLength(MinMaxNaosol, length(MatrizSol));
    SetLength(MaxMinNaosol, length(MatrizSol));
    for i := 0 to length(MatrizSol) - 1 do
    begin
      MediaGeral[i] := 0;
      MaximoGeral[i] := 0;
      MinimoGeral[i] := 999999999999;
      MinMax[i] := 9999999999999999;
      MaxMin[i] := 0;
      MediaGeralNaosol[i] := 0;
      MaximoGeralNaosol[i] := 0;
      MinimoGeralNaosol[i] := 999999999999;
      MinMaxNaosol[i] := 9999999999999999;
      MaxMinNaosol[i] := 0;
      for j := 0 to length(MatrizSol) - 1 do
      begin
        if i = j then
          continue;
        MaxDifSol[i, j] := 0;
        MinDifSol[i, j] := 9999999999;
        for k := 0 to length(casesol) - 1 do
        begin
          if casesol[k] <> i + 1 then
          begin
            for t := 0 to Numalt - 1 do
            begin

              TIWUserSession(WebApplication.Data).MediaGeralNaosol[i] :=
                MediaGeralNaosol[i] + TIWUserSession(WebApplication.Data)
                .MatrizSol[i, t] * TIWUserSession(WebApplication.Data)
                .MatrizDifVG[k, t];

            end; // t
          end;
          if casesol[k] = i + 1 then
          begin
            for t := 0 to Numalt - 1 do
            begin
              MediaDifSol[i, j] := MediaDifSol[i, j] + TIWUserSession
                (WebApplication.Data).MatrizSol[j, t] * TIWUserSession
                (WebApplication.Data).MatrizDifVG[k, t];

              MediaGeral[i] := MediaGeral[i] + TIWUserSession
                (WebApplication.Data).MatrizSol[j, t] * TIWUserSession
                (WebApplication.Data).MatrizDifVG[k, t];

              if MaxDifSol[i, j] < (MatrizSol[j, t] * TIWUserSession
                  (WebApplication.Data).MatrizDifVG[k, t]) then
              begin
                TIWUserSession(WebApplication.Data).MaxDifSol[i, j] := MatrizSol
                  [j, t] * TIWUserSession(WebApplication.Data)
                  .MatrizDifVG[k, t];

                { if MaxDifSol[i,j] < MinMax[i] then
                  TIWUserSession(WebApplication.Data).MinMax[i] := MaxDifSol[i,j];
                  } end;

              if (MatrizSol[j, t] = 1) and
                (MinDifSol[i, j] > (MatrizSol[j, t] * TIWUserSession
                    (WebApplication.Data).MatrizDifVG[k, t])) then
              begin
                TIWUserSession(WebApplication.Data).MinDifSol[i, j] := MatrizSol
                  [j, t] * TIWUserSession(WebApplication.Data)
                  .MatrizDifVG[k, t];

                { if MinDifSol[i,j] > MaxMin[i] then
                  TIWUserSession(WebApplication.Data).MaxMin[i] := MinDifSol[i,j];
                  } end;

              { if MaximoGeral[i] < (MatrizSol[j, t] * TIWUserSession
                (WebApplication.Data).MatrizDifVG[k, t]) then
                TIWUserSession(WebApplication.Data).MaximoGeral[i] := MatrizSol
                [j, t] * TIWUserSession(WebApplication.Data)
                .MatrizDifVG[k, t];

                if (MatrizSol[j, t] = 1) and
                (MinimoGeral[i] > (MatrizSol[j, t] * TIWUserSession
                (WebApplication.Data).MatrizDifVG[k, t])) then
                TIWUserSession(WebApplication.Data).MinimoGeral[i] := MatrizSol
                [j, t] * TIWUserSession(WebApplication.Data)
                .MatrizDifVG[k, t]; }

            end;
          end; // t
        end; // k
        auxresultado := 0;
        for t := 0 to Numalt - 1 do
          auxresultado := auxresultado + MatrizSol[j, t];

        MediaDifSol[i, j] := TIWUserSession(WebApplication.Data)
          .MediaDifSol[i, j] / (auxresultado * TIWUserSession
            (WebApplication.Data).ResultSol[i]);
      end; // j
      TIWUserSession(WebApplication.Data).MediaGeral[i] := MediaGeral[i] /
        (auxresultado * TIWUserSession(WebApplication.Data).ResultSol[i] *
          (length(TIWUserSession(WebApplication.Data).MatrizSol)));

      MaxMin[i] := 0;
      MinMax[i] := 9999999999;
      for j := 0 to length(MatrizSol) - 1 do
      begin
        if i = j then
          continue;
        if MaxDifSol[i, j] < MinMax[i] then
          MinMax[i] := MaxDifSol[i, j];
        if MinDifSol[i, j] > MaxMin[i] then
          MaxMin[i] := MinDifSol[i, j];
        if MaxDifSol[i, j] > MaximoGeral[i] then
          TIWUserSession(WebApplication.Data).MaximoGeral[i] := MaxDifSol[i, j];
        if MinDifSol[i, j] < MinimoGeral[i] then
          TIWUserSession(WebApplication.Data).MinimoGeral[i] := MinDifSol[i, j];
      end;
    end; // i
    for i := 0 to length(MatrizSol) - 1 do
    begin
      Fatorial := 1;
      for j := 1 to Numcrit do
        Fatorial := Fatorial * j;

      Fatorial := Fatorial + 1;

      TIWUserSession(WebApplication.Data).MediaGeralNaosol[i] :=
        MediaGeralNaosol[i] / (auxresultado * (Fatorial - TIWUserSession
            (WebApplication.Data).ResultSol[i]) *
          (length(TIWUserSession(WebApplication.Data).MatrizSol)));

      for j := 0 to length(MatrizSol) - 1 do
      begin
        if i = j then
          continue;
        if MaxDifSol[i, j] < MinMaxNaosol[j] then
          TIWUserSession(WebApplication.Data).MinMaxNaosol[j] := MaxDifSol
            [i, j];
        if MinDifSol[i, j] > MaxMinNaosol[j] then
          TIWUserSession(WebApplication.Data).MaxMinNaosol[j] := MinDifSol
            [i, j];

        if MaxDifSol[i, j] > MaximoGeralNaosol[j] then
          TIWUserSession(WebApplication.Data).MaximoGeralNaosol[j] := MaxDifSol
            [i, j];
        if MinDifSol[i, j] < MinimoGeralNaosol[j] then
          TIWUserSession(WebApplication.Data).MinimoGeralNaosol[j] :=
            TIWUserSession(WebApplication.Data).MinDifSol[i, j];
      end;
    end;

    for i := 0 to length(MatrizSol) - 1 do
    begin
      DesvioPadraoGeral[i] := 0;
      for j := 0 to length(MatrizSol) - 1 do
      begin
        if i = j then
          continue;

        DesvioPadraoDifSol[i, j] := 0;
        for k := 0 to length(casesol) - 1 do
        begin
          if casesol[k] <> i + 1 then
          begin
            { DesvioPadraoGeralNaosol[i] := DesvioPadraoGeralNaosol[i] + sqr
              (abs(MediaDifSol[i, j] - TIWUserSession(WebApplication.Data)
              .MatrizDifVG[k, t]));
              { DesvioPadraoGeralNaosol[i] := DesvioPadraoGeralNaosol[i] + TIWUserSession
              (WebApplication.Data).MatrizSol[i, t] * sqr
              (abs(MediaGeral[i] - TIWUserSession(WebApplication.Data)
              .MatrizDifVG[k, t])); }
          end;

          if casesol[k] = i + 1 then
          begin
            for t := 0 to Numalt - 1 do
            begin
              DesvioPadraoDifSol[i, j] := DesvioPadraoDifSol[i, j]
                + TIWUserSession(WebApplication.Data).MatrizSol[j, t] * sqr
                (abs(MediaDifSol[i, j] - TIWUserSession(WebApplication.Data)
                    .MatrizDifVG[k, t]));

              DesvioPadraoGeral[i] := DesvioPadraoGeral[i] + TIWUserSession
                (WebApplication.Data).MatrizSol[j, t] * sqr
                (abs(MediaGeral[i] - TIWUserSession(WebApplication.Data)
                    .MatrizDifVG[k, t]));
            end;
          end;

        end; // k
        auxresultado := 0;
        for t := 0 to Numalt - 1 do
          auxresultado := auxresultado + MatrizSol[j, t];

        TIWUserSession(WebApplication.Data).DesvioPadraoDifSol[i, j] := sqrt
          (DesvioPadraoDifSol[i, j] / (auxresultado * TIWUserSession
              (WebApplication.Data).ResultSol[i]));
      end; // j
      TIWUserSession(WebApplication.Data).DesvioPadraoGeral[i] := sqrt
        (DesvioPadraoGeral[i] / (auxresultado * TIWUserSession
            (WebApplication.Data).ResultSol[i] * (length(MatrizSol))));
    end; // i
    for i := 0 to length(MatrizSol) - 1 do
      TIWUserSession(WebApplication.Data).DesvioPadraoGeralNaosol[i] := sqrt
        (DesvioPadraoGeralNaosol[i] / (auxresultado * (Fatorial - TIWUserSession
              (WebApplication.Data).ResultSol[i]) *
            (length(TIWUserSession(WebApplication.Data).MatrizSol))));

    Fatorial := 1;
    for j := 1 to Numcrit - 1 do
      Fatorial := Fatorial * j;
    for i := 0 to Numcrit - 1 do // crit10
    begin
      k := 0;
      for j := 0 to length(TIWUserSession(WebApplication.Data).casesol) - 1 do
      begin
        if TIWUserSession(WebApplication.Data).CasesOrdemcrit[j, i] = 1 then
        begin
          if k = 0 then
            auxsol := casesol[j];

          if casesol[j] = auxsol then
            inc(k);
        end;
      end;
      if k = Fatorial then
        TIWUserSession(WebApplication.Data).crit10[i] := crit10[i] + k;
    end;

    for i := 0 to Numcrit - 1 do // crit20
    begin
      k := 0;
      for j := 0 to length(TIWUserSession(WebApplication.Data).casesol) - 1 do
      begin
        if TIWUserSession(WebApplication.Data).CasesOrdemcrit[j, i] = 2 then
        begin
          if k = 0 then
            auxsol := casesol[j];

          if casesol[j] = auxsol then
            inc(k);
        end;
      end;
      if k = Fatorial then
        TIWUserSession(WebApplication.Data).crit20[i] := crit20[i] + k;
    end;

    for i := 0 to Numcrit - 1 do // crit12
      for k := 0 to Numcrit - 1 do
      begin
        S := 0;
        if i = k then
          continue;
        for j := 0 to length(TIWUserSession(WebApplication.Data).casesol) - 1 do
        begin
          if (TIWUserSession(WebApplication.Data).CasesOrdemcrit[j, i] = 1) and
            (TIWUserSession(WebApplication.Data).CasesOrdemcrit[j, k] = 2) then
          begin
            if S = 0 then
              auxsol := casesol[j];

            if casesol[j] = auxsol then
              inc(S);
          end;
        end;
        if S = Round(Fatorial / (Numcrit - 1)) then
          TIWUserSession(WebApplication.Data).crit12[i] := crit12[i] + S;
      end;

    for i := 0 to Numcrit - 1 do // crit21
      for k := 0 to Numcrit - 1 do
      begin
        S := 0;
        if i = k then
          continue;
        for j := 0 to length(TIWUserSession(WebApplication.Data).casesol) - 1 do
        begin
          if (TIWUserSession(WebApplication.Data).CasesOrdemcrit[j, i] = 2) and
            (TIWUserSession(WebApplication.Data).CasesOrdemcrit[j, k] = 1) then
          begin
            if S = 0 then
              auxsol := casesol[j];

            if casesol[j] = auxsol then
              inc(S);
          end;
        end;
        if S = Round(Fatorial / (Numcrit - 1)) then
          TIWUserSession(WebApplication.Data).crit21[i] := crit21[i] + S;
      end;

    for i := 0 to Numcrit - 1 do // crit13
      for k := 0 to Numcrit - 1 do
        for S := 0 to Numcrit - 1 do
        begin
          t := 0;
          if (i = k) or (i = S) or (k = S) then
            continue;
          for j := 0 to length(TIWUserSession(WebApplication.Data).casesol) - 1
            do
          begin
            if (TIWUserSession(WebApplication.Data).CasesOrdemcrit[j, i] = 1)
              and (TIWUserSession(WebApplication.Data).CasesOrdemcrit[j, k]
                = 2) and (TIWUserSession(WebApplication.Data)
                .CasesOrdemcrit[j, S] = 3) then
            begin
              if t = 0 then
                auxsol := casesol[j];

              if casesol[j] = auxsol then
                inc(t);
            end;
          end;
          if t = Round(Fatorial / ((Numcrit - 1) * (Numcrit - 2))) then
            TIWUserSession(WebApplication.Data).crit13[i] := crit13[i] + t;
        end;

    for i := 0 to Numcrit - 1 do // crit23
      for k := 0 to Numcrit - 1 do
        for S := 0 to Numcrit - 1 do
        begin
          t := 0;
          if (i = k) or (i = S) or (k = S) then
            continue;
          for j := 0 to length(TIWUserSession(WebApplication.Data).casesol) - 1
            do
          begin
            if (TIWUserSession(WebApplication.Data).CasesOrdemcrit[j, i] = 2)
              and (TIWUserSession(WebApplication.Data).CasesOrdemcrit[j, k]
                = 1) and (TIWUserSession(WebApplication.Data)
                .CasesOrdemcrit[j, S] = 3) then
            begin
              if t = 0 then
                auxsol := casesol[j];

              if casesol[j] = auxsol then
                inc(t);
            end;
          end;
          if t = Round(Fatorial / ((Numcrit - 1) * (Numcrit - 2))) then
            TIWUserSession(WebApplication.Data).crit23[i] := crit23[i] + t;
        end;

    for i := 0 to Numcrit - 1 do // crit14
      for k := 0 to Numcrit - 1 do
        for S := 0 to Numcrit - 1 do
          for t := 0 to Numcrit - 1 do
          begin
            q := 0;
            if (i = k) or (i = S) or (i = t) or (k = S) or (k = t) or (S = t)
              then
              continue;
            for j := 0 to length(TIWUserSession(WebApplication.Data).casesol)
              - 1 do
            begin
              if (TIWUserSession(WebApplication.Data).CasesOrdemcrit[j, i] = 1)
                and (TIWUserSession(WebApplication.Data).CasesOrdemcrit[j, k]
                  = 2) and (TIWUserSession(WebApplication.Data)
                  .CasesOrdemcrit[j, S] = 3) and
                (TIWUserSession(WebApplication.Data).CasesOrdemcrit[j, t] = 4)
                then
              begin
                if q = 0 then
                  auxsol := casesol[j];

                if casesol[j] = auxsol then
                  inc(q);
              end;
            end;
            if q = Round
              (Fatorial / ((Numcrit - 1) * (Numcrit - 2) * (Numcrit - 3))) then
              TIWUserSession(WebApplication.Data).crit14[i] := crit14[i] + q;
          end;

    for i := 0 to Numcrit - 1 do // crit24
      for k := 0 to Numcrit - 1 do
        for S := 0 to Numcrit - 1 do
          for t := 0 to Numcrit - 1 do
          begin
            q := 0;
            if (i = k) or (i = S) or (i = t) or (k = S) or (k = t) or (S = t)
              then
              continue;
            for j := 0 to length(TIWUserSession(WebApplication.Data).casesol)
              - 1 do
            begin
              if (TIWUserSession(WebApplication.Data).CasesOrdemcrit[j, i] = 2)
                and (TIWUserSession(WebApplication.Data).CasesOrdemcrit[j, k]
                  = 1) and (TIWUserSession(WebApplication.Data)
                  .CasesOrdemcrit[j, S] = 3) and
                (TIWUserSession(WebApplication.Data).CasesOrdemcrit[j, t] = 4)
                then
              begin
                if q = 0 then
                  auxsol := casesol[j];

                if casesol[j] = auxsol then
                  inc(q);
              end;
            end;
            if q = Round
              (Fatorial / ((Numcrit - 1) * (Numcrit - 2) * (Numcrit - 3))) then
              TIWUserSession(WebApplication.Data).crit24[i] := crit24[i] + q;
          end;

    Fatorial := 1;
    for j := 1 to Numcrit - 1 do
      Fatorial := Fatorial * j;

  end; // with

end;

procedure TIWForm1.GerarCases;
var
  Vetor, Indices: array of Integer;
  Fatorial, Linha, i, j, k, kk, Temp: Integer;
  N: Integer;
  Matriz: array of array of Integer;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    // Calcular o fatorial de N (número total de permutações)
    Fatorial := 1;
    N := Numcrit;
    for i := 2 to N do
      Fatorial := Fatorial * i;

    // Redimensionar a matriz para armazenar todas as permutações
    SetLength(Matriz, Fatorial, N);

    // Inicializar o vetor de entrada e índices
    SetLength(Vetor, N);
    for i := 0 to N - 1 do
      Vetor[i] := i + 1;

    SetLength(Indices, N);
    for i := 0 to N - 1 do
      Indices[i] := 0;

    // Copiar a primeira permutação
    Linha := 0;
    for i := 0 to N - 1 do
      Matriz[Linha][i] := Vetor[i];
    inc(Linha);

    // Algoritmo Heap's para gerar as permutações
    i := 0;
    while i < N do
    begin
      if Indices[i] < i then
      begin
        // Trocar elementos
        if i mod 2 = 0 then
          Temp := Vetor[0]
        else
          Temp := Vetor[Indices[i]];

        if i mod 2 = 0 then
          Vetor[0] := Vetor[i]
        else
          Vetor[Indices[i]] := Vetor[i];

        Vetor[i] := Temp;

        // Armazenar a nova permutação na matriz
        for j := 0 to N - 1 do
          Matriz[Linha][j] := Vetor[j];
        inc(Linha);

        // Incrementar o índice
        inc(Indices[i]);
        i := 0;
      end
      else
      begin
        Indices[i] := 0;
        inc(i);
      end;
    end;

    if Numcrit <= 7 then
    begin
      // Ordenar a matriz de permutações em ordem lexicográfica
      for i := 0 to Fatorial - 2 do
        for j := i + 1 to Fatorial - 1 do
        begin
          for k := 0 to N - 1 do
          begin
            if Matriz[i][k] < Matriz[j][k] then
              break
            else if Matriz[i][k] > Matriz[j][k] then
            begin
              // Trocar as linhas
              for kk := 0 to N - 1 do
              begin
                Temp := Matriz[i][kk];
                Matriz[i][kk] := Matriz[j][kk];
                Matriz[j][kk] := Temp;
              end;
              break;
            end;
          end;
        end;
    end;

    SetLength(CasesOrdemcrit, Fatorial + 1, Numcrit);
    for i := 0 to Fatorial - 1 do
      for j := 0 to Numcrit - 1 do
        CasesOrdemcrit[i, j] := Matriz[i, j];
  end;
end;

procedure TIWForm1.grdlimiaresComboChange(Sender: TObject;
  RowIndex, ColumnIndex: Integer; AValue: string);
var
  i: Integer;
  aux: string;
begin
  if StrToFloat(grdlimiares.DynEdits[1, 0]) < 0.03 then
    grdlimiares.DynEdits[1, 0] := '0.03';
  if StrToFloat(grdlimiares.DynEdits[1, 1]) < 0.02 then
    grdlimiares.DynEdits[1, 1] := '0.02';
  if StrToFloat(grdlimiares.DynEdits[1, 2]) < 0.01 then
    grdlimiares.DynEdits[1, 2] := '0.01';

  if StrToFloat(grdlimiares.DynEdits[1, 3]) > 0.97 then
    grdlimiares.DynEdits[1, 3] := '0.97';
  if StrToFloat(grdlimiares.DynEdits[1, 2]) > 0.98 then
    grdlimiares.DynEdits[1, 2] := '0.98';
  if StrToFloat(grdlimiares.DynEdits[1, 1]) > 0.99 then
    grdlimiares.DynEdits[1, 1] := '0.99';

  for i := grdlimiares.rowcount - 1 downto 1 do
  begin
    if StrToFloat(grdlimiares.DynEdits[1, i]) > StrToFloat
      (grdlimiares.DynEdits[1, -1]) then
    begin
      aux := grdlimiares.DynEdits[1, i];
      grdlimiares.DynEdits[1, i] := grdlimiares.DynEdits[1, i - 1];
      grdlimiares.DynEdits[1, i - 1] := aux;
    end;
    if StrToFloat(grdlimiares.DynEdits[1, i]) = StrToFloat
      (grdlimiares.DynEdits[1, -1]) then
    begin
      grdlimiares.DynEdits[i, i - 1] := formatfloat
        ('0.00', StrToFloat(grdlimiares.DynEdits[i, i]) + 0.01);
    end;
  end;
end;

procedure TIWForm1.grdlimiaresEdit(Sender: TObject);
var
  i: Integer;
  aux: string;
begin
  if StrToFloat(grdlimiares.DynEdits[1, 0]) < 0.03 then
    grdlimiares.DynEdits[1, 0] := '0.03';
  if StrToFloat(grdlimiares.DynEdits[1, 1]) < 0.02 then
    grdlimiares.DynEdits[1, 1] := '0.02';
  if StrToFloat(grdlimiares.DynEdits[1, 2]) < 0.01 then
    grdlimiares.DynEdits[1, 2] := '0.01';

  if StrToFloat(grdlimiares.DynEdits[1, 3]) > 0.97 then
    grdlimiares.DynEdits[1, 3] := '0.97';
  if StrToFloat(grdlimiares.DynEdits[1, 2]) > 0.98 then
    grdlimiares.DynEdits[1, 2] := '0.98';
  if StrToFloat(grdlimiares.DynEdits[1, 1]) > 0.99 then
    grdlimiares.DynEdits[1, 1] := '0.99';

  for i := grdlimiares.rowcount - 1 downto 1 do
  begin
    if StrToFloat(grdlimiares.DynEdits[1, i]) > StrToFloat
      (grdlimiares.DynEdits[1, -1]) then
    begin
      aux := grdlimiares.DynEdits[1, i];
      grdlimiares.DynEdits[1, i] := grdlimiares.DynEdits[1, i - 1];
      grdlimiares.DynEdits[1, i - 1] := aux;
    end;
    if StrToFloat(grdlimiares.DynEdits[1, i]) = StrToFloat
      (grdlimiares.DynEdits[1, -1]) then
    begin
      grdlimiares.DynEdits[i, i - 1] := formatfloat
        ('0.00', StrToFloat(grdlimiares.DynEdits[i, i]) + 0.01);
    end;
  end;
end;

procedure TIWForm1.LimpaVariaveis;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    // Finalize(Tipocrit);
    // Finalize(niveis);
    // Finalize(MatrizConseq);
    // Finalize(Max);
    // Finalize(Min);
    // Finalize(MatrizConseqNorm);
    Finalize(ResultSol);
    Finalize(ResultAlt);
    // finalize(edtNomeCrit);
    // finalize(edtNomeAlt);
    // finalize(cmbbxNiveisCrit);
    // finalize(cmbbxTipoCrit);
    // finalize(edtMatrizConseq);
    // finalize(NomeCrit);
    // finalize(NomeAlt);
    // finalize(Tipocrit);
    // finalize(Niveis);
    // finalize(MatrizConseq);
    // finalize(MatrizConseqNorm);
    Finalize(PesoCrit);
    Finalize(CasesOrdemcrit);
    // finalize(Max);
    // finalize(Min);
    Finalize(MatrizPOA);
    Finalize(ResultadoROC);
    Finalize(MatrizSol);
    Finalize(ResultAlt);
    Finalize(ResultSol);
    Finalize(coresalt);
    Finalize(coresalt1);
    Finalize(casesol);
    Finalize(CasesOrdemcrit);
    Finalize(MatrizPOA);
    Finalize(ResultadoROC);
    Finalize(PesoCrit);
    // Finalize(NomeAlt);
    // Finalize(Nomecrit);
    Finalize(MatrizSol);
    Finalize(coresalt);
    Finalize(coresalt1);
    Finalize(crit10);
    Finalize(crit12);
    Finalize(crit13);
    Finalize(crit14);
    Finalize(crit20);
    Finalize(crit21);
    Finalize(crit23);
    Finalize(crit24);
    Finalize(ResultSol);
    Finalize(ResultAlt);

    // Finalize(PesoCritPromethee);
    Finalize(CasesOrdemcrit);
    Finalize(matrizpoaPromethee);
    Finalize(ResultadoPromethee);
    Finalize(matrizsolPromethee);
    Finalize(ResultAltPromethee);
    Finalize(ResultSolPromethee);
    Finalize(casesolPromethee);
    Finalize(matrizpoaPromethee);
    Finalize(matrizsolPromethee);
    Finalize(MatrizDifVG);
    Finalize(MatrizDifVGPromethee);
    Finalize(MediaDifSol);
    Finalize(MaxDifSol);
    Finalize(DesvioPadraoDifSol);
    Finalize(MinDifSol);
    Finalize(MediaDifSolPromethee);
    Finalize(MaxDifSolPromethee);
    Finalize(DesvioPadraoDifSolPromethee);
    Finalize(MinDifSolPromethee);
    Finalize(SobClassMatrix);
    Finalize(PositiveFlow);
    Finalize(NegativeFlow);
    Finalize(NetFlow);
    Finalize(matrizParaPar);
    Finalize(OrdemSol);
    Finalize(ordemsolPromethee);
    Finalize(MediaGeral);
    Finalize(DesvioPadraoGeral);
    Finalize(MaximoGeral);
    Finalize(MinimoGeral);
    Finalize(MaxMin);
    Finalize(MinMax);
    Finalize(MediaGeralPromethee);
    Finalize(DesvioPadraoGeralPromethee);
    Finalize(MaximoGeralPromethee);
    Finalize(MinimoGeralPromethee);
    Finalize(MaxMinPromethee);
    Finalize(MinMaxPromethee);
    Finalize(MediaGeralNaosol);
    Finalize(DesvioPadraoGeralNaosol);
    Finalize(MaximoGeralNaosol);
    Finalize(MinimoGeralNaosol);
    Finalize(MaxMinNaosol);
    Finalize(MinMaxNaosol);
    Finalize(MediaGeralNaoSolPromethee);
    Finalize(DesvioPadraoGeralNaosolPromethee);
    Finalize(MaximoGeralNaosolPromethee);
    Finalize(MinimoGeralNaoSolPromethee);
    Finalize(MaxMinNaoSolPromethee);
    Finalize(MinMaxNaoSolPromethee);
    Finalize(SimilaridadeAlt);
    Finalize(SomenteAditivo);
    Finalize(SomenteSobreclassificacao);

  end;
end;

// Transforma os valores da consMatrix em uma escala de razão
procedure TIWForm1.EscalaRazao;
var
  i, j: Integer;
  Maxl, Minl, MaxA, MinA, MaxB, MinB: Array of Double;
  AuxVetPOA: array of Integer;
  // ConsRazao: array of array of Double;
begin

  with TIWUserSession(WebApplication.Data) do
  begin
   SetLength(MatrizConseqNorm, Numalt, Numcrit);
   SetLength(Max, Numcrit);
   SetLength(Min, Numcrit);
   // ESCALA INTERVALAR
       for i := 0 to Numcrit - 1 do
      begin
      if (tipocrit[i] <> 2) and (tipocrit[i] <> 3) then
      begin
      Max[i] := -9999999999;
      Min[i] := 9999999999;
      for j := 0 to Numalt - 1 do
      begin
      if Max[i] < MatrizConseq[j, i] then
      Max[i] := MatrizConseq[j, i];
      if Min[i] > MatrizConseq[j, i] then
      Min[i] := MatrizConseq[j, i];
      end; // j
      end // if
      else
      begin
      if niveis[i] = 2 then
      begin
      Max[i] := 1;
      Min[i] := 0;
      end;
      if niveis[i] > 2 then
      begin
      TIWUserSession(WebApplication.Data).Max[i] := niveis[i];
      TIWUserSession(WebApplication.Data).Min[i] := 1;
      end;
      end;
      end; // i
      for j := 0 to Numalt - 1 do
      for i := 0 to Numcrit - 1 do
      begin
      if (tipocrit[i] = 1) or (tipocrit[i] = 3) or (tipocrit[i] = 5) then
      MatrizConseqNorm[j, i] := (MatrizConseq[j, i] - Min[i]) /
      (Max[i] - Min[i]);
      /// Normalização para os critérios de minimização //////////////////////
      if (tipocrit[i] = 0) or (tipocrit[i] = 2) or (tipocrit[i] = 4) then
      TIWUserSession(WebApplication.Data).MatrizConseqNorm[j, i] :=
      (MatrizConseq[j, i] - Max[i]) / (-Max[i] + Min[i]);
      end;


   {
   // ESCALA DE RAZÃO
    SetLength(MatrizConseqNorm, Numalt, Numcrit);
    SetLength(Maxl, Numcrit);
    SetLength(Minl, Numcrit);

    SetLength(MaxA, Numcrit);
    SetLength(MinA, Numcrit);
    SetLength(MaxB, Numcrit);
    SetLength(MinB, Numcrit);

    // ** Buscando maximos e minimos locais (Escala de Razão) **//
    for j := 0 to Numcrit - 1 do
    begin
      Maxl[j] := MatrizConseq[0, j];
      Minl[j] := MatrizConseq[0, j];
    end;
    for j := 0 to Numcrit - 1 do
      for i := 0 to Numalt - 1 do
      begin
        if MatrizConseq[i, j] > Maxl[j] then
          Maxl[j] := MatrizConseq[i, j];
        if MatrizConseq[i, j] < Minl[j] then
          Minl[j] := MatrizConseq[i, j];
      end;

    for j := 0 to Numcrit - 1 do
      for i := 0 to Numalt - 1 do
      begin
        if (tipocrit[j] = 1) or (tipocrit[j] = 3) or (tipocrit[j] = 5) then
        begin // dos criterios de maximização

          // * ConsequenciaPositivas ou zero *//
          if (Maxl[j] > 0) and (Minl[j] >= 0) then
            TIWUserSession(WebApplication.Data).MatrizConseqNorm[i, j] :=
              (TIWUserSession(WebApplication.Data).MatrizConseq[i, j] / Maxl[j]
              );
          // fim caso para consequencias são diferente de zero

          // * Consequencias mistas (+, -, 0)*//
          if (Maxl[j] >= 0) and (Minl[j] < 0) then
            TIWUserSession(WebApplication.Data).MatrizConseqNorm[i, j] :=
              TIWUserSession(WebApplication.Data).MatrizConseq[i, j] /
              (Maxl[j] - Minl[j]) + 1 - Maxl[j] / (Maxl[j] - Minl[j]);

          // * Consequencias negativas *//
          if (Minl[j] < 0) and (Maxl[j] < 0) then
            TIWUserSession(WebApplication.Data).MatrizConseqNorm[i, j] := Maxl
              [j] / TIWUserSession(WebApplication.Data).MatrizConseq[i, j];

        end; // Fimdoscritérios de maximização

        if (tipocrit[j] = 0) or (tipocrit[j] = 2) or (tipocrit[j] = 4) then
        begin // inicio dos critérios de minimização

          // * ConsequenciaPositivas *//
          if (MatrizConseq[i, j] <> 0) and (Minl[j] > 0) then
            TIWUserSession(WebApplication.Data).MatrizConseqNorm[i, j] :=
              (Minl[j] / TIWUserSession(WebApplication.Data)
                .MatrizConseq[i, j]);
          // para consequencias diferentes de zero

          // * Consequencias mistas (+, -, 0)*//
          if (Minl[j] <= 0) and (Maxl[j] >= 0) then
            TIWUserSession(WebApplication.Data).MatrizConseqNorm[i, j] :=
              TIWUserSession(WebApplication.Data).MatrizConseq[i, j] /
              (Minl[j] - Maxl[j]) + 1 - Minl[j] / (Minl[j] - Maxl[j]);

          // * Consequencias negativas *//
          if (Minl[j] < 0) and (Maxl[j] < 0) then
            TIWUserSession(WebApplication.Data).MatrizConseqNorm[i, j] :=
              TIWUserSession(WebApplication.Data).MatrizConseq[i, j] / Minl[j];

        end; // fim dos critérios de minimização
      end; // i
              }
  end; // with
end; // fim do procedimento

procedure TIWForm1.DecrementaValorLimiar(j, k: Integer);
var
  valorAtual, novoValor: Double;
begin
  if grdlimiares.DynEdits[j, k] = '1.00' then
    grdlimiares.DynEdits[j, k] := '0.99'
  else if grdlimiares.DynEdits[j, k] = '0.99' then
    grdlimiares.DynEdits[j, k] := '0.98'
  else if grdlimiares.DynEdits[j, k] = '0.98' then
    grdlimiares.DynEdits[j, k] := '0.97'
  else if grdlimiares.DynEdits[j, k] = '0.97' then
    grdlimiares.DynEdits[j, k] := '0.96'
  else if grdlimiares.DynEdits[j, k] = '0.96' then
    grdlimiares.DynEdits[j, k] := '0.95'
  else if grdlimiares.DynEdits[j, k] = '0.95' then
    grdlimiares.DynEdits[j, k] := '0.94'
  else if grdlimiares.DynEdits[j, k] = '0.94' then
    grdlimiares.DynEdits[j, k] := '0.93'
  else if grdlimiares.DynEdits[j, k] = '0.93' then
    grdlimiares.DynEdits[j, k] := '0.92'
  else if grdlimiares.DynEdits[j, k] = '0.92' then
    grdlimiares.DynEdits[j, k] := '0.91'
  else if grdlimiares.DynEdits[j, k] = '0.91' then
    grdlimiares.DynEdits[j, k] := '0.90'
  else if grdlimiares.DynEdits[j, k] = '0.90' then
    grdlimiares.DynEdits[j, k] := '0.89'
  else if grdlimiares.DynEdits[j, k] = '0.89' then
    grdlimiares.DynEdits[j, k] := '0.88'
  else if grdlimiares.DynEdits[j, k] = '0.88' then
    grdlimiares.DynEdits[j, k] := '0.87'
  else if grdlimiares.DynEdits[j, k] = '0.87' then
    grdlimiares.DynEdits[j, k] := '0.86'
  else if grdlimiares.DynEdits[j, k] = '0.86' then
    grdlimiares.DynEdits[j, k] := '0.85'
  else if grdlimiares.DynEdits[j, k] = '0.85' then
    grdlimiares.DynEdits[j, k] := '0.84'
  else if grdlimiares.DynEdits[j, k] = '0.84' then
    grdlimiares.DynEdits[j, k] := '0.83'
  else if grdlimiares.DynEdits[j, k] = '0.83' then
    grdlimiares.DynEdits[j, k] := '0.82'
  else if grdlimiares.DynEdits[j, k] = '0.82' then
    grdlimiares.DynEdits[j, k] := '0.81'
  else if grdlimiares.DynEdits[j, k] = '0.81' then
    grdlimiares.DynEdits[j, k] := '0.80'
  else if grdlimiares.DynEdits[j, k] = '0.80' then
    grdlimiares.DynEdits[j, k] := '0.79'
  else if grdlimiares.DynEdits[j, k] = '0.79' then
    grdlimiares.DynEdits[j, k] := '0.78'
  else if grdlimiares.DynEdits[j, k] = '0.78' then
    grdlimiares.DynEdits[j, k] := '0.77'
  else if grdlimiares.DynEdits[j, k] = '0.77' then
    grdlimiares.DynEdits[j, k] := '0.76'
  else if grdlimiares.DynEdits[j, k] = '0.76' then
    grdlimiares.DynEdits[j, k] := '0.75'
  else if grdlimiares.DynEdits[j, k] = '0.75' then
    grdlimiares.DynEdits[j, k] := '0.74'
  else if grdlimiares.DynEdits[j, k] = '0.74' then
    grdlimiares.DynEdits[j, k] := '0.73'
  else if grdlimiares.DynEdits[j, k] = '0.73' then
    grdlimiares.DynEdits[j, k] := '0.72'
  else if grdlimiares.DynEdits[j, k] = '0.72' then
    grdlimiares.DynEdits[j, k] := '0.71'
  else if grdlimiares.DynEdits[j, k] = '0.71' then
    grdlimiares.DynEdits[j, k] := '0.70'
  else if grdlimiares.DynEdits[j, k] = '0.70' then
    grdlimiares.DynEdits[j, k] := '0.69'
  else if grdlimiares.DynEdits[j, k] = '0.69' then
    grdlimiares.DynEdits[j, k] := '0.68'
  else if grdlimiares.DynEdits[j, k] = '0.68' then
    grdlimiares.DynEdits[j, k] := '0.67'
  else if grdlimiares.DynEdits[j, k] = '0.67' then
    grdlimiares.DynEdits[j, k] := '0.66'
  else if grdlimiares.DynEdits[j, k] = '0.66' then
    grdlimiares.DynEdits[j, k] := '0.65'
  else if grdlimiares.DynEdits[j, k] = '0.65' then
    grdlimiares.DynEdits[j, k] := '0.64'
  else if grdlimiares.DynEdits[j, k] = '0.64' then
    grdlimiares.DynEdits[j, k] := '0.63'
  else if grdlimiares.DynEdits[j, k] = '0.63' then
    grdlimiares.DynEdits[j, k] := '0.62'
  else if grdlimiares.DynEdits[j, k] = '0.62' then
    grdlimiares.DynEdits[j, k] := '0.61'
  else if grdlimiares.DynEdits[j, k] = '0.61' then
    grdlimiares.DynEdits[j, k] := '0.60'
  else if grdlimiares.DynEdits[j, k] = '0.60' then
    grdlimiares.DynEdits[j, k] := '0.59'
  else if grdlimiares.DynEdits[j, k] = '0.59' then
    grdlimiares.DynEdits[j, k] := '0.58'
  else if grdlimiares.DynEdits[j, k] = '0.58' then
    grdlimiares.DynEdits[j, k] := '0.57'
  else if grdlimiares.DynEdits[j, k] = '0.57' then
    grdlimiares.DynEdits[j, k] := '0.56'
  else if grdlimiares.DynEdits[j, k] = '0.56' then
    grdlimiares.DynEdits[j, k] := '0.55'
  else if grdlimiares.DynEdits[j, k] = '0.55' then
    grdlimiares.DynEdits[j, k] := '0.54'
  else if grdlimiares.DynEdits[j, k] = '0.54' then
    grdlimiares.DynEdits[j, k] := '0.53'
  else if grdlimiares.DynEdits[j, k] = '0.53' then
    grdlimiares.DynEdits[j, k] := '0.52'
  else if grdlimiares.DynEdits[j, k] = '0.52' then
    grdlimiares.DynEdits[j, k] := '0.51'
  else if grdlimiares.DynEdits[j, k] = '0.51' then
    grdlimiares.DynEdits[j, k] := '0.50'
  else if grdlimiares.DynEdits[j, k] = '0.50' then
    grdlimiares.DynEdits[j, k] := '0.49'
  else if grdlimiares.DynEdits[j, k] = '0.49' then
    grdlimiares.DynEdits[j, k] := '0.48'
  else if grdlimiares.DynEdits[j, k] = '0.48' then
    grdlimiares.DynEdits[j, k] := '0.47'
  else if grdlimiares.DynEdits[j, k] = '0.47' then
    grdlimiares.DynEdits[j, k] := '0.46'
  else if grdlimiares.DynEdits[j, k] = '0.46' then
    grdlimiares.DynEdits[j, k] := '0.45'
  else if grdlimiares.DynEdits[j, k] = '0.45' then
    grdlimiares.DynEdits[j, k] := '0.44'
  else if grdlimiares.DynEdits[j, k] = '0.44' then
    grdlimiares.DynEdits[j, k] := '0.43'
  else if grdlimiares.DynEdits[j, k] = '0.43' then
    grdlimiares.DynEdits[j, k] := '0.42'
  else if grdlimiares.DynEdits[j, k] = '0.42' then
    grdlimiares.DynEdits[j, k] := '0.41'
  else if grdlimiares.DynEdits[j, k] = '0.41' then
    grdlimiares.DynEdits[j, k] := '0.40'
  else if grdlimiares.DynEdits[j, k] = '0.40' then
    grdlimiares.DynEdits[j, k] := '0.39'
  else if grdlimiares.DynEdits[j, k] = '0.39' then
    grdlimiares.DynEdits[j, k] := '0.38'
  else if grdlimiares.DynEdits[j, k] = '0.38' then
    grdlimiares.DynEdits[j, k] := '0.37'
  else if grdlimiares.DynEdits[j, k] = '0.37' then
    grdlimiares.DynEdits[j, k] := '0.36'
  else if grdlimiares.DynEdits[j, k] = '0.36' then
    grdlimiares.DynEdits[j, k] := '0.35'
  else if grdlimiares.DynEdits[j, k] = '0.35' then
    grdlimiares.DynEdits[j, k] := '0.34'
  else if grdlimiares.DynEdits[j, k] = '0.34' then
    grdlimiares.DynEdits[j, k] := '0.33'
  else if grdlimiares.DynEdits[j, k] = '0.33' then
    grdlimiares.DynEdits[j, k] := '0.32'
  else if grdlimiares.DynEdits[j, k] = '0.32' then
    grdlimiares.DynEdits[j, k] := '0.31'
  else if grdlimiares.DynEdits[j, k] = '0.31' then
    grdlimiares.DynEdits[j, k] := '0.30'
  else if grdlimiares.DynEdits[j, k] = '0.30' then
    grdlimiares.DynEdits[j, k] := '0.29'
  else if grdlimiares.DynEdits[j, k] = '0.29' then
    grdlimiares.DynEdits[j, k] := '0.28'
  else if grdlimiares.DynEdits[j, k] = '0.28' then
    grdlimiares.DynEdits[j, k] := '0.27'
  else if grdlimiares.DynEdits[j, k] = '0.27' then
    grdlimiares.DynEdits[j, k] := '0.26'
  else if grdlimiares.DynEdits[j, k] = '0.26' then
    grdlimiares.DynEdits[j, k] := '0.25'
  else if grdlimiares.DynEdits[j, k] = '0.25' then
    grdlimiares.DynEdits[j, k] := '0.24'
  else if grdlimiares.DynEdits[j, k] = '0.24' then
    grdlimiares.DynEdits[j, k] := '0.23'
  else if grdlimiares.DynEdits[j, k] = '0.23' then
    grdlimiares.DynEdits[j, k] := '0.22'
  else if grdlimiares.DynEdits[j, k] = '0.22' then
    grdlimiares.DynEdits[j, k] := '0.21'
  else if grdlimiares.DynEdits[j, k] = '0.21' then
    grdlimiares.DynEdits[j, k] := '0.20'
  else if grdlimiares.DynEdits[j, k] = '0.20' then
    grdlimiares.DynEdits[j, k] := '0.19'
  else if grdlimiares.DynEdits[j, k] = '0.19' then
    grdlimiares.DynEdits[j, k] := '0.18'
  else if grdlimiares.DynEdits[j, k] = '0.18' then
    grdlimiares.DynEdits[j, k] := '0.17'
  else if grdlimiares.DynEdits[j, k] = '0.17' then
    grdlimiares.DynEdits[j, k] := '0.16'
  else if grdlimiares.DynEdits[j, k] = '0.16' then
    grdlimiares.DynEdits[j, k] := '0.15'
  else if grdlimiares.DynEdits[j, k] = '0.15' then
    grdlimiares.DynEdits[j, k] := '0.14'
  else if grdlimiares.DynEdits[j, k] = '0.14' then
    grdlimiares.DynEdits[j, k] := '0.13'
  else if grdlimiares.DynEdits[j, k] = '0.13' then
    grdlimiares.DynEdits[j, k] := '0.12'
  else if grdlimiares.DynEdits[j, k] = '0.12' then
    grdlimiares.DynEdits[j, k] := '0.11'
  else if grdlimiares.DynEdits[j, k] = '0.11' then
    grdlimiares.DynEdits[j, k] := '0.10'
  else if grdlimiares.DynEdits[j, k] = '0.10' then
    grdlimiares.DynEdits[j, k] := '0.09'
  else if grdlimiares.DynEdits[j, k] = '0.09' then
    grdlimiares.DynEdits[j, k] := '0.08'
  else if grdlimiares.DynEdits[j, k] = '0.08' then
    grdlimiares.DynEdits[j, k] := '0.07'
  else if grdlimiares.DynEdits[j, k] = '0.07' then
    grdlimiares.DynEdits[j, k] := '0.06'
  else if grdlimiares.DynEdits[j, k] = '0.06' then
    grdlimiares.DynEdits[j, k] := '0.05'
  else if grdlimiares.DynEdits[j, k] = '0.05' then
    grdlimiares.DynEdits[j, k] := '0.04'
  else if grdlimiares.DynEdits[j, k] = '0.04' then
    grdlimiares.DynEdits[j, k] := '0.03'
  else if grdlimiares.DynEdits[j, k] = '0.03' then
    grdlimiares.DynEdits[j, k] := '0.02'
  else if grdlimiares.DynEdits[j, k] = '0.02' then
    grdlimiares.DynEdits[j, k] := '0.01'
  else
    grdlimiares.DynEdits[j, k] := '0.00';
end;

procedure TIWForm1.DecrementaValorLimiarPromethee(j, k: Integer);
var
  valorAtual, novoValor: Double;
begin
  if grdlimiarespromethee.DynEdits[j, k] = '1.00' then
    grdlimiarespromethee.DynEdits[j, k] := '0.99'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.99' then
    grdlimiarespromethee.DynEdits[j, k] := '0.98'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.98' then
    grdlimiarespromethee.DynEdits[j, k] := '0.97'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.97' then
    grdlimiarespromethee.DynEdits[j, k] := '0.96'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.96' then
    grdlimiarespromethee.DynEdits[j, k] := '0.95'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.95' then
    grdlimiarespromethee.DynEdits[j, k] := '0.94'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.94' then
    grdlimiarespromethee.DynEdits[j, k] := '0.93'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.93' then
    grdlimiarespromethee.DynEdits[j, k] := '0.92'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.92' then
    grdlimiarespromethee.DynEdits[j, k] := '0.91'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.91' then
    grdlimiarespromethee.DynEdits[j, k] := '0.90'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.90' then
    grdlimiarespromethee.DynEdits[j, k] := '0.89'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.89' then
    grdlimiarespromethee.DynEdits[j, k] := '0.88'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.88' then
    grdlimiarespromethee.DynEdits[j, k] := '0.87'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.87' then
    grdlimiarespromethee.DynEdits[j, k] := '0.86'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.86' then
    grdlimiarespromethee.DynEdits[j, k] := '0.85'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.85' then
    grdlimiarespromethee.DynEdits[j, k] := '0.84'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.84' then
    grdlimiarespromethee.DynEdits[j, k] := '0.83'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.83' then
    grdlimiarespromethee.DynEdits[j, k] := '0.82'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.82' then
    grdlimiarespromethee.DynEdits[j, k] := '0.81'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.81' then
    grdlimiarespromethee.DynEdits[j, k] := '0.80'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.80' then
    grdlimiarespromethee.DynEdits[j, k] := '0.79'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.79' then
    grdlimiarespromethee.DynEdits[j, k] := '0.78'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.78' then
    grdlimiarespromethee.DynEdits[j, k] := '0.77'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.77' then
    grdlimiarespromethee.DynEdits[j, k] := '0.76'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.76' then
    grdlimiarespromethee.DynEdits[j, k] := '0.75'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.75' then
    grdlimiarespromethee.DynEdits[j, k] := '0.74'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.74' then
    grdlimiarespromethee.DynEdits[j, k] := '0.73'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.73' then
    grdlimiarespromethee.DynEdits[j, k] := '0.72'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.72' then
    grdlimiarespromethee.DynEdits[j, k] := '0.71'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.71' then
    grdlimiarespromethee.DynEdits[j, k] := '0.70'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.70' then
    grdlimiarespromethee.DynEdits[j, k] := '0.69'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.69' then
    grdlimiarespromethee.DynEdits[j, k] := '0.68'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.68' then
    grdlimiarespromethee.DynEdits[j, k] := '0.67'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.67' then
    grdlimiarespromethee.DynEdits[j, k] := '0.66'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.66' then
    grdlimiarespromethee.DynEdits[j, k] := '0.65'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.65' then
    grdlimiarespromethee.DynEdits[j, k] := '0.64'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.64' then
    grdlimiarespromethee.DynEdits[j, k] := '0.63'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.63' then
    grdlimiarespromethee.DynEdits[j, k] := '0.62'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.62' then
    grdlimiarespromethee.DynEdits[j, k] := '0.61'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.61' then
    grdlimiarespromethee.DynEdits[j, k] := '0.60'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.60' then
    grdlimiarespromethee.DynEdits[j, k] := '0.59'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.59' then
    grdlimiarespromethee.DynEdits[j, k] := '0.58'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.58' then
    grdlimiarespromethee.DynEdits[j, k] := '0.57'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.57' then
    grdlimiarespromethee.DynEdits[j, k] := '0.56'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.56' then
    grdlimiarespromethee.DynEdits[j, k] := '0.55'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.55' then
    grdlimiarespromethee.DynEdits[j, k] := '0.54'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.54' then
    grdlimiarespromethee.DynEdits[j, k] := '0.53'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.53' then
    grdlimiarespromethee.DynEdits[j, k] := '0.52'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.52' then
    grdlimiarespromethee.DynEdits[j, k] := '0.51'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.51' then
    grdlimiarespromethee.DynEdits[j, k] := '0.50'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.50' then
    grdlimiarespromethee.DynEdits[j, k] := '0.49'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.49' then
    grdlimiarespromethee.DynEdits[j, k] := '0.48'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.48' then
    grdlimiarespromethee.DynEdits[j, k] := '0.47'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.47' then
    grdlimiarespromethee.DynEdits[j, k] := '0.46'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.46' then
    grdlimiarespromethee.DynEdits[j, k] := '0.45'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.45' then
    grdlimiarespromethee.DynEdits[j, k] := '0.44'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.44' then
    grdlimiarespromethee.DynEdits[j, k] := '0.43'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.43' then
    grdlimiarespromethee.DynEdits[j, k] := '0.42'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.42' then
    grdlimiarespromethee.DynEdits[j, k] := '0.41'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.41' then
    grdlimiarespromethee.DynEdits[j, k] := '0.40'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.40' then
    grdlimiarespromethee.DynEdits[j, k] := '0.39'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.39' then
    grdlimiarespromethee.DynEdits[j, k] := '0.38'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.38' then
    grdlimiarespromethee.DynEdits[j, k] := '0.37'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.37' then
    grdlimiarespromethee.DynEdits[j, k] := '0.36'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.36' then
    grdlimiarespromethee.DynEdits[j, k] := '0.35'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.35' then
    grdlimiarespromethee.DynEdits[j, k] := '0.34'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.34' then
    grdlimiarespromethee.DynEdits[j, k] := '0.33'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.33' then
    grdlimiarespromethee.DynEdits[j, k] := '0.32'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.32' then
    grdlimiarespromethee.DynEdits[j, k] := '0.31'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.31' then
    grdlimiarespromethee.DynEdits[j, k] := '0.30'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.30' then
    grdlimiarespromethee.DynEdits[j, k] := '0.29'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.29' then
    grdlimiarespromethee.DynEdits[j, k] := '0.28'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.28' then
    grdlimiarespromethee.DynEdits[j, k] := '0.27'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.27' then
    grdlimiarespromethee.DynEdits[j, k] := '0.26'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.26' then
    grdlimiarespromethee.DynEdits[j, k] := '0.25'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.25' then
    grdlimiarespromethee.DynEdits[j, k] := '0.24'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.24' then
    grdlimiarespromethee.DynEdits[j, k] := '0.23'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.23' then
    grdlimiarespromethee.DynEdits[j, k] := '0.22'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.22' then
    grdlimiarespromethee.DynEdits[j, k] := '0.21'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.21' then
    grdlimiarespromethee.DynEdits[j, k] := '0.20'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.20' then
    grdlimiarespromethee.DynEdits[j, k] := '0.19'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.19' then
    grdlimiarespromethee.DynEdits[j, k] := '0.18'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.18' then
    grdlimiarespromethee.DynEdits[j, k] := '0.17'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.17' then
    grdlimiarespromethee.DynEdits[j, k] := '0.16'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.16' then
    grdlimiarespromethee.DynEdits[j, k] := '0.15'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.15' then
    grdlimiarespromethee.DynEdits[j, k] := '0.14'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.14' then
    grdlimiarespromethee.DynEdits[j, k] := '0.13'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.13' then
    grdlimiarespromethee.DynEdits[j, k] := '0.12'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.12' then
    grdlimiarespromethee.DynEdits[j, k] := '0.11'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.11' then
    grdlimiarespromethee.DynEdits[j, k] := '0.10'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.10' then
    grdlimiarespromethee.DynEdits[j, k] := '0.09'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.09' then
    grdlimiarespromethee.DynEdits[j, k] := '0.08'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.08' then
    grdlimiarespromethee.DynEdits[j, k] := '0.07'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.07' then
    grdlimiarespromethee.DynEdits[j, k] := '0.06'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.06' then
    grdlimiarespromethee.DynEdits[j, k] := '0.05'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.05' then
    grdlimiarespromethee.DynEdits[j, k] := '0.04'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.04' then
    grdlimiarespromethee.DynEdits[j, k] := '0.03'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.03' then
    grdlimiarespromethee.DynEdits[j, k] := '0.02'
  else if grdlimiarespromethee.DynEdits[j, k] = '0.02' then
    grdlimiarespromethee.DynEdits[j, k] := '0.01'
  else
    grdlimiarespromethee.DynEdits[j, k] := '0.00';
end;

procedure TIWForm1.edtnumaltAsyncExit(Sender: TObject;
  EventParams: TStringList);
var
  valor: Integer;
begin
  if not trystrtoint(edtnumalt.Text, valor) then
    edtnumalt.Clear;
end;

procedure TIWForm1.edtnumaltSubmit(Sender: TObject);
var
  valor: Integer;
begin
  if not trystrtoint(edtnumalt.Text, valor) then
    edtnumalt.Clear;
end;

procedure TIWForm1.btnaltxaltClick(Sender: TObject);
begin
rgnaltxalt.Visible := true;
rgnelicitacaoalt.visible := false;
end;

procedure TIWForm1.btnatualizarlimiaresClick(Sender: TObject);
var
  i, j, k, qtdpareto: Integer;
  auxregra, aux1, aux2, aux3: real;
  bar: array of TBarSeries;
  SolRegra1, SolRegra2, SolRegra3: Integer;
  epsilon, omega: array of real;
  totalacumulado: real;
  recomendacao: boolean;
  Px: real;
  faixaprob: array of real;
  aux: string;
  mudança: boolean;
begin
  with TIWUserSession(WebApplication.Data) do
  begin
    SetLength(epsilon, 4);
    SetLength(omega, 4);
    SetLength(faixaprob, 4);
    for j := 1 to 2 do
    begin
      if (trocapv(grdlimiares.DynEdits[j, 0])) < 0.03 then
        grdlimiares.DynEdits[j, 0] := '0.03';
      if (trocapv(grdlimiares.DynEdits[j, 1])) < 0.02 then
        grdlimiares.DynEdits[j, 1] := '0.02';
      if (trocapv(grdlimiares.DynEdits[j, 2])) < 0.01 then
        grdlimiares.DynEdits[j, 2] := '0.01';

      if (trocapv(grdlimiares.DynEdits[j, 3])) > 0.97 then
        grdlimiares.DynEdits[j, 3] := '0.97';
      if (trocapv(grdlimiares.DynEdits[j, 2])) > 0.98 then
        grdlimiares.DynEdits[j, 2] := '0.98';
      if (trocapv(grdlimiares.DynEdits[j, 1])) > 0.99 then
        grdlimiares.DynEdits[j, 1] := '0.99';

      for i := 0 to grdlimiares.rowcount - 2 do
        for k := i + 1 to grdlimiares.rowcount - 1 do
        begin
          if (trocapv(grdlimiares.DynEdits[j, i])) <
            (trocapv(grdlimiares.DynEdits[j, k])) then
          begin
            aux := grdlimiares.DynEdits[j, i];
            grdlimiares.DynEdits[j, i] := grdlimiares.DynEdits[j, k];
            grdlimiares.DynEdits[j, k] := aux;
          end;
          if (trocapv(grdlimiares.DynEdits[j, i])) =
            (trocapv(grdlimiares.DynEdits[j, k])) then
          begin
            DecrementaValorLimiar(j, k);
          end;
        end;

    end;
    for i := 0 to 3 do
    begin
      epsilon[i] := trocapv(grdlimiares.DynEdits[1, i]);
      omega[i] := trocapv(grdlimiares.DynEdits[2, i]);
    end;

    faixaprob[0] := StrToFloat(cmbbxp1.Items[cmbbxp1.ItemIndex]);
    faixaprob[1] := StrToFloat(cmbbxp2.Items[cmbbxp2.ItemIndex]);
    faixaprob[2] := StrToFloat(cmbbxp3.Items[cmbbxp3.ItemIndex]);
    faixaprob[3] := StrToFloat(cmbbxp4.Items[cmbbxp4.ItemIndex]);

    Px := ((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
          [OrdemSol[2]]) / Total) * 100;

    recomendacao := false;

    grdrecomendacao.Clear;
    grdrecomendacao.rowcount := 4;
    grdrecomendacao.ColumnCount := 2;
    for i := 0 to grdrecomendacao.ColumnCount - 1 do
    begin

      grdrecomendacao.Cell[0, i].Font.Color := ClWebRed;
      grdrecomendacao.Cell[0, i].Font.Style := [fsbold];
      grdrecomendacao.Cell[0, i].BGColor := clWebWhiteSmoke;
      for j := 0 to grdrecomendacao.rowcount - 1 do
      begin
        grdrecomendacao.Cell[j, i].Alignment := taCenter;
        grdrecomendacao.Cell[j, i].wrap := True;
      end;
    end;

    grdrecomendacao.Cell[0, 0].Text := 'Solution';
    grdrecomendacao.Cell[0, 1].Text := 'Possibles Recommendations';
    grdrecomendacao.Cell[1, 1].Text := 'Best Alternative';
    grdrecomendacao.Cell[2, 1].Text := 'Two/Three Alternatives are Competitive';
    grdrecomendacao.Cell[3, 1].Text := 'Unable to make it';
    txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' + grdperdamediaemaxima.Cell
      [1, 0].Text + ')' + ' + ' + 'P(' + grdperdamediaemaxima.Cell[2, 0]
      .Text + ')' + ' + ' + 'P(' + grdperdamediaemaxima.Cell[3, 0].Text + ')';
    // Para alt1
    if (TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]] / Total) >=
      (faixaprob[0] / 100) then // Se a alt[1] tiver uma probabilidade >= 80%, ela é recomendada
    begin
      Px := (TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
          / Total) * 100;
      if (TIWUserSession(WebApplication.Data).MaximoGeralNaosol[OrdemSol[0]]
          < epsilon[0]) and (TIWUserSession(WebApplication.Data)
          .MediaGeralNaosol[OrdemSol[0]] < omega[0]) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[1, 0].Text = '' then
              grdrecomendacao.Cell[1, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
        recomendacao := false;
    end;
    if ((ResultSol[OrdemSol[0]] / Total) < (faixaprob[0] / 100)) and
      ((ResultSol[OrdemSol[0]] / Total) >= (faixaprob[1] / 100)) and
      (recomendacao = false) then
    begin
      Px := (TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
          / Total) * 100;
      if (MaximoGeralNaosol[OrdemSol[0]] < epsilon[1]) and
        (MediaGeralNaosol[OrdemSol[0]] < omega[1]) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[1, 0].Text = '' then
              grdrecomendacao.Cell[1, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if ((ResultSol[OrdemSol[0]] / Total) < (faixaprob[1] / 100)) and
      ((ResultSol[OrdemSol[0]] / Total) >= (faixaprob[2] / 100)) and
      (recomendacao = false) then
    begin
      Px := (TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
          / Total) * 100;
      if (MaximoGeralNaosol[OrdemSol[0]] < epsilon[2]) and
        (MediaGeralNaosol[OrdemSol[0]] < omega[2]) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[1, 0].Text = '' then
              grdrecomendacao.Cell[1, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if ((ResultSol[OrdemSol[0]] / Total) < (faixaprob[2] / 100)) and
      ((ResultSol[OrdemSol[0]] / Total) >= (faixaprob[3] / 100)) and
      (recomendacao = false) then
    begin
      Px := (TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
          / Total) * 100;
      if (MaximoGeralNaosol[OrdemSol[0]] < epsilon[3]) and
        (MediaGeralNaosol[OrdemSol[0]] < omega[3]) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[1, 0].Text = '' then
              grdrecomendacao.Cell[1, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;

    // Para alt1 e alt2
    if (((TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]] + ResultSol
            [OrdemSol[1]]) / Total) >= (faixaprob[0] / 100)) and
      (recomendacao = false) then
    // Se a alt[1] tiver uma probabilidade >= 80%, ela é recomendada
    begin
      Px := ((TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
            + ResultSol[OrdemSol[1]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[0]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[0])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[0]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[0])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if TIWUserSession(WebApplication.Data).MatrizSol[OrdemSol[0], j]
            = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaxima.Cell[2, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
        recomendacao := false;
    end;
    if (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]]) / Total) <
        (faixaprob[0] / 100)) and
      (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]]) / Total) >=
        (faixaprob[1] / 100)) and (recomendacao = false) then
    begin
      Px := ((TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
            + ResultSol[OrdemSol[1]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[1]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[1])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[1]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[1])) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaxima.Cell[2, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]]) / Total) <
        (faixaprob[1] / 100)) and
      (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]]) / Total) >=
        (faixaprob[2] / 100)) and (recomendacao = false) then
    begin
      Px := ((TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
            + ResultSol[OrdemSol[1]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[2]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[2])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[2]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[2])) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaxima.Cell[2, 0].Text + ')';

        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]]) / Total) <
        (faixaprob[2] / 100)) and
      (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]]) / Total) >=
        (faixaprob[3] / 100)) and (recomendacao = false) then
    begin
      Px := ((TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
            + ResultSol[OrdemSol[1]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[3]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[3])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[3]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[3])) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaxima.Cell[2, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;

    // Para alt1, alt2 e alt3
    if ((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
          [OrdemSol[2]]) / Total) >= (faixaprob[0] / 100) then
    // Se a alt[1] tiver uma probabilidade >= 80%, ela é recomendada
    begin
      Px := ((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[0]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[0])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[0]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[0])) and
        ((MaximoGeralNaosol[OrdemSol[2]] < epsilon[0]) and
          (MediaGeralNaosol[OrdemSol[2]] < omega[0])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[2], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        recomendacao := True;
        exit;
      end
      else
        recomendacao := false;
    end;
    if (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) < (faixaprob[0] / 100)) and
      (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) >= (faixaprob[1] / 100)) and
      (recomendacao = false) then
    begin
      Px := ((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[1]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[1])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[1]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[1])) and
        ((MaximoGeralNaosol[OrdemSol[2]] < epsilon[1]) and
          (MediaGeralNaosol[OrdemSol[2]] < omega[1])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[2], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) < (faixaprob[1] / 100)) and
      (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) >= (faixaprob[2] / 100)) and
      (recomendacao = false) then
    begin
      Px := ((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[2]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[2])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[2]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[2])) and
        ((MaximoGeralNaosol[OrdemSol[2]] < epsilon[2]) and
          (MediaGeralNaosol[OrdemSol[2]] < omega[2])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[2], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) < (faixaprob[2] / 100)) and
      (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) >= (faixaprob[3] / 100)) and
      (recomendacao = false) then
    begin
      Px := ((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[3]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[3])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[3]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[3])) and
        ((MaximoGeralNaosol[OrdemSol[2]] < epsilon[3]) and
          (MediaGeralNaosol[OrdemSol[2]] < omega[3])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[2], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;

    // caso não passe em nenhuma regra
    if recomendacao = false then
    begin
      grdrecomendacao.Cell[3, 0].Text := 'Not Possible';
      grdrecomendacao.Cell[3, 0].Text := grdrecomendacao.Cell[3, 0]
        .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
    end;

  end;
end;

procedure TIWForm1.btnatualizarlimiaresPrometheeClick(Sender: TObject);
var
  i, j, k, qtdpareto: Integer;
  auxregra, aux1, aux2, aux3: real;
  bar: array of TBarSeries;
  SolRegra1, SolRegra2, SolRegra3: Integer;
  epsilon, omega, faixaprob: array of real;
  totalacumulado: real;
  recomendacao: boolean;
  Px: real;
  aux: string;
begin
  with TIWUserSession(WebApplication.Data) do
  begin
    SetLength(epsilon, 4);
    SetLength(omega, 4);
    SetLength(faixaprob, 4);
    for j := 1 to 2 do
    begin
      if (trocapv(grdlimiarespromethee.DynEdits[j, 0])) < 0.03 then
        grdlimiarespromethee.DynEdits[j, 0] := '0.03';
      if (trocapv(grdlimiarespromethee.DynEdits[j, 1])) < 0.02 then
        grdlimiarespromethee.DynEdits[j, 1] := '0.02';
      if (trocapv(grdlimiarespromethee.DynEdits[j, 2])) < 0.01 then
        grdlimiarespromethee.DynEdits[j, 2] := '0.01';

      if (trocapv(grdlimiarespromethee.DynEdits[j, 3])) > 0.97 then
        grdlimiarespromethee.DynEdits[j, 3] := '0.97';
      if (trocapv(grdlimiarespromethee.DynEdits[j, 2])) > 0.98 then
        grdlimiarespromethee.DynEdits[j, 2] := '0.98';
      if (trocapv(grdlimiarespromethee.DynEdits[j, 1])) > 0.99 then
        grdlimiarespromethee.DynEdits[j, 1] := '0.99';

      for i := 0 to grdlimiarespromethee.rowcount - 2 do
        for k := i + 1 to grdlimiarespromethee.rowcount - 1 do
        begin
          if (trocapv(grdlimiarespromethee.DynEdits[j, i])) <
            (trocapv(grdlimiarespromethee.DynEdits[j, k])) then
          begin
            aux := grdlimiarespromethee.DynEdits[j, i];
            grdlimiarespromethee.DynEdits[j, i] := grdlimiarespromethee.DynEdits
              [j, k];
            grdlimiarespromethee.DynEdits[j, k] := aux;
          end;
          if (trocapv(grdlimiarespromethee.DynEdits[j, i])) =
            (trocapv(grdlimiarespromethee.DynEdits[j, k])) then
          begin
            DecrementaValorLimiarPromethee(j, k);
          end;
        end;
    end;

    for i := 0 to 3 do
    begin
      epsilon[i] := trocapv(grdlimiarespromethee.DynEdits[1, i]);
      omega[i] := trocapv(grdlimiarespromethee.DynEdits[2, i]);
    end;

    faixaprob[0] := StrToFloat
      (cmbbxp1promethee.Items[cmbbxp1promethee.ItemIndex]);
    faixaprob[1] := StrToFloat
      (cmbbxp2promethee.Items[cmbbxp2promethee.ItemIndex]);
    faixaprob[2] := StrToFloat
      (cmbbxp3promethee.Items[cmbbxp3promethee.ItemIndex]);
    faixaprob[3] := StrToFloat
      (cmbbxp4promethee.Items[cmbbxp4promethee.ItemIndex]);
    Px := ((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
          [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
        / Total) * 100;

    recomendacao := false;

    grdrecomendacaopromethee.Clear;
    grdrecomendacaopromethee.rowcount := 4;
    grdrecomendacaopromethee.ColumnCount := 2;
    for i := 0 to grdrecomendacaopromethee.ColumnCount - 1 do
    begin

      grdrecomendacaopromethee.Cell[0, i].Font.Color := ClWebRed;
      grdrecomendacaopromethee.Cell[0, i].Font.Style := [fsbold];
      grdrecomendacaopromethee.Cell[0, i].BGColor := clWebWhiteSmoke;
      for j := 0 to grdrecomendacaopromethee.rowcount - 1 do
      begin
        grdrecomendacaopromethee.Cell[j, i].Alignment := taCenter;
        grdrecomendacaopromethee.Cell[j, i].wrap := True;
      end;
    end;

    grdrecomendacaopromethee.Cell[0, 0].Text := 'Solution';
    grdrecomendacaopromethee.Cell[0, 1].Text := 'Possibles Recommendations';
    grdrecomendacaopromethee.Cell[1, 1].Text := 'Best Alternative';
    grdrecomendacaopromethee.Cell[2, 1].Text :=
      'Two/Three Alternatives are Competitive';
    grdrecomendacaopromethee.Cell[3, 1].Text := 'Unable to make it';
    txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
      grdperdamediaemaximapromethee.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
      grdperdamediaemaximapromethee.Cell[2, 0].Text + ')' + ' + ' + 'P(' +
      grdperdamediaemaxima.Cell[3, 0].Text + ')';
    // Para alt1

    Px := ((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
          [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
        / Total) * 100;
    if (TIWUserSession(WebApplication.Data).ResultSolPromethee
        [ordemsolPromethee[0]] / Total) >= faixaprob[0] then // Se a alt[1] tiver uma probabilidade >= 80%, ela é recomendada
    begin
      Px := (TIWUserSession(WebApplication.Data).ResultSolPromethee
          [ordemsolPromethee[0]] / Total) * 100;
      if (TIWUserSession(WebApplication.Data).MaximoGeralNaosolPromethee
          [ordemsolPromethee[0]] < epsilon[0]) and
        (TIWUserSession(WebApplication.Data).MediaGeralNaoSolPromethee
          [ordemsolPromethee[0]] < omega[0]) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[1, 0].Text = '' then
              grdrecomendacaopromethee.Cell[1, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[1, 0].Text :=
                grdrecomendacaopromethee.Cell[1, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[1, 0].Text :=
          grdrecomendacaopromethee.Cell[1, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
        recomendacao := false;
    end;
    if ((ResultSolPromethee[ordemsolPromethee[0]] / Total) < faixaprob[0]) and
      ((ResultSolPromethee[ordemsolPromethee[0]] / Total) >= faixaprob[1]) and
      (recomendacao = false) then
    begin
      Px := (TIWUserSession(WebApplication.Data).ResultSolPromethee
          [ordemsolPromethee[0]] / Total) * 100;
      if (MaximoGeralNaosolPromethee[ordemsolPromethee[0]] < epsilon[1]) and
        (MediaGeralNaoSolPromethee[ordemsolPromethee[0]] < omega[1]) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[1, 0].Text = '' then
              grdrecomendacaopromethee.Cell[1, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[1, 0].Text :=
                grdrecomendacaopromethee.Cell[1, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[1, 0].Text :=
          grdrecomendacaopromethee.Cell[1, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if ((ResultSolPromethee[ordemsolPromethee[0]] / Total) < faixaprob[1]) and
      ((ResultSolPromethee[ordemsolPromethee[0]] / Total) >= faixaprob[2]) and
      (recomendacao = false) then
    begin
      Px := (TIWUserSession(WebApplication.Data).ResultSolPromethee
          [ordemsolPromethee[0]] / Total) * 100;
      if (MaximoGeralNaosolPromethee[ordemsolPromethee[0]] < epsilon[2]) and
        (MediaGeralNaoSolPromethee[ordemsolPromethee[0]] < omega[2]) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[1, 0].Text = '' then
              grdrecomendacaopromethee.Cell[1, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[1, 0].Text :=
                grdrecomendacaopromethee.Cell[1, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[1, 0].Text :=
          grdrecomendacaopromethee.Cell[1, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if ((ResultSolPromethee[ordemsolPromethee[0]] / Total) < faixaprob[2]) and
      ((ResultSolPromethee[ordemsolPromethee[0]] / Total) >= faixaprob[3]) and
      (recomendacao = false) then
    begin
      Px := (TIWUserSession(WebApplication.Data).ResultSolPromethee
          [ordemsolPromethee[0]] / Total) * 100;
      if (MaximoGeralNaosolPromethee[ordemsolPromethee[0]] < epsilon[3]) and
        (MediaGeralNaoSolPromethee[ordemsolPromethee[0]] < omega[3]) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[1, 0].Text = '' then
              grdrecomendacaopromethee.Cell[1, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[1, 0].Text :=
                grdrecomendacaopromethee.Cell[1, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[1, 0].Text :=
          grdrecomendacaopromethee.Cell[1, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;

    // Para alt1 e alt2
    if (((TIWUserSession(WebApplication.Data).ResultSolPromethee
            [ordemsolPromethee[0]] + ResultSolPromethee[ordemsolPromethee[1]])
          / Total) >= faixaprob[0]) and (recomendacao = false) then
    // Se a alt[1] tiver uma probabilidade >= 80%, ela é recomendada
    begin
      Px := ((TIWUserSession(WebApplication.Data).ResultSolPromethee
            [ordemsolPromethee[0]] + ResultSolPromethee[ordemsolPromethee[1]])
          / Total) * 100;
      if ((MaximoGeralNaosolPromethee[ordemsolPromethee[0]] < epsilon[0]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[0]] < omega[0])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[1]] < epsilon[0]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[1]] < omega[0])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if TIWUserSession(WebApplication.Data).matrizsolPromethee
            [ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;

        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[1], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[2, 0].Text :=
          grdrecomendacaopromethee.Cell[2, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[2, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
        recomendacao := false;
    end;
    if (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]]) / Total) < faixaprob[0]) and
      (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]]) / Total) >= faixaprob[1]) and
      (recomendacao = false) then
    begin
      Px := ((TIWUserSession(WebApplication.Data).ResultSolPromethee
            [ordemsolPromethee[0]] + ResultSolPromethee[ordemsolPromethee[1]])
          / Total) * 100;
      if ((MaximoGeralNaosolPromethee[ordemsolPromethee[0]] < epsilon[1]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[0]] < omega[1])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[1]] < epsilon[1]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[1]] < omega[1])) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;

        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[1], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[2, 0].Text :=
          grdrecomendacaopromethee.Cell[2, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[2, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]]) / Total) < faixaprob[1]) and
      (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]]) / Total) >= faixaprob[2]) and
      (recomendacao = false) then
    begin
      Px := ((TIWUserSession(WebApplication.Data).ResultSolPromethee
            [ordemsolPromethee[0]] + ResultSolPromethee[ordemsolPromethee[1]])
          / Total) * 100;
      if ((MaximoGeralNaosolPromethee[ordemsolPromethee[0]] < epsilon[2]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[0]] < omega[2])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[1]] < epsilon[2]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[1]] < omega[2])) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;

        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[1], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[2, 0].Text :=
          grdrecomendacaopromethee.Cell[2, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[2, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]]) / Total) < faixaprob[2]) and
      (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]]) / Total) >= faixaprob[3]) and
      (recomendacao = false) then
    begin
      Px := ((TIWUserSession(WebApplication.Data).ResultSolPromethee
            [ordemsolPromethee[0]] + ResultSolPromethee[ordemsolPromethee[1]])
          / Total) * 100;
      if ((MaximoGeralNaosolPromethee[ordemsolPromethee[0]] < epsilon[3]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[0]] < omega[3])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[1]] < epsilon[3]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[1]] < omega[3])) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;

        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[1], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[2, 0].Text :=
          grdrecomendacaopromethee.Cell[2, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[2, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;

    // Para alt1, alt2 e alt3
    if ((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
          [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
        / Total) >= faixaprob[0] then
    // Se a alt[1] tiver uma probabilidade >= 80%, ela é recomendada
    begin
      Px := ((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
          / Total) * 100;
      if ((MaximoGeralNaosolPromethee[ordemsolPromethee[0]] < epsilon[0]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[0]] < omega[0])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[1]] < epsilon[0]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[1]] < omega[0])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[2]] < epsilon[0]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[2]] < omega[0])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[1], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[2], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[2, 0].Text :=
          grdrecomendacaopromethee.Cell[2, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[2, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[3, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
        recomendacao := false;
    end;
    if (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
          / Total) < faixaprob[0]) and
      (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
          / Total) >= faixaprob[1]) and (recomendacao = false) then
    begin
      Px := ((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
          / Total) * 100;
      if ((MaximoGeralNaosolPromethee[ordemsolPromethee[0]] < epsilon[1]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[0]] < omega[1])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[1]] < epsilon[1]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[1]] < omega[1])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[2]] < epsilon[1]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[2]] < omega[1])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[1], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[2], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[2, 0].Text :=
          grdrecomendacaopromethee.Cell[2, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[2, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[3, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
          / Total) < faixaprob[1]) and
      (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
          / Total) >= faixaprob[2]) and (recomendacao = false) then
    begin
      Px := ((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
          / Total) * 100;
      if ((MaximoGeralNaosolPromethee[ordemsolPromethee[0]] < epsilon[2]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[0]] < omega[2])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[1]] < epsilon[2]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[1]] < omega[2])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[2]] < epsilon[2]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[2]] < omega[2])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[1], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[2], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[2, 0].Text :=
          grdrecomendacaopromethee.Cell[2, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[2, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[3, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
          / Total) < faixaprob[2]) and
      (((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
          / Total) >= faixaprob[3]) and (recomendacao = false) then
    begin
      Px := ((ResultSolPromethee[ordemsolPromethee[0]] + ResultSolPromethee
            [ordemsolPromethee[1]] + ResultSolPromethee[ordemsolPromethee[2]])
          / Total) * 100;
      if ((MaximoGeralNaosolPromethee[ordemsolPromethee[0]] < epsilon[3]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[0]] < omega[3])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[1]] < epsilon[3]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[1]] < omega[3])) and
        ((MaximoGeralNaosolPromethee[ordemsolPromethee[2]] < epsilon[3]) and
          (MediaGeralNaoSolPromethee[ordemsolPromethee[2]] < omega[3])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[0], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[1], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if matrizsolPromethee[ordemsolPromethee[2], j] = 1 then
          begin
            if grdrecomendacaopromethee.Cell[2, 0].Text = '' then
              grdrecomendacaopromethee.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacaopromethee.Cell[2, 0].Text :=
                grdrecomendacaopromethee.Cell[2, 0].Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacaopromethee.Cell[2, 0].Text :=
          grdrecomendacaopromethee.Cell[2, 0].Text + '(P(x) = ' + formatfloat
          ('0', Px) + '%)';
        txtdescricaoregraPromethee.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[2, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaximapromethee.Cell[3, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;

    // caso não passe em nenhuma regra
    if recomendacao = false then
      grdrecomendacaopromethee.Cell[3, 0].Text := 'Not Possible' + '(P(x) = ' +
        formatfloat('0', Px) + '%)';

  end;
end;

procedure TIWForm1.AtualizarPopulacao;
var
  i, j, k, y, w: Integer;
  auxnovapopulacaoOrdemCrit: array of array of Integer;
  auxNovaPopulacaoResultadoRoc: array of array of Double;
  auxNovaPopulacaoMatrizPOA: array of array of Integer;
  auxnperg: Integer;
begin
  with TIWUserSession(WebApplication.Data) do
  begin
    if nperg = 0 then // antes de gerar a primeira pergunta
    begin
      SetLength(tamanhonovapopulacao, nperg + 1);
      tamanhonovapopulacao[nperg] := length(CasesOrdemcrit) - 1;
      SetLength(TIWUserSession(WebApplication.Data).novapopulacaoOrdemCrit,
        tamanhonovapopulacao[nperg], Numcrit);
      SetLength(TIWUserSession(WebApplication.Data).NovaPopulacaoResultadoRoc,
        tamanhonovapopulacao[nperg], Numalt);
      SetLength(TIWUserSession(WebApplication.Data).NovaPopulacaoMatrizPOA,
        tamanhonovapopulacao[nperg], Numalt);

      for k := 0 to tamanhonovapopulacao[nperg] - 1 do
      begin
        for i := 0 to Numcrit - 1 do
          novapopulacaoOrdemCrit[k, i] := CasesOrdemcrit[k, i];
        for j := 0 to Numalt - 1 do
        begin
          NovaPopulacaoResultadoRoc[k, j] := ResultadoROC[k, j];
          NovaPopulacaoMatrizPOA[k, j] := MatrizPOA[k, j];
        end; // for j
      end; // for k
    end; // if nperg = 0
    if nperg > 0 then
    begin
      auxnperg := nperg - 1;
      SetLength(tamanhonovapopulacao, nperg + 1);
      SetLength(auxnovapopulacaoOrdemCrit, length(novapopulacaoOrdemCrit),
        Numcrit);
      SetLength(auxNovaPopulacaoResultadoRoc, length(novapopulacaoOrdemCrit),
        Numalt);
      SetLength(auxNovaPopulacaoMatrizPOA, length(novapopulacaoOrdemCrit),
        Numalt);

      if TIWUserSession(WebApplication.Data).VetorRespostaElicitacao
        [auxnperg, 0] = 0 then // a pergunta respondida foi crit x crit
      begin
        y := 0;
        for k := 0 to length(novapopulacaoOrdemCrit) - 1 do
        begin
          if novapopulacaoOrdemCrit[k, VetorRespostaElicitacao[auxnperg, 1]]
            < novapopulacaoOrdemCrit[k, VetorRespostaElicitacao[auxnperg, 2]]
            then
          begin
            for i := 0 to Numcrit - 1 do
              auxnovapopulacaoOrdemCrit[y, i] := novapopulacaoOrdemCrit[k, i];
            for j := 0 to Numalt - 1 do
            begin
              auxNovaPopulacaoResultadoRoc[y, j] := TIWUserSession
                (WebApplication.Data).NovaPopulacaoResultadoRoc[k, j];
              auxNovaPopulacaoMatrizPOA[y, j] := NovaPopulacaoMatrizPOA[k, j];
            end; // for j
            inc(y);
          end; // if
        end; // for k

        tamanhonovapopulacao[nperg] := y;

        Finalize(novapopulacaoOrdemCrit);
        Finalize(NovaPopulacaoResultadoRoc);
        Finalize(NovaPopulacaoMatrizPOA);

        SetLength(novapopulacaoOrdemCrit, tamanhonovapopulacao[nperg], Numcrit);
        SetLength(NovaPopulacaoResultadoRoc, tamanhonovapopulacao[nperg],
          Numalt);
        SetLength(NovaPopulacaoMatrizPOA, tamanhonovapopulacao[nperg], Numalt);

        for k := 0 to tamanhonovapopulacao[nperg] - 1 do
        begin
          for i := 0 to Numcrit - 1 do
            novapopulacaoOrdemCrit[k, i] := auxnovapopulacaoOrdemCrit[k, i];
          for j := 0 to Numalt - 1 do
          begin
            NovaPopulacaoResultadoRoc[k, j] := auxNovaPopulacaoResultadoRoc
              [k, j];
            NovaPopulacaoMatrizPOA[k, j] := auxNovaPopulacaoMatrizPOA[k, j];
          end; // for j
        end; // for k

      end;
      if VetorRespostaElicitacao[auxnperg, 0] = 1 then
      // a pergunta respondida foi alt x alt
      begin
        y := 0;
        for k := 0 to length(novapopulacaoOrdemCrit) - 1 do
        begin
          if NovaPopulacaoResultadoRoc[k, VetorRespostaElicitacao[auxnperg, 1]]
            > NovaPopulacaoResultadoRoc[k,
            VetorRespostaElicitacao[auxnperg, 2]] then
          begin
            for i := 0 to Numcrit - 1 do
              auxnovapopulacaoOrdemCrit[y, i] := novapopulacaoOrdemCrit[k, i];
            for j := 0 to Numalt - 1 do
            begin
              auxNovaPopulacaoResultadoRoc[y, j] := NovaPopulacaoResultadoRoc
                [k, j];
              auxNovaPopulacaoMatrizPOA[y, j] := NovaPopulacaoMatrizPOA[k, j];
            end; // for j
            inc(y);
          end; // if
        end; // for k

        tamanhonovapopulacao[nperg] := y;

        Finalize(novapopulacaoOrdemCrit);
        Finalize(NovaPopulacaoResultadoRoc);
        Finalize(NovaPopulacaoMatrizPOA);

        SetLength(novapopulacaoOrdemCrit, tamanhonovapopulacao[nperg], Numcrit);
        SetLength(NovaPopulacaoResultadoRoc, tamanhonovapopulacao[nperg],
          Numalt);
        SetLength(NovaPopulacaoMatrizPOA, tamanhonovapopulacao[nperg], Numalt);

        for k := 0 to tamanhonovapopulacao[nperg] - 1 do
        begin
          for i := 0 to Numcrit - 1 do
            TIWUserSession(WebApplication.Data).novapopulacaoOrdemCrit[k, i] :=
              auxnovapopulacaoOrdemCrit[k, i];
          for j := 0 to Numalt - 1 do
          begin
            TIWUserSession(WebApplication.Data).NovaPopulacaoResultadoRoc[k, j]
              := auxNovaPopulacaoResultadoRoc[k, j];
            TIWUserSession(WebApplication.Data).NovaPopulacaoMatrizPOA[k, j] :=
              auxNovaPopulacaoMatrizPOA[k, j];
          end; // for j
        end; // for k
      end;
    end;

  end; // with

end;

procedure TIWForm1.AnaliseParaElicitacao;
var
  i, j, k, y, w: Integer;
  auxaltcc, auxaltaa: array of Double;
  auxtotal, aux: Double;
begin
  With TIWUserSession(WebApplication.Data) do
  begin

    Naltsel := 0;
    for i := 0 to length(ResultAlt) - 1 do
    begin
      if ResultAlt[i] <> 0 then
        inc(Naltsel);
    end;

    AtualizarPopulacao;
    SetLength(auxaltcc, Numalt);
    SetLength(auxaltaa, Numalt);
    setlength(combcritaltx, numcrit, numcrit);
    setlength(combcritaltz, numcrit, numcrit);
    setlength(combcritalto, numcrit, numcrit);
    setlength(combaltaltx, Numalt, Numalt);
    setlength(combaltaltz, Numalt, Numalt);
    setlength(combaltalto, Numalt, Numalt);

    // Comparação Critério x Posição do ranking
    for i := 0 to Numcrit - 1 do
      for j := 0 to Numcrit - 1 do
      begin
        if i = j then
         continue;

        auxtotal := 0;
        for k := 0 to length(novapopulacaoOrdemCrit) - 1 do
        begin
          if TIWUserSession(WebApplication.Data).novapopulacaoOrdemCrit[k, i] < TIWUserSession(WebApplication.Data).novapopulacaoOrdemCrit[k, j] then
          begin
            for y := 0 to Numalt - 1 do
            begin
              if NovaPopulacaoMatrizPOA[k, y] = 1 then
                auxaltcc[y] := auxaltcc[y] + 1;
            end;
            //for y := 0 to Numalt - 1 do
            //begin
              if NovaPopulacaoMatrizPOA[k, altx] = 1 then
                combcritaltx[i,j] := combcritaltx[i,j] + 1
              else if NovaPopulacaoMatrizPOA[k, altz] = 1 then
                combcritaltz[i,j] := combcritaltz[i,j] + 1
              else
                combcritalto[i,j] := combcritalto[i,j] + 1;
            //end;

            auxtotal := auxtotal + 1;
          end;
          //
        end;

        aux := 0;
        if auxtotal > 0 then
        begin
        TIWUserSession(WebApplication.Data).combcritaltx[i,j] := combcritaltx[i,j] / auxtotal;
        TIWUserSession(WebApplication.Data).combcritaltz[i,j] := combcritaltz[i,j] / auxtotal;
        TIWUserSession(WebApplication.Data).combcritalto[i,j] := combcritalto[i,j] / auxtotal;

          for y := 0 to Numalt - 1 do
          begin
            auxaltcc[y] := auxaltcc[y] / (auxtotal);

            if auxaltcc[y] > aux then
            begin
              // w := y;
              aux := auxaltcc[y];
            end;
          end;
        end
        else
          auxaltcc[y] := 0;

        TIWUserSession(WebApplication.Data).combinacaocrit[i, j] := aux;

        for y := 0 to numcrit - 1 do
        auxaltcc[y] := 0;

      end;
    // Fim da comparação Critério x Critério

    // Comparação Alternativa x Alternativa
    for i := 0 to Numalt - 1 do
      for j := 0 to Numalt - 1 do
      begin
        if i = j then
          continue;

        auxtotal := 0;
        for k := 0 to length(NovaPopulacaoResultadoRoc) - 1 do
        begin
          if TIWUserSession(WebApplication.Data).NovaPopulacaoResultadoRoc[k, i] > NovaPopulacaoResultadoRoc[k, j]
            then
          begin
            for y := 0 to Numalt - 1 do
            begin
              if NovaPopulacaoMatrizPOA[k, y] = 1 then
                auxaltaa[y] := auxaltaa[y] + 1;
            end;

             if NovaPopulacaoMatrizPOA[k, altx] = 1 then
                combaltaltx[i,j] := combaltaltx[i,j] + 1
              else if NovaPopulacaoMatrizPOA[k, altz] = 1 then
                combaltaltz[i,j] := combaltaltz[i,j] + 1
              else
                combaltalto[i,j] := combaltalto[i,j] + 1;


            auxtotal := auxtotal + 1;
          end;
          //
        end;

        // w := -1;
        aux := 0;
        if auxtotal > 0 then
        begin
          TIWUserSession(WebApplication.Data).combaltaltx[i,j] := combaltaltx[i,j] / auxtotal;
          TIWUserSession(WebApplication.Data).combaltaltz[i,j] := combaltaltz[i,j] / auxtotal;
          TIWUserSession(WebApplication.Data).combaltalto[i,j] := combaltalto[i,j] / auxtotal;

          for y := 0 to Numalt - 1 do
          begin
            auxaltaa[y] := auxaltaa[y] / (auxtotal);

            if auxaltaa[y] > aux then
            begin
              // w := y;
              aux := auxaltaa[y];
            end;
          end
        end
        else
          auxaltaa[y] := 0;

        TIWUserSession(WebApplication.Data).combinacaoalt[i, j] := aux;

      for y := 0 to Numalt - 1 do
      auxaltaa[y] := 0;

      end;



    // Fim da comparação Alternativa x Alternativa

    {SelecionarPergunta;
  end;
  TIWUserSession(WebApplication.Data).problemafinalizado := True;
  if TIWUserSession(WebApplication.Data).VetorPerguntaElicitacao
    [TIWUserSession(WebApplication.Data).nperg - 1, 3] < 1 then
    TIWUserSession(WebApplication.Data).problemafinalizado := false;}



  // Plota gráfico da comparação entre os critérios,
  // os resultados são em termos da probabilidade de vitoria das alternativas X e Z
  grdcomparacaocrit.RowCount := numcrit +2;
  grdcomparacaocrit.ColumnCount := numcrit +1;


  for j := 0 to grdcomparacaocrit.ColumnCount - 1 do
      for i := 0 to grdcomparacaocrit.rowcount - 1 do
      begin
        grdcomparacaocrit.Cell[i, j].Alignment := taCenter;
        grdcomparacaocrit.Cell[i, j].wrap := True;
        // grdresultados.Cell[0,j].Font.Style := [fsbold];
      end;
    for j := 0 to grdcomparacaocrit.ColumnCount - 1 do
      grdcomparacaocrit.Cell[0, j].Font.Style := [fsbold];

    for i := 0 to grdcomparacaocrit.rowcount - 1 do
      grdcomparacaocrit.Cell[i, 0].Font.Style := [fsbold];

  for i := 0 to numcrit - 1 do
  begin
    grdcomparacaocrit.Cell[i+2,0].Text := NomeCrit[i];
    grdcomparacaocrit.Cell[0,i+1].Text := NomeCrit[i];
    grdcomparacaocrit.Cell[1,i+1].Text := 'X|Z|Other';

    for j := 0 to numcrit - 1 do
    begin
      if i = j then continue;

      grdcomparacaocrit.Cell[2+i,j+1].Text := FormatFloat('0',100*combcritaltx[i,j]) + '%|'+
                                          FormatFloat('0',100*combcritaltz[i,j]) + '%|'+
                                       FormatFloat('0',100*combcritalto[i,j]) + '%';
    end; // for j
  end; // for i


  grdcritxcrit.RowCount := numcrit +1;
  grdcritxcrit.ColumnCount := numcrit +1;


  for j := 0 to grdcritxcrit.ColumnCount - 1 do
      for i := 0 to grdcritxcrit.rowcount - 1 do
      begin
        grdcritxcrit.Cell[i, j].Alignment := taCenter;
        grdcritxcrit.Cell[i, j].wrap := True;
        // grdresultados.Cell[0,j].Font.Style := [fsbold];
      end;
    for j := 0 to grdcritxcrit.ColumnCount - 1 do
      grdcritxcrit.Cell[0, j].Font.Style := [fsbold];

    for i := 0 to grdcritxcrit.rowcount - 1 do
      grdcritxcrit.Cell[i, 0].Font.Style := [fsbold];

  for i := 0 to numcrit - 1 do
  begin
    grdcritxcrit.Cell[i+1,0].Text := NomeCrit[i];
    grdcritxcrit.Cell[0,i+1].Text := NomeCrit[i];

    for j := 0 to numcrit - 1 do
    begin
      if i = j then continue;

      grdcritxcrit.Cell[1+i,j+1].Text := FormatFloat('0',100*combinacaocrit[i,j]) + '%';
    end;
  end; // for i

  // Plota gráfico da comparação entre as alternativas,
  // os resultados são em termos da probabilidade de vitoria das alternativas X e Z
  grdcomparacaoalt.RowCount := naltsel +2;
  grdcomparacaoalt.ColumnCount := naltsel +1;


  for j := 0 to grdcomparacaoalt.ColumnCount - 1 do
      for i := 0 to grdcomparacaoalt.rowcount - 1 do
      begin
        grdcomparacaoalt.Cell[i, j].Alignment := taCenter;
        grdcomparacaoalt.Cell[i, j].wrap := True;
        // grdresultados.Cell[0,j].Font.Style := [fsbold];
      end;
    for j := 0 to grdcomparacaoalt.ColumnCount - 1 do
      grdcomparacaoalt.Cell[0, j].Font.Style := [fsbold];

    for i := 0 to grdcomparacaoalt.rowcount - 1 do
      grdcomparacaoalt.Cell[i, 0].Font.Style := [fsbold];

  for i := 0 to naltsel - 1 do
  begin
    grdcomparacaoalt.Cell[i+2,0].Text := NomeAlt[ordemalt[i]];
    grdcomparacaoalt.Cell[0,i+1].Text := NomeAlt[ordemalt[i]];
    grdcomparacaoalt.Cell[1,i+1].Text := 'X|Z|Other';

    for j := 0 to naltsel - 1 do
    begin
      if i = j then continue;

      grdcomparacaoalt.Cell[2+i,j+1].Text := FormatFloat('0',100*combaltaltx[ordemalt[i],ordemalt[j]]) + '%|'+
                                          FormatFloat('0',100*combaltaltz[ordemalt[i],ordemalt[j]]) + '%|'+
                                       FormatFloat('0',100*combaltalto[ordemalt[i],ordemalt[j]]) + '%';
    end;  // for j
  end;  // for i

  grdaltxalt.RowCount := naltsel +1;
  grdaltxalt.ColumnCount := naltsel +1;


  for j := 0 to grdaltxalt.ColumnCount - 1 do
      for i := 0 to grdaltxalt.rowcount - 1 do
      begin
        grdaltxalt.Cell[i, j].Alignment := taCenter;
        grdaltxalt.Cell[i, j].wrap := True;
        // grdresultados.Cell[0,j].Font.Style := [fsbold];
      end;
    for j := 0 to grdaltxalt.ColumnCount - 1 do
      grdaltxalt.Cell[0, j].Font.Style := [fsbold];

    for i := 0 to grdaltxalt.rowcount - 1 do
      grdaltxalt.Cell[i, 0].Font.Style := [fsbold];

  for i := 0 to naltsel - 1 do
  begin
    grdaltxalt.Cell[i+1,0].Text := Nomealt[ordemalt[i]];
    grdaltxalt.Cell[0,i+1].Text := Nomealt[ordemalt[i]];

    for j := 0 to naltsel - 1 do
    begin
      if i = j then continue;

      grdaltxalt.Cell[1+i,j+1].Text := FormatFloat('0',100*combinacaoalt[ordemalt[i],ordemalt[j]]) + '%';
    end;
  end; // for i



  end;  // with


end;    // procedure

procedure TIWForm1.SelecionarPergunta;
var
  i, j, k, y: Integer;
  parcomparado: array [0 .. 2] of Integer; // se parcomparado[0] = 0 -> crit x crit, se parcomparado[0] = 1 -> alt x alt
  aux, aux2: Double;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    SetLength(VetorPerguntaElicitacao, nperg + 1, 4);
    SetLength(VetorRespostaElicitacao, nperg + 1, 3);
    aux := 0;


    for i := 0 to numcrit - 1 do
    for j := 0 to numcrit - 1 do
    begin
      if pularcomparacao[i,j] = true then
      continue;

      if aux < combinacaocrit[i,j] then
      begin
        aux := combinacaocrit[i,j];
        parcomparado[0] := 0;
        parcomparado[1] := i;
        parcomparado[2] := j;
      end;
    end;


    // aux2 := 0;
    // Procurando entre os pares de critérios, aquele que possui a maior probabilidade
    {for i := 0 to Numcrit - 1 do
      for j := i + 1 to Numcrit - 1 do
      begin
        if i = j then
          continue;

        if (combinacaocrit[i, j] <= combinacaocrit[j, i]) and
          (combinacaocrit[i, j] > aux) then
        begin
          aux := combinacaocrit[i, j];
          parcomparado[0] := 0;
          parcomparado[1] := i;
          parcomparado[2] := j;
          continue;
        end;

        if (TIWUserSession(WebApplication.Data).combinacaocrit[i, j]
            > combinacaocrit[j, i]) and (combinacaocrit[j, i] > aux) then
        begin
          aux := combinacaocrit[j, i];
          parcomparado[0] := 0;
          parcomparado[1] := i;
          parcomparado[2] := j;
        end;
      end;

    // Procurando entre os pares de alternativas, aquele que possui a maior probabilidade
    {for i := 0 to Numalt - 2 do
      for j := i + 1 to Numalt - 1 do
      begin
        if i = j then
          continue;

        if (combinacaoalt[i, j] <= combinacaoalt[j, i]) and
          (combinacaoalt[i, j] > aux) then
        begin
          aux := combinacaoalt[i, j];
          parcomparado[0] := 1;
          parcomparado[1] := i;
          parcomparado[2] := j;
          continue;
        end;

        if (combinacaoalt[i, j] > TIWUserSession(WebApplication.Data)
            .combinacaoalt[j, i]) and (combinacaoalt[j, i] > aux) then
        begin
          aux := combinacaoalt[j, i];
          parcomparado[0] := 1;
          parcomparado[1] := i;
          parcomparado[2] := j;
        end;
      end;  }
    for i := 0 to 2 do
      TIWUserSession(WebApplication.Data).VetorPerguntaElicitacao[nperg, i] :=
        parcomparado[i];

    pularcomparacao[parcomparado[1],parcomparado[2]] := true;

    VetorPerguntaElicitacao[nperg, 3] := aux;
    for i := 0 to 2 do
      TIWUserSession(WebApplication.Data).VetorRespostaElicitacao[nperg, i] :=
        Round(TIWUserSession(WebApplication.Data).VetorPerguntaElicitacao
          [nperg, i]);

    inc(nperg);
    AtualizarPopulacao;
  end;

end;

procedure TIWForm1.btnauxchooseClick(Sender: TObject);
var
  i, j, jj, k, qtdpareto, S, r: Integer;
  auxregra, aux1, aux2, aux3: real;
  bar, barpoa: array of TBarSeries;
  SolRegra1, SolRegra2, SolRegra3: Integer;
  epsilon, omega: array of real;
  totalacumulado: real;
  recomendacao: boolean;
  Px: real;
  faixaprob: array of real;
  pular: array of boolean;
  pular2: boolean;
  auxresultado: real;
  Linha: array of TBarSeries;

begin
  With TIWUserSession(WebApplication.Data) do
  begin
    LimpaVariaveis;
    with UserSession.ZQuery1 do
    begin
      if lstbxProjects.ItemIndex = -1 then
      begin
        WebApplication.showmessage('Select a problem to continue!');
        exit;
      end;

      TIWUserSession(WebApplication.Data).NameProblem := TIWUserSession
        (WebApplication.Data).problems[lstbxProjects.ItemIndex];
      TIWUserSession(WebApplication.Data).CurrentID_problem := StrToInt
        (TIWUserSession(WebApplication.Data).id_problem[lstbxProjects.ItemIndex]
        );
      SQL.Text :=
        'SELECT ID_alternativa FROM alternativa WHERE ID_problema=' + inttostr
        (TIWUserSession(WebApplication.Data).CurrentID_problem);
      ExecSQL;
      Open;
      First;
      TIWUserSession(WebApplication.Data).Numalt := RecordCount;
      SetLength(TIWUserSession(WebApplication.Data).altcode, TIWUserSession
          (WebApplication.Data).Numalt);
      for S := 1 to TIWUserSession(WebApplication.Data).Numalt do
      begin
        TIWUserSession(WebApplication.Data).altcode[S - 1] := FieldByName
          ('ID_alternativa').AsInteger;
        Next;
      end;
      Close;
      SQL.Text :=
        'SELECT ID_criterio FROM criterio WHERE ID_problema=' + inttostr
        (TIWUserSession(WebApplication.Data).CurrentID_problem);
      ExecSQL;
      Open;
      First;
      TIWUserSession(WebApplication.Data).Numcrit := RecordCount;
      SetLength(TIWUserSession(WebApplication.Data).critcode, TIWUserSession
          (WebApplication.Data).Numcrit);
      for r := 1 to TIWUserSession(WebApplication.Data).Numcrit do
      begin
        TIWUserSession(WebApplication.Data).critcode[r - 1] := FieldByName
          ('ID_criterio').AsInteger;
        Next;
      end;
      Close;
    end;
    // >>>>
    // Gravação de dados dos pesos, limiares e tipos dos critérios

    SetLength(TIWUserSession(WebApplication.Data).Nomecrit, TIWUserSession
        (WebApplication.Data).Numcrit);
    SetLength(TIWUserSession(WebApplication.Data).CritSigla, TIWUserSession
        (WebApplication.Data).Numcrit);
    SetLength(TIWUserSession(WebApplication.Data).tipocrit, TIWUserSession
        (WebApplication.Data).Numcrit);
    SetLength(TIWUserSession(WebApplication.Data).niveis, TIWUserSession
        (WebApplication.Data).Numcrit);

    with UserSession.ZQuery1 do
    begin
      for i := 1 to TIWUserSession(WebApplication.Data).Numcrit do
      begin
        SQL.Text := 'SELECT * FROM criterio WHERE ID_criterio=' + inttostr
          (TIWUserSession(WebApplication.Data).critcode[i - 1]);
        ExecSQL;
        Open;
        TIWUserSession(WebApplication.Data).CritSigla[i - 1] :=
          'Crit' + inttostr(i);
        TIWUserSession(WebApplication.Data).Nomecrit[i - 1] := FieldByName
          ('nome_criterio').AsString;

        if FieldByName('tipo_criterio').AsString <> '' then
          TIWUserSession(WebApplication.Data).tipocrit[i - 1] := StrToIntDef
            (FieldByName('tipo_criterio').AsString, 2)
        else
          TIWUserSession(WebApplication.Data).tipocrit[i - 1] := -1;

        if FieldByName('niveis').AsString <> '' then
          TIWUserSession(WebApplication.Data).niveis[i - 1] := StrToIntDef
            (FieldByName('niveis').AsString, 2)
        else
          TIWUserSession(WebApplication.Data).niveis[i - 1] := -1;

        Close;
      end;
      SetLength(TIWUserSession(WebApplication.Data).nomealt, TIWUserSession
          (WebApplication.Data).Numalt);

      for j := 1 to TIWUserSession(WebApplication.Data).Numalt do
      begin
        SQL.Text := 'SELECT * FROM alternativa WHERE ID_alternativa=' + inttostr
          (TIWUserSession(WebApplication.Data).altcode[j - 1]);
        ExecSQL;
        Open;
        TIWUserSession(WebApplication.Data).nomealt[j - 1] := FieldByName
          ('nome_alternativa').AsString;
        Close;
      end;

    end;

    // Gravação de dados da matriz consequência
    SetLength(TIWUserSession(WebApplication.Data).MatrizConseq, TIWUserSession
        (WebApplication.Data).Numalt, TIWUserSession(WebApplication.Data)
        .Numcrit);

    with UserSession.ZQuery1 do
    begin
      SQL.Text :=
        'SELECT valor_performance FROM matrizconsequencia WHERE ID_problema=' +
        inttostr(TIWUserSession(WebApplication.Data).CurrentID_problem);
      ExecSQL;
      Open;
      First;
      for i := 1 to TIWUserSession(WebApplication.Data).Numcrit do
        for j := 1 to TIWUserSession(WebApplication.Data).Numalt do
        begin
          TIWUserSession(WebApplication.Data).MatrizConseq[j - 1, i - 1] :=
            StrToFloat(FieldByName('valor_performance').AsString);
          Next;
        end;
      Close;
    end;
    GerarCases;
    EscalaRazao;
    Surrogate;
    TIWUserSession(WebApplication.Data).problemafinalizado := false;
    SetLength(TIWUserSession(WebApplication.Data).combinacaocrit, Numcrit,
      Numcrit);
    SetLength(TIWUserSession(WebApplication.Data).combinacaoalt, Numalt,
      Numalt);
    setlength(pularcomparacao, numcrit, numcrit);
    for i := 0 to numcrit - 1 do
    for j := 0 to numcrit - 1 do
    pularcomparacao[i,j] := false;

    altx := Ordemalt[0];
    altz := Ordemalt[1];
   // while TIWUserSession(WebApplication.Data).problemafinalizado = false do
    //begin
      AnaliseParaElicitacao;
    //end; }
    SetarCores;
    ResultadosPromethee;
    ComparacaoMetodos;
    chrtResultado.RemoveAllSeries;
    chrtResultado.Legend.Items.Clear;
    // chrtresultado.ClearChart;
    chrtResultado.BottomAxis.Items.Clear;
    chrtResultado.View3D := false;
    SetLength(bar, Numalt);
    i := 0;
    Total := length(CasesOrdemcrit);
    totalacumulado := 0;
    for j := 0 to Numalt - 1 do
    begin
      if totalacumulado > 0.8 * Total then
      begin
        bar[i] := TBarSeries.Create(self);
        chrtResultado.AddSeries(bar[i]);
        bar[i].AddXY(0, Total - totalacumulado, formatfloat
            ('0', Total - totalacumulado), coresalt[i]);
        // chrtResultado.BottomAxis.Items.Add(i, 'S' + inttostr(i + 1));

        chrtResultado.Legend.Item[i].Text := 'Other Alternatives';
        bar[i].Color := coresalt[i];
        break;
      end;
      if ResultAlt[Ordemalt[j]] <> 0 then
      begin
        bar[i] := TBarSeries.Create(self);
        chrtResultado.AddSeries(bar[i]);
        bar[i].AddXY(0, ResultAlt[Ordemalt[j]], formatfloat
            ('0', ResultAlt[Ordemalt[j]]), coresalt[i]);
        // chrtResultado.BottomAxis.Items.Add(i, 'S' + inttostr(i + 1));
        chrtResultado.Legend.Item[i].Text := nomealt[Ordemalt[j]];
        bar[i].Color := coresalt[i];
        totalacumulado := totalacumulado + ResultAlt[Ordemalt[j]];
        inc(i);
      end;
    end;
    Total := 0;
    bar[1] := TBarSeries.Create(self);
    chrtResultado.AddSeries(bar[1]);
    for i := 0 to length(ResultSol) - 1 do
      Total := Total + ResultSol[i];
    bar[1].AddXY(0, Total, '', clWhite);
    bar[1].Visible := false;
    // chrtResultado.LeftAxis.Increment := 0.2; // ?
    // chrtResultado.Legend.Visible := false;
    chrtResultado.BottomAxis.LabelsAlternate;
    chrtResultado.BottomAxis.Items.Automatic := false;
    chrtResultado.LeftAxis.Maximum := Total;
    chrtResultado.LeftAxis.AutomaticMaximum := false;
    /// chrtResultado.LeftAxis.MaximumOffset := 16;
    chrtResultado.LeftAxis.Minimum := 0;
    // chrtresultado.leftaxis.Items.Automatic := false;
    chrtResultado.LeftAxis.Increment := Total;

    imgResultado.Picture.Bitmap.Assign
      (chrtResultado.TeeCreateBitmap(clWebPURPLE, rect(0, 0,
          imgResultado.Width, imgResultado.Height)));

    grdresultados.Clear;
    Naltsel := 0;
    for i := 0 to length(ResultAlt) - 1 do
    begin
      if ResultAlt[i] <> 0 then
        inc(Naltsel);
    end;
    grdresultados.rowcount := 2 + Naltsel;
    grdresultados.ColumnCount := 3;
    { if grdresultados.ColumnCount * 80 > 832 then
      grdresultados.Width := grdresultados.ColumnCount * 80
      else
      grdresultados.Width := 832; }
    grdresultados.Cell[0, 0].Text := 'Alternative';
    grdresultados.Cell[0, 1].Text := 'Frequency';
    grdresultados.Cell[0, 2].Text := '%';
    i := 0;
    for j := 0 to Numalt - 1 do
    begin
      if ResultAlt[Ordemalt[j]] <> 0 then
      begin
        grdresultados.Cell[i + 1, 0].Text := nomealt[Ordemalt[j]];
        grdresultados.Cell[i + 1, 1].Text := FloatToStr(ResultAlt[Ordemalt[j]]);
        grdresultados.Cell[i + 1, 2].Text := formatfloat
          ('0', 100 * ResultAlt[Ordemalt[j]] / Total) + ' %';
        inc(i);
      end;
    end;
    // grdresultados.Cell[0, Numalt+1].Text := 'Cases';
    // grdresultados.Cell[0, Naltsel + 1].Text := 'Total';
    { for i := 0 to length(ResultSol) - 1 do
      grdresultados.Cell[i + 1, 0].Text := 'Sol ' + inttostr(i + 1) + ':'; }
    grdresultados.Cell[1 + i, 0].Text := 'Total:';
    grdresultados.Cell[1 + i, 1].Text := FloatToStr(Total);
    grdresultados.Cell[1 + i, 2].Text := '100%';
    { for i := 0 to length(ResultSol) - 1 do
      begin
      k := 0;
      for j := 0 to Numalt - 1 do
      begin
      if ResultAlt[j] <> 0 then
      begin
      if MatrizSol[OrdemSol][i], j] = 1 then
      grdresultados.Cell[i + 1, k + 1].Text := 'X'; // 'μ|Max|σ';
      inc(k);
      end;
      end;
      end; }

    {
      for i := 0 to length(ResultSol) - 1 do
      grdresultados.Cell[i + 1, Naltsel + 1].Text := floattostr
      (ResultSol[OrdemSol[i]]);
      }
    for j := 0 to grdresultados.ColumnCount - 1 do
      for i := 0 to grdresultados.rowcount - 1 do
      begin
        grdresultados.Cell[i, j].Alignment := taCenter;
        grdresultados.Cell[i, j].wrap := True;
        // grdresultados.Cell[0,j].Font.Style := [fsbold];
      end;

    for j := 0 to grdresultados.ColumnCount - 1 do
    begin
      grdresultados.Cell[0, j].Font.Style := [fsbold];
      grdresultados.Cell[1 + Naltsel, j].Font.Color := clwebgray;
      grdresultados.Cell[1 + Naltsel, j].BGColor := clWebWhiteSmoke;
    end;
    grdresultados.Cell[1 + Naltsel, 0].Font.Style := [fsbold];

    { for I := 0 to length(ResultSol) - 1 do
      begin
      grdresultados.Cell[i+1, Numalt+1].Font.Size := 8;
      for j := 0 to Round(Total) - 1 do
      begin
      if Casesol[j] = i+1 then
      begin
      if grdresultados.Cell[i+1, Numalt+1].Text = '' then
      grdresultados.Cell[i+1, Numalt+1].Text := inttostr(j+1)
      else
      grdresultados.Cell[i+1, Numalt+1].Text := grdresultados.Cell[i+1, Numalt+1].Text + ', ' + inttostr(j+1);
      end;
      end;
      end; }

    grdcases.rowcount := Round(Total) + 1;
    grdcases.ColumnCount := 2 + Numcrit;

    for j := 0 to grdcases.ColumnCount - 1 do
      for i := 0 to grdcases.rowcount - 1 do
      begin
        grdcases.Cell[i, j].Alignment := taCenter;
        grdcases.Cell[i, j].wrap := True;
        // grdresultados.Cell[0,j].Font.Style := [fsbold];
      end;
    for j := 0 to grdcases.ColumnCount - 1 do
      grdcases.Cell[0, j].Font.Style := [fsbold];

    grdcases.Cell[0, 0].Text := 'Surrogate';
    for i := 0 to Numcrit - 1 do
      grdcases.Cell[0, i + 1].Text := Nomecrit[i];

    grdcases.Cell[0, Numcrit + 1].Text := 'Solution';

    for j := 0 to Round(Total) - 1 do
    begin
      grdcases.Cell[j + 1, 0].Text := 'S' + inttostr(j + 1);
      for i := 0 to Numalt - 1 do
      begin
        if MatrizPOA[j, i] = 1 then
        begin
          grdcases.Cell[j + 1, Numcrit + 1].Text := nomealt[i] + ' ';
        end;
      end;
      for i := 0 to Numcrit - 1 do
        grdcases.Cell[j + 1, i + 1].Text := inttostr(CasesOrdemcrit[j, i]) + 'º'
    end;

    rgnLoginproject.Visible := false;
    rgnresultado.Visible := True;
    // btnshowresults.Enabled := true;

    Total := 0;
    for i := 0 to length(ResultSol) - 1 do
      Total := Total + ResultSol[i];

    grdestatistica.Clear;
    grdestatistica.rowcount := 3 + length(MediaDifSol);
    grdestatistica.ColumnCount := 4 + length(MediaDifSol);

    for i := 0 to length(MediaDifSol) - 1 do
    begin
      grdestatistica.Cell[0, i + 4].Alignment := taCenter;
      // grdestatistica.Cell[0, i+1].wrap := True;
      grdestatistica.Cell[0, i + 4].Font.Style := [fsbold];
      grdestatistica.Cell[1, i + 4].Alignment := taCenter;
      // grdestatistica.Cell[1, i+1].wrap := True;
      grdestatistica.Cell[1, i + 4].Font.Style := [fsbold];
      // grdestatistica.Cell[0, i + 4].Text := 'Sol ' + inttostr(i + 1);

      grdestatistica.Cell[0, i + 4].Font.Color := clwebgray;
      grdestatistica.Cell[1, i + 4].Text := ' μ | σ ';

      grdestatistica.Cell[i + 2, 3].Alignment := taCenter;
      grdestatistica.Cell[i + 2, 3].wrap := True;
      grdestatistica.Cell[i + 2, 3].Font.Style := [fsbold];
      grdestatistica.Cell[i + 2, 3].Font.Color := clwebgray;
      for j := 0 to Numalt - 1 do
      begin
        if MatrizSol[OrdemSol[i], j] = 1 then
        begin
          if grdestatistica.Cell[i + 2, 3].Text = '' then
            grdestatistica.Cell[i + 2, 3].Text := nomealt[j]
          else
            grdestatistica.Cell[i + 2, 3].Text := grdestatistica.Cell[i + 2, 3]
              .Text + ' + ' + nomealt[j];
        end;
      end;
      grdestatistica.Cell[0, 4 + i].Text := grdestatistica.Cell[i + 2, 3].Text;
      // grdestatistica.Cell[i + 2, 3].Text := 'Sol ' + inttostr(i + 1);

      grdestatistica.Cell[i + 2, 1].Alignment := taCenter;
      grdestatistica.Cell[i + 2, 1].wrap := True;
      // grdestatistica.Cell[i + 2, 1].Font.Style := [fsbold];
      grdestatistica.Cell[i + 2, 1].Font.Color := clWebBlue;
      grdestatistica.Cell[i + 2, 1].Text := formatfloat
        ('0.00', MediaGeral[OrdemSol[i]]);

      grdestatistica.Cell[i + 2, 2].Alignment := taCenter;
      grdestatistica.Cell[i + 2, 2].wrap := True;
      // grdestatistica.Cell[i + 2, 1].Font.Style := [fsbold];
      grdestatistica.Cell[i + 2, 2].Font.Color := clWebBlue;
      grdestatistica.Cell[i + 2, 2].Text := formatfloat
        ('0.00', DesvioPadraoGeral[OrdemSol[i]]);

      grdestatistica.Cell[i + 2, 0].Alignment := taCenter;
      grdestatistica.Cell[i + 2, 0].wrap := True;
      // grdestatistica.Cell[i + 2, 1].Font.Style := [fsbold];
      grdestatistica.Cell[i + 2, 0].Font.Color := clWebBlue;
      grdestatistica.Cell[i + 2, 0].Text := formatfloat
        ('0.00', (ResultSol[OrdemSol[i]] / Total) * 100) + '%';
    end;

    grdestatistica.Cell[1, 1].Alignment := taCenter;
    grdestatistica.Cell[1, 1].wrap := True;
    grdestatistica.Cell[1, 1].Font.Style := [fsbold];
    grdestatistica.Cell[1, 1].Font.Color := clWebBlue;
    grdestatistica.Cell[1, 1].Text := 'μ';

    grdestatistica.Cell[1, 2].Alignment := taCenter;
    grdestatistica.Cell[1, 2].wrap := True;
    grdestatistica.Cell[1, 2].Font.Style := [fsbold];
    grdestatistica.Cell[1, 2].Font.Color := clWebBlue;
    grdestatistica.Cell[1, 2].Text := 'σ';

    grdestatistica.Cell[1, 0].Alignment := taCenter;
    grdestatistica.Cell[1, 0].wrap := True;
    grdestatistica.Cell[1, 0].Font.Style := [fsbold];
    grdestatistica.Cell[1, 0].Font.Color := clWebBlue;
    grdestatistica.Cell[1, 0].Text := '%';

    grdestatistica.Cell[2 + length(MediaDifSol), 3].Alignment := taCenter;
    grdestatistica.Cell[2 + length(MediaDifSol), 3].wrap := True;
    grdestatistica.Cell[2 + length(MediaDifSol), 3].Font.Style := [fsbold];
    grdestatistica.Cell[2 + length(MediaDifSol), 3].Font.Color := clWebBlue;
    grdestatistica.Cell[2 + length(MediaDifSol), 3].Text := 'μ:';

    { grdestatistica.Cell[3+length(MediaDifSol), 3].Alignment := taCenter;
      grdestatistica.Cell[3+length(MediaDifSol), 3].wrap := True;
      grdestatistica.Cell[3+length(MediaDifSol), 3].Font.Style := [fsbold];
      grdestatistica.Cell[3+length(MediaDifSol), 3].Font.Color := clWebBlue;
      grdestatistica.Cell[3+length(MediaDifSol), 3].Text := 'σ:'; }

    for i := 0 to length(MediaDifSol) - 1 do
    begin
      grdestatistica.Cell[2 + length(MediaDifSol), 4 + i].Text := formatfloat
        ('0.00', MediaGeralNaosol[OrdemSol[i]]);
      grdestatistica.Cell[2 + length(MediaDifSol), 4 + i].Alignment := taCenter;
      grdestatistica.Cell[2 + length(MediaDifSol), 4 + i].wrap := True;
      grdestatistica.Cell[2 + length(MediaDifSol), 4 + i].Font.Color :=
        clWebBlue;

      { grdestatistica.Cell[3+length(MediaDifSol), 4+i].Text := formatfloat('0.00', DesvioPadraoGeralNaoSol[OrdemSol[i]]);
        grdestatistica.Cell[3+length(MediaDifSol), 4+i].Alignment := taCenter;
        grdestatistica.Cell[3+length(MediaDifSol), 4+i].wrap := True;
        grdestatistica.Cell[3+length(MediaDifSol), 4+i].Font.Color := clWebBlue; }
    end;

    for i := 0 to length(MediaDifSol) - 1 do
      for j := 0 to length(MediaDifSol) - 1 do
      begin
        grdestatistica.Cell[i + 2, j + 4].Alignment := taCenter;
        grdestatistica.Cell[i + 2, j + 4].wrap := True;
        if i = j then
          grdestatistica.Cell[i + 2, j + 4].Text := ''
        else
        begin
          grdestatistica.Cell[i + 2, j + 4].Text := formatfloat
            ('0.00', MediaDifSol[OrdemSol[i], OrdemSol[j]]) + '|' + formatfloat
            ('0.00', DesvioPadraoDifSol[OrdemSol[i], OrdemSol[j]]);
        end;
      end;

    grdmaxmin.Clear;
    grdmaxmin.rowcount := 7 + length(MediaDifSol);
    grdmaxmin.ColumnCount := 5 + length(MediaDifSol);

    for i := 0 to length(MediaDifSol) - 1 do
    begin
      grdmaxmin.Cell[0, i + 5].Alignment := taCenter;
      // grdmaxmin.Cell[0, i+1].wrap := True;
      grdmaxmin.Cell[0, i + 5].Font.Style := [fsbold];
      grdmaxmin.Cell[1, i + 5].Alignment := taCenter;
      // grdmaxmin.Cell[1, i+1].wrap := True;
      grdmaxmin.Cell[1, i + 5].Font.Style := [fsbold];
      // grdmaxmin.Cell[0, i + 5].Text := 'Sol ' + inttostr(i + 1);
      grdmaxmin.Cell[0, i + 5].Font.Color := clwebgray;
      grdmaxmin.Cell[1, i + 5].Text := '  Max | Min ';

      grdmaxmin.Cell[i + 2, 4].Alignment := taCenter;
      grdmaxmin.Cell[i + 2, 4].wrap := True;
      grdmaxmin.Cell[i + 2, 4].Font.Style := [fsbold];
      grdmaxmin.Cell[i + 2, 4].Font.Color := clwebgray;

      for j := 0 to Numalt - 1 do
      begin
        if MatrizSol[OrdemSol[i], j] = 1 then
        begin
          if grdmaxmin.Cell[i + 2, 4].Text = '' then
            grdmaxmin.Cell[i + 2, 4].Text := nomealt[j]
          else
            grdmaxmin.Cell[i + 2, 4].Text := grdmaxmin.Cell[i + 2, 4]
              .Text + ' + ' + nomealt[j];
        end;
      end;
      grdmaxmin.Cell[0, 5 + i].Text := grdmaxmin.Cell[i + 2, 4].Text;

      // grdmaxmin.Cell[i + 2, 4].Text := 'Sol ' + inttostr(i + 1);

      grdmaxmin.Cell[i + 2, 1].Alignment := taCenter;
      grdmaxmin.Cell[i + 2, 1].wrap := True;
      // grdestatistica.Cell[i + 2, 1].Font.Style := [fsbold];
      grdmaxmin.Cell[i + 2, 1].Font.Color := clWebBlue;
      grdmaxmin.Cell[i + 2, 1].Text := formatfloat
        ('0.00', MinimoGeral[OrdemSol[i]]);

      grdmaxmin.Cell[i + 2, 2].Alignment := taCenter;
      grdmaxmin.Cell[i + 2, 2].wrap := True;
      // grdestatistica.Cell[i + 2, 1].Font.Style := [fsbold];
      grdmaxmin.Cell[i + 2, 2].Font.Color := clWebBlue;
      grdmaxmin.Cell[i + 2, 2].Text := formatfloat('0.00', MaxMin[OrdemSol[i]]);

      grdmaxmin.Cell[i + 2, 3].Alignment := taCenter;
      grdmaxmin.Cell[i + 2, 3].wrap := True;
      // grdestatistica.Cell[i + 2, 1].Font.Style := [fsbold];
      grdmaxmin.Cell[i + 2, 3].Font.Color := clWebBlue;
      grdmaxmin.Cell[i + 2, 3].Text := formatfloat('0.00', MinMax[OrdemSol[i]]);

      grdmaxmin.Cell[i + 2, 0].Alignment := taCenter;
      grdmaxmin.Cell[i + 2, 0].wrap := True;
      // grdestatistica.Cell[i + 2, 1].Font.Style := [fsbold];
      grdmaxmin.Cell[i + 2, 0].Font.Color := clWebBlue;
      grdmaxmin.Cell[i + 2, 0].Text := formatfloat
        ('0.00', MaximoGeral[OrdemSol[i]]);

      grdmaxmin.Cell[1, 1].Alignment := taCenter;
      grdmaxmin.Cell[1, 1].wrap := True;
      grdmaxmin.Cell[1, 1].Font.Style := [fsbold];
      grdmaxmin.Cell[1, 1].Font.Color := clWebBlue;
      grdmaxmin.Cell[1, 1].Text := 'Min-Min';

      grdmaxmin.Cell[1, 2].Alignment := taCenter;
      grdmaxmin.Cell[1, 2].wrap := True;
      grdmaxmin.Cell[1, 2].Font.Style := [fsbold];
      grdmaxmin.Cell[1, 2].Font.Color := clWebBlue;
      grdmaxmin.Cell[1, 2].Text := 'Max-Min';

      grdmaxmin.Cell[1, 3].Alignment := taCenter;
      grdmaxmin.Cell[1, 3].wrap := True;
      grdmaxmin.Cell[1, 3].Font.Style := [fsbold];
      grdmaxmin.Cell[1, 3].Font.Color := clWebBlue;
      grdmaxmin.Cell[1, 3].Text := 'Min-Max';

      grdmaxmin.Cell[1, 0].Alignment := taCenter;
      grdmaxmin.Cell[1, 0].wrap := True;
      grdmaxmin.Cell[1, 0].Font.Style := [fsbold];
      grdmaxmin.Cell[1, 0].Font.Color := clWebBlue;
      grdmaxmin.Cell[1, 0].Text := 'Max-Max';

    end;

    grdmaxmin.Cell[2 + length(MediaDifSol), 4].Alignment := taCenter;
    grdmaxmin.Cell[2 + length(MediaDifSol), 4].wrap := True;
    grdmaxmin.Cell[2 + length(MediaDifSol), 4].Font.Style := [fsbold];
    grdmaxmin.Cell[2 + length(MediaDifSol), 4].Font.Color := clWebBlue;
    grdmaxmin.Cell[2 + length(MediaDifSol), 4].Text := 'Max-Max:';

    grdmaxmin.Cell[3 + length(MediaDifSol), 4].Alignment := taCenter;
    grdmaxmin.Cell[3 + length(MediaDifSol), 4].wrap := True;
    grdmaxmin.Cell[3 + length(MediaDifSol), 4].Font.Style := [fsbold];
    grdmaxmin.Cell[3 + length(MediaDifSol), 4].Font.Color := clWebBlue;
    grdmaxmin.Cell[3 + length(MediaDifSol), 4].Text := 'Min-Min:';

    grdmaxmin.Cell[4 + length(MediaDifSol), 4].Alignment := taCenter;
    grdmaxmin.Cell[4 + length(MediaDifSol), 4].wrap := True;
    grdmaxmin.Cell[4 + length(MediaDifSol), 4].Font.Style := [fsbold];
    grdmaxmin.Cell[4 + length(MediaDifSol), 4].Font.Color := clWebBlue;
    grdmaxmin.Cell[4 + length(MediaDifSol), 4].Text := 'Max-Min:';

    grdmaxmin.Cell[5 + length(MediaDifSol), 4].Alignment := taCenter;
    grdmaxmin.Cell[5 + length(MediaDifSol), 4].wrap := True;
    grdmaxmin.Cell[5 + length(MediaDifSol), 4].Font.Style := [fsbold];
    grdmaxmin.Cell[5 + length(MediaDifSol), 4].Font.Color := clWebBlue;
    grdmaxmin.Cell[5 + length(MediaDifSol), 4].Text := 'Min-Max:';

    grdmaxmin.Cell[6 + length(MediaDifSol), 4].Alignment := taCenter;
    grdmaxmin.Cell[6 + length(MediaDifSol), 4].wrap := True;
    grdmaxmin.Cell[6 + length(MediaDifSol), 4].Font.Style := [fsbold];
    grdmaxmin.Cell[6 + length(MediaDifSol), 4].Font.Color := clWebBlue;
    grdmaxmin.Cell[6 + length(MediaDifSol), 4].Text := 'Min-Max-Max:';

    minmaxmaxnaosol := 99999999999;
    for i := 0 to length(MediaDifSol) - 1 do
    begin
      if minmaxmaxnaosol > MaximoGeralNaosol[OrdemSol[i]] then
      begin
        TIWUserSession(WebApplication.Data).minmaxmaxnaosol := MaximoGeralNaosol
          [OrdemSol[i]];
        j := i;
      end;
    end;
    grdmaxmin.Cell[6 + length(MediaDifSol), 5 + j].Text := formatfloat
      ('0.00', minmaxmaxnaosol);
    grdmaxmin.Cell[6 + length(MediaDifSol), 5 + j].Alignment := taCenter;
    grdmaxmin.Cell[6 + length(MediaDifSol), 5 + j].wrap := True;
    grdmaxmin.Cell[6 + length(MediaDifSol), 5 + j].Font.Color := clWebBlue;

    for i := 0 to length(MediaDifSol) - 1 do
    begin
      grdmaxmin.Cell[2 + length(MediaDifSol), 5 + i].Text := formatfloat
        ('0.00', MaximoGeralNaosol[OrdemSol[i]]);
      grdmaxmin.Cell[2 + length(MediaDifSol), 5 + i].Alignment := taCenter;
      grdmaxmin.Cell[2 + length(MediaDifSol), 5 + i].wrap := True;
      grdmaxmin.Cell[2 + length(MediaDifSol), 5 + i].Font.Color := clWebBlue;

      grdmaxmin.Cell[3 + length(MediaDifSol), 5 + i].Text := formatfloat
        ('0.00', MinimoGeralNaosol[OrdemSol[i]]);
      grdmaxmin.Cell[3 + length(MediaDifSol), 5 + i].Alignment := taCenter;
      grdmaxmin.Cell[3 + length(MediaDifSol), 5 + i].wrap := True;
      grdmaxmin.Cell[3 + length(MediaDifSol), 5 + i].Font.Color := clWebBlue;

      grdmaxmin.Cell[4 + length(MediaDifSol), 5 + i].Text := formatfloat
        ('0.00', MaxMinNaosol[OrdemSol[i]]);
      grdmaxmin.Cell[4 + length(MediaDifSol), 5 + i].Alignment := taCenter;
      grdmaxmin.Cell[4 + length(MediaDifSol), 5 + i].wrap := True;
      grdmaxmin.Cell[4 + length(MediaDifSol), 5 + i].Font.Color := clWebBlue;

      grdmaxmin.Cell[5 + length(MediaDifSol), 5 + i].Text := formatfloat
        ('0.00', MinMaxNaosol[OrdemSol[i]]);
      grdmaxmin.Cell[5 + length(MediaDifSol), 5 + i].Alignment := taCenter;
      grdmaxmin.Cell[5 + length(MediaDifSol), 5 + i].wrap := True;
      grdmaxmin.Cell[5 + length(MediaDifSol), 5 + i].Font.Color := clWebBlue;

      { grdmaxmin.Cell[3+length(MediaDifSol), 4+i].Text := formatfloat('0.00', DesvioPadraoGeralNaoSol[OrdemSol[i]]);
        grdmaxmin.Cell[3+length(MediaDifSol), 4+i].Alignment := taCenter;
        grdmaxmin.Cell[3+length(MediaDifSol), 4+i].wrap := True;
        grdmaxmin.Cell[3+length(MediaDifSol), 4+i].Font.Color := clWebBlue; }
    end;

    for i := 0 to length(MediaDifSol) - 1 do
      for j := 0 to length(MediaDifSol) - 1 do
      begin
        grdmaxmin.Cell[i + 2, j + 5].Alignment := taCenter;
        grdmaxmin.Cell[i + 2, j + 5].wrap := True;
        if i = j then
          grdmaxmin.Cell[i + 2, j + 5].Text := ''
        else
        begin
          grdmaxmin.Cell[i + 2, j + 5].Text := formatfloat
            ('0.00', MaxDifSol[OrdemSol[i], OrdemSol[j]]) + '|' + formatfloat
            ('0.00', MinDifSol[OrdemSol[i], OrdemSol[j]]);
        end;
      end;

    auxregra := 0;
    qtdpareto := 0;
    while auxregra < 0.8 * Total do
    begin
      auxregra := auxregra + ResultSol[OrdemSol[qtdpareto]];
      inc(qtdpareto);
    end;

    grdperdamediaemaxima.Clear;
    grdperdamediaemaxima.ColumnCount := 4;
    grdperdamediaemaxima.rowcount := 1 + qtdpareto;
    for i := 0 to grdperdamediaemaxima.ColumnCount - 1 do
    begin
      for j := 0 to grdperdamediaemaxima.rowcount - 1 do
      begin
        grdperdamediaemaxima.Cell[j, i].Alignment := taCenter;
        grdperdamediaemaxima.Cell[j, i].wrap := True;
        if j = 0 then
        begin
          grdperdamediaemaxima.Cell[j, i].BGColor := clWebWhiteSmoke;
          grdperdamediaemaxima.Cell[j, i].Font.Style := [fsbold];
        end;
      end;
    end;
    grdperdamediaemaxima.Cell[0, 0].Text := 'Alternatives';
    grdperdamediaemaxima.Cell[0, 1].Text := 'Probability[P(Ai)]';
    grdperdamediaemaxima.Cell[0, 2].Text := 'Average Loss[μ(Aj,Ai)]:';
    grdperdamediaemaxima.Cell[0, 3].Text := 'Maximum Loss[Dmax(Aj, Ai)]:';
    for j := 0 to qtdpareto - 1 do
    begin
      for i := 0 to Numalt - 1 do
      begin
        if MatrizSol[OrdemSol[j], i] = 1 then
        begin
          if grdperdamediaemaxima.Cell[1 + j, 0].Text = '' then
            grdperdamediaemaxima.Cell[1 + j, 0].Text := nomealt[i]
          else
            grdperdamediaemaxima.Cell[1 + j, 0].Text :=
              grdperdamediaemaxima.Cell[1 + j, 0].Text + ' + ' + nomealt[i];
        end;
        grdperdamediaemaxima.Cell[j + 1, 1].Text := formatfloat
          ('0', (ResultSol[OrdemSol[j]] / Total) * 100) + '%';
        grdperdamediaemaxima.Cell[j + 1, 2].Text := formatfloat
          ('0.00', MediaGeralNaosol[OrdemSol[j]]);
        grdperdamediaemaxima.Cell[j + 1, 3].Text := formatfloat
          ('0.00', MaximoGeralNaosol[OrdemSol[j]]);
      end;
    end;

    grdregradedecisao.Clear;
    grdregradedecisao.ColumnCount := 1 + qtdpareto;
    grdregradedecisao.rowcount := 11;
    for i := 0 to grdregradedecisao.ColumnCount - 1 do
    begin
      if i >= 1 + qtdpareto then
      begin
        grdregradedecisao.Cell[0, i].Font.Color := ClWebRed;
        grdregradedecisao.Cell[0, i].Font.Style := [fsbold];
      end;
      for j := 0 to grdregradedecisao.rowcount - 1 do
      begin
        grdregradedecisao.Cell[j, i].Alignment := taCenter;
        grdregradedecisao.Cell[j, i].wrap := True;
        if i = 0 then
          grdregradedecisao.Cell[j, 0].Font.Style := [fsbold];

        if j = 0 then
          grdregradedecisao.Cell[j, i].BGColor := clWebWhiteSmoke;

        if j >= 8 then
        begin
          grdregradedecisao.Cell[j, i].Font.Color := ClWebRed;
          grdregradedecisao.Cell[j, i].BGColor := clWebWhiteSmoke;
        end;
      end;

    end;

    grdregradedecisao.Cell[0, 0].Text := 'Alternative:';
    grdregradedecisao.Cell[1, 0].Text := 'Probability [P(Ai)]:';
    grdregradedecisao.Cell[2, 0].Text := 'Average Loss[μ(Aj,Ai)]:';
    grdregradedecisao.Cell[3, 0].Text := 'Average Gain[μ(Ai,Aj)]:';
    grdregradedecisao.Cell[4, 0].Text := 'Expected Loss[[1-P(Ai)]*μ(Aj,Ai)]:';
    grdregradedecisao.Cell[5, 0].Text := 'Expected Gain[P(Ai)* μ(Aj,Ai)]:';
    grdregradedecisao.Cell[6, 0].Text := 'Maximum Loss[Dmax(Aj, Ai)]:';
    grdregradedecisao.Cell[7, 0].Text := 'Global Performance[G(Ai)]:';

    grdregradedecisao.Cell[8, 0].Text := 'Minimum Avarage Loss:';
    grdregradedecisao.Cell[9, 0].Text := 'Minimum Maximum Loss:';
    grdregradedecisao.Cell[10, 0].Text := 'Maximum Global Performance:';

    aux1 := 99999999999;
    aux2 := 99999999999;
    aux3 := -9999999999;

    for i := 0 to qtdpareto - 1 do
    begin

      for j := 0 to Numalt - 1 do
      begin
        if MatrizSol[OrdemSol[i], j] = 1 then
        begin
          if grdregradedecisao.Cell[0, 1 + i].Text = '' then
            grdregradedecisao.Cell[0, 1 + i].Text := nomealt[j]
          else
            grdregradedecisao.Cell[0, 1 + i].Text := grdregradedecisao.Cell
              [0, 1 + i].Text + ' + ' + nomealt[j];
        end;
      end;

      // grdregradedecisao.Cell[0,1 + i].Text := 'Sol ' + inttostr(i + 1);
      grdregradedecisao.Cell[1, 1 + i].Text := formatfloat
        ('0', (ResultSol[OrdemSol[i]] / Total) * 100) + '%';
      grdregradedecisao.Cell[2, 1 + i].Text := formatfloat
        ('0.00', MediaGeralNaosol[OrdemSol[i]]);
      grdregradedecisao.Cell[3, 1 + i].Text := formatfloat
        ('0.00', MediaGeral[OrdemSol[i]]);
      grdregradedecisao.Cell[4, 1 + i].Text := formatfloat
        ('0.00', MediaGeralNaosol[OrdemSol[i]] *
          (1 - (ResultSol[OrdemSol[i]] / Total)));
      grdregradedecisao.Cell[5, 1 + i].Text := formatfloat
        ('0.00', MediaGeral[OrdemSol[i]] * (ResultSol[OrdemSol[i]] / Total));
      grdregradedecisao.Cell[6, 1 + i].Text := formatfloat
        ('0.00', MaximoGeralNaosol[OrdemSol[i]]);
      auxregra := 0;
      auxregra := MediaGeral[OrdemSol[i]] * (ResultSol[OrdemSol[i]] / Total)
        - MediaGeralNaosol[OrdemSol[i]] *
        ((Total - ResultSol[OrdemSol[i]]) / Total);
      grdregradedecisao.Cell[7, 1 + i].Text := formatfloat('0.00', auxregra);

      if MediaGeralNaosol[OrdemSol[i]] < aux1 then
      begin

        aux1 := MediaGeralNaosol[OrdemSol[i]];
        SolRegra1 := i;
      end;
      if MaximoGeralNaosol[OrdemSol[i]] < aux2 then
      begin
        aux2 := MaximoGeralNaosol[OrdemSol[i]];
        SolRegra2 := i;
      end;
      if (MediaGeral[OrdemSol[i]] * (ResultSol[OrdemSol[i]] / Total)
          - MediaGeralNaosol[OrdemSol[i]] * ((Total - ResultSol[OrdemSol[i]])
            / Total)) > aux3 then
      begin
        aux3 := (MediaGeral[OrdemSol[i]] * (ResultSol[OrdemSol[i]] / Total)
            - MediaGeralNaosol[OrdemSol[i]] * ((Total - ResultSol[OrdemSol[i]])
              / Total));
        SolRegra3 := i;
      end;

    end;

    grdregradedecisao.Cell[8, 1 + SolRegra1].Text := formatfloat
      ('0.00', MediaGeralNaosol[OrdemSol[SolRegra1]]);

    grdregradedecisao.Cell[9, 1 + SolRegra2].Text := formatfloat
      ('0.00', MaximoGeralNaosol[OrdemSol[SolRegra2]]);

    grdregradedecisao.Cell[10, 1 + SolRegra3].Text := formatfloat
      ('0.00', (MediaGeral[OrdemSol[SolRegra3]] *
          (ResultSol[OrdemSol[SolRegra3]] / Total) - MediaGeralNaosol
          [OrdemSol[SolRegra3]] * ((Total - ResultSol[OrdemSol[SolRegra3]])
            / Total)));

    SetLength(epsilon, 4);
    SetLength(omega, 4);
    SetLength(faixaprob, 4);

    faixaprob[0] := 80;
    faixaprob[1] := 70;
    faixaprob[2] := 60;
    faixaprob[3] := 50;

    epsilon[0] := 0.5;
    epsilon[1] := 0.4;
    epsilon[2] := 0.3;
    epsilon[3] := 0.2;

    omega[0] := 0.25;
    omega[1] := 0.2;
    omega[2] := 0.15;
    omega[3] := 0.10;

    { grdlimiares.Cells[0, 0] := '80% < P(x)';
      grdlimiares.Cells[0, 1] := '70% < P(x) < 80%';
      grdlimiares.Cells[0, 2] := '60% < P(x) < 70%';
      grdlimiares.Cells[0, 3] := '50% < P(x) < 60%';
      }

    // voltar pra ca
    grdlimiares.DynEdits[1, 0] := '0.50';
    grdlimiares.DynEdits[1, 1] := '0.40';
    grdlimiares.DynEdits[1, 2] := '0.30';
    grdlimiares.DynEdits[1, 3] := '0.20';

    grdlimiares.DynEdits[2, 0] := '0.25';
    grdlimiares.DynEdits[2, 1] := '0.20';
    grdlimiares.DynEdits[2, 2] := '0.15';
    grdlimiares.DynEdits[2, 3] := '0.10';

    recomendacao := false;

    grdrecomendacao.Clear;
    grdrecomendacao.rowcount := 4;
    grdrecomendacao.ColumnCount := 2;
    for i := 0 to grdrecomendacao.ColumnCount - 1 do
    begin

      grdrecomendacao.Cell[0, i].Font.Color := ClWebRed;
      grdrecomendacao.Cell[0, i].Font.Style := [fsbold];
      grdrecomendacao.Cell[0, i].BGColor := clWebWhiteSmoke;
      for j := 0 to grdrecomendacao.rowcount - 1 do
      begin
        grdrecomendacao.Cell[j, i].Alignment := taCenter;
        grdrecomendacao.Cell[j, i].wrap := True;
      end;
    end;

    grdrecomendacao.Cell[0, 0].Text := 'Solution';
    grdrecomendacao.Cell[0, 1].Text := 'Possibles Recommendations';
    grdrecomendacao.Cell[1, 1].Text := 'Best Alternative';
    grdrecomendacao.Cell[2, 1].Text := 'Two/Three Alternatives are Competitive';
    grdrecomendacao.Cell[3, 1].Text := 'Unable to make it';
    // txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' + grdperdamediaemaxima.Cell[1, 0].Text + ')' + ' + ' + 'P(' + grdperdamediaemaxima.Cell[2, 0].Text + ')' + ' + ' + 'P(' + grdperdamediaemaxima.Cell[3, 0].Text + ')';
    // Para alt1
    if (TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]] / Total) >=
      (faixaprob[0] / 100) then // Se a alt[1] tiver uma probabilidade >= 80%, ela é recomendada
    begin
      Px := (TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
          / Total) * 100;
      if (TIWUserSession(WebApplication.Data).MaximoGeralNaosol[OrdemSol[0]]
          < epsilon[0]) and (TIWUserSession(WebApplication.Data)
          .MediaGeralNaosol[OrdemSol[0]] < omega[0]) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[1, 0].Text = '' then
              grdrecomendacao.Cell[1, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
        recomendacao := false;
    end;
    if ((ResultSol[OrdemSol[0]] / Total) < (faixaprob[0] / 100)) and
      ((ResultSol[OrdemSol[0]] / Total) >= (faixaprob[1] / 100)) and
      (recomendacao = false) then
    begin
      Px := (TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
          / Total) * 100;
      if (MaximoGeralNaosol[OrdemSol[0]] < epsilon[1]) and
        (MediaGeralNaosol[OrdemSol[0]] < omega[1]) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[1, 0].Text = '' then
              grdrecomendacao.Cell[1, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if ((ResultSol[OrdemSol[0]] / Total) < (faixaprob[1] / 100)) and
      ((ResultSol[OrdemSol[0]] / Total) >= (faixaprob[2] / 100)) and
      (recomendacao = false) then
    begin
      Px := (TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
          / Total) * 100;
      if (MaximoGeralNaosol[OrdemSol[0]] < epsilon[2]) and
        (MediaGeralNaosol[OrdemSol[0]] < omega[2]) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[1, 0].Text = '' then
              grdrecomendacao.Cell[1, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if ((ResultSol[OrdemSol[0]] / Total) < (faixaprob[2] / 100)) and
      ((ResultSol[OrdemSol[0]] / Total) >= (faixaprob[3] / 100)) and
      (recomendacao = false) then
    begin
      Px := (TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
          / Total) * 100;
      if (MaximoGeralNaosol[OrdemSol[0]] < epsilon[3]) and
        (MediaGeralNaosol[OrdemSol[0]] < omega[3]) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[1, 0].Text = '' then
              grdrecomendacao.Cell[1, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;

    // Para alt1 e alt2
    if (((TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]] + ResultSol
            [OrdemSol[1]]) / Total) >= (faixaprob[0] / 100)) and
      (recomendacao = false) then
    // Se a alt[1] tiver uma probabilidade >= 80%, ela é recomendada
    begin
      Px := ((TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
            + ResultSol[OrdemSol[1]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[0]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[0])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[0]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[0])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if TIWUserSession(WebApplication.Data).MatrizSol[OrdemSol[0], j]
            = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaxima.Cell[2, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
        recomendacao := false;
    end;
    if (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]]) / Total) <
        (faixaprob[0] / 100)) and
      (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]]) / Total) >=
        (faixaprob[1] / 100)) and (recomendacao = false) then
    begin
      Px := ((TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
            + ResultSol[OrdemSol[1]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[1]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[1])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[1]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[1])) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaxima.Cell[2, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]]) / Total) <
        (faixaprob[1] / 100)) and
      (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]]) / Total) >=
        (faixaprob[2] / 100)) and (recomendacao = false) then
    begin
      Px := ((TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
            + ResultSol[OrdemSol[1]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[2]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[2])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[2]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[2])) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaxima.Cell[2, 0].Text + ')';

        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]]) / Total) <
        (faixaprob[2] / 100)) and
      (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]]) / Total) >=
        (faixaprob[3] / 100)) and (recomendacao = false) then
    begin
      Px := ((TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
            + ResultSol[OrdemSol[1]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[3]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[3])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[3]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[3])) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaxima.Cell[2, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;

    // Para alt1, alt2 e alt3
    if ((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
          [OrdemSol[2]]) / Total) >= (faixaprob[0] / 100) then
    // Se a alt[1] tiver uma probabilidade >= 80%, ela é recomendada
    begin
      Px := ((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[0]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[0])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[0]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[0])) and
        ((MaximoGeralNaosol[OrdemSol[2]] < epsilon[0]) and
          (MediaGeralNaosol[OrdemSol[2]] < omega[0])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[2], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        recomendacao := True;
        exit;
      end
      else
        recomendacao := false;
    end;
    if (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) < (faixaprob[0] / 100)) and
      (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) >= (faixaprob[1] / 100)) and
      (recomendacao = false) then
    begin
      Px := ((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[1]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[1])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[1]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[1])) and
        ((MaximoGeralNaosol[OrdemSol[2]] < epsilon[1]) and
          (MediaGeralNaosol[OrdemSol[2]] < omega[1])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[2], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) < (faixaprob[1] / 100)) and
      (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) >= (faixaprob[2] / 100)) and
      (recomendacao = false) then
    begin
      Px := ((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[2]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[2])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[2]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[2])) and
        ((MaximoGeralNaosol[OrdemSol[2]] < epsilon[2]) and
          (MediaGeralNaosol[OrdemSol[2]] < omega[2])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[2], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) < (faixaprob[2] / 100)) and
      (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) >= (faixaprob[3] / 100)) and
      (recomendacao = false) then
    begin
      Px := ((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[3]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[3])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[3]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[3])) and
        ((MaximoGeralNaosol[OrdemSol[2]] < epsilon[3]) and
          (MediaGeralNaosol[OrdemSol[2]] < omega[3])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[2], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;

    // caso não passe em nenhuma regra
    if recomendacao = false then
    begin
      grdrecomendacao.Cell[3, 0].Text := 'Not Possible';
      grdrecomendacao.Cell[3, 0].Text := grdrecomendacao.Cell[3, 0]
        .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
    end;

  end;
end;

procedure TIWForm1.btnBackRU2Click(Sender: TObject);
begin
  iwrgnCaixaDialogoRU2.Visible := false;
end;

procedure TIWForm1.btnBackRUClick(Sender: TObject);
begin
  iwrgnCaixaDialogoRU.Visible := false;
end;

procedure TIWForm1.btncomparacaoClick(Sender: TObject);
begin
  rgnresultado.Visible := false;
  rgnresultadopromethee.Visible := false;
  rgncomparacao.Visible := True;
end;

procedure TIWForm1.IWButton2Click(Sender: TObject);
begin
  rgnresultado.Visible := false;
  rgnresultadopromethee.Visible := false;
  rgncomparacao.Visible := True;
end;

procedure TIWForm1.cmbbxp1Change(Sender: TObject);
var
  i, aux: Integer;
begin
  if StrToInt(cmbbxp1.Items[cmbbxp1.ItemIndex]) <= StrToInt
    (cmbbxp2.Items[cmbbxp2.ItemIndex]) then
  begin
    aux := cmbbxp2.ItemIndex;
    cmbbxp2.ItemIndex := cmbbxp1.ItemIndex;
    cmbbxp1.ItemIndex := aux;
    if cmbbxp2.ItemIndex <= 1 then
      cmbbxp2.ItemIndex := 1;
    if cmbbxp1.ItemIndex >= 47 then
      cmbbxp1.ItemIndex := 47;

    if cmbbxp1.ItemIndex = cmbbxp2.ItemIndex then
      cmbbxp2.ItemIndex := cmbbxp2.ItemIndex + 1;

    cmbbxp2Change(self);
  end;

  lblp2.Caption := '% < P(x) < ' + cmbbxp1.Items[cmbbxp1.ItemIndex] + '%';

end;

procedure TIWForm1.cmbbxp1prometheeChange(Sender: TObject);
var
  i, aux: Integer;
begin
  if StrToInt(cmbbxp1promethee.Items[cmbbxp1promethee.ItemIndex]) <= StrToInt
    (cmbbxp2promethee.Items[cmbbxp2promethee.ItemIndex]) then
  begin
    aux := cmbbxp2promethee.ItemIndex;
    cmbbxp2promethee.ItemIndex := cmbbxp1promethee.ItemIndex;
    cmbbxp1promethee.ItemIndex := aux;
    if cmbbxp2promethee.ItemIndex <= 1 then
      cmbbxp2promethee.ItemIndex := 1;
    if cmbbxp1promethee.ItemIndex >= 47 then
      cmbbxp1promethee.ItemIndex := 47;

    if cmbbxp1promethee.ItemIndex = cmbbxp2promethee.ItemIndex then
      cmbbxp2promethee.ItemIndex := cmbbxp2promethee.ItemIndex + 1;

    cmbbxp2prometheeChange(self);
  end;

  lblp2promethee.Caption := '% < P(x) < ' + cmbbxp1promethee.Items
    [cmbbxp1promethee.ItemIndex] + '%';

end;

procedure TIWForm1.cmbbxp2Change(Sender: TObject);
var
  i, aux: Integer;
begin
  if StrToInt(cmbbxp1.Items[cmbbxp1.ItemIndex]) <= StrToInt
    (cmbbxp2.Items[cmbbxp2.ItemIndex]) then
  begin
    aux := cmbbxp2.ItemIndex;
    cmbbxp2.ItemIndex := cmbbxp1.ItemIndex;
    cmbbxp1.ItemIndex := aux;
    if cmbbxp2.ItemIndex <= 1 then
      cmbbxp2.ItemIndex := 1;
    if cmbbxp1.ItemIndex >= 47 then
      cmbbxp1.ItemIndex := 47;

    if cmbbxp1.ItemIndex = cmbbxp2.ItemIndex then
      cmbbxp2.ItemIndex := cmbbxp2.ItemIndex + 1;

    cmbbxp1Change(self);
  end;

  if StrToInt(cmbbxp2.Items[cmbbxp2.ItemIndex]) <= StrToInt
    (cmbbxp3.Items[cmbbxp3.ItemIndex]) then
  begin
    aux := cmbbxp3.ItemIndex;
    cmbbxp3.ItemIndex := cmbbxp2.ItemIndex;
    cmbbxp2.ItemIndex := aux;
    if cmbbxp3.ItemIndex <= 2 then
      cmbbxp3.ItemIndex := 2;
    if cmbbxp2.ItemIndex >= 48 then
      cmbbxp2.ItemIndex := 48;

    if cmbbxp2.ItemIndex = cmbbxp3.ItemIndex then
      cmbbxp3.ItemIndex := cmbbxp3.ItemIndex + 1;

    cmbbxp3Change(self);
  end;

  lblp3.Caption := '% < P(x) < ' + cmbbxp2.Items[cmbbxp2.ItemIndex] + '%';
end;

procedure TIWForm1.cmbbxp2prometheeChange(Sender: TObject);
var
  i, aux: Integer;
begin
  if StrToInt(cmbbxp1promethee.Items[cmbbxp1promethee.ItemIndex]) <= StrToInt
    (cmbbxp2promethee.Items[cmbbxp2promethee.ItemIndex]) then
  begin
    aux := cmbbxp2promethee.ItemIndex;
    cmbbxp2promethee.ItemIndex := cmbbxp1promethee.ItemIndex;
    cmbbxp1promethee.ItemIndex := aux;
    if cmbbxp2promethee.ItemIndex <= 1 then
      cmbbxp2promethee.ItemIndex := 1;
    if cmbbxp1promethee.ItemIndex >= 47 then
      cmbbxp1promethee.ItemIndex := 47;

    if cmbbxp1promethee.ItemIndex = cmbbxp2promethee.ItemIndex then
      cmbbxp2promethee.ItemIndex := cmbbxp2promethee.ItemIndex + 1;

    cmbbxp1prometheeChange(self);
  end;

  if StrToInt(cmbbxp2promethee.Items[cmbbxp2promethee.ItemIndex]) <= StrToInt
    (cmbbxp3promethee.Items[cmbbxp3promethee.ItemIndex]) then
  begin
    aux := cmbbxp3promethee.ItemIndex;
    cmbbxp3promethee.ItemIndex := cmbbxp2promethee.ItemIndex;
    cmbbxp2promethee.ItemIndex := aux;
    if cmbbxp3promethee.ItemIndex <= 2 then
      cmbbxp3promethee.ItemIndex := 2;
    if cmbbxp2promethee.ItemIndex >= 48 then
      cmbbxp2promethee.ItemIndex := 48;

    if cmbbxp2promethee.ItemIndex = cmbbxp3promethee.ItemIndex then
      cmbbxp3promethee.ItemIndex := cmbbxp3promethee.ItemIndex + 1;

    cmbbxp3prometheeChange(self);
  end;

  lblp3promethee.Caption := '% < P(x) < ' + cmbbxp2promethee.Items
    [cmbbxp2promethee.ItemIndex] + '%';
end;

procedure TIWForm1.cmbbxp3Change(Sender: TObject);
var
  i, aux: Integer;
begin
  if StrToInt(cmbbxp2.Items[cmbbxp2.ItemIndex]) <= StrToInt
    (cmbbxp3.Items[cmbbxp3.ItemIndex]) then
  begin
    aux := cmbbxp3.ItemIndex;
    cmbbxp3.ItemIndex := cmbbxp2.ItemIndex;
    cmbbxp2.ItemIndex := aux;
    if cmbbxp3.ItemIndex <= 2 then
      cmbbxp3.ItemIndex := 2;
    if cmbbxp2.ItemIndex >= 48 then
      cmbbxp2.ItemIndex := 48;

    if cmbbxp2.ItemIndex = cmbbxp3.ItemIndex then
      cmbbxp3.ItemIndex := cmbbxp3.ItemIndex + 1;

    cmbbxp2Change(self);
  end;

  if StrToInt(cmbbxp3.Items[cmbbxp3.ItemIndex]) <= StrToInt
    (cmbbxp4.Items[cmbbxp4.ItemIndex]) then
  begin
    aux := cmbbxp4.ItemIndex;
    cmbbxp4.ItemIndex := cmbbxp3.ItemIndex;
    cmbbxp3.ItemIndex := aux;
    if cmbbxp4.ItemIndex <= 3 then
      cmbbxp4.ItemIndex := 3;
    if cmbbxp3.ItemIndex >= 49 then
      cmbbxp3.ItemIndex := 49;

    if cmbbxp3.ItemIndex = cmbbxp4.ItemIndex then
      cmbbxp4.ItemIndex := cmbbxp4.ItemIndex + 1;

    cmbbxp4Change(self);
  end;

  lblp4.Caption := '% < P(x) < ' + cmbbxp3.Items[cmbbxp3.ItemIndex] + '%';
end;

procedure TIWForm1.cmbbxp3prometheeChange(Sender: TObject);
var
  i, aux: Integer;
begin
  if StrToInt(cmbbxp2promethee.Items[cmbbxp2promethee.ItemIndex]) <= StrToInt
    (cmbbxp3promethee.Items[cmbbxp3promethee.ItemIndex]) then
  begin
    aux := cmbbxp3promethee.ItemIndex;
    cmbbxp3promethee.ItemIndex := cmbbxp2promethee.ItemIndex;
    cmbbxp2promethee.ItemIndex := aux;
    if cmbbxp3promethee.ItemIndex <= 2 then
      cmbbxp3promethee.ItemIndex := 2;
    if cmbbxp2promethee.ItemIndex >= 48 then
      cmbbxp2promethee.ItemIndex := 48;

    if cmbbxp2promethee.ItemIndex = cmbbxp3promethee.ItemIndex then
      cmbbxp3promethee.ItemIndex := cmbbxp3promethee.ItemIndex + 1;

    cmbbxp2prometheeChange(self);
  end;

  if StrToInt(cmbbxp3promethee.Items[cmbbxp3promethee.ItemIndex]) <= StrToInt
    (cmbbxp4promethee.Items[cmbbxp4promethee.ItemIndex]) then
  begin
    aux := cmbbxp4promethee.ItemIndex;
    cmbbxp4promethee.ItemIndex := cmbbxp3promethee.ItemIndex;
    cmbbxp3promethee.ItemIndex := aux;
    if cmbbxp4promethee.ItemIndex <= 3 then
      cmbbxp4promethee.ItemIndex := 3;
    if cmbbxp3promethee.ItemIndex >= 49 then
      cmbbxp3promethee.ItemIndex := 49;

    if cmbbxp3promethee.ItemIndex = cmbbxp4promethee.ItemIndex then
      cmbbxp4promethee.ItemIndex := cmbbxp4promethee.ItemIndex + 1;

    cmbbxp4prometheeChange(self);
  end;

  lblp4promethee.Caption := '% < P(x) < ' + cmbbxp3promethee.Items
    [cmbbxp3promethee.ItemIndex] + '%';
end;

procedure TIWForm1.cmbbxp4Change(Sender: TObject);
var
  i, aux: Integer;
begin

  if StrToInt(cmbbxp3.Items[cmbbxp3.ItemIndex]) <= StrToInt
    (cmbbxp4.Items[cmbbxp4.ItemIndex]) then
  begin
    aux := cmbbxp4.ItemIndex;
    cmbbxp4.ItemIndex := cmbbxp3.ItemIndex;
    cmbbxp3.ItemIndex := aux;
    if cmbbxp4.ItemIndex <= 3 then
      cmbbxp4.ItemIndex := 3;
    if cmbbxp3.ItemIndex >= 49 then
      cmbbxp3.ItemIndex := 49;

    if cmbbxp3.ItemIndex = cmbbxp4.ItemIndex then
      cmbbxp4.ItemIndex := cmbbxp4.ItemIndex + 1;

    cmbbxp3Change(self);
  end;
end;

procedure TIWForm1.cmbbxp4prometheeChange(Sender: TObject);
var
  i, aux: Integer;
begin

  if StrToInt(cmbbxp3promethee.Items[cmbbxp3promethee.ItemIndex]) <= StrToInt
    (cmbbxp4promethee.Items[cmbbxp4promethee.ItemIndex]) then
  begin
    aux := cmbbxp4promethee.ItemIndex;
    cmbbxp4promethee.ItemIndex := cmbbxp3promethee.ItemIndex;
    cmbbxp3promethee.ItemIndex := aux;
    if cmbbxp4promethee.ItemIndex <= 3 then
      cmbbxp4promethee.ItemIndex := 3;
    if cmbbxp3promethee.ItemIndex >= 49 then
      cmbbxp3promethee.ItemIndex := 49;

    if cmbbxp3promethee.ItemIndex = cmbbxp4promethee.ItemIndex then
      cmbbxp4promethee.ItemIndex := cmbbxp4promethee.ItemIndex + 1;

    cmbbxp3prometheeChange(self);
  end;
end;

procedure TIWForm1.btnforcacritClick(Sender: TObject);
Var
  i, j: Integer;
  Fatorial: real;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    {cmbbxfiltro.Clear;
    for i := 0 to length(ResultSol) - 1 do
      cmbbxfiltro.Items.Add('Solution ' + inttostr(i + 1));

    cmbbxfiltro.ItemIndex := 0;
    cmbbxfiltroChange(self);

    rgnfiltro2.Visible := True;
    rgnresultado.Visible := false;
    if Numcrit = 7 then
    begin
      grdprimeiro.rowcount := 6;
      grdsegundo.rowcount := 6;
    end;
    if (Numcrit = 6) or (Numcrit = 5) then
    begin
      grdprimeiro.rowcount := 5;
      grdsegundo.rowcount := 5;
    end;
    if Numcrit = 4 then
    begin
      grdprimeiro.rowcount := 4;
      grdsegundo.rowcount := 4;
    end;
    if Numcrit = 3 then
    begin
      grdprimeiro.rowcount := 3;
      grdsegundo.rowcount := 3;
    end;

    grdprimeiro.ColumnCount := 1 + Numcrit;
    grdsegundo.ColumnCount := 1 + Numcrit;

    for j := 0 to grdprimeiro.ColumnCount - 1 do
      for i := 0 to grdprimeiro.rowcount - 1 do
      begin
        grdprimeiro.Cell[i, j].Alignment := taCenter;
        grdprimeiro.Cell[i, j].wrap := True;
        // grdresultados.Cell[0,j].Font.Style := [fsbold];
      end;
    for j := 0 to grdprimeiro.ColumnCount - 1 do
      grdprimeiro.Cell[0, j].Font.Style := [fsbold];

    for i := 0 to grdprimeiro.rowcount - 1 do
      grdprimeiro.Cell[i, 0].Font.Style := [fsbold];

    Fatorial := 1;
    for i := 1 to Numcrit - 1 do
      Fatorial := Fatorial * i;
    if Numcrit = 7 then
    begin
      for i := 0 to Numcrit - 1 do
      begin
        grdprimeiro.Cell[0, i + 1].Text := Nomecrit[i];
        grdprimeiro.Cell[1, i + 1].Text := formatfloat
          ('0', crit10[i] / Fatorial * 100) + ' %';
        grdprimeiro.Cell[2, i + 1].Text := formatfloat
          ('0', crit12[i] / Fatorial * 100) + ' %';
        grdprimeiro.Cell[3, i + 1].Text := formatfloat
          ('0', crit13[i] / Fatorial * 100) + ' %';
        grdprimeiro.Cell[4, i + 1].Text := formatfloat
          ('0', crit14[i] / Fatorial * 100) + ' %';
        grdprimeiro.Cell[5, i + 1].Text := formatfloat
          ('0', Fatorial / Fatorial * 100) + ' %';
      end;

      grdprimeiro.Cell[0, 0].Text := 'Criteria:';
      grdprimeiro.Cell[1, 0].Text := '1th:';
      grdprimeiro.Cell[2, 0].Text := '1th + 2th:';
      grdprimeiro.Cell[3, 0].Text := '1th + 2th + 3th:';
      grdprimeiro.Cell[4, 0].Text := '1th + 2th + 3th + 4th:';
      grdprimeiro.Cell[5, 0].Text := 'Others:';

    end;

    if (Numcrit = 6) or (Numcrit = 5) then
    begin
      for i := 0 to Numcrit - 1 do
      begin
        grdprimeiro.Cell[0, i + 1].Text := Nomecrit[i];
        grdprimeiro.Cell[1, i + 1].Text := formatfloat
          ('0', crit10[i] / Fatorial * 100) + ' %';
        grdprimeiro.Cell[2, i + 1].Text := formatfloat
          ('0', crit12[i] / Fatorial * 100) + ' %';
        grdprimeiro.Cell[3, i + 1].Text := formatfloat
          ('0', crit13[i] / Fatorial * 100) + ' %';
        // grdprimeiro.Cell[4, i+1].Text := formatfloat('0.00', crit14[i]/fatorial *100) + ' %';
        grdprimeiro.Cell[4, i + 1].Text := formatfloat
          ('0', Fatorial / Fatorial * 100) + ' %';
      end;

      grdprimeiro.Cell[0, 0].Text := 'Criteria:';
      grdprimeiro.Cell[1, 0].Text := '1th:';
      grdprimeiro.Cell[2, 0].Text := '1th + 2th:';
      grdprimeiro.Cell[3, 0].Text := '1th + 2th + 3th:';
      // grdprimeiro.Cell[4, 0].Text := '1th + 2th + 3th + 4th:';
      grdprimeiro.Cell[4, 0].Text := 'Others:';
    end;
    if (Numcrit = 4) then
    begin
      for i := 0 to Numcrit - 1 do
      begin
        grdprimeiro.Cell[0, i + 1].Text := Nomecrit[i];
        grdprimeiro.Cell[1, i + 1].Text := formatfloat
          ('0', crit10[i] / Fatorial * 100) + ' %';
        grdprimeiro.Cell[2, i + 1].Text := formatfloat
          ('0', crit12[i] / Fatorial * 100) + ' %';
        // grdprimeiro.Cell[3, i+1].Text := formatfloat('0.00', crit13[i]/fatorial *100) + ' %';
        // grdprimeiro.Cell[4, i+1].Text := formatfloat('0.00', crit14[i]/fatorial *100) + ' %';
        grdprimeiro.Cell[3, i + 1].Text := formatfloat
          ('0', Fatorial / Fatorial * 100) + ' %';
      end;

      grdprimeiro.Cell[0, 0].Text := 'Criteria:';
      grdprimeiro.Cell[1, 0].Text := '1th:';
      grdprimeiro.Cell[2, 0].Text := '1th + 2th:';
      // grdprimeiro.Cell[3, 0].Text := '1th + 2th + 3th:';
      // grdprimeiro.Cell[4, 0].Text := '1th + 2th + 3th + 4th:';
      grdprimeiro.Cell[3, 0].Text := 'Others:';
    end;
    if (Numcrit = 3) then
    begin
      for i := 0 to Numcrit - 1 do
      begin
        grdprimeiro.Cell[0, i + 1].Text := Nomecrit[i];
        grdprimeiro.Cell[1, i + 1].Text := formatfloat
          ('0', crit10[i] / Fatorial * 100) + ' %';
        // grdprimeiro.Cell[2, i+1].Text := formatfloat('0.00', crit12[i]/fatorial *100) + ' %';
        // grdprimeiro.Cell[3, i+1].Text := formatfloat('0.00', crit13[i]/fatorial *100) + ' %';
        // grdprimeiro.Cell[4, i+1].Text := formatfloat('0.00', crit14[i]/fatorial *100) + ' %';
        grdprimeiro.Cell[2, i + 1].Text := formatfloat
          ('0', Fatorial / Fatorial * 100) + ' %';
      end;

      grdprimeiro.Cell[0, 0].Text := 'Criteria:';
      grdprimeiro.Cell[1, 0].Text := '1th:';
      // grdprimeiro.Cell[2, 0].Text := '1th + 2th:';
      // grdprimeiro.Cell[3, 0].Text := '1th + 2th + 3th:';
      // grdprimeiro.Cell[4, 0].Text := '1th + 2th + 3th + 4th:';
      grdprimeiro.Cell[2, 0].Text := 'Others:';
    end;
    /// ///////////////////////////////////////////////////////////////////

    for j := 0 to grdsegundo.ColumnCount - 1 do
      for i := 0 to grdsegundo.rowcount - 1 do
      begin
        grdsegundo.Cell[i, j].Alignment := taCenter;
        grdsegundo.Cell[i, j].wrap := True;
        // grdresultados.Cell[0,j].Font.Style := [fsbold];
      end;
    for j := 0 to grdsegundo.ColumnCount - 1 do
      grdsegundo.Cell[0, j].Font.Style := [fsbold];

    for i := 0 to grdsegundo.rowcount - 1 do
      grdsegundo.Cell[i, 0].Font.Style := [fsbold];

    Fatorial := 1;
    for i := 1 to Numcrit - 1 do
      Fatorial := Fatorial * i;

    if Numcrit = 7 then
    begin
      for i := 0 to Numcrit - 1 do
      begin
        grdsegundo.Cell[0, i + 1].Text := Nomecrit[i];
        grdsegundo.Cell[1, i + 1].Text := formatfloat
          ('0', crit20[i] / Fatorial * 100) + ' %';
        grdsegundo.Cell[2, i + 1].Text := formatfloat
          ('0', crit21[i] / Fatorial * 100) + ' %';
        grdsegundo.Cell[3, i + 1].Text := formatfloat
          ('0', crit23[i] / Fatorial * 100) + ' %';
        grdsegundo.Cell[4, i + 1].Text := formatfloat
          ('0', crit24[i] / Fatorial * 100) + ' %';
        grdsegundo.Cell[5, i + 1].Text := formatfloat
          ('0', Fatorial / Fatorial * 100) + ' %';
      end;

      grdsegundo.Cell[0, 0].Text := 'Criteria:';
      grdsegundo.Cell[1, 0].Text := '2th:';
      grdsegundo.Cell[2, 0].Text := '2th + 1th:';
      grdsegundo.Cell[3, 0].Text := '2th + 1th + 3th:';
      grdsegundo.Cell[4, 0].Text := '2th + 1th + 3th + 4th:';
      grdsegundo.Cell[5, 0].Text := 'Others:';
    end;
    if (Numcrit = 6) or (Numcrit = 5) then
    begin
      for i := 0 to Numcrit - 1 do
      begin
        grdsegundo.Cell[0, i + 1].Text := Nomecrit[i];
        grdsegundo.Cell[1, i + 1].Text := formatfloat
          ('0', crit20[i] / Fatorial * 100) + ' %';
        grdsegundo.Cell[2, i + 1].Text := formatfloat
          ('0', crit21[i] / Fatorial * 100) + ' %';
        grdsegundo.Cell[3, i + 1].Text := formatfloat
          ('0', crit23[i] / Fatorial * 100) + ' %';
        // grdsegundo.Cell[4, i+1].Text := formatfloat('0.00', crit24[i]/fatorial *100) + ' %';
        grdsegundo.Cell[4, i + 1].Text := formatfloat
          ('0', Fatorial / Fatorial * 100) + ' %';
      end;

      grdsegundo.Cell[0, 0].Text := 'Criteria:';
      grdsegundo.Cell[1, 0].Text := '2th:';
      grdsegundo.Cell[2, 0].Text := '2th + 1th:';
      grdsegundo.Cell[3, 0].Text := '2th + 1th + 3th:';
      // grdsegundo.Cell[4, 0].Text := '2th + 1th + 3th + 4th:';
      grdsegundo.Cell[4, 0].Text := 'Others:';
    end;
    if (Numcrit = 4) then
    begin
      for i := 0 to Numcrit - 1 do
      begin
        grdsegundo.Cell[0, i + 1].Text := Nomecrit[i];
        grdsegundo.Cell[1, i + 1].Text := formatfloat
          ('0', crit20[i] / Fatorial * 100) + ' %';
        grdsegundo.Cell[2, i + 1].Text := formatfloat
          ('0', crit21[i] / Fatorial * 100) + ' %';
        // grdsegundo.Cell[3, i+1].Text := formatfloat('0.00', crit23[i]/fatorial *100) + ' %';
        // grdsegundo.Cell[4, i+1].Text := formatfloat('0.00', crit24[i]/fatorial *100) + ' %';
        grdsegundo.Cell[3, i + 1].Text := formatfloat
          ('0', Fatorial / Fatorial * 100) + ' %';
      end;

      grdsegundo.Cell[0, 0].Text := 'Criteria:';
      grdsegundo.Cell[1, 0].Text := '2th:';
      grdsegundo.Cell[2, 0].Text := '2th + 1th:';
      // grdsegundo.Cell[3, 0].Text := '2th + 1th + 3th:';
      // grdsegundo.Cell[4, 0].Text := '2th + 1th + 3th + 4th:';
      grdsegundo.Cell[3, 0].Text := 'Others:';
    end;
    if (Numcrit = 3) then
    begin
      for i := 0 to Numcrit - 1 do
      begin
        grdsegundo.Cell[0, i + 1].Text := Nomecrit[i];
        grdsegundo.Cell[1, i + 1].Text := formatfloat
          ('0', crit20[i] / Fatorial * 100) + ' %';
        // grdsegundo.Cell[2, i+1].Text := formatfloat('0.00', crit21[i]/fatorial *100) + ' %';
        // grdsegundo.Cell[3, i+1].Text := formatfloat('0.00', crit23[i]/fatorial *100) + ' %';
        // grdsegundo.Cell[4, i+1].Text := formatfloat('0.00', crit24[i]/fatorial *100) + ' %';
        grdsegundo.Cell[2, i + 1].Text := formatfloat
          ('0', Fatorial / Fatorial * 100) + ' %';
      end;

      grdsegundo.Cell[0, 0].Text := 'Criteria:';
      grdsegundo.Cell[1, 0].Text := '2th:';
      // grdsegundo.Cell[2, 0].Text := '2th + 1th:';
      // grdsegundo.Cell[3, 0].Text := '2th + 1th + 3th:';
      // grdsegundo.Cell[4, 0].Text := '2th + 1th + 3th + 4th:';
      grdsegundo.Cell[2, 0].Text := 'Others:';
    end;
      }
   rgnelicitacaocrit.Visible := true;
   rgnresultado.Visible := false;
  end;

end;

procedure TIWForm1.btncomparacaoprometheeClick(Sender: TObject);
begin
  rgnresultado.Visible := false;
  rgnresultadopromethee.Visible := false;
  rgncomparacao.Visible := True;
end;

procedure TIWForm1.btnContinueregistClick(Sender: TObject);
var
  numproblemas, i, j: Integer;
begin
  rgnOpcaoInput.Visible := false;
  rgnLoginproject.Visible := True;

  lstbxProjects.Clear;
  With TIWUserSession(WebApplication.Data) do
  begin
    novoproblema := false;
    retomada := True;
    with UserSession.ZQuery1 do
    begin
      SQL.Text :=
        'SELECT COUNT(nome_problema) FROM problema WHERE ID_usuario = "' +
        (TIWUserSession(WebApplication.Data).ID_user) + '";';
      ExecSQL;
      Open;
      numproblemas := FieldByName('COUNT(nome_problema)').AsInteger;
      Close;

      SetLength(problems, numproblemas);
      SetLength(id_problem, numproblemas);
      SetLength(date_problem, numproblemas);

      SQL.Text := 'SELECT * FROM problema WHERE ID_usuario = "' +
        (TIWUserSession(WebApplication.Data).ID_user) + '";';
      ExecSQL;
      Open;
      First;
      i := 0;
      while i < numproblemas do
      begin
        problems[i] := FieldByName('nome_problema').AsString;
        id_problem[i] := FieldByName('ID_problema').AsString;
        date_problem[i] := FieldByName('data_problema').AsString;
        lstbxProjects.Items.Add(problems[i] + '(' + date_problem[i] + ')');
        Next;
        inc(i);
      end;
      Close;

    end;
  end;

end;

procedure TIWForm1.btncontinueRU2Click(Sender: TObject);
begin
  WebApplication.TerminateAndRedirect('http://15.229.155.120/CDSID-Access/');
  WebApplication.Terminate();
end;

procedure TIWForm1.btncontinueRUClick(Sender: TObject);
begin
  WebApplication.TerminateAndRedirect('http://15.229.155.120/CDSID-Access/');
  WebApplication.Terminate();
end;

procedure TIWForm1.btncritxcritClick(Sender: TObject);
begin
rgncritxcrit.Visible := true;
rgnelicitacaocrit.Visible := false;
end;

procedure TIWForm1.btnEnterClick(Sender: TObject);
// procedimento de login no sistema
var
  email, id_email: string;
  senha, validation: string;
  quantregis: Integer;

begin
  lblemail.Visible := false;
  lblpassword.Visible := false;
  TIWUserSession(WebApplication.Data).primeiroacesso := false;
  // inicialmente fazendo um selectcount para verificar se já existe algum cadastro com esse email
  with UserSession.ZQuery1 do
  begin

    SQL.Text := 'SELECT count(*) FROM usuario WHERE email="' + (edtemail.Text)
      + '"';
    ExecSQL;
    Open;
    quantregis := FieldByName('count(*)').AsInteger;
    Close;
  end;
  if quantregis = 0 then
  begin
    lblemail.Visible := True;
    lnkbacklogin.Visible := True;
  end
  else
  begin

    TIWUserSession(WebApplication.Data).email := edtemail.Text;
    with UserSession.ZQuery1 do
    begin

      SQL.Text := 'SELECT validacao FROM usuario WHERE email ="' +
        (edtemail.Text) + '"';
      ExecSQL;
      Open;
      validation := FieldByName('validacao').AsString;
      Close;

      SQL.Text := 'SELECT password FROM usuario WHERE email ="' +
        (edtemail.Text) + '"';
      ExecSQL;
      Open;
      senha := FieldByName('password').AsString;
      Close;
    end;
    if senha = edtSenha.Text then
    begin
      with UserSession.ZQuery1 do
      begin
        SQL.Text := 'SELECT id FROM usuario WHERE email="' + (edtemail.Text)
          + '"';
        ExecSQL;
        Open;
        TIWUserSession(WebApplication.Data).ID_user := FieldByName('id')
          .AsString;
        Close;

        SQL.Text := 'SELECT analyst FROM usuario WHERE email="' +
          (edtemail.Text) + '"';
        ExecSQL;
        Open;
        TIWUserSession(WebApplication.Data).analyst := 0;
        Close;

        if TIWUserSession(WebApplication.Data).analyst = 0 then
        begin
          TIWUserSession(WebApplication.Data).email := edtemail.Text;
          TIWUserSession(WebApplication.Data).senha := edtSenha.Text;
        end;

      end;

      // * Caso o usuário esteja trocando a senha de cadastro *//
      { if validation = 'password' then
        begin
        WebApplication.ShowMessage('Temporary password has been found. Access http://15.229.155.120/CDSID-ACCESS/ to proceed with changing your password.');
        edtSenha.Text := '';
        //Exit;
        end; }

      // * Caso seja o primeiro acesso *//
      if validation <> 'validado' then
      begin
        WebApplication.showmessage(
          'User has not been validated. Access http://15.229.155.120/CDSID-ACCESS/ to proceed with user validation.');
        edtSenha.Clear;
        exit;
      end;

      rgnEnterLogin.Visible := false;
      rgnOpcaoInput.Visible := True;
      // rgnOpcaoInput.Visible := false;
    end
    else
    begin
      lblpassword.Visible := True;
      lblemail.Visible := false;
      lnkbacklogin.Visible := True;
    end;

  end;

end;

procedure TIWForm1.btnestatisticaprometheeClick(Sender: TObject);
begin
  rgnresultadopromethee.Visible := false;
  rgnAnaliseEstatisticaPromethee.Visible := True;
end;

procedure TIWForm1.btnupdatevisualizacaoPrometheeClick(Sender: TObject);
var
  i, j, k: Integer;
  bar: array of TBarSeries;
  totalacumulado: real;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    if rdgVisualizacaoPromethee.ItemIndex = 0 then
    begin
      chrtresultadopromethee.RemoveAllSeries;
      chrtresultadopromethee.Legend.Items.Clear;
      // chrtresultadoPromethee.ClearChart;
      chrtresultadopromethee.BottomAxis.Items.Clear;
      chrtresultadopromethee.View3D := false;
      SetLength(bar, Numalt);
      i := 0;
      Total := length(CasesOrdemcrit);
      totalacumulado := 0;
      for j := 0 to Numalt - 1 do
      begin
        if totalacumulado > 0.8 * Total then
        begin
          bar[i] := TBarSeries.Create(self);
          chrtresultadopromethee.AddSeries(bar[i]);
          bar[i].AddXY(0, Total - totalacumulado, formatfloat
              ('0', Total - totalacumulado), coresalt[i]);
          // chrtResultadopromethee.BottomAxis.Items.Add(i, 'S' + inttostr(i + 1));

          chrtresultadopromethee.Legend.Item[i].Text := 'Other Alternatives';
          bar[i].Color := coresalt[i];
          break;
        end;
        if ResultAltPromethee[OrdemaltPromethee[j]] <> 0 then
        begin
          bar[i] := TBarSeries.Create(self);
          chrtresultadopromethee.AddSeries(bar[i]);
          bar[i].AddXY(0, ResultAltPromethee[OrdemaltPromethee[j]], formatfloat
              ('0', ResultAltPromethee[OrdemaltPromethee[j]]), coresalt[i]);
          // chrtResultadoPromethee.BottomAxis.Items.Add(i, 'S' + inttostr(i + 1));
          chrtresultadopromethee.Legend.Item[i].Text := nomealt
            [OrdemaltPromethee[j]];
          bar[i].Color := coresalt[i];
          totalacumulado := totalacumulado + ResultAltPromethee
            [OrdemaltPromethee[j]];
          inc(i);
        end;
      end;
      Total := 0;
      bar[1] := TBarSeries.Create(self);
      chrtresultadopromethee.AddSeries(bar[1]);
      for i := 0 to length(ResultSolPromethee) - 1 do
        Total := Total + ResultSolPromethee[i];
      bar[1].AddXY(0, Total, '', clWhite);
      bar[1].Visible := false;
      // chrtResultadoPromethee.LeftAxis.Increment := 0.2; // ?
      // chrtResultadoPromethee.Legend.Visible := false;
      chrtresultadopromethee.BottomAxis.LabelsAlternate;
      chrtresultadopromethee.BottomAxis.Items.Automatic := false;
      chrtresultadopromethee.LeftAxis.Maximum := Total;
      chrtresultadopromethee.LeftAxis.AutomaticMaximum := false;
      /// chrtResultadoPromethee.LeftAxis.MaximumOffset := 16;
      chrtresultadopromethee.LeftAxis.Minimum := 0;
      // chrtresultadoPromethee.leftaxis.Items.Automatic := false;
      chrtresultadopromethee.LeftAxis.Increment := Total;

    end
    else
    begin
      chrtresultadopromethee.RemoveAllSeries;
      chrtresultadopromethee.Legend.Items.Clear;
      // chrtresultadoPromethee.ClearChart;
      chrtresultadopromethee.BottomAxis.Items.Clear;
      chrtresultadopromethee.View3D := false;
      SetLength(bar, Numalt);
      i := 0;
      Total := length(CasesOrdemcrit);
      totalacumulado := 0;
      for j := 0 to Numalt - 1 do
      begin
        if totalacumulado > 0.8 * Total then
        begin
          bar[i] := TBarSeries.Create(self);
          chrtresultadopromethee.AddSeries(bar[i]);
          bar[i].AddXY(0, 100 * (Total - totalacumulado) / Total, formatfloat
              ('0', (100 * (Total - totalacumulado)) / Total), coresalt[i]);
          // chrtResultado.BottomAxis.Items.Add(i, 'S' + inttostr(i + 1));

          chrtresultadopromethee.Legend.Item[i].Text := 'Other Alternatives';
          bar[i].Color := coresalt[i];
          break;
        end;
        if ResultAltPromethee[OrdemaltPromethee[j]] <> 0 then
        begin
          bar[i] := TBarSeries.Create(self);
          chrtresultadopromethee.AddSeries(bar[i]);
          bar[i].AddXY(0, 100 * ResultAltPromethee[OrdemaltPromethee[j]]
              / Total, formatfloat('0', 100 * ResultAltPromethee
                [OrdemaltPromethee[j]] / Total), coresalt[i]);
          // chrtResultadoPromethee.BottomAxis.Items.Add(i, 'S' + inttostr(i + 1));
          chrtresultadopromethee.Legend.Item[i].Text := nomealt
            [OrdemaltPromethee[j]];
          bar[i].Color := coresalt[i];
          totalacumulado := totalacumulado + ResultAltPromethee
            [OrdemaltPromethee[j]];
          inc(i);
        end;
      end;
      Total := 0;
      bar[1] := TBarSeries.Create(self);
      chrtresultadopromethee.AddSeries(bar[1]);
      for i := 0 to length(ResultSolPromethee) - 1 do
        Total := Total + ResultSolPromethee[i];
      bar[1].AddXY(0, Total, '', clWhite);
      bar[1].Visible := false;
      // chrtResultadoPromethee.LeftAxis.Increment := 0.2; // ?
      // chrtResultadoPromethee.Legend.Visible := false;
      chrtresultadopromethee.BottomAxis.LabelsAlternate;
      chrtresultadopromethee.BottomAxis.Items.Automatic := false;
      chrtresultadopromethee.LeftAxis.Maximum := 100;
      chrtresultadopromethee.LeftAxis.AutomaticMaximum := false;
      /// chrtResultadoPromethee.LeftAxis.MaximumOffset := 16;
      chrtresultadopromethee.LeftAxis.Minimum := 0;
      // chrtresultadoPromethee.leftaxis.Items.Automatic := false;
      chrtresultadopromethee.LeftAxis.Increment := 100;

    end;
    imgResultadoPromethee.Picture.Bitmap.Assign
      (chrtresultadopromethee.TeeCreateBitmap(clWebPURPLE, rect(0, 0,
          imgResultadoPromethee.Width, imgResultadoPromethee.Height)));
  end;
end;

procedure TIWForm1.IWLink10Click(Sender: TObject);
begin
  rgnresultadopromethee.Visible := True;
  rgnAnaliseEstatisticaPromethee.Visible := false;
end;

procedure TIWForm1.IWLink11Click(Sender: TObject);
begin
  rgncomparacao.Visible := false;
  rgnresultado.Visible := True;
end;

procedure TIWForm1.IWLink12Click(Sender: TObject);
begin
  rgnregradedecisao.Visible := false;
  rgnAnaliseEstatistica.Visible := True;
end;

procedure TIWForm1.IWLink13Click(Sender: TObject);
begin
  rgnregradedecisao.Visible := True;
  rgnAnaliseEstatistica.Visible := false;
end;

procedure TIWForm1.IWLink14Click(Sender: TObject);
begin
  rgnmatcons.Visible := false;
  rgnOpcaoInput.Visible := True;
end;

procedure TIWForm1.IWLink15Click(Sender: TObject);
begin
  rgnAnaliseEstatisticaPromethee.Visible := false;
  rgnregradedecisaoPromethee.Visible := True;
end;

procedure TIWForm1.IWLink16Click(Sender: TObject);
begin
  rgnresultado.Visible := True;
  rgnresultadopromethee.Visible := false;
end;

procedure TIWForm1.IWLink17Click(Sender: TObject);
begin
rgnelicitacaoalt.Visible := false;
rgnelicitacaocrit.Visible := true;
end;

procedure TIWForm1.IWLink18Click(Sender: TObject);
begin
rgnelicitacaocrit.Visible := false;
rgnresultado.Visible := true;
end;

procedure TIWForm1.IWLink19Click(Sender: TObject);
begin
rgnelicitacaoalt.Visible := true;
rgnelicitacaocrit.Visible := false;
end;

procedure TIWForm1.IWLink1Click(Sender: TObject);
begin
  With TIWUserSession(WebApplication.Data) do
  begin

    rgnresultado.Visible := false;
    rgnresultadopromethee.Visible := false;
    rgnopcaoinput.Visible := True;
    novoproblema := false;
  end;
end;

procedure TIWForm1.IWLink20Click(Sender: TObject);
begin
rgncritxcrit.Visible := false;
rgnelicitacaocrit.Visible := true;
end;

procedure TIWForm1.IWLink22Click(Sender: TObject);
begin
rgnaltxalt.Visible := false;
rgnelicitacaoalt.visible := true;
end;

procedure TIWForm1.IWLink2Click(Sender: TObject);
Var
  i, j: Integer;
  Fatorial: real;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    cmbbxfiltro.Clear;
    for i := 0 to length(ResultSol) - 1 do
      cmbbxfiltro.Items.Add('Solution ' + inttostr(i + 1));

    cmbbxfiltro.ItemIndex := 0;
    cmbbxfiltroChange(self);

    rgnfiltro2.Visible := True;
    rgnresultado.Visible := false;
    if Numcrit = 7 then
    begin
      grdprimeiro.rowcount := 6;
      grdsegundo.rowcount := 6;
    end;
    if (Numcrit = 6) or (Numcrit = 5) then
    begin
      grdprimeiro.rowcount := 5;
      grdsegundo.rowcount := 5;
    end;
    if Numcrit = 4 then
    begin
      grdprimeiro.rowcount := 4;
      grdsegundo.rowcount := 4;
    end;
    if Numcrit = 3 then
    begin
      grdprimeiro.rowcount := 3;
      grdsegundo.rowcount := 3;
    end;

    grdprimeiro.ColumnCount := 1 + Numcrit;
    grdsegundo.ColumnCount := 1 + Numcrit;

    for j := 0 to grdprimeiro.ColumnCount - 1 do
      for i := 0 to grdprimeiro.rowcount - 1 do
      begin
        grdprimeiro.Cell[i, j].Alignment := taCenter;
        grdprimeiro.Cell[i, j].wrap := True;
        // grdresultados.Cell[0,j].Font.Style := [fsbold];
      end;
    for j := 0 to grdprimeiro.ColumnCount - 1 do
      grdprimeiro.Cell[0, j].Font.Style := [fsbold];

    for i := 0 to grdprimeiro.rowcount - 1 do
      grdprimeiro.Cell[i, 0].Font.Style := [fsbold];

    Fatorial := 1;
    for i := 1 to Numcrit - 1 do
      Fatorial := Fatorial * i;
    if Numcrit = 7 then
    begin
      for i := 0 to Numcrit - 1 do
      begin
        grdprimeiro.Cell[0, i + 1].Text := Nomecrit[i];
        grdprimeiro.Cell[1, i + 1].Text := formatfloat
          ('0', crit10[i] / Fatorial * 100) + ' %';
        grdprimeiro.Cell[2, i + 1].Text := formatfloat
          ('0', crit12[i] / Fatorial * 100) + ' %';
        grdprimeiro.Cell[3, i + 1].Text := formatfloat
          ('0', crit13[i] / Fatorial * 100) + ' %';
        grdprimeiro.Cell[4, i + 1].Text := formatfloat
          ('0', crit14[i] / Fatorial * 100) + ' %';
        grdprimeiro.Cell[5, i + 1].Text := formatfloat
          ('0', Fatorial / Fatorial * 100) + ' %';
      end;

      grdprimeiro.Cell[0, 0].Text := 'Criteria:';
      grdprimeiro.Cell[1, 0].Text := '1th:';
      grdprimeiro.Cell[2, 0].Text := '1th + 2th:';
      grdprimeiro.Cell[3, 0].Text := '1th + 2th + 3th:';
      grdprimeiro.Cell[4, 0].Text := '1th + 2th + 3th + 4th:';
      grdprimeiro.Cell[5, 0].Text := 'Others:';

    end;

    if (Numcrit = 6) or (Numcrit = 5) then
    begin
      for i := 0 to Numcrit - 1 do
      begin
        grdprimeiro.Cell[0, i + 1].Text := Nomecrit[i];
        grdprimeiro.Cell[1, i + 1].Text := formatfloat
          ('0', crit10[i] / Fatorial * 100) + ' %';
        grdprimeiro.Cell[2, i + 1].Text := formatfloat
          ('0', crit12[i] / Fatorial * 100) + ' %';
        grdprimeiro.Cell[3, i + 1].Text := formatfloat
          ('0', crit13[i] / Fatorial * 100) + ' %';
        // grdprimeiro.Cell[4, i+1].Text := formatfloat('0.00', crit14[i]/fatorial *100) + ' %';
        grdprimeiro.Cell[4, i + 1].Text := formatfloat
          ('0', Fatorial / Fatorial * 100) + ' %';
      end;

      grdprimeiro.Cell[0, 0].Text := 'Criteria:';
      grdprimeiro.Cell[1, 0].Text := '1th:';
      grdprimeiro.Cell[2, 0].Text := '1th + 2th:';
      grdprimeiro.Cell[3, 0].Text := '1th + 2th + 3th:';
      // grdprimeiro.Cell[4, 0].Text := '1th + 2th + 3th + 4th:';
      grdprimeiro.Cell[4, 0].Text := 'Others:';
    end;
    if (Numcrit = 4) then
    begin
      for i := 0 to Numcrit - 1 do
      begin
        grdprimeiro.Cell[0, i + 1].Text := Nomecrit[i];
        grdprimeiro.Cell[1, i + 1].Text := formatfloat
          ('0', crit10[i] / Fatorial * 100) + ' %';
        grdprimeiro.Cell[2, i + 1].Text := formatfloat
          ('0', crit12[i] / Fatorial * 100) + ' %';
        // grdprimeiro.Cell[3, i+1].Text := formatfloat('0.00', crit13[i]/fatorial *100) + ' %';
        // grdprimeiro.Cell[4, i+1].Text := formatfloat('0.00', crit14[i]/fatorial *100) + ' %';
        grdprimeiro.Cell[3, i + 1].Text := formatfloat
          ('0', Fatorial / Fatorial * 100) + ' %';
      end;

      grdprimeiro.Cell[0, 0].Text := 'Criteria:';
      grdprimeiro.Cell[1, 0].Text := '1th:';
      grdprimeiro.Cell[2, 0].Text := '1th + 2th:';
      // grdprimeiro.Cell[3, 0].Text := '1th + 2th + 3th:';
      // grdprimeiro.Cell[4, 0].Text := '1th + 2th + 3th + 4th:';
      grdprimeiro.Cell[3, 0].Text := 'Others:';
    end;
    if (Numcrit = 3) then
    begin
      for i := 0 to Numcrit - 1 do
      begin
        grdprimeiro.Cell[0, i + 1].Text := Nomecrit[i];
        grdprimeiro.Cell[1, i + 1].Text := formatfloat
          ('0', crit10[i] / Fatorial * 100) + ' %';
        // grdprimeiro.Cell[2, i+1].Text := formatfloat('0.00', crit12[i]/fatorial *100) + ' %';
        // grdprimeiro.Cell[3, i+1].Text := formatfloat('0.00', crit13[i]/fatorial *100) + ' %';
        // grdprimeiro.Cell[4, i+1].Text := formatfloat('0.00', crit14[i]/fatorial *100) + ' %';
        grdprimeiro.Cell[2, i + 1].Text := formatfloat
          ('0', Fatorial / Fatorial * 100) + ' %';
      end;

      grdprimeiro.Cell[0, 0].Text := 'Criteria:';
      grdprimeiro.Cell[1, 0].Text := '1th:';
      // grdprimeiro.Cell[2, 0].Text := '1th + 2th:';
      // grdprimeiro.Cell[3, 0].Text := '1th + 2th + 3th:';
      // grdprimeiro.Cell[4, 0].Text := '1th + 2th + 3th + 4th:';
      grdprimeiro.Cell[2, 0].Text := 'Others:';
    end;
    /// ///////////////////////////////////////////////////////////////////

    for j := 0 to grdsegundo.ColumnCount - 1 do
      for i := 0 to grdsegundo.rowcount - 1 do
      begin
        grdsegundo.Cell[i, j].Alignment := taCenter;
        grdsegundo.Cell[i, j].wrap := True;
        // grdresultados.Cell[0,j].Font.Style := [fsbold];
      end;
    for j := 0 to grdsegundo.ColumnCount - 1 do
      grdsegundo.Cell[0, j].Font.Style := [fsbold];

    for i := 0 to grdsegundo.rowcount - 1 do
      grdsegundo.Cell[i, 0].Font.Style := [fsbold];

    Fatorial := 1;
    for i := 1 to Numcrit - 1 do
      Fatorial := Fatorial * i;

    if Numcrit = 7 then
    begin
      for i := 0 to Numcrit - 1 do
      begin
        grdsegundo.Cell[0, i + 1].Text := Nomecrit[i];
        grdsegundo.Cell[1, i + 1].Text := formatfloat
          ('0', crit20[i] / Fatorial * 100) + ' %';
        grdsegundo.Cell[2, i + 1].Text := formatfloat
          ('0', crit21[i] / Fatorial * 100) + ' %';
        grdsegundo.Cell[3, i + 1].Text := formatfloat
          ('0', crit23[i] / Fatorial * 100) + ' %';
        grdsegundo.Cell[4, i + 1].Text := formatfloat
          ('0', crit24[i] / Fatorial * 100) + ' %';
        grdsegundo.Cell[5, i + 1].Text := formatfloat
          ('0', Fatorial / Fatorial * 100) + ' %';
      end;

      grdsegundo.Cell[0, 0].Text := 'Criteria:';
      grdsegundo.Cell[1, 0].Text := '2th:';
      grdsegundo.Cell[2, 0].Text := '2th + 1th:';
      grdsegundo.Cell[3, 0].Text := '2th + 1th + 3th:';
      grdsegundo.Cell[4, 0].Text := '2th + 1th + 3th + 4th:';
      grdsegundo.Cell[5, 0].Text := 'Others:';
    end;
    if (Numcrit = 6) or (Numcrit = 5) then
    begin
      for i := 0 to Numcrit - 1 do
      begin
        grdsegundo.Cell[0, i + 1].Text := Nomecrit[i];
        grdsegundo.Cell[1, i + 1].Text := formatfloat
          ('0', crit20[i] / Fatorial * 100) + ' %';
        grdsegundo.Cell[2, i + 1].Text := formatfloat
          ('0', crit21[i] / Fatorial * 100) + ' %';
        grdsegundo.Cell[3, i + 1].Text := formatfloat
          ('0', crit23[i] / Fatorial * 100) + ' %';
        // grdsegundo.Cell[4, i+1].Text := formatfloat('0.00', crit24[i]/fatorial *100) + ' %';
        grdsegundo.Cell[4, i + 1].Text := formatfloat
          ('0', Fatorial / Fatorial * 100) + ' %';
      end;

      grdsegundo.Cell[0, 0].Text := 'Criteria:';
      grdsegundo.Cell[1, 0].Text := '2th:';
      grdsegundo.Cell[2, 0].Text := '2th + 1th:';
      grdsegundo.Cell[3, 0].Text := '2th + 1th + 3th:';
      // grdsegundo.Cell[4, 0].Text := '2th + 1th + 3th + 4th:';
      grdsegundo.Cell[4, 0].Text := 'Others:';
    end;
    if (Numcrit = 4) then
    begin
      for i := 0 to Numcrit - 1 do
      begin
        grdsegundo.Cell[0, i + 1].Text := Nomecrit[i];
        grdsegundo.Cell[1, i + 1].Text := formatfloat
          ('0', crit20[i] / Fatorial * 100) + ' %';
        grdsegundo.Cell[2, i + 1].Text := formatfloat
          ('0', crit21[i] / Fatorial * 100) + ' %';
        // grdsegundo.Cell[3, i+1].Text := formatfloat('0.00', crit23[i]/fatorial *100) + ' %';
        // grdsegundo.Cell[4, i+1].Text := formatfloat('0.00', crit24[i]/fatorial *100) + ' %';
        grdsegundo.Cell[3, i + 1].Text := formatfloat
          ('0', Fatorial / Fatorial * 100) + ' %';
      end;

      grdsegundo.Cell[0, 0].Text := 'Criteria:';
      grdsegundo.Cell[1, 0].Text := '2th:';
      grdsegundo.Cell[2, 0].Text := '2th + 1th:';
      // grdsegundo.Cell[3, 0].Text := '2th + 1th + 3th:';
      // grdsegundo.Cell[4, 0].Text := '2th + 1th + 3th + 4th:';
      grdsegundo.Cell[3, 0].Text := 'Others:';
    end;
    if (Numcrit = 3) then
    begin
      for i := 0 to Numcrit - 1 do
      begin
        grdsegundo.Cell[0, i + 1].Text := Nomecrit[i];
        grdsegundo.Cell[1, i + 1].Text := formatfloat
          ('0', crit20[i] / Fatorial * 100) + ' %';
        // grdsegundo.Cell[2, i+1].Text := formatfloat('0.00', crit21[i]/fatorial *100) + ' %';
        // grdsegundo.Cell[3, i+1].Text := formatfloat('0.00', crit23[i]/fatorial *100) + ' %';
        // grdsegundo.Cell[4, i+1].Text := formatfloat('0.00', crit24[i]/fatorial *100) + ' %';
        grdsegundo.Cell[2, i + 1].Text := formatfloat
          ('0', Fatorial / Fatorial * 100) + ' %';
      end;

      grdsegundo.Cell[0, 0].Text := 'Criteria:';
      grdsegundo.Cell[1, 0].Text := '2th:';
      // grdsegundo.Cell[2, 0].Text := '2th + 1th:';
      // grdsegundo.Cell[3, 0].Text := '2th + 1th + 3th:';
      // grdsegundo.Cell[4, 0].Text := '2th + 1th + 3th + 4th:';
      grdsegundo.Cell[2, 0].Text := 'Others:';
    end;

  end;

end;

procedure TIWForm1.IWLink3Click(Sender: TObject);
begin
  rgnfiltro.Visible := false;
  rgnresultado.Visible := True;
end;

procedure TIWForm1.IWLink4Click(Sender: TObject);
begin
  rgnfiltro2.Visible := false;
  rgnresultado.Visible := True;
end;

procedure TIWForm1.IWLink5Click(Sender: TObject);
begin
  rgnresultado.Visible := True;
  rgnAnaliseEstatistica.Visible := false;
end;

procedure TIWForm1.IWLink6Click(Sender: TObject);
begin
  rgnregradedecisaoPromethee.Visible := false;
  rgnAnaliseEstatisticaPromethee.Visible := True;
end;

procedure TIWForm1.IWLink7Click(Sender: TObject);
begin
  rgnOpcaoInput.Visible := false;
  rgnEnterLogin.Visible := True;
end;

procedure TIWForm1.IWLink8Click(Sender: TObject);
begin
  rgnOpcaoInput.Visible := True;
  rgnLoginproject.Visible := false;
end;

procedure TIWForm1.lnkregrasdedecisaoPrometheeClick(Sender: TObject);
begin
  if grdperdamediaemaximapromethee.Visible = True then
  begin
    lnkregrasdedecisaoPromethee.Caption :=
      'Click to view others statistical indexes';
    grdperdamediaemaximapromethee.Visible := false;
    grdregradedecisaopromethee.Visible := True;
    exit;
  end;
  if grdperdamediaemaximapromethee.Visible = false then
  begin
    lnkregrasdedecisaoPromethee.Caption :=
      'Click to view decision rules indexes';
    grdperdamediaemaximapromethee.Visible := True;
    grdregradedecisaopromethee.Visible := false;
    exit;
  end;
end;

procedure TIWForm1.IWLink9Click(Sender: TObject);
begin
  rgnresultado.Visible := false;
  rgnresultadopromethee.Visible := True;
end;

procedure TIWForm1.lnkbackloginClick(Sender: TObject);
begin
  rgnEnterLogin.Visible := false;
  rgnintro.Visible := True;
end;

procedure TIWForm1.lnkdescricaolimiarClick(Sender: TObject);
begin
  if rgndescricao.Visible = True then
  begin
    lnkdescricaolimiar.Caption :=
      'Click here to view the decision rules thresholds';
    rgndescricao.Visible := false;
    rgnlimiares.Visible := True;
    exit;
  end;
  if rgnlimiares.Visible = True then
  begin
    lnkdescricaolimiar.Caption := 'Click to view the conceptual explanations';
    rgndescricao.Visible := True;
    rgnlimiares.Visible := false;
    exit;
  end;
end;

procedure TIWForm1.lnkdescricaolimiarPrometheeClick(Sender: TObject);
begin
  if rgndescricaoPromethee.Visible = True then
  begin
    lnkdescricaolimiarPromethee.Caption :=
      'Click here to view the decision rules thresholds';
    rgndescricaoPromethee.Visible := false;
    rgnlimiaresPromethee.Visible := True;
    exit;
  end;
  if rgnlimiaresPromethee.Visible = True then
  begin
    lnkdescricaolimiarPromethee.Caption :=
      'Click to view the conceptual explanations';
    rgndescricaoPromethee.Visible := True;
    rgnlimiaresPromethee.Visible := false;
    exit;
  end;
end;

procedure TIWForm1.lnkesqueceusenhaClick(Sender: TObject);
begin
  iwrgnCaixaDialogoRU2.Visible := True;
end;

procedure TIWForm1.lnkestatisticaClick(Sender: TObject);
begin
  rgnresultado.Visible := false;
  rgnAnaliseEstatistica.Visible := True;
end;

procedure TIWForm1.lnkestatisticaPrometheeClick(Sender: TObject);
begin
  rgnresultadopromethee.Visible := false;
  rgnAnaliseEstatisticaPromethee.Visible := True;
end;

procedure TIWForm1.lnkexportacaoprometheeClick(Sender: TObject);
var
  i, j, k: Integer;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    grdexportacao.Columns.Clear;
    grdexportacao.totalrows := 9 + grdestatisticaPromethee.rowcount +
      grdmaxminPromethee.rowcount + grdresultadospromethee.rowcount;
    if grdestatisticaPromethee.ColumnCount >=
      grdmaxminPromethee.ColumnCount then
    begin
      for i := 0 to grdestatisticaPromethee.ColumnCount + 1 do
        grdexportacao.Columns.Add;
    end
    else
    begin
      for i := 0 to grdmaxminPromethee.ColumnCount + 1 do
        grdexportacao.Columns.Add;
    end;

    grdexportacao.Cells[0, 0] := 'Outranking Model - Statistical Analysis';
    grdexportacao.Cells[0, 2] := 'Legend of Solutions:';
    k := 3;
    for i := 0 to grdresultadospromethee.rowcount - 1 do
      for j := 0 to grdresultadospromethee.ColumnCount - 1 do
      begin
        grdexportacao.Cells[j, i + k + 1] := grdresultadospromethee.Cell[i, j]
          .Text;
      end;
    k := k + 1 + grdresultadospromethee.rowcount;

    grdexportacao.Cells[0, k + 1] :=
      'Mean Performance Difference between Solutions:';
    k := k + 2;

    // grdexportacao.Cells[0,k] := 'Consequences Variation percentage values:';
    for i := 0 to grdestatisticaPromethee.rowcount - 1 do
      for j := 0 to grdestatisticaPromethee.ColumnCount - 1 do
      begin
        grdexportacao.Cells[j, i + k + 1] := grdestatisticaPromethee.Cell[i, j]
          .Text;
      end;
    k := k + 1 + grdestatisticaPromethee.rowcount;
    grdexportacao.Cells[0, k + 1] :=
      'Maximum and Minimum Performance Difference between Solutions:';
    k := k + 1;
    // grdexportacao.TotalRows := K + 1 + grdcrivarWeight.ColumnCount;
    // grdexportacao.Cells[0,k] := 'Weights Variation percentage values:';
    for i := 0 to grdmaxminPromethee.rowcount - 1 do
      for j := 0 to grdmaxminPromethee.ColumnCount - 1 do
      begin
        grdexportacao.Cells[j, i + k + 1] := grdmaxminPromethee.Cell[i, j].Text;
      end;
    k := k + 2 + grdmaxminPromethee.rowcount;

    TIWAdvWebGridExcelIO1.AdvWebGrid := grdexportacao;
    TIWAdvWebGridExcelIO1.XLSExport
      ('Outranking Model - Statistical Analysis' + '.xls', WebApplication);

  end;
end;

procedure TIWForm1.lnkExportarClick(Sender: TObject);
var
  i, j, k: Integer;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    grdexportacao.Columns.Clear;
    grdexportacao.totalrows := 9 + grdestatistica.rowcount +
      grdmaxmin.rowcount + grdresultados.rowcount;
    if grdestatistica.ColumnCount >= grdmaxmin.ColumnCount then
    begin
      for i := 0 to grdestatistica.ColumnCount + 1 do
        grdexportacao.Columns.Add;
    end
    else
    begin
      for i := 0 to grdmaxmin.ColumnCount + 1 do
        grdexportacao.Columns.Add;
    end;

    grdexportacao.Cells[0, 0] := 'Additive Model - Statistical Analysis';
    grdexportacao.Cells[0, 2] := 'Legend of Solutions:';
    k := 3;
    for i := 0 to grdresultados.rowcount - 1 do
      for j := 0 to grdresultados.ColumnCount - 1 do
      begin
        grdexportacao.Cells[j, i + k + 1] := grdresultados.Cell[i, j].Text;
      end;
    k := k + 1 + grdresultados.rowcount;

    grdexportacao.Cells[0, k + 1] :=
      'Mean Performance Difference between Solutions:';
    k := k + 2;

    // grdexportacao.Cells[0,k] := 'Consequences Variation percentage values:';
    for i := 0 to grdestatistica.rowcount - 1 do
      for j := 0 to grdestatistica.ColumnCount - 1 do
      begin
        grdexportacao.Cells[j, i + k + 1] := grdestatistica.Cell[i, j].Text;
      end;
    k := k + 1 + grdestatistica.rowcount;
    grdexportacao.Cells[0, k + 1] :=
      'Maximum and Minimum Performance Difference between Solutions:';
    k := k + 1;
    // grdexportacao.TotalRows := K + 1 + grdcrivarWeight.ColumnCount;
    // grdexportacao.Cells[0,k] := 'Weights Variation percentage values:';
    for i := 0 to grdmaxmin.rowcount - 1 do
      for j := 0 to grdmaxmin.ColumnCount - 1 do
      begin
        grdexportacao.Cells[j, i + k + 1] := grdmaxmin.Cell[i, j].Text;
      end;
    k := k + 2 + grdmaxmin.rowcount;

    TIWAdvWebGridExcelIO1.AdvWebGrid := grdexportacao;
    TIWAdvWebGridExcelIO1.XLSExport
      ('Additive Model - Statistical Analysis' + '.xls', WebApplication);

  end;
end;

procedure TIWForm1.lnkregrasdedecisaoClick(Sender: TObject);
begin
  if grdperdamediaemaxima.Visible = True then
  begin
    lnkregrasdedecisao.Caption := 'Click to view others statistical indexes';
    grdperdamediaemaxima.Visible := false;
    grdregradedecisao.Visible := True;
    exit;
  end;
  if grdperdamediaemaxima.Visible = false then
  begin
    lnkregrasdedecisao.Caption := 'Click to view decision rules indexes';
    grdperdamediaemaxima.Visible := True;
    grdregradedecisao.Visible := false;
    exit;
  end;
end;

procedure TIWForm1.SetarCores;
Var
  i, j: Integer;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    SetLength(coresalt, 130);
    SetLength(coresalt1, TIWUserSession(WebApplication.Data).Numalt);
    coresalt[0] := $000000CC; // vermelho
    coresalt[1] := $0033BBFF; // amarelo
    coresalt[2] := $00CC9900; // azul
    coresalt[3] := $00BBBBFF; // rosa
    coresalt[4] := $005AE6C8; // verde
    coresalt[5] := $00FF00FF;
    coresalt[6] := $00808000;
    coresalt[7] := $00800000;
    coresalt[8] := $00000080;
    coresalt[9] := $0000FF00;
    coresalt[10] := $00008080;
    coresalt[11] := $00800080;
    coresalt[12] := $00FFFF00;
    coresalt[13] := $00C0DCC0;
    coresalt[14] := $00A4A0A0;
    coresalt[15] := $00E049BB;
    coresalt[16] := $005353FF;
    coresalt[17] := $000080FF;
    coresalt[18] := $00800040;
    coresalt[19] := $00434321;
    coresalt[20] := $00189AD8;
    coresalt[21] := $000000C1;
    coresalt[22] := $00DC0E65;
    coresalt[23] := $0000AE00;
    coresalt[24] := $002676CE;
    coresalt[25] := $008000FF; // índigo
    coresalt[26] := $00808080; // cinza claro
    coresalt[27] := $00FF4500; // laranja avermelhado
    coresalt[28] := $00FFD700; // ouro claro
    coresalt[29] := $00640000; // marrom escuro
    coresalt[30] := $008B4513; // marrom de sela
    coresalt[31] := $00A52A2A; // marrom de fogo
    coresalt[32] := $00D2691E; // chocolate
    coresalt[33] := $008BB000; // marrom escuro
    coresalt[34] := $00B22222; // tijolo
    coresalt[35] := $00FF6347; // tomate
    coresalt[36] := $00FF8C00; // laranja escuro
    coresalt[37] := $00FFA07A; // salmão claro
    coresalt[38] := $00FF7F50; // coral claro
    coresalt[39] := $00FF0000; // vermelho claro
    coresalt[40] := $00CD5C5C; // marrom claro
    coresalt[41] := $008B0000; // marrom escuro
    coresalt[42] := $00A52A2A; // marrom de fogo
    coresalt[43] := $00B22222; // tijolo
    coresalt[44] := $008B0000; // marrom escuro
    coresalt[45] := $00A52A2A; // marrom de fogo
    coresalt[46] := $00B22222; // tijolo
    coresalt[47] := $008B0000; // marrom escuro
    coresalt[48] := $00A52A2A; // marrom de fogo
    coresalt[49] := $00B22222; // tijolo
    coresalt[50] := $008B0000; // marrom escuro
    coresalt[51] := $00A52A2A; // marrom de fogo
    coresalt[52] := $00B22222; // tijolo
    coresalt[53] := $008B0000; // marrom escuro
    coresalt[54] := $00A52A2A; // marrom de fogo
    coresalt[55] := $00B22222; // tijolo
    coresalt[56] := $00FF69B4; // rosa quente
    coresalt[57] := $00FFE4E1; // rosa claro
    coresalt[58] := $00FF1493; // rosa profundo
    coresalt[59] := $00C71585; // vermelho violeta
    coresalt[60] := $00DB7093; // roxo claro
    coresalt[61] := $00DA70D6; // orquídea
    coresalt[62] := $009932CC; // índigo escuro
    coresalt[63] := $009ACD32; // verde amarelo
    coresalt[64] := $00C0FF3E; // verde limão brilhante
    coresalt[65] := $00AEEEEE; // azul celeste
    coresalt[66] := $000000CC; // vermelho
    coresalt[67] := $0033BBFF; // amarelo
    coresalt[68] := $00CC9900; // azul
    coresalt[69] := $00BBBBFF; // rosa
    coresalt[70] := $005AE6C8; // verde
    coresalt[71] := $00FF00FF;
    coresalt[72] := $00808000;
    coresalt[73] := $00800000;
    coresalt[74] := $00000080;
    coresalt[75] := $0000FF00;
    coresalt[76] := $00008080;
    coresalt[77] := $00800080;
    coresalt[78] := $00FFFF00;
    coresalt[79] := $00C0DCC0;
    coresalt[80] := $00A4A0A0;
    coresalt[81] := $00E049BB;
    coresalt[82] := $005353FF;
    coresalt[83] := $000080FF;
    coresalt[84] := $00800040;
    coresalt[85] := $00434321;
    coresalt[86] := $00189AD8;
    coresalt[87] := $000000C1;
    coresalt[88] := $00DC0E65;
    coresalt[89] := $0000AE00;
    coresalt[90] := $002676CE;
    coresalt[91] := $008000FF; // índigo
    coresalt[92] := $00808080; // cinza claro
    coresalt[93] := $00FF4500; // laranja avermelhado
    coresalt[94] := $00FFD700; // ouro claro
    coresalt[95] := $00640000; // marrom escuro
    coresalt[96] := $008B4513; // marrom de sela
    coresalt[97] := $00A52A2A; // marrom de fogo
    coresalt[98] := $00D2691E; // chocolate
    coresalt[99] := $008B0000; // marrom escuro
    coresalt[100] := $00B22222; // tijolo
    coresalt[101] := $00FF6347; // tomate
    coresalt[102] := $00FF8C00; // laranja escuro
    coresalt[103] := $00FFA07A; // salmão claro
    coresalt[104] := $00FF7F50; // coral claro
    coresalt[105] := $00FF0000; // vermelho claro
    coresalt[106] := $00CD5C5C; // marrom claro
    coresalt[107] := $008B0000; // marrom escuro
    coresalt[108] := $00A52A2A; // marrom de fogo
    coresalt[109] := $00B22222; // tijolo
    coresalt[110] := $008B0000; // marrom escuro
    coresalt[111] := $00A52A2A; // marrom de fogo
    coresalt[112] := $00B22222; // tijolo
    coresalt[113] := $008B0000; // marrom escuro
    coresalt[114] := $00A52A2A; // marrom de fogo
    coresalt[115] := $00B22222; // tijolo
    coresalt[116] := $008B0000; // marrom escuro
    coresalt[117] := $00A52A2A; // marrom de fogo
    coresalt[118] := $00B22222; // tijolo
    coresalt[119] := $008B0000; // marrom escuro
    coresalt[120] := $00A52A2A; // marrom de fogo
    coresalt[121] := $00B22222; // tijolo
    coresalt[122] := $00FF69B4; // rosa quente
    coresalt[123] := $00FFE4E1; // rosa claro
    coresalt[124] := $00FF1493; // rosa profundo
    coresalt[125] := $00C71585; // vermelho violeta
    coresalt[126] := $00DB7093; // roxo claro
    coresalt[127] := $00DA70D6; // orquídea
    coresalt[128] := $009932CC; // índigo escuro
    coresalt[129] := $009ACD32; // verde amarelo
  end;

end;

procedure TIWForm1.btnestatisticaClick(Sender: TObject);
begin
  rgnresultado.Visible := false;
  rgnAnaliseEstatistica.Visible := True;
end;

procedure TIWForm1.btnimportClick(Sender: TObject);
var
  i, j: Integer;
  cred: boolean;
  vetchar: array [0 .. 9] of string;
  Linha, coluna, ncrit, nalt: Integer;
  tipoarquivo: string;
begin

  if (FileArquivo.Filename = '') then
  begin
    if TIWUserSession(WebApplication.Data).excelverif = false then
      WebApplication.showmessage('You must choose an excel file');
    exit;
  end
  else if FileArquivo.Filename <> '' then
  begin
    { vetchar[0] := '!';
      vetchar[1] := '@';
      vetchar[2] := '#';
      vetchar[3] := '$';
      vetchar[4] := '%';
      vetchar[5] := '&';
      vetchar[6] := '*';
      vetchar[7] := '^';
      vetchar[8] := '-';
      vetchar[9] := '+'; }
    cred := false;
    { for i := 0 to 9 do
      for j := 0 to length(edteditarnome.Text) do
      begin
      if vetchar[i] = edteditarnome.Text[j] then
      cred := true;
      end; }

    if cred = True then
    begin
      WebApplication.showmessage(
        'Special characters are not allowed. You must choose a different name for your problem.');
      exit;
    end
    else
    begin

      tipoarquivo := FileArquivo.ContentType;
      if (tipoarquivo = ' application/vnd.ms-excel') then
      // verifica se o arquivo é compatível
      begin
        FileArquivo.SaveToFile('C:\Users\Public\Documents' +
            (FileArquivo.Filename));
        TIWAdvWebGridExcelIO1.XLSImport('C:\Users\Public\Documents' +
            (FileArquivo.Filename));

        TIWAdvWebGridExcelIO1.AdvWebGrid := grdImport;
      end
      else
      begin
        WebApplication.showmessage('Select an excel file to continue!');
        exit;
      end;

      ncrit := grdImport.Columns.Count;
      nalt := grdImport.totalrows;

      for i := ncrit - 1 downto 1 do
      begin
        if grdImport.Cells[i, 0] = '' then
          grdImport.Columns.delete(i);
      end;
      coluna := 0;
      for i := 1 to grdImport.Columns.Count - 1 do
      begin
        if grdImport.Cells[i, 0] <> '' then
          inc(coluna);
      end;
      Linha := 8;
      for j := 8 to grdImport.totalrows - 1 do
      begin
        if grdImport.Cells[0, j] <> '' then
          inc(Linha);
      end;

      if coluna = 3 then
        cmbbxcriteria.ItemIndex := 0;
      if coluna = 4 then
        cmbbxcriteria.ItemIndex := 1;
      if coluna = 5 then
        cmbbxcriteria.ItemIndex := 2;
      if coluna = 6 then
        cmbbxcriteria.ItemIndex := 3;
      if coluna >= 7 then
        cmbbxcriteria.ItemIndex := 4;

      edtnumalt.Text := inttostr(Linha - 8);
      btnupdateClick(self);

      With TIWUserSession(WebApplication.Data) do
      begin
        for i := 0 to TIWUserSession(WebApplication.Data).Numcrit - 1 do
        begin
          if grdImport.Cells[i + 1, 1] = '0' then
            cmbbxTipoCrit[i].ItemIndex := 0;
          if grdImport.Cells[i + 1, 1] = '1' then
            cmbbxTipoCrit[i].ItemIndex := 1;
          if grdImport.Cells[i + 1, 1] = '2' then
            cmbbxTipoCrit[i].ItemIndex := 2;
          if grdImport.Cells[i + 1, 1] = '3' then
            cmbbxTipoCrit[i].ItemIndex := 3;
          if grdImport.Cells[i + 1, 1] = '4' then
            cmbbxTipoCrit[i].ItemIndex := 4;
          if grdImport.Cells[i + 1, 1] = '5' then
            cmbbxTipoCrit[i].ItemIndex := 5;

          cmbbxauxiliarChange(self);

          if (TIWUserSession(WebApplication.Data).cmbbxTipoCrit[i]
              .ItemIndex = 2) or (cmbbxTipoCrit[i].ItemIndex = 3) then
          begin
            if grdImport.Cells[i + 1, 6] = '2' then
              cmbbxNiveisCrit[i].ItemIndex := 0;
            if grdImport.Cells[i + 1, 6] = '3' then
              cmbbxNiveisCrit[i].ItemIndex := 1;
            if grdImport.Cells[i + 1, 6] = '4' then
              cmbbxNiveisCrit[i].ItemIndex := 2;
            if grdImport.Cells[i + 1, 6] = '5' then
              cmbbxNiveisCrit[i].ItemIndex := 3;
            if grdImport.Cells[i + 1, 6] = '6' then
              cmbbxNiveisCrit[i].ItemIndex := 4;
            if grdImport.Cells[i + 1, 6] = '7' then
              cmbbxNiveisCrit[i].ItemIndex := 5;
          end;

          cmbbxNiveisCrit[i].Enabled := false;
          for j := 0 to Numalt - 1 do
          begin
            edtMatrizConseq[j, i].Text := grdImport.Cells[1 + i, 8 + j];
            edtMatrizConseq[j, i].Enabled := false;
          end;
        end;
        for i := 0 to Numcrit - 1 do
        begin
          edtNomeCrit[i].Text := grdImport.Cells[i + 1, 0];
          Nomecrit[i] := grdImport.Cells[i + 1, 0];

        end;
        for j := 0 to Numalt - 1 do
        begin
          edtnomealt[j].Text := grdImport.Cells[0, j + 8];
          nomealt[j] := grdImport.Cells[0, j + 8];
        end;
      end;

    end;
    TIWUserSession(WebApplication.Data).excelverif := True;
  end;
end;

procedure TIWForm1.btnLoginClick(Sender: TObject);
begin
  rgnintro.Visible := false;
  rgnEnterLogin.Visible := True;
end;

procedure TIWForm1.btnRegisterClick(Sender: TObject);
begin
  iwrgnCaixaDialogoRU.Visible := True;
end;

procedure TIWForm1.btnRegisternewClick(Sender: TObject);
var
  i, j: Integer;
begin
  rgnOpcaoInput.Visible := false;
  rgnmatcons.Visible := True;

  LimpaVariaveis;
  edtnumalt.Text := '';
  cmbbxcriteria.ItemIndex := -1;

  edtnameproblem.Clear;
  btnshowresults.Enabled := false;
  With TIWUserSession(WebApplication.Data) do
  begin
    novoproblema := True;
    retomada := false;
    // Numalt := strtoint(edtnumalt.Text);
    // Numcrit := strtoint(cmbbxcriteria.Items[cmbbxcriteria.ItemIndex]);
    if length(edtNomeCrit) <> 0 then
    begin
      for i := 0 to length(edtNomeCrit) - 1 do
        edtNomeCrit[i].Destroy;
      Finalize(edtNomeCrit);
    end;
    if length(edtnomealt) <> 0 then
    begin
      for i := 0 to length(edtnomealt) - 1 do
        edtnomealt[i].Destroy;
      Finalize(edtnomealt);
    end;
    if length(cmbbxTipoCrit) <> 0 then
    begin
      for i := 0 to length(cmbbxTipoCrit) - 1 do
        cmbbxTipoCrit[i].Destroy;
      Finalize(cmbbxTipoCrit);
    end;
    if length(cmbbxNiveisCrit) <> 0 then
    begin
      for i := 0 to length(cmbbxNiveisCrit) - 1 do
        cmbbxNiveisCrit[i].Destroy;
      Finalize(cmbbxNiveisCrit);
    end;
    if length(edtMatrizConseq) <> 0 then
    begin
      for i := 0 to length(edtMatrizConseq) - 1 do
        for j := 0 to length(edtMatrizConseq[i]) - 1 do
          edtMatrizConseq[i, j].Destroy;
      Finalize(edtMatrizConseq);
    end;
  end;

end;

procedure TIWForm1.btnshowresultsClick(Sender: TObject);
var
  i, j, k, qtdpareto: Integer;
  auxregra, aux1, aux2, aux3: real;
  bar: array of TBarSeries;
  SolRegra1, SolRegra2, SolRegra3: Integer;
  epsilon, omega: array of real;
  totalacumulado: real;
  recomendacao: boolean;
  Px: real;
  faixaprob: array of real;
  num : double;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    if edtnameproblem.Text = '' then // verifica se foi inserido um nome pro arquivo
    begin // se não tiver nome do problema
      WebApplication.showmessage('Please, enter a problem''s name!');
      exit;
    end; // se não tiver nome do problema
    for i := 0 to Numcrit - 1 do
    begin
    if cmbbxTipoCrit[i].ItemIndex = -1 then
    begin
    WebApplication.ShowMessage('Fill the fields correctly!');
    exit;
    end;
    if ((cmbbxTipoCrit[i].ItemIndex = 2) or (cmbbxTipoCrit[i].ItemIndex = 3)) and (cmbbxNiveisCrit[i].ItemIndex = -1) then
    begin
    WebApplication.ShowMessage('Fill the fields correctly!');
    exit;
    end;
    {for j := 0 to numalt - 1 do
    begin
    if TryStrToFloat(edtMatrizConseq[j, i].Text, num) = false then
    begin
    WebApplication.ShowMessage('Fill the fields correctly!');
    exit;
    end;
    end;}
    end;
    LimpaVariaveis;
    GerarCases;
    NameProblem := edtnameproblem.Text;
    for i := 0 to Numcrit - 1 do
    begin
      TIWUserSession(WebApplication.Data).tipocrit[i] := cmbbxTipoCrit[i]
        .ItemIndex;
      if (tipocrit[i] = 2) or (tipocrit[i] = 3) then
        TIWUserSession(WebApplication.Data).niveis[i] := StrToInt
          (cmbbxNiveisCrit[i].Items[cmbbxNiveisCrit[i].ItemIndex]);
    end;
    for j := 0 to Numalt - 1 do
      for i := 0 to Numcrit - 1 do
        begin
        TIWUserSession(WebApplication.Data).MatrizConseq[j, i] := trocapv
          (edtMatrizConseq[j, i].Text);
        if (tipocrit[i] = 2) or (tipocrit[i] = 3) then
        begin
          if (niveis[i] = 2) and (MatrizConseq[j, i] <> 0) and (MatrizConseq[j, i] <> 1) then
          begin
          WebApplication.ShowMessage('Fill the fields correctly!');
            exit;
          end;
          if (niveis[i] > 2) and ((MatrizConseq[j, i] < 1) or (MatrizConseq[j, i] > niveis[i])) then
          begin
          WebApplication.ShowMessage('Fill the fields correctly!');
            exit;
          end;
        end;
        end;

 {    for i := 0 to Numcrit - 1 do
      begin
      if (tipocrit[i] <> 2) and (tipocrit[i] <> 3) then
      begin
      Max[i] := -9999999999;
      Min[i] := 9999999999;
      for j := 0 to Numalt - 1 do
      begin
      if Max[i] < MatrizConseq[j, i] then
      Max[i] := MatrizConseq[j, i];
      if Min[i] > MatrizConseq[j, i] then
      Min[i] := MatrizConseq[j, i];
      end; // j
      end // if
      else
      begin
      if niveis[i] = 2 then
      begin
      Max[i] := 1;
      Min[i] := 0;
      end;
      if niveis[i] > 2 then
      begin
      TIWUserSession(WebApplication.Data).Max[i] := niveis[i];
      TIWUserSession(WebApplication.Data).Min[i] := 1;
      end;
      end;
      end; // i
      for j := 0 to Numalt - 1 do
      for i := 0 to Numcrit - 1 do
      begin
      if (tipocrit[i] = 1) or (tipocrit[i] = 3) or (tipocrit[i] = 5) then
      MatrizConseqNorm[j, i] := (MatrizConseq[j, i] - Min[i]) /
      (Max[i] - Min[i]);
      /// Normalização para os critérios de minimização //////////////////////
      if (tipocrit[i] = 0) or (tipocrit[i] = 2) or (tipocrit[i] = 4) then
      TIWUserSession(WebApplication.Data).MatrizConseqNorm[j, i] :=
      (MatrizConseq[j, i] - Max[i]) / (-Max[i] + Min[i]);
      end; }
    EscalaRazao;
    Surrogate;
    TIWUserSession(WebApplication.Data).problemafinalizado := false;
    SetLength(TIWUserSession(WebApplication.Data).combinacaocrit, Numcrit,
      Numcrit);
    SetLength(TIWUserSession(WebApplication.Data).combinacaoalt, Numalt,
      Numalt);

    setlength(pularcomparacao, numcrit, numcrit);
    for i := 0 to numcrit - 1 do
    for j := 0 to numcrit - 1 do
    pularcomparacao[i,j] := false;


    altx := Ordemalt[0];
    altz := Ordemalt[1];

    Naltsel := 0;
    for i := 0 to length(ResultAlt) - 1 do
    begin
      if ResultAlt[i] <> 0 then
        inc(Naltsel);
    end;

    //while TIWUserSession(WebApplication.Data).problemafinalizado = false do
    //begin
      AnaliseParaElicitacao;
   // end;

   // Comparação dos critérios para o Promethee
   TIWUserSession(WebApplication.Data).problemafinalizadoPromethee := false;
    SetLength(TIWUserSession(WebApplication.Data).combinacaocritPromethee, Numcrit,
      Numcrit);
    SetLength(TIWUserSession(WebApplication.Data).combinacaoaltPromethee, Numalt,
      Numalt);

    setlength(pularcomparacaoPromethee, numcrit, numcrit);
    for i := 0 to numcrit - 1 do
    for j := 0 to numcrit - 1 do
    pularcomparacaoPromethee[i,j] := false;


    altxPromethee := OrdemaltPromethee[0];
    altzPromethee := OrdemaltPromethee[1];
    //while TIWUserSession(WebApplication.Data).problemafinalizado = false do
    //begin
      //AnaliseParaElicitacaoPromethee;
   // end;

    SetarCores;
    ResultadosPromethee;
    ComparacaoMetodos;
    if novoproblema = True then
      SaveProblem;
    chrtResultado.RemoveAllSeries;
    chrtResultado.Legend.Items.Clear;
    // chrtresultado.ClearChart;
    chrtResultado.BottomAxis.Items.Clear;
    chrtResultado.View3D := false;
    SetLength(bar, Numalt);
    i := 0;
    Total := length(CasesOrdemcrit);
    totalacumulado := 0;
    for j := 0 to Numalt - 1 do
    begin
      if totalacumulado > 0.8 * Total then
      begin
        bar[i] := TBarSeries.Create(self);
        chrtResultado.AddSeries(bar[i]);
        bar[i].AddXY(0, Total - totalacumulado, formatfloat
            ('0', Total - totalacumulado), coresalt[i]);
        // chrtResultado.BottomAxis.Items.Add(i, 'S' + inttostr(i + 1));

        chrtResultado.Legend.Item[i].Text := 'Other Alternatives';
        bar[i].Color := coresalt[i];
        break;
      end;
      if ResultAlt[Ordemalt[j]] <> 0 then
      begin
        bar[i] := TBarSeries.Create(self);
        chrtResultado.AddSeries(bar[i]);
        bar[i].AddXY(0, ResultAlt[Ordemalt[j]], formatfloat
            ('0', ResultAlt[Ordemalt[j]]), coresalt[i]);
        // chrtResultado.BottomAxis.Items.Add(i, 'S' + inttostr(i + 1));
        chrtResultado.Legend.Item[i].Text := nomealt[Ordemalt[j]];
        bar[i].Color := coresalt[i];
        totalacumulado := totalacumulado + ResultAlt[Ordemalt[j]];
        inc(i);
      end;
    end;
    Total := 0;
    bar[1] := TBarSeries.Create(self);
    chrtResultado.AddSeries(bar[1]);
    for i := 0 to length(ResultSol) - 1 do
      Total := Total + ResultSol[i];
    bar[1].AddXY(0, Total, '', clWhite);
    bar[1].Visible := false;
    // chrtResultado.LeftAxis.Increment := 0.2; // ?
    // chrtResultado.Legend.Visible := false;
    chrtResultado.BottomAxis.LabelsAlternate;
    chrtResultado.BottomAxis.Items.Automatic := false;
    chrtResultado.LeftAxis.Maximum := Total;
    chrtResultado.LeftAxis.AutomaticMaximum := false;
    /// chrtResultado.LeftAxis.MaximumOffset := 16;
    chrtResultado.LeftAxis.Minimum := 0;
    // chrtresultado.leftaxis.Items.Automatic := false;
    chrtResultado.LeftAxis.Increment := Total;

    imgResultado.Picture.Bitmap.Assign
      (chrtResultado.TeeCreateBitmap(clWebPURPLE, rect(0, 0,
          imgResultado.Width, imgResultado.Height)));

    grdresultados.Clear;

    grdresultados.rowcount := 2 + Naltsel;
    grdresultados.ColumnCount := 3;
    { if grdresultados.ColumnCount * 80 > 832 then
      grdresultados.Width := grdresultados.ColumnCount * 80
      else
      grdresultados.Width := 832; }
    grdresultados.Cell[0, 0].Text := 'Alternative';
    grdresultados.Cell[0, 1].Text := 'Frequency';
    grdresultados.Cell[0, 2].Text := '%';
    i := 0;
    for j := 0 to Numalt - 1 do
    begin
      if ResultAlt[Ordemalt[j]] <> 0 then
      begin
        grdresultados.Cell[i + 1, 0].Text := nomealt[Ordemalt[j]];
        grdresultados.Cell[i + 1, 1].Text := FloatToStr(ResultAlt[Ordemalt[j]]);
        grdresultados.Cell[i + 1, 2].Text := formatfloat
          ('0', 100 * ResultAlt[Ordemalt[j]] / Total) + ' %';
        inc(i);
      end;
    end;
    // grdresultados.Cell[0, Numalt+1].Text := 'Cases';
    // grdresultados.Cell[0, Naltsel + 1].Text := 'Total';
    { for i := 0 to length(ResultSol) - 1 do
      grdresultados.Cell[i + 1, 0].Text := 'Sol ' + inttostr(i + 1) + ':'; }
    grdresultados.Cell[1 + i, 0].Text := 'Total:';
    grdresultados.Cell[1 + i, 1].Text := FloatToStr(Total);
    grdresultados.Cell[1 + i, 2].Text := '100%';
    { for i := 0 to length(ResultSol) - 1 do
      begin
      k := 0;
      for j := 0 to Numalt - 1 do
      begin
      if ResultAlt[j] <> 0 then
      begin
      if MatrizSol[OrdemSol[i], j] = 1 then
      grdresultados.Cell[i + 1, k + 1].Text := 'X'; // 'μ|Max|σ';
      inc(k);
      end;
      end;
      end; }

    {
      for i := 0 to length(ResultSol) - 1 do
      grdresultados.Cell[i + 1, Naltsel + 1].Text := floattostr
      (ResultSol[OrdemSol[i]]);
      }
    for j := 0 to grdresultados.ColumnCount - 1 do
      for i := 0 to grdresultados.rowcount - 1 do
      begin
        grdresultados.Cell[i, j].Alignment := taCenter;
        grdresultados.Cell[i, j].wrap := True;
        // grdresultados.Cell[0,j].Font.Style := [fsbold];
      end;

    for j := 0 to grdresultados.ColumnCount - 1 do
    begin
      grdresultados.Cell[0, j].Font.Style := [fsbold];
      grdresultados.Cell[1 + Naltsel, j].Font.Color := clwebgray;
      grdresultados.Cell[1 + Naltsel, j].BGColor := clWebWhiteSmoke;
    end;
    grdresultados.Cell[1 + Naltsel, 0].Font.Style := [fsbold];

    { for I := 0 to length(ResultSol) - 1 do
      begin
      grdresultados.Cell[i+1, Numalt+1].Font.Size := 8;
      for j := 0 to Round(Total) - 1 do
      begin
      if Casesol[j] = i+1 then
      begin
      if grdresultados.Cell[i+1, Numalt+1].Text = '' then
      grdresultados.Cell[i+1, Numalt+1].Text := inttostr(j+1)
      else
      grdresultados.Cell[i+1, Numalt+1].Text := grdresultados.Cell[i+1, Numalt+1].Text + ', ' + inttostr(j+1);
      end;
      end;
      end; }

    grdcases.rowcount := Round(Total) + 1;
    grdcases.ColumnCount := 2 + Numcrit;

    for j := 0 to grdcases.ColumnCount - 1 do
      for i := 0 to grdcases.rowcount - 1 do
      begin
        grdcases.Cell[i, j].Alignment := taCenter;
        grdcases.Cell[i, j].wrap := True;
        // grdresultados.Cell[0,j].Font.Style := [fsbold];
      end;
    for j := 0 to grdcases.ColumnCount - 1 do
      grdcases.Cell[0, j].Font.Style := [fsbold];

    grdcases.Cell[0, 0].Text := 'Surrogate';
    for i := 0 to Numcrit - 1 do
      grdcases.Cell[0, i + 1].Text := Nomecrit[i];

    grdcases.Cell[0, Numcrit + 1].Text := 'Solution';

    for j := 0 to Round(Total) - 1 do
    begin
      grdcases.Cell[j + 1, 0].Text := 'S' + inttostr(j + 1);
      for i := 0 to Numalt - 1 do
      begin
        if MatrizPOA[j, i] = 1 then
        begin
          grdcases.Cell[j + 1, Numcrit + 1].Text := nomealt[i] + ' ';
        end;
      end;
      for i := 0 to Numcrit - 1 do
        grdcases.Cell[j + 1, i + 1].Text := inttostr(CasesOrdemcrit[j, i]) + 'º'
    end;

    rgnmatcons.Visible := false;
    rgnresultado.Visible := True;
    // btnshowresults.Enabled := true;

    Total := 0;
    for i := 0 to length(ResultSol) - 1 do
      Total := Total + ResultSol[i];

    grdestatistica.Clear;
    grdestatistica.rowcount := 3 + length(MediaDifSol);
    grdestatistica.ColumnCount := 4 + length(MediaDifSol);

    for i := 0 to length(MediaDifSol) - 1 do
    begin
      grdestatistica.Cell[0, i + 4].Alignment := taCenter;
      // grdestatistica.Cell[0, i+1].wrap := True;
      grdestatistica.Cell[0, i + 4].Font.Style := [fsbold];
      grdestatistica.Cell[1, i + 4].Alignment := taCenter;
      // grdestatistica.Cell[1, i+1].wrap := True;
      grdestatistica.Cell[1, i + 4].Font.Style := [fsbold];
      // grdestatistica.Cell[0, i + 4].Text := 'Sol ' + inttostr(i + 1);

      grdestatistica.Cell[0, i + 4].Font.Color := clwebgray;
      grdestatistica.Cell[1, i + 4].Text := ' μ | σ ';

      grdestatistica.Cell[i + 2, 3].Alignment := taCenter;
      grdestatistica.Cell[i + 2, 3].wrap := True;
      grdestatistica.Cell[i + 2, 3].Font.Style := [fsbold];
      grdestatistica.Cell[i + 2, 3].Font.Color := clwebgray;
      for j := 0 to Numalt - 1 do
      begin
        if MatrizSol[OrdemSol[i], j] = 1 then
        begin
          if grdestatistica.Cell[i + 2, 3].Text = '' then
            grdestatistica.Cell[i + 2, 3].Text := nomealt[j]
          else
            grdestatistica.Cell[i + 2, 3].Text := grdestatistica.Cell[i + 2, 3]
              .Text + ' + ' + nomealt[j];
        end;
      end;
      grdestatistica.Cell[0, 4 + i].Text := grdestatistica.Cell[i + 2, 3].Text;
      // grdestatistica.Cell[i + 2, 3].Text := 'Sol ' + inttostr(i + 1);

      grdestatistica.Cell[i + 2, 1].Alignment := taCenter;
      grdestatistica.Cell[i + 2, 1].wrap := True;
      // grdestatistica.Cell[i + 2, 1].Font.Style := [fsbold];
      grdestatistica.Cell[i + 2, 1].Font.Color := clWebBlue;
      grdestatistica.Cell[i + 2, 1].Text := formatfloat
        ('0.00', MediaGeral[OrdemSol[i]]);

      grdestatistica.Cell[i + 2, 2].Alignment := taCenter;
      grdestatistica.Cell[i + 2, 2].wrap := True;
      // grdestatistica.Cell[i + 2, 1].Font.Style := [fsbold];
      grdestatistica.Cell[i + 2, 2].Font.Color := clWebBlue;
      grdestatistica.Cell[i + 2, 2].Text := formatfloat
        ('0.00', DesvioPadraoGeral[OrdemSol[i]]);

      grdestatistica.Cell[i + 2, 0].Alignment := taCenter;
      grdestatistica.Cell[i + 2, 0].wrap := True;
      // grdestatistica.Cell[i + 2, 1].Font.Style := [fsbold];
      grdestatistica.Cell[i + 2, 0].Font.Color := clWebBlue;
      grdestatistica.Cell[i + 2, 0].Text := formatfloat
        ('0.00', (ResultSol[OrdemSol[i]] / Total) * 100) + '%';
    end;

    grdestatistica.Cell[1, 1].Alignment := taCenter;
    grdestatistica.Cell[1, 1].wrap := True;
    grdestatistica.Cell[1, 1].Font.Style := [fsbold];
    grdestatistica.Cell[1, 1].Font.Color := clWebBlue;
    grdestatistica.Cell[1, 1].Text := 'μ';

    grdestatistica.Cell[1, 2].Alignment := taCenter;
    grdestatistica.Cell[1, 2].wrap := True;
    grdestatistica.Cell[1, 2].Font.Style := [fsbold];
    grdestatistica.Cell[1, 2].Font.Color := clWebBlue;
    grdestatistica.Cell[1, 2].Text := 'σ';

    grdestatistica.Cell[1, 0].Alignment := taCenter;
    grdestatistica.Cell[1, 0].wrap := True;
    grdestatistica.Cell[1, 0].Font.Style := [fsbold];
    grdestatistica.Cell[1, 0].Font.Color := clWebBlue;
    grdestatistica.Cell[1, 0].Text := '%';

    grdestatistica.Cell[2 + length(MediaDifSol), 3].Alignment := taCenter;
    grdestatistica.Cell[2 + length(MediaDifSol), 3].wrap := True;
    grdestatistica.Cell[2 + length(MediaDifSol), 3].Font.Style := [fsbold];
    grdestatistica.Cell[2 + length(MediaDifSol), 3].Font.Color := clWebBlue;
    grdestatistica.Cell[2 + length(MediaDifSol), 3].Text := 'μ:';

    { grdestatistica.Cell[3+length(MediaDifSol), 3].Alignment := taCenter;
      grdestatistica.Cell[3+length(MediaDifSol), 3].wrap := True;
      grdestatistica.Cell[3+length(MediaDifSol), 3].Font.Style := [fsbold];
      grdestatistica.Cell[3+length(MediaDifSol), 3].Font.Color := clWebBlue;
      grdestatistica.Cell[3+length(MediaDifSol), 3].Text := 'σ:'; }

    for i := 0 to length(MediaDifSol) - 1 do
    begin
      grdestatistica.Cell[2 + length(MediaDifSol), 4 + i].Text := formatfloat
        ('0.00', MediaGeralNaosol[OrdemSol[i]]);
      grdestatistica.Cell[2 + length(MediaDifSol), 4 + i].Alignment := taCenter;
      grdestatistica.Cell[2 + length(MediaDifSol), 4 + i].wrap := True;
      grdestatistica.Cell[2 + length(MediaDifSol), 4 + i].Font.Color :=
        clWebBlue;

      { grdestatistica.Cell[3+length(MediaDifSol), 4+i].Text := formatfloat('0.00', DesvioPadraoGeralNaoSol[OrdemSol[i]]);
        grdestatistica.Cell[3+length(MediaDifSol), 4+i].Alignment := taCenter;
        grdestatistica.Cell[3+length(MediaDifSol), 4+i].wrap := True;
        grdestatistica.Cell[3+length(MediaDifSol), 4+i].Font.Color := clWebBlue; }
    end;

    for i := 0 to length(MediaDifSol) - 1 do
      for j := 0 to length(MediaDifSol) - 1 do
      begin
        grdestatistica.Cell[i + 2, j + 4].Alignment := taCenter;
        grdestatistica.Cell[i + 2, j + 4].wrap := True;
        if i = j then
          grdestatistica.Cell[i + 2, j + 4].Text := ''
        else
        begin
          grdestatistica.Cell[i + 2, j + 4].Text := formatfloat
            ('0.00', MediaDifSol[OrdemSol[i], OrdemSol[j]]) + '|' + formatfloat
            ('0.00', DesvioPadraoDifSol[OrdemSol[i], OrdemSol[j]]);
        end;
      end;

    grdmaxmin.Clear;
    grdmaxmin.rowcount := 7 + length(MediaDifSol);
    grdmaxmin.ColumnCount := 5 + length(MediaDifSol);

    for i := 0 to length(MediaDifSol) - 1 do
    begin
      grdmaxmin.Cell[0, i + 5].Alignment := taCenter;
      // grdmaxmin.Cell[0, i+1].wrap := True;
      grdmaxmin.Cell[0, i + 5].Font.Style := [fsbold];
      grdmaxmin.Cell[1, i + 5].Alignment := taCenter;
      // grdmaxmin.Cell[1, i+1].wrap := True;
      grdmaxmin.Cell[1, i + 5].Font.Style := [fsbold];
      // grdmaxmin.Cell[0, i + 5].Text := 'Sol ' + inttostr(i + 1);
      grdmaxmin.Cell[0, i + 5].Font.Color := clwebgray;
      grdmaxmin.Cell[1, i + 5].Text := '  Max | Min ';

      grdmaxmin.Cell[i + 2, 4].Alignment := taCenter;
      grdmaxmin.Cell[i + 2, 4].wrap := True;
      grdmaxmin.Cell[i + 2, 4].Font.Style := [fsbold];
      grdmaxmin.Cell[i + 2, 4].Font.Color := clwebgray;

      for j := 0 to Numalt - 1 do
      begin
        if MatrizSol[OrdemSol[i], j] = 1 then
        begin
          if grdmaxmin.Cell[i + 2, 4].Text = '' then
            grdmaxmin.Cell[i + 2, 4].Text := nomealt[j]
          else
            grdmaxmin.Cell[i + 2, 4].Text := grdmaxmin.Cell[i + 2, 4]
              .Text + ' + ' + nomealt[j];
        end;
      end;
      grdmaxmin.Cell[0, 5 + i].Text := grdmaxmin.Cell[i + 2, 4].Text;

      // grdmaxmin.Cell[i + 2, 4].Text := 'Sol ' + inttostr(i + 1);

      grdmaxmin.Cell[i + 2, 1].Alignment := taCenter;
      grdmaxmin.Cell[i + 2, 1].wrap := True;
      // grdestatistica.Cell[i + 2, 1].Font.Style := [fsbold];
      grdmaxmin.Cell[i + 2, 1].Font.Color := clWebBlue;
      grdmaxmin.Cell[i + 2, 1].Text := formatfloat
        ('0.00', MinimoGeral[OrdemSol[i]]);

      grdmaxmin.Cell[i + 2, 2].Alignment := taCenter;
      grdmaxmin.Cell[i + 2, 2].wrap := True;
      // grdestatistica.Cell[i + 2, 1].Font.Style := [fsbold];
      grdmaxmin.Cell[i + 2, 2].Font.Color := clWebBlue;
      grdmaxmin.Cell[i + 2, 2].Text := formatfloat('0.00', MaxMin[OrdemSol[i]]);

      grdmaxmin.Cell[i + 2, 3].Alignment := taCenter;
      grdmaxmin.Cell[i + 2, 3].wrap := True;
      // grdestatistica.Cell[i + 2, 1].Font.Style := [fsbold];
      grdmaxmin.Cell[i + 2, 3].Font.Color := clWebBlue;
      grdmaxmin.Cell[i + 2, 3].Text := formatfloat('0.00', MinMax[OrdemSol[i]]);

      grdmaxmin.Cell[i + 2, 0].Alignment := taCenter;
      grdmaxmin.Cell[i + 2, 0].wrap := True;
      // grdestatistica.Cell[i + 2, 1].Font.Style := [fsbold];
      grdmaxmin.Cell[i + 2, 0].Font.Color := clWebBlue;
      grdmaxmin.Cell[i + 2, 0].Text := formatfloat
        ('0.00', MaximoGeral[OrdemSol[i]]);

      grdmaxmin.Cell[1, 1].Alignment := taCenter;
      grdmaxmin.Cell[1, 1].wrap := True;
      grdmaxmin.Cell[1, 1].Font.Style := [fsbold];
      grdmaxmin.Cell[1, 1].Font.Color := clWebBlue;
      grdmaxmin.Cell[1, 1].Text := 'Min-Min';

      grdmaxmin.Cell[1, 2].Alignment := taCenter;
      grdmaxmin.Cell[1, 2].wrap := True;
      grdmaxmin.Cell[1, 2].Font.Style := [fsbold];
      grdmaxmin.Cell[1, 2].Font.Color := clWebBlue;
      grdmaxmin.Cell[1, 2].Text := 'Max-Min';

      grdmaxmin.Cell[1, 3].Alignment := taCenter;
      grdmaxmin.Cell[1, 3].wrap := True;
      grdmaxmin.Cell[1, 3].Font.Style := [fsbold];
      grdmaxmin.Cell[1, 3].Font.Color := clWebBlue;
      grdmaxmin.Cell[1, 3].Text := 'Min-Max';

      grdmaxmin.Cell[1, 0].Alignment := taCenter;
      grdmaxmin.Cell[1, 0].wrap := True;
      grdmaxmin.Cell[1, 0].Font.Style := [fsbold];
      grdmaxmin.Cell[1, 0].Font.Color := clWebBlue;
      grdmaxmin.Cell[1, 0].Text := 'Max-Max';

    end;

    grdmaxmin.Cell[2 + length(MediaDifSol), 4].Alignment := taCenter;
    grdmaxmin.Cell[2 + length(MediaDifSol), 4].wrap := True;
    grdmaxmin.Cell[2 + length(MediaDifSol), 4].Font.Style := [fsbold];
    grdmaxmin.Cell[2 + length(MediaDifSol), 4].Font.Color := clWebBlue;
    grdmaxmin.Cell[2 + length(MediaDifSol), 4].Text := 'Max-Max:';

    grdmaxmin.Cell[3 + length(MediaDifSol), 4].Alignment := taCenter;
    grdmaxmin.Cell[3 + length(MediaDifSol), 4].wrap := True;
    grdmaxmin.Cell[3 + length(MediaDifSol), 4].Font.Style := [fsbold];
    grdmaxmin.Cell[3 + length(MediaDifSol), 4].Font.Color := clWebBlue;
    grdmaxmin.Cell[3 + length(MediaDifSol), 4].Text := 'Min-Min:';

    grdmaxmin.Cell[4 + length(MediaDifSol), 4].Alignment := taCenter;
    grdmaxmin.Cell[4 + length(MediaDifSol), 4].wrap := True;
    grdmaxmin.Cell[4 + length(MediaDifSol), 4].Font.Style := [fsbold];
    grdmaxmin.Cell[4 + length(MediaDifSol), 4].Font.Color := clWebBlue;
    grdmaxmin.Cell[4 + length(MediaDifSol), 4].Text := 'Max-Min:';

    grdmaxmin.Cell[5 + length(MediaDifSol), 4].Alignment := taCenter;
    grdmaxmin.Cell[5 + length(MediaDifSol), 4].wrap := True;
    grdmaxmin.Cell[5 + length(MediaDifSol), 4].Font.Style := [fsbold];
    grdmaxmin.Cell[5 + length(MediaDifSol), 4].Font.Color := clWebBlue;
    grdmaxmin.Cell[5 + length(MediaDifSol), 4].Text := 'Min-Max:';

    grdmaxmin.Cell[6 + length(MediaDifSol), 4].Alignment := taCenter;
    grdmaxmin.Cell[6 + length(MediaDifSol), 4].wrap := True;
    grdmaxmin.Cell[6 + length(MediaDifSol), 4].Font.Style := [fsbold];
    grdmaxmin.Cell[6 + length(MediaDifSol), 4].Font.Color := clWebBlue;
    grdmaxmin.Cell[6 + length(MediaDifSol), 4].Text := 'Min-Max-Max:';

    minmaxmaxnaosol := 99999999999;
    for i := 0 to length(MediaDifSol) - 1 do
    begin
      if minmaxmaxnaosol > MaximoGeralNaosol[OrdemSol[i]] then
      begin
        TIWUserSession(WebApplication.Data).minmaxmaxnaosol := MaximoGeralNaosol
          [OrdemSol[i]];
        j := i;
      end;
    end;
    grdmaxmin.Cell[6 + length(MediaDifSol), 5 + j].Text := formatfloat
      ('0.00', minmaxmaxnaosol);
    grdmaxmin.Cell[6 + length(MediaDifSol), 5 + j].Alignment := taCenter;
    grdmaxmin.Cell[6 + length(MediaDifSol), 5 + j].wrap := True;
    grdmaxmin.Cell[6 + length(MediaDifSol), 5 + j].Font.Color := clWebBlue;

    for i := 0 to length(MediaDifSol) - 1 do
    begin
      grdmaxmin.Cell[2 + length(MediaDifSol), 5 + i].Text := formatfloat
        ('0.00', MaximoGeralNaosol[OrdemSol[i]]);
      grdmaxmin.Cell[2 + length(MediaDifSol), 5 + i].Alignment := taCenter;
      grdmaxmin.Cell[2 + length(MediaDifSol), 5 + i].wrap := True;
      grdmaxmin.Cell[2 + length(MediaDifSol), 5 + i].Font.Color := clWebBlue;

      grdmaxmin.Cell[3 + length(MediaDifSol), 5 + i].Text := formatfloat
        ('0.00', MinimoGeralNaosol[OrdemSol[i]]);
      grdmaxmin.Cell[3 + length(MediaDifSol), 5 + i].Alignment := taCenter;
      grdmaxmin.Cell[3 + length(MediaDifSol), 5 + i].wrap := True;
      grdmaxmin.Cell[3 + length(MediaDifSol), 5 + i].Font.Color := clWebBlue;

      grdmaxmin.Cell[4 + length(MediaDifSol), 5 + i].Text := formatfloat
        ('0.00', MaxMinNaosol[OrdemSol[i]]);
      grdmaxmin.Cell[4 + length(MediaDifSol), 5 + i].Alignment := taCenter;
      grdmaxmin.Cell[4 + length(MediaDifSol), 5 + i].wrap := True;
      grdmaxmin.Cell[4 + length(MediaDifSol), 5 + i].Font.Color := clWebBlue;

      grdmaxmin.Cell[5 + length(MediaDifSol), 5 + i].Text := formatfloat
        ('0.00', MinMaxNaosol[OrdemSol[i]]);
      grdmaxmin.Cell[5 + length(MediaDifSol), 5 + i].Alignment := taCenter;
      grdmaxmin.Cell[5 + length(MediaDifSol), 5 + i].wrap := True;
      grdmaxmin.Cell[5 + length(MediaDifSol), 5 + i].Font.Color := clWebBlue;

      { grdmaxmin.Cell[3+length(MediaDifSol), 4+i].Text := formatfloat('0.00', DesvioPadraoGeralNaoSol[OrdemSol[i]]);
        grdmaxmin.Cell[3+length(MediaDifSol), 4+i].Alignment := taCenter;
        grdmaxmin.Cell[3+length(MediaDifSol), 4+i].wrap := True;
        grdmaxmin.Cell[3+length(MediaDifSol), 4+i].Font.Color := clWebBlue; }
    end;

    for i := 0 to length(MediaDifSol) - 1 do
      for j := 0 to length(MediaDifSol) - 1 do
      begin
        grdmaxmin.Cell[i + 2, j + 5].Alignment := taCenter;
        grdmaxmin.Cell[i + 2, j + 5].wrap := True;
        if i = j then
          grdmaxmin.Cell[i + 2, j + 5].Text := ''
        else
        begin
          grdmaxmin.Cell[i + 2, j + 5].Text := formatfloat
            ('0.00', MaxDifSol[OrdemSol[i], OrdemSol[j]]) + '|' + formatfloat
            ('0.00', MinDifSol[OrdemSol[i], OrdemSol[j]]);
        end;
      end;

    auxregra := 0;
    qtdpareto := 0;
    while auxregra < 0.8 * Total do
    begin
      auxregra := auxregra + ResultSol[OrdemSol[qtdpareto]];
      inc(qtdpareto);
    end;

    grdperdamediaemaxima.Clear;
    grdperdamediaemaxima.ColumnCount := 4;
    grdperdamediaemaxima.rowcount := 1 + qtdpareto;
    for i := 0 to grdperdamediaemaxima.ColumnCount - 1 do
    begin
      for j := 0 to grdperdamediaemaxima.rowcount - 1 do
      begin
        grdperdamediaemaxima.Cell[j, i].Alignment := taCenter;
        grdperdamediaemaxima.Cell[j, i].wrap := True;
        if j = 0 then
        begin
          grdperdamediaemaxima.Cell[j, i].BGColor := clWebWhiteSmoke;
          grdperdamediaemaxima.Cell[j, i].Font.Style := [fsbold];
        end;
      end;
    end;
    grdperdamediaemaxima.Cell[0, 0].Text := 'Alternatives';
    grdperdamediaemaxima.Cell[0, 1].Text := 'Probability[P(Ai)]';
    grdperdamediaemaxima.Cell[0, 2].Text := 'Average Loss[μ(Aj,Ai)]:';
    grdperdamediaemaxima.Cell[0, 3].Text := 'Maximum Loss[Dmax(Aj, Ai)]:';
    for j := 0 to qtdpareto - 1 do
    begin
      for i := 0 to Numalt - 1 do
      begin
        if MatrizSol[OrdemSol[j], i] = 1 then
        begin
          if grdperdamediaemaxima.Cell[1 + j, 0].Text = '' then
            grdperdamediaemaxima.Cell[1 + j, 0].Text := nomealt[i]
          else
            grdperdamediaemaxima.Cell[1 + j, 0].Text :=
              grdperdamediaemaxima.Cell[1 + j, 0].Text + ' + ' + nomealt[i];
        end;
        grdperdamediaemaxima.Cell[j + 1, 1].Text := formatfloat
          ('0', (ResultSol[OrdemSol[j]] / Total) * 100) + '%';
        grdperdamediaemaxima.Cell[j + 1, 2].Text := formatfloat
          ('0.00', MediaGeralNaosol[OrdemSol[j]]);
        grdperdamediaemaxima.Cell[j + 1, 3].Text := formatfloat
          ('0.00', MaximoGeralNaosol[OrdemSol[j]]);
      end;
    end;

    grdregradedecisao.Clear;
    grdregradedecisao.ColumnCount := 1 + qtdpareto;
    grdregradedecisao.rowcount := 11;
    for i := 0 to grdregradedecisao.ColumnCount - 1 do
    begin
      if i >= 1 + qtdpareto then
      begin
        grdregradedecisao.Cell[0, i].Font.Color := ClWebRed;
        grdregradedecisao.Cell[0, i].Font.Style := [fsbold];
      end;
      for j := 0 to grdregradedecisao.rowcount - 1 do
      begin
        grdregradedecisao.Cell[j, i].Alignment := taCenter;
        grdregradedecisao.Cell[j, i].wrap := True;
        if i = 0 then
          grdregradedecisao.Cell[j, 0].Font.Style := [fsbold];

        if j = 0 then
          grdregradedecisao.Cell[j, i].BGColor := clWebWhiteSmoke;

        if j >= 8 then
        begin
          grdregradedecisao.Cell[j, i].Font.Color := ClWebRed;
          grdregradedecisao.Cell[j, i].BGColor := clWebWhiteSmoke;
        end;
      end;

    end;

    grdregradedecisao.Cell[0, 0].Text := 'Alternative:';
    grdregradedecisao.Cell[1, 0].Text := 'Probability [P(Ai)]:';
    grdregradedecisao.Cell[2, 0].Text := 'Average Loss[μ(Aj,Ai)]:';
    grdregradedecisao.Cell[3, 0].Text := 'Average Gain[μ(Ai,Aj)]:';
    grdregradedecisao.Cell[4, 0].Text := 'Expected Loss[[1-P(Ai)]*μ(Aj,Ai)]:';
    grdregradedecisao.Cell[5, 0].Text := 'Expected Gain[P(Ai)* μ(Aj,Ai)]:';
    grdregradedecisao.Cell[6, 0].Text := 'Maximum Loss[Dmax(Aj, Ai)]:';
    grdregradedecisao.Cell[7, 0].Text := 'Global Performance[G(Ai)]:';

    grdregradedecisao.Cell[8, 0].Text := 'Minimum Avarage Loss:';
    grdregradedecisao.Cell[9, 0].Text := 'Minimum Maximum Loss:';
    grdregradedecisao.Cell[10, 0].Text := 'Maximum Global Performance:';

    aux1 := 99999999999;
    aux2 := 99999999999;
    aux3 := -9999999999;

    for i := 0 to qtdpareto - 1 do
    begin

      for j := 0 to Numalt - 1 do
      begin
        if MatrizSol[OrdemSol[i], j] = 1 then
        begin
          if grdregradedecisao.Cell[0, 1 + i].Text = '' then
            grdregradedecisao.Cell[0, 1 + i].Text := nomealt[j]
          else
            grdregradedecisao.Cell[0, 1 + i].Text := grdregradedecisao.Cell
              [0, 1 + i].Text + ' + ' + nomealt[j];
        end;
      end;

      // grdregradedecisao.Cell[0,1 + i].Text := 'Sol ' + inttostr(i + 1);
      grdregradedecisao.Cell[1, 1 + i].Text := formatfloat
        ('0', (ResultSol[OrdemSol[i]] / Total) * 100) + '%';
      grdregradedecisao.Cell[2, 1 + i].Text := formatfloat
        ('0.00', MediaGeralNaosol[OrdemSol[i]]);
      grdregradedecisao.Cell[3, 1 + i].Text := formatfloat
        ('0.00', MediaGeral[OrdemSol[i]]);
      grdregradedecisao.Cell[4, 1 + i].Text := formatfloat
        ('0.00', MediaGeralNaosol[OrdemSol[i]] *
          (1 - (ResultSol[OrdemSol[i]] / Total)));
      grdregradedecisao.Cell[5, 1 + i].Text := formatfloat
        ('0.00', MediaGeral[OrdemSol[i]] * (ResultSol[OrdemSol[i]] / Total));
      grdregradedecisao.Cell[6, 1 + i].Text := formatfloat
        ('0.00', MaximoGeralNaosol[OrdemSol[i]]);
      auxregra := 0;
      auxregra := MediaGeral[OrdemSol[i]] * (ResultSol[OrdemSol[i]] / Total)
        - MediaGeralNaosol[OrdemSol[i]] *
        ((Total - ResultSol[OrdemSol[i]]) / Total);
      grdregradedecisao.Cell[7, 1 + i].Text := formatfloat('0.00', auxregra);

      if MediaGeralNaosol[OrdemSol[i]] < aux1 then
      begin

        aux1 := MediaGeralNaosol[OrdemSol[i]];
        SolRegra1 := i;
      end;
      if MaximoGeralNaosol[OrdemSol[i]] < aux2 then
      begin
        aux2 := MaximoGeralNaosol[OrdemSol[i]];
        SolRegra2 := i;
      end;
      if (MediaGeral[OrdemSol[i]] * (ResultSol[OrdemSol[i]] / Total)
          - MediaGeralNaosol[OrdemSol[i]] * ((Total - ResultSol[OrdemSol[i]])
            / Total)) > aux3 then
      begin
        aux3 := (MediaGeral[OrdemSol[i]] * (ResultSol[OrdemSol[i]] / Total)
            - MediaGeralNaosol[OrdemSol[i]] * ((Total - ResultSol[OrdemSol[i]])
              / Total));
        SolRegra3 := i;
      end;

    end;

    grdregradedecisao.Cell[8, 1 + SolRegra1].Text := formatfloat
      ('0.00', MediaGeralNaosol[OrdemSol[SolRegra1]]);

    grdregradedecisao.Cell[9, 1 + SolRegra2].Text := formatfloat
      ('0.00', MaximoGeralNaosol[OrdemSol[SolRegra2]]);

    grdregradedecisao.Cell[10, 1 + SolRegra3].Text := formatfloat
      ('0.00', (MediaGeral[OrdemSol[SolRegra3]] *
          (ResultSol[OrdemSol[SolRegra3]] / Total) - MediaGeralNaosol
          [OrdemSol[SolRegra3]] * ((Total - ResultSol[OrdemSol[SolRegra3]])
            / Total)));

    SetLength(epsilon, 4);
    SetLength(omega, 4);
    SetLength(faixaprob, 4);

    faixaprob[0] := 80;
    faixaprob[1] := 70;
    faixaprob[2] := 60;
    faixaprob[3] := 50;

    epsilon[0] := 0.5;
    epsilon[1] := 0.4;
    epsilon[2] := 0.3;
    epsilon[3] := 0.2;

    omega[0] := 0.25;
    omega[1] := 0.2;
    omega[2] := 0.15;
    omega[3] := 0.10;

    { grdlimiares.Cells[0, 0] := '80% < P(x)';
      grdlimiares.Cells[0, 1] := '70% < P(x) < 80%';
      grdlimiares.Cells[0, 2] := '60% < P(x) < 70%';
      grdlimiares.Cells[0, 3] := '50% < P(x) < 60%';
      }

    // voltar pra ca
    grdlimiares.DynEdits[1, 0] := '0.50';
    grdlimiares.DynEdits[1, 1] := '0.40';
    grdlimiares.DynEdits[1, 2] := '0.30';
    grdlimiares.DynEdits[1, 3] := '0.20';

    grdlimiares.DynEdits[2, 0] := '0.25';
    grdlimiares.DynEdits[2, 1] := '0.20';
    grdlimiares.DynEdits[2, 2] := '0.15';
    grdlimiares.DynEdits[2, 3] := '0.10';

    recomendacao := false;

    grdrecomendacao.Clear;
    grdrecomendacao.rowcount := 4;
    grdrecomendacao.ColumnCount := 2;
    for i := 0 to grdrecomendacao.ColumnCount - 1 do
    begin

      grdrecomendacao.Cell[0, i].Font.Color := ClWebRed;
      grdrecomendacao.Cell[0, i].Font.Style := [fsbold];
      grdrecomendacao.Cell[0, i].BGColor := clWebWhiteSmoke;
      for j := 0 to grdrecomendacao.rowcount - 1 do
      begin
        grdrecomendacao.Cell[j, i].Alignment := taCenter;
        grdrecomendacao.Cell[j, i].wrap := True;
      end;
    end;

    grdrecomendacao.Cell[0, 0].Text := 'Solution';
    grdrecomendacao.Cell[0, 1].Text := 'Possibles Recommendations';
    grdrecomendacao.Cell[1, 1].Text := 'Best Alternative';
    grdrecomendacao.Cell[2, 1].Text := 'Two/Three Alternatives are Competitive';
    grdrecomendacao.Cell[3, 1].Text := 'Unable to make it';
    // txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' + grdperdamediaemaxima.Cell[1, 0].Text + ')' + ' + ' + 'P(' + grdperdamediaemaxima.Cell[2, 0].Text + ')' + ' + ' + 'P(' + grdperdamediaemaxima.Cell[3, 0].Text + ')';
    // Para alt1
    if (TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]] / Total) >=
      (faixaprob[0] / 100) then // Se a alt[1] tiver uma probabilidade >= 80%, ela é recomendada
    begin
      Px := (TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
          / Total) * 100;
      if (TIWUserSession(WebApplication.Data).MaximoGeralNaosol[OrdemSol[0]]
          < epsilon[0]) and (TIWUserSession(WebApplication.Data)
          .MediaGeralNaosol[OrdemSol[0]] < omega[0]) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[1, 0].Text = '' then
              grdrecomendacao.Cell[1, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
        recomendacao := false;
    end;
    if ((ResultSol[OrdemSol[0]] / Total) < (faixaprob[0] / 100)) and
      ((ResultSol[OrdemSol[0]] / Total) >= (faixaprob[1] / 100)) and
      (recomendacao = false) then
    begin
      Px := (TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
          / Total) * 100;
      if (MaximoGeralNaosol[OrdemSol[0]] < epsilon[1]) and
        (MediaGeralNaosol[OrdemSol[0]] < omega[1]) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[1, 0].Text = '' then
              grdrecomendacao.Cell[1, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if ((ResultSol[OrdemSol[0]] / Total) < (faixaprob[1] / 100)) and
      ((ResultSol[OrdemSol[0]] / Total) >= (faixaprob[2] / 100)) and
      (recomendacao = false) then
    begin
      Px := (TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
          / Total) * 100;
      if (MaximoGeralNaosol[OrdemSol[0]] < epsilon[2]) and
        (MediaGeralNaosol[OrdemSol[0]] < omega[2]) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[1, 0].Text = '' then
              grdrecomendacao.Cell[1, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if ((ResultSol[OrdemSol[0]] / Total) < (faixaprob[2] / 100)) and
      ((ResultSol[OrdemSol[0]] / Total) >= (faixaprob[3] / 100)) and
      (recomendacao = false) then
    begin
      Px := (TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
          / Total) * 100;
      if (MaximoGeralNaosol[OrdemSol[0]] < epsilon[3]) and
        (MediaGeralNaosol[OrdemSol[0]] < omega[3]) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[1, 0].Text = '' then
              grdrecomendacao.Cell[1, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[1, 0].Text := grdrecomendacao.Cell[1, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;

    // Para alt1 e alt2
    if (((TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]] + ResultSol
            [OrdemSol[1]]) / Total) >= (faixaprob[0] / 100)) and
      (recomendacao = false) then
    // Se a alt[1] tiver uma probabilidade >= 80%, ela é recomendada
    begin
      Px := ((TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
            + ResultSol[OrdemSol[1]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[0]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[0])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[0]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[0])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if TIWUserSession(WebApplication.Data).MatrizSol[OrdemSol[0], j]
            = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaxima.Cell[2, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
        recomendacao := false;
    end;
    if (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]]) / Total) <
        (faixaprob[0] / 100)) and
      (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]]) / Total) >=
        (faixaprob[1] / 100)) and (recomendacao = false) then
    begin
      Px := ((TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
            + ResultSol[OrdemSol[1]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[1]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[1])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[1]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[1])) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaxima.Cell[2, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]]) / Total) <
        (faixaprob[1] / 100)) and
      (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]]) / Total) >=
        (faixaprob[2] / 100)) and (recomendacao = false) then
    begin
      Px := ((TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
            + ResultSol[OrdemSol[1]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[2]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[2])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[2]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[2])) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaxima.Cell[2, 0].Text + ')';

        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]]) / Total) <
        (faixaprob[2] / 100)) and
      (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]]) / Total) >=
        (faixaprob[3] / 100)) and (recomendacao = false) then
    begin
      Px := ((TIWUserSession(WebApplication.Data).ResultSol[OrdemSol[0]]
            + ResultSol[OrdemSol[1]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[3]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[3])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[3]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[3])) then
      begin

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;

        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        txtdescricaoregra.Lines[2] := 'P(x) = ' + 'P(' +
          grdperdamediaemaxima.Cell[1, 0].Text + ')' + ' + ' + 'P(' +
          grdperdamediaemaxima.Cell[2, 0].Text + ')';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;

    // Para alt1, alt2 e alt3
    if ((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
          [OrdemSol[2]]) / Total) >= (faixaprob[0] / 100) then
    // Se a alt[1] tiver uma probabilidade >= 80%, ela é recomendada
    begin
      Px := ((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[0]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[0])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[0]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[0])) and
        ((MaximoGeralNaosol[OrdemSol[2]] < epsilon[0]) and
          (MediaGeralNaosol[OrdemSol[2]] < omega[0])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[2], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        recomendacao := True;
        exit;
      end
      else
        recomendacao := false;
    end;
    if (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) < (faixaprob[0] / 100)) and
      (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) >= (faixaprob[1] / 100)) and
      (recomendacao = false) then
    begin
      Px := ((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[1]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[1])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[1]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[1])) and
        ((MaximoGeralNaosol[OrdemSol[2]] < epsilon[1]) and
          (MediaGeralNaosol[OrdemSol[2]] < omega[1])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[2], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) < (faixaprob[1] / 100)) and
      (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) >= (faixaprob[2] / 100)) and
      (recomendacao = false) then
    begin
      Px := ((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[2]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[2])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[2]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[2])) and
        ((MaximoGeralNaosol[OrdemSol[2]] < epsilon[2]) and
          (MediaGeralNaosol[OrdemSol[2]] < omega[2])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[2], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;
    if (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) < (faixaprob[2] / 100)) and
      (((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) >= (faixaprob[3] / 100)) and
      (recomendacao = false) then
    begin
      Px := ((ResultSol[OrdemSol[0]] + ResultSol[OrdemSol[1]] + ResultSol
            [OrdemSol[2]]) / Total) * 100;
      if ((MaximoGeralNaosol[OrdemSol[0]] < epsilon[3]) and
          (MediaGeralNaosol[OrdemSol[0]] < omega[3])) and
        ((MaximoGeralNaosol[OrdemSol[1]] < epsilon[3]) and
          (MediaGeralNaosol[OrdemSol[1]] < omega[3])) and
        ((MaximoGeralNaosol[OrdemSol[2]] < epsilon[3]) and
          (MediaGeralNaosol[OrdemSol[2]] < omega[3])) then
      begin
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[0], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[1], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        for j := 0 to Numalt - 1 do
        begin
          if MatrizSol[OrdemSol[2], j] = 1 then
          begin
            if grdrecomendacao.Cell[2, 0].Text = '' then
              grdrecomendacao.Cell[2, 0].Text := nomealt[j]
            else
              grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
                .Text + ' + ' + nomealt[j];
          end;
        end;
        grdrecomendacao.Cell[2, 0].Text := grdrecomendacao.Cell[2, 0]
          .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
        recomendacao := True;
        exit;
      end
      else
      begin
        recomendacao := false;
      end;
    end;

    // caso não passe em nenhuma regra
    if recomendacao = false then
    begin
      grdrecomendacao.Cell[3, 0].Text := 'Not Possible';
      grdrecomendacao.Cell[3, 0].Text := grdrecomendacao.Cell[3, 0]
        .Text + ' (P(x) = ' + formatfloat('0', Px) + '%)';
    end;

  end;

end;

procedure TIWForm1.SaveProblem;
var
  i, j, x, k: Integer;
  numero: Double;
begin



  // IWButton16.Cursor := crWait;

  With TIWUserSession(WebApplication.Data) do
  begin
    SetLength(TIWUserSession(WebApplication.Data).critcode, TIWUserSession
        (WebApplication.Data).Numcrit);
    SetLength(TIWUserSession(WebApplication.Data).altcode, TIWUserSession
        (WebApplication.Data).Numalt);
    // Setando a matriz de consequencias
    with UserSession.ZQuery1 do
    begin
      SQL.Text :=
        'INSERT INTO problema (nome_problema, data_problema,hora_problema, ID_usuario) VALUES ("' + TIWUserSession(WebApplication.Data).NameProblem + '", "' + datetostr(date) + '", "' + datetostr(now) + '","' + ID_user + '")';
      ExecSQL;
      SQL.Text :=
        'SELECT ID_problema from problema WHERE nome_problema = "'
        + TIWUserSession(WebApplication.Data).NameProblem + '"';
      ExecSQL;
      Open;
      Last;
      TIWUserSession(WebApplication.Data).CurrentID_problem := FieldByName
        ('ID_problema').AsInteger;
      Close;
      SQL.Text :=
        'INSERT INTO usuario_problema (ID_usuario, ID_problema) VALUES ( ' +
        ID_user + ' ,' + inttostr(TIWUserSession(WebApplication.Data)
          .CurrentID_problem) + ')';
      ExecSQL;

    end;

    with UserSession.ZQuery1 do
    begin
      for i := 0 to TIWUserSession(WebApplication.Data).Numcrit - 1 do
      begin
        with TIWUserSession(WebApplication.Data) do
        begin
          SQL.Text :=
            'INSERT INTO criterio (nome_criterio, tipo_criterio, niveis, ID_problema) VALUES ("' + Nomecrit[i] + '","' + inttostr(tipocrit[i]) + '","' + inttostr(niveis[i]) + '","' + inttostr(TIWUserSession(WebApplication.Data).CurrentID_problem) + '")';
          ExecSQL;
          // prefdir[i] := 1 - prefdir[i];
        end; // with usersession
      end; // for i
      for j := 0 to TIWUserSession(WebApplication.Data).Numalt - 1 do
      begin
        with TIWUserSession(WebApplication.Data) do
        begin
          SQL.Text :=
            'INSERT INTO alternativa (nome_alternativa, ID_problema) VALUES ("'
            + nomealt[j] + '","' + inttostr(TIWUserSession(WebApplication.Data)
              .CurrentID_problem) + '")';
          ExecSQL;
        end; // with usersession
      end; // for j
    end; // with zquery

    with UserSession.ZQuery1 do
    begin
      for i := 0 to TIWUserSession(WebApplication.Data).Numcrit - 1 do
      begin
        with TIWUserSession(WebApplication.Data) do
        begin
          SQL.Text :=
            'SELECT ID_criterio FROM criterio WHERE nome_criterio="' + Nomecrit
            [i] + '" AND ID_problema= "' + inttostr
            (TIWUserSession(WebApplication.Data).CurrentID_problem) + '"';
          ExecSQL;
          Open;
          critcode[i] := UserSession.ZQuery1.FieldByName('ID_criterio')
            .AsInteger;
          Close;
        end; // with usersession
      end; // for i
      for j := 0 to TIWUserSession(WebApplication.Data).Numalt - 1 do
      begin
        with TIWUserSession(WebApplication.Data) do
        begin
          SQL.Text :=
            'SELECT ID_alternativa FROM alternativa WHERE nome_alternativa="' +
            nomealt[j] + '" AND ID_problema="' + inttostr
            (TIWUserSession(WebApplication.Data).CurrentID_problem) + '"';
          ExecSQL;
          Open;
          altcode[j] := UserSession.ZQuery1.FieldByName('ID_alternativa')
            .AsInteger;
          Close;
        end; // with usersession
      end; // for j
    end; // with zquery

    // Essa parte do código tem que ir para o botão de submeter a matriz
    UserSession.ZQuery1.SQL.Text :=
      'SELECT ID_alternativa FROM alternativa WHERE ID_problema=' + inttostr
      (TIWUserSession(WebApplication.Data).CurrentID_problem);
    UserSession.ZQuery1.ExecSQL;
    UserSession.ZQuery1.Open;
    for j := 1 to TIWUserSession(WebApplication.Data).Numalt do
    begin
      TIWUserSession(WebApplication.Data).altcode[j - 1] :=
        UserSession.ZQuery1.FieldByName('ID_alternativa').AsInteger;
      UserSession.ZQuery1.Next;
    end;
    UserSession.ZQuery1.Close;
    //

    // Essa parte do código tem que ir para o botão de submeter a matriz
    UserSession.ZQuery1.SQL.Text :=
      'SELECT ID_criterio FROM criterio WHERE ID_problema=' + inttostr
      (TIWUserSession(WebApplication.Data).CurrentID_problem);
    UserSession.ZQuery1.ExecSQL;
    UserSession.ZQuery1.Open;
    for k := 1 to TIWUserSession(WebApplication.Data).Numcrit do
    begin
      TIWUserSession(WebApplication.Data).critcode[k - 1] :=
        UserSession.ZQuery1.FieldByName('ID_criterio').AsInteger;
      UserSession.ZQuery1.Next;
    end;
    UserSession.ZQuery1.Close;
    //
    x := 0;
    { with UserSession.ZQuery1 do
      begin
      SQL.Text := 'SELECT * from matrizconsequencia WHERE ID_problema=' + IntToStr
      (TIWUserSession(WebApplication.Data).CurrentID_problem);
      ExecSQL;
      Open;
      x := RecordCount;
      Close;
      end; }
    if (x = 0) then
    begin
      with UserSession.ZQuery1 do
      begin

        for i := 0 to TIWUserSession(WebApplication.Data).Numcrit - 1 do
        begin
          for j := 0 to TIWUserSession(WebApplication.Data).Numalt - 1 do
          begin

            SQL.Text :=
              'INSERT INTO matrizconsequencia (ID_alternativa, ID_criterio, valor_performance, ID_problema) VALUES (' + inttostr(TIWUserSession(WebApplication.Data).altcode[j]) + ',' + inttostr(TIWUserSession(WebApplication.Data).critcode[i]) + ',"' + FloatToStr(MatrizConseq[j, i]) + '",' + inttostr(TIWUserSession(WebApplication.Data).CurrentID_problem) + ')';
            ExecSQL;
          end;
        end;

      end;
    end;
  end;

  WebApplication.showmessage('Data collection submitted', smAlert);
end;

procedure TIWForm1.ComparacaoMetodos;
var
  i, j: Integer;
  Total: real;
  TotalAlt, auxresultado: real;
  bar: array of TBarSeries;
  nalt: Integer;
  pular: array of boolean;
  pular2: boolean;
  ordemsimilaridade: array of Integer;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    Total := 1;
    for i := 1 to Numcrit do
      Total := Total * i;
    Total := Total + 1;
    SetLength(SimilaridadeAlt, Numalt);
    SetLength(SomenteAditivo, Numalt);
    SetLength(SomenteSobreclassificacao, Numalt);
    SimilaridadeGeral := 0;
    nalt := 0;
    for j := 0 to Numalt - 1 do
    begin
      if (ResultAlt[j] <> 0) or (ResultAltPromethee[j] <> 0) then
      begin
        inc(nalt);
      end;
    end;
    for i := 0 to Numalt - 1 do
    begin
      SimilaridadeAlt[i] := 0;
      SomenteAditivo[i] := 0;
      SomenteSobreclassificacao[i] := 0;

      if (ResultAlt[i] = 0) and (ResultAltPromethee[i] <> 0) then
      begin
        SimilaridadeAlt[i] := 0;
        SomenteAditivo[i] := 0;
        SomenteSobreclassificacao[i] := ResultAltPromethee[i];
      end;
      if (ResultAlt[i] <> 0) and (ResultAltPromethee[i] = 0) then
      begin
        SimilaridadeAlt[i] := 0;
        SomenteAditivo[i] := ResultAlt[i];
        SomenteSobreclassificacao[i] := 0;
      end;

      if (TIWUserSession(WebApplication.Data).ResultAlt[i] = 0) and
        (ResultAltPromethee[i] = 0) then
        continue;

      if (ResultAlt[i] <> 0) and (TIWUserSession(WebApplication.Data)
          .ResultAltPromethee[i] <> 0) then
      begin
        for j := 0 to Round(Total) - 1 do
        begin

          if (MatrizPOA[j, i] = 1) and (matrizpoaPromethee[j, i] = 1) then
          begin
            SimilaridadeAlt[i] := SimilaridadeAlt[i] + 1;
            SimilaridadeGeral := SimilaridadeGeral + 1;
          end;

          if (MatrizPOA[j, i] = 1) and (matrizpoaPromethee[j, i] = 0) then
            SomenteAditivo[i] := SomenteAditivo[i] + 1;

          if (MatrizPOA[j, i] = 0) and (matrizpoaPromethee[j, i] = 1) then
            SomenteSobreclassificacao[i] := SomenteSobreclassificacao[i] + 1;
        end;
      end;
      TotalAlt := ResultAlt[i] + ResultAltPromethee[i] - SimilaridadeAlt[i];
      TIWUserSession(WebApplication.Data).SimilaridadeAlt[i] := SimilaridadeAlt
        [i] / (Total);
      TIWUserSession(WebApplication.Data).SomenteAditivo[i] := SomenteAditivo
        [i] / (Total);
      TIWUserSession(WebApplication.Data).SomenteSobreclassificacao[i] :=
        SomenteSobreclassificacao[i] / (Total);
    end;
    TIWUserSession(WebApplication.Data).SimilaridadeGeral :=
      SimilaridadeGeral / (Total);

    SetLength(ordemsimilaridade, Numalt);
    SetLength(pular, Numalt);
    for i := 0 to Numalt - 1 do
      pular[i] := false;

    for i := 0 to Numalt - 1 do
    begin
      auxresultado := 0;
      for j := 0 to Numalt - 1 do
      begin
        if pular[j] = True then
          continue;
        if TIWUserSession(WebApplication.Data).SimilaridadeAlt[j]
          > auxresultado then
          auxresultado := SimilaridadeAlt[j];
      end;
      pular2 := false;
      for j := 0 to Numalt - 1 do
      begin
        if (pular2 = True) or (pular[j] = True) then
          continue;
        if auxresultado = SimilaridadeAlt[j] then
        begin
          pular[j] := True;
          ordemsimilaridade[i] := j;
          pular2 := True;
        end;
      end;
    end;
    grdcomparacao.Clear;

    grdcomparacao.rowcount := 2 + nalt;
    grdcomparacao.ColumnCount := 6;

    for i := 0 to grdcomparacao.ColumnCount - 1 do
    begin
      grdcomparacao.Cell[0, i].Alignment := taCenter;
      grdcomparacao.Cell[0, i].wrap := True;
      grdcomparacao.Cell[0, i].Font.Style := [fsbold];
    end;

    grdcomparacao.Cell[0, 0].Text := 'Alternative';
    grdcomparacao.Cell[0, 1].Text := 'Additive Recommendation (%)';
    grdcomparacao.Cell[0, 2].Text := 'Outranking Recommendation (%)';
    grdcomparacao.Cell[0, 3].Text := 'Similarity (%)';
    grdcomparacao.Cell[0, 4].Text := 'Additive Only (%)';
    grdcomparacao.Cell[0, 5].Text := 'Outranking Only (%)';
    i := 0;
    for j := 0 to Numalt - 1 do
    begin
      if (ResultAlt[ordemsimilaridade[j]] <> 0) or
        (ResultAltPromethee[ordemsimilaridade[j]] <> 0) then
      begin
        grdcomparacao.Cell[i + 1, 0].Text := nomealt[ordemsimilaridade[j]];
        grdcomparacao.Cell[i + 1, 1].Text := formatfloat
          ('0.00', (ResultAlt[ordemsimilaridade[j]] * 100) / Total) + '%';
        grdcomparacao.Cell[i + 1, 2].Text := formatfloat
          ('0.00', (ResultAltPromethee[ordemsimilaridade[j]] * 100) / Total)
          + '%';
        grdcomparacao.Cell[i + 1, 3].Text := formatfloat
          ('0.00', (SimilaridadeAlt[ordemsimilaridade[j]] * 100)) + '%';
        grdcomparacao.Cell[i + 1, 4].Text := formatfloat
          ('0.00', (SomenteAditivo[ordemsimilaridade[j]] * 100)) + '%';
        grdcomparacao.Cell[i + 1, 5].Text := formatfloat
          ('0.00', (SomenteSobreclassificacao[ordemsimilaridade[j]] * 100))
          + '%';

        inc(i);
      end;
    end;
    grdcomparacao.Cell[nalt + 1, 0].Text := 'Global Similarity:';
    grdcomparacao.Cell[nalt + 1, 0].Alignment := taCenter;
    grdcomparacao.Cell[nalt + 1, 0].wrap := True;
    grdcomparacao.Cell[nalt + 1, 0].Font.Style := [fsbold];
    grdcomparacao.Cell[nalt + 1, 0].Font.Color := clwebgray;
    lblsimilaridade.Caption := formatfloat('0.00', (SimilaridadeGeral * 100))
      + '%';
    grdcomparacao.Cell[nalt + 1, 3].Text := formatfloat
      ('0.00', (SimilaridadeGeral * 100)) + '%';
    grdcomparacao.Cell[nalt + 1, 3].Alignment := taCenter;
    grdcomparacao.Cell[nalt + 1, 3].wrap := True;
    grdcomparacao.Cell[nalt + 1, 3].Font.Color := clwebgray;

    chrtsimilaridade.RemoveAllSeries;
    // chrtsimilaridade.ClearChart;
    chrtsimilaridade.View3D := false;
    chrtsimilaridade.Legend.Items.Clear;
    chrtsimilaridade.BottomAxis.Items.Clear;
    SetLength(bar, 3);

    for i := 0 to 2 do
    begin
      bar[i] := TBarSeries.Create(self);
      chrtsimilaridade.AddSeries(bar[i]);
      bar[i].Color := coresalt[i];
      bar[i].Marks.Visible := false;

    end;
    // chrtResultado.BottomAxis.Items.Add(i, 'S' + inttostr(i + 1));
    i := 0;
    for j := 0 to Numalt - 1 do
    begin
      if (ResultAlt[ordemsimilaridade[j]] <> 0) or
        (ResultAltPromethee[ordemsimilaridade[j]] <> 0) then
      begin
        bar[0].AddXY(i, (SimilaridadeAlt[ordemsimilaridade[j]] * 100),
          formatfloat('0.00', (SimilaridadeAlt[ordemsimilaridade[j]] * 100))
            + '%', coresalt[0]);
        bar[1].AddXY(i, (SomenteAditivo[ordemsimilaridade[j]] * 100),
          formatfloat('0.00', (SomenteAditivo[ordemsimilaridade[j]] * 100))
            + '%', coresalt[1]);
        bar[2].AddXY(i, (SomenteSobreclassificacao[ordemsimilaridade[j]] * 100)
            , formatfloat
            ('0.00', (SomenteSobreclassificacao[ordemsimilaridade[j]] * 100))
            + '%', coresalt[2]);
        chrtsimilaridade.BottomAxis.Items.Add(i, nomealt[ordemsimilaridade[j]]);
        inc(i);
      end;
    end;
    chrtsimilaridade.Legend.Item[0].Text := 'Similarity';
    chrtsimilaridade.Legend.Item[1].Text := 'Additive Only';
    chrtsimilaridade.Legend.Item[2].Text := 'Outranking Only';
    // chrtsimilaridade.Legend.Item[3].Text := 'Global Similarity';
    // chrtsimilaridade.LeftAxis.Maximum := 100;
    chrtsimilaridade.LeftAxis.AutomaticMaximum := True;
    /// chrtResultado.LeftAxis.MaximumOffset := 16;
    chrtsimilaridade.LeftAxis.Minimum := 0;

    { bar[3].AddXY(i, (SimilaridadeGeral * 100), formatfloat
      ('0.00', (SimilaridadeGeral * 100)) + '%', coresalt[3]);
      }
    // chrtsimilaridade.Legend.Visible := false;

    imgcomparacao.Picture.Bitmap.Assign(chrtsimilaridade.TeeCreateBitmap
        (clWebPURPLE, rect(0, 0, imgcomparacao.Width, imgcomparacao.Height)));
    grdcasescomparacao.Clear;
    grdcasescomparacao.rowcount := Round(Total) + 1;
    grdcasescomparacao.ColumnCount := 3 + Numcrit;

    if grdcasescomparacao.ColumnCount * 60 > 473 then
      grdcasescomparacao.Width := grdcasescomparacao.ColumnCount * 60
    else
      grdcasescomparacao.Width := 473;

    for j := 0 to grdcasescomparacao.ColumnCount - 1 do
      for i := 0 to grdcasescomparacao.rowcount - 1 do
      begin
        grdcasescomparacao.Cell[i, j].Alignment := taCenter;
        grdcasescomparacao.Cell[i, j].wrap := True;
        // grdcasescomparacao.Cell[0,j].Font.Style := [fsbold];
      end;
    for j := 0 to grdcasescomparacao.ColumnCount - 1 do
      grdcasescomparacao.Cell[0, j].Font.Style := [fsbold];

    grdcasescomparacao.Cell[0, 0].Text := 'Decisor';
    for i := 0 to Numcrit - 1 do
      grdcasescomparacao.Cell[0, i + 1].Text := Nomecrit[i];

    grdcasescomparacao.Cell[0, Numcrit + 2].Text := 'Outranking Solution';
    grdcasescomparacao.Cell[0, Numcrit + 1].Text := 'Additive Solution';
    for j := 0 to Round(Total) - 1 do
    begin
      grdcasescomparacao.Cell[j + 1, 0].Text := 'D' + inttostr(j + 1);
      for i := 0 to Numalt - 1 do
      begin
        if MatrizPOA[j, i] = 1 then
        begin
          grdcasescomparacao.Cell[j + 1, Numcrit + 1].Text :=
            grdcasescomparacao.Cell[j + 1, Numcrit + 1].Text + ' ' + nomealt[i];
        end;
        if matrizpoaPromethee[j, i] = 1 then
        begin
          grdcasescomparacao.Cell[j + 1, Numcrit + 2].Text :=
            grdcasescomparacao.Cell[j + 1, Numcrit + 2].Text + ' ' + nomealt[i];
        end;
      end;
      for i := 0 to Numcrit - 1 do
        grdcasescomparacao.Cell[j + 1, i + 1].Text := inttostr
          (CasesOrdemcrit[j, i]) + 'º'
    end;
  end;
end;

procedure TIWForm1.btnupdateClick(Sender: TObject);
var
  i, j: Integer;
  valor: Integer;
begin
  if not trystrtoint(edtnumalt.Text, valor) then
  begin
    edtnumalt.Text := '';
    WebApplication.showmessage('Fill the number of alternatives correctly!');
    exit;
  end;
  if cmbbxcriteria.ItemIndex = -1 then
  begin
    WebApplication.showmessage('Fill the number of criteria correctly!');
    exit;
  end;
  if edtnumalt.Text = '1' then
  begin
    edtnumalt.Text := '';
    WebApplication.showmessage(
      'The number of alternative must be greater than 1!');
    exit;
  end;

  for j := 1 to grdmatrizconseq.Columns.Count - 1 do
    grdmatrizconseq.Columns.delete(1);

  grdmatrizconseq.rowcount := StrToInt(edtnumalt.Text) + 4;

  for i := 0 to StrToInt(cmbbxcriteria.Items[cmbbxcriteria.ItemIndex]) - 1 do
  begin
    grdmatrizconseq.Columns.Add;
    grdmatrizconseq.Columns.Items[i + 1].Width := 200;
  end;
  With TIWUserSession(WebApplication.Data) do
  begin
    Numalt := StrToInt(edtnumalt.Text);
    Numcrit := StrToInt(cmbbxcriteria.Items[cmbbxcriteria.ItemIndex]);
    if length(edtNomeCrit) <> 0 then
    begin
      for i := 0 to length(edtNomeCrit) - 1 do
        edtNomeCrit[i].Destroy;
      Finalize(edtNomeCrit);
    end;
    if length(edtnomealt) <> 0 then
    begin
      for i := 0 to length(edtnomealt) - 1 do
        edtnomealt[i].Destroy;
      Finalize(edtnomealt);
    end;
    if length(cmbbxTipoCrit) <> 0 then
    begin
      for i := 0 to length(cmbbxTipoCrit) - 1 do
        cmbbxTipoCrit[i].Destroy;
      Finalize(cmbbxTipoCrit);
    end;
    if length(cmbbxNiveisCrit) <> 0 then
    begin
      for i := 0 to length(cmbbxNiveisCrit) - 1 do
        cmbbxNiveisCrit[i].Destroy;
      Finalize(cmbbxNiveisCrit);
    end;
    if length(edtMatrizConseq) <> 0 then
    begin
      for i := 0 to length(edtMatrizConseq) - 1 do
        for j := 0 to length(edtMatrizConseq[i]) - 1 do
          edtMatrizConseq[i, j].Destroy;
      Finalize(edtMatrizConseq);
    end;

    SetLength(edtNomeCrit, Numcrit);
    SetLength(edtnomealt, Numalt);
    SetLength(cmbbxTipoCrit, Numcrit);
    SetLength(cmbbxNiveisCrit, Numcrit);
    SetLength(edtMatrizConseq, Numalt, Numcrit);

    Finalize(Nomecrit);
    Finalize(nomealt);
    Finalize(tipocrit);
    Finalize(niveis);
    Finalize(MatrizConseq);
    Finalize(MatrizConseqNorm);
    Finalize(Max);
    Finalize(Min);

    SetLength(Nomecrit, Numcrit);
    SetLength(nomealt, Numalt);
    SetLength(tipocrit, Numcrit);
    SetLength(niveis, Numcrit);
    SetLength(MatrizConseq, Numalt, Numcrit);
    SetLength(MatrizConseqNorm, Numalt, Numcrit);
    SetLength(Max, Numcrit);
    SetLength(Min, Numcrit);

    for i := 0 to Numcrit - 1 do
    begin
      edtNomeCrit[i] := TIWEdit.Create(self);
      edtNomeCrit[i].Parent := rgnmatrizdeconsequencias;
      edtNomeCrit[i].Top := 55;
      edtNomeCrit[i].Left := 17 + 201 * (i + 1);
      edtNomeCrit[i].Width := 200;
      edtNomeCrit[i].Height := 23;
      edtNomeCrit[i].Text := 'Crit ' + inttostr(i + 1);
      Nomecrit[i] := 'Crit ' + inttostr(i + 1);
      edtNomeCrit[i].Font.FontFamily := 'Verdana, Helvetica, Sans-Serif';
      edtNomeCrit[i].Enabled := false;

      cmbbxTipoCrit[i] := TIWComboBox.Create(self);
      cmbbxTipoCrit[i].Parent := rgnmatrizdeconsequencias;
      cmbbxTipoCrit[i].Top := 78;
      cmbbxTipoCrit[i].Left := 17 + 201 * (i + 1);
      cmbbxTipoCrit[i].Width := 200;
      cmbbxTipoCrit[i].Height := 23;
      cmbbxTipoCrit[i].Items.Add('0: Continuous(Min)');
      cmbbxTipoCrit[i].Items.Add('1: Continuous(Max)');
      cmbbxTipoCrit[i].Items.Add('2: Discreet(Min)');
      cmbbxTipoCrit[i].Items.Add('3: Discreet(Max)');
      cmbbxTipoCrit[i].Items.Add('4: Integer(Min)');
      cmbbxTipoCrit[i].Items.Add('5: Integer(Max)');
      cmbbxTipoCrit[i].ZIndex := 1000;
      cmbbxTipoCrit[i].OnChange := cmbbxauxiliar.OnChange;

      cmbbxNiveisCrit[i] := TIWComboBox.Create(self);
      cmbbxNiveisCrit[i].Parent := rgnmatrizdeconsequencias;
      cmbbxNiveisCrit[i].Top := 101;
      cmbbxNiveisCrit[i].Left := 17 + 201 * (i + 1);
      cmbbxNiveisCrit[i].Width := 200;
      cmbbxNiveisCrit[i].Height := 33;
      cmbbxNiveisCrit[i].Items.Add('2');
      cmbbxNiveisCrit[i].Items.Add('3');
      cmbbxNiveisCrit[i].Items.Add('4');
      cmbbxNiveisCrit[i].Items.Add('5');
      cmbbxNiveisCrit[i].Items.Add('6');
      cmbbxNiveisCrit[i].Items.Add('7');
      cmbbxNiveisCrit[i].ZIndex := 1000;
      cmbbxNiveisCrit[i].Visible := false;

      // cmbbxTipocrit[i] := TiwComboBox.Create(self);
    end;
    for j := 0 to Numalt - 1 do
    begin
      edtnomealt[j] := TIWEdit.Create(self);
      edtnomealt[j].Parent := rgnmatrizdeconsequencias;
      edtnomealt[j].Top := 66 + 23 * (j + 4);
      edtnomealt[j].Left := 17;
      edtnomealt[j].Width := 200;
      edtnomealt[j].Height := 23;
      edtnomealt[j].Text := 'Alt ' + inttostr(j + 1);
      nomealt[j] := 'Alt ' + inttostr(j + 1);
      edtnomealt[j].Font.FontFamily := 'Verdana, Helvetica, Sans-Serif';
      edtnomealt[j].Enabled := false;
    end;

    for i := 0 to Numcrit - 1 do
      for j := 0 to Numalt - 1 do
      begin
        edtMatrizConseq[j, i] := TIWEdit.Create(self);
        edtMatrizConseq[j, i].Parent := rgnmatrizdeconsequencias;
        edtMatrizConseq[j, i].Top := 66 + 23 * (j + 4);
        edtMatrizConseq[j, i].Left := 17 + 201 * (i + 1);
        edtMatrizConseq[j, i].Width := 200;
        edtMatrizConseq[j, i].Height := 23;
        edtMatrizConseq[j, i].Text := '';
        edtMatrizConseq[j, i].Font.FontFamily :=
          'Verdana, Helvetica, Sans-Serif';
        edtMatrizConseq[j, i].Enabled := True;
        edtMatrizConseq[j, i].ZIndex := 1000;
      end;
  end;
  btnshowresults.Enabled := True;
end;

procedure TIWForm1.btnupdatevisualizacaoClick(Sender: TObject);
var
  i, j, k: Integer;
  bar: array of TBarSeries;
  totalacumulado: real;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    if rdgVisualizacao.ItemIndex = 0 then
    begin
      chrtResultado.RemoveAllSeries;
      chrtResultado.Legend.Items.Clear;
      // chrtresultado.ClearChart;
      chrtResultado.BottomAxis.Items.Clear;
      chrtResultado.View3D := false;
      SetLength(bar, Numalt);
      i := 0;
      Total := length(CasesOrdemcrit);
      totalacumulado := 0;
      for j := 0 to Numalt - 1 do
      begin
        if totalacumulado > 0.8 * Total then
        begin
          bar[i] := TBarSeries.Create(self);
          chrtResultado.AddSeries(bar[i]);
          bar[i].AddXY(0, Total - totalacumulado, formatfloat
              ('0', Total - totalacumulado), coresalt[i]);
          // chrtResultado.BottomAxis.Items.Add(i, 'S' + inttostr(i + 1));

          chrtResultado.Legend.Item[i].Text := 'Other Alternatives';
          bar[i].Color := coresalt[i];
          break;
        end;
        if ResultAlt[Ordemalt[j]] <> 0 then
        begin
          bar[i] := TBarSeries.Create(self);
          chrtResultado.AddSeries(bar[i]);
          bar[i].AddXY(0, ResultAlt[Ordemalt[j]], formatfloat
              ('0', ResultAlt[Ordemalt[j]]), coresalt[i]);
          // chrtResultado.BottomAxis.Items.Add(i, 'S' + inttostr(i + 1));
          chrtResultado.Legend.Item[i].Text := nomealt[Ordemalt[j]];
          bar[i].Color := coresalt[i];
          totalacumulado := totalacumulado + ResultAlt[Ordemalt[j]];
          inc(i);
        end;
      end;
      Total := 0;
      bar[1] := TBarSeries.Create(self);
      chrtResultado.AddSeries(bar[1]);
      for i := 0 to length(ResultSol) - 1 do
        Total := Total + ResultSol[i];
      bar[1].AddXY(0, Total, '', clWhite);
      bar[1].Visible := false;
      // chrtResultado.LeftAxis.Increment := 0.2; // ?
      // chrtResultado.Legend.Visible := false;
      chrtResultado.BottomAxis.LabelsAlternate;
      chrtResultado.BottomAxis.Items.Automatic := false;
      chrtResultado.LeftAxis.Maximum := Total;
      chrtResultado.LeftAxis.AutomaticMaximum := false;
      /// chrtResultado.LeftAxis.MaximumOffset := 16;
      chrtResultado.LeftAxis.Minimum := 0;
      // chrtresultado.leftaxis.Items.Automatic := false;
      chrtResultado.LeftAxis.Increment := Total;

    end
    else
    begin
      chrtResultado.RemoveAllSeries;
      chrtResultado.Legend.Items.Clear;
      // chrtresultadoPromethee.ClearChart;
      chrtResultado.BottomAxis.Items.Clear;
      chrtResultado.View3D := false;
      SetLength(bar, Numalt);
      i := 0;
      Total := length(CasesOrdemcrit);
      totalacumulado := 0;
      for j := 0 to Numalt - 1 do
      begin
        if totalacumulado > 0.8 * Total then
        begin
          bar[i] := TBarSeries.Create(self);
          chrtResultado.AddSeries(bar[i]);
          bar[i].AddXY(0, 100 * (Total - totalacumulado) / Total, formatfloat
              ('0', (100 * (Total - totalacumulado)) / Total), coresalt[i]);
          // chrtResultado.BottomAxis.Items.Add(i, 'S' + inttostr(i + 1));

          chrtResultado.Legend.Item[i].Text := 'Other Alternatives';
          bar[i].Color := coresalt[i];
          break;
        end;
        if ResultAlt[Ordemalt[j]] <> 0 then
        begin
          bar[i] := TBarSeries.Create(self);
          chrtResultado.AddSeries(bar[i]);
          bar[i].AddXY(0, 100 * ResultAlt[Ordemalt[j]] / Total, formatfloat
              ('0', 100 * ResultAlt[Ordemalt[j]] / Total), coresalt[i]);
          // chrtResultado.BottomAxis.Items.Add(i, 'S' + inttostr(i + 1));
          chrtResultado.Legend.Item[i].Text := nomealt[Ordemalt[j]];
          bar[i].Color := coresalt[i];
          totalacumulado := totalacumulado + ResultAlt[Ordemalt[j]];
          inc(i);
        end;
      end;
      Total := 0;
      bar[1] := TBarSeries.Create(self);
      chrtResultado.AddSeries(bar[1]);
      for i := 0 to length(ResultSol) - 1 do
        Total := Total + ResultSol[i];
      bar[1].AddXY(0, Total, '', clWhite);
      bar[1].Visible := false;
      // chrtResultado.LeftAxis.Increment := 0.2; // ?
      // chrtResultado.Legend.Visible := false;
      chrtResultado.BottomAxis.LabelsAlternate;
      chrtResultado.BottomAxis.Items.Automatic := false;
      chrtResultado.LeftAxis.Maximum := 100;
      chrtResultado.LeftAxis.AutomaticMaximum := false;
      /// chrtResultado.LeftAxis.MaximumOffset := 16;
      chrtResultado.LeftAxis.Minimum := 0;
      // chrtresultado.leftaxis.Items.Automatic := false;
      chrtResultado.LeftAxis.Increment := 100;
    end;
    imgResultado.Picture.Bitmap.Assign
      (chrtResultado.TeeCreateBitmap(clWebPURPLE, rect(0, 0,
          imgResultado.Width, imgResultado.Height)));
  end;
end;

procedure TIWForm1.cmbbxauxiliarAsyncChange(Sender: TObject;
  EventParams: TStringList);
var
  i: Integer;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    for i := 0 to Numcrit - 1 do
    begin
      if (cmbbxTipoCrit[i].ItemIndex = 2) or (cmbbxTipoCrit[i].ItemIndex = 3)
        then
        cmbbxNiveisCrit[i].Visible := True
      else
        cmbbxNiveisCrit[i].Visible := false;
    end;
  end;
end;

procedure TIWForm1.cmbbxauxiliarChange(Sender: TObject);
var
  i: Integer;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    for i := 0 to Numcrit - 1 do
    begin
      if (cmbbxTipoCrit[i].ItemIndex = 2) or (cmbbxTipoCrit[i].ItemIndex = 3)
        then
        cmbbxNiveisCrit[i].Visible := True
      else
        cmbbxNiveisCrit[i].Visible := false;
    end;
  end;
end;

procedure TIWForm1.cmbbxfiltroChange(Sender: TObject);
var
  i, j, k: Integer;
  tonsdeverde: array of tiwcolor;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    SetLength(tonsdeverde, 11);
    tonsdeverde[0] := $FFFFFF;
    tonsdeverde[1] := $DBEAD5;
    tonsdeverde[2] := $B7D5AC;
    tonsdeverde[3] := $93BF85;
    tonsdeverde[4] := $6EAA5E;
    tonsdeverde[5] := $469536;
    tonsdeverde[6] := $008000;
    tonsdeverde[7] := $217517;
    tonsdeverde[8] := $2E6B23;
    tonsdeverde[9] := $36602C;
    tonsdeverde[10] := $3B5534;

    grdfiltro.Clear;

    grdfiltro.ColumnCount := Numcrit + 1;
    grdfiltro.rowcount := Round(ResultSol[cmbbxfiltro.ItemIndex]) + 1;

    for j := 0 to grdfiltro.ColumnCount - 1 do
      for i := 0 to grdfiltro.rowcount - 1 do
      begin
        grdfiltro.Cell[i, j].Alignment := taCenter;
        grdfiltro.Cell[i, j].wrap := True;
        // grdresultados.Cell[0,j].Font.Style := [fsbold];
      end;
    for j := 0 to grdfiltro.ColumnCount - 1 do
      grdfiltro.Cell[0, j].Font.Style := [fsbold];

    grdfiltro.Cell[0, 0].Text := 'Decisor';
    for i := 0 to Numcrit - 1 do
      grdfiltro.Cell[0, i + 1].Text := Nomecrit[i];

    k := 0;
    for j := 0 to Round(Total) - 1 do
    begin
      if TIWUserSession(WebApplication.Data)
        .casesol[j] = cmbbxfiltro.ItemIndex + 1 then
      begin
        grdfiltro.Cell[k + 1, 0].Text := 'D' + inttostr(j + 1);
        for i := 0 to Numcrit - 1 do
        begin
          grdfiltro.Cell[k + 1, i + 1].Text := inttostr(CasesOrdemcrit[j, i])
            + 'º';
          grdfiltro.Cell[k + 1, i + 1].BGColor := tonsdeverde
            [Numcrit - CasesOrdemcrit[j, i] + 1];
        end;
        inc(k);
      end;
    end;

  end;

end;

procedure TIWForm1.rgnaltxaltCreate(Sender: TObject);
begin
rgnaltxalt.Align := AlClient;
end;

procedure TIWForm1.rgnAnaliseEstatisticaCreate(Sender: TObject);
begin
  rgnAnaliseEstatistica.Align := AlClient;
end;

procedure TIWForm1.rgnAnaliseEstatisticaPrometheeCreate(Sender: TObject);
begin
  rgnAnaliseEstatisticaPromethee.Align := AlClient;
end;

procedure TIWForm1.rgncomparacaoCreate(Sender: TObject);
begin
  rgncomparacao.Align := AlClient;
end;

procedure TIWForm1.rgncritxcritCreate(Sender: TObject);
begin
rgncritxcrit.Align := AlClient;
end;

procedure TIWForm1.rgnelicitacaoaltCreate(Sender: TObject);
begin
rgnelicitacaoalt.Align := AlClient;
end;

procedure TIWForm1.rgnelicitacaocritCreate(Sender: TObject);
begin
rgnelicitacaocrit.Align := alclient;
end;

procedure TIWForm1.rgnEnterLoginCreate(Sender: TObject);
begin
  rgnEnterLogin.Align := AlClient;
end;

procedure TIWForm1.rgnfiltro2Create(Sender: TObject);
begin
  rgnfiltro2.Align := AlClient;
end;

procedure TIWForm1.rgnfiltroCreate(Sender: TObject);
begin
  rgnfiltro.Align := AlClient;
end;

procedure TIWForm1.rgnintroCreate(Sender: TObject);
begin
  rgnintro.Align := AlClient;
end;

procedure TIWForm1.rgnLoginprojectCreate(Sender: TObject);
begin
  rgnLoginproject.Align := AlClient;
end;

procedure TIWForm1.rgnmatconsCreate(Sender: TObject);
begin
  rgnmatcons.Align := AlClient;
  grdmatrizconseq.Cells[0, 0] := 'Criteria:';
  grdmatrizconseq.Cells[0, 1] := 'Criterion Type:';
  grdmatrizconseq.Cells[0, 2] := 'Number of levels of discrete criteria:';
  grdmatrizconseq.Cells[0, 3] := 'Alternatives:';
  grdmatrizconseq.Cells[1, 3] := 'Consequence Matrix:';
end;

procedure TIWForm1.rgnmatconsRender(Sender: TObject);
begin
  with TIWUserSession(WebApplication.Data) do
  begin
    if novoproblema = True then
      edtnameproblem.Enabled := True
    else
      edtnameproblem.Enabled := false;
  end;
end;

procedure TIWForm1.rgnOpcaoInputCreate(Sender: TObject);
begin
  rgnOpcaoInput.Align := AlClient;
end;

procedure TIWForm1.rgnregradedecisaoCreate(Sender: TObject);
begin
  rgnregradedecisao.Align := AlClient;
end;

procedure TIWForm1.rgnregradedecisaoprometheeCreate(Sender: TObject);
begin
  rgnregradedecisaoPromethee.Align := AlClient;
end;

procedure TIWForm1.rgnresultadoCreate(Sender: TObject);
begin
  rgnresultado.Align := AlClient;
end;

procedure TIWForm1.rgnresultadoPrometheeCreate(Sender: TObject);
begin
  rgnresultadopromethee.Align := AlClient;
end;

initialization

TIWForm1.SetAsMainForm;

end.
