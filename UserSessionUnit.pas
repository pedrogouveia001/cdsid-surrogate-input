unit UserSessionUnit;

{
  This is a DataModule where you can add components or declare fields that are specific to
  ONE user. Instead of creating global variables, it is better to use this datamodule. You can then
  access the it using UserSession.
}
interface

uses
  IWUserSessionBase, Classes, SysUtils, IWAppForm, IWApplication, IWColor, IWTypes, IWCompEdit,
  IWVCLBaseControl, IWBaseControl, IWBaseHTMLControl, IWControl, IWCompLabel,
  Controls, Forms, IWVCLBaseContainer, IWContainer, IWHTMLContainer,
  IWHTML40Container, IWRegion, TeeGDIPlus, IWCompButton, TeEngine, ExtCtrls,
  TeeProcs, Chart, IWExtCtrls, IWWebGrid, IWAdvWebGrid, BubbleCh, TeePolar,
  IWCompListbox, IWCompRadioButton, IWTMSCheckList, series, IWGrids, TeeLisB,
  TeeSeriesTextEd, Buttons, Math, pngimage, CheckLst,
  ShellApi, TeeLegendScrollBar, Spin, ToolWin, StdActns, ActnList, ImgList,
  TeeURL, TeeXML, AppEvnts, IWCompDynamicChart, IWVCLComponent, TeeProCo,
  Windows, Messages, Variants, Graphics, Dialogs, ComCtrls, Menus, StdCtrls,
  jpeg, ComObj, Grids, IWCompText, IWHTMLControls, ErrorBar,
  IWAdvWebGridExcel, IWProgressIndicator, IWDBGrids,
  IWCompProgressBar, IWBaseComponent, IWBaseHTMLComponent,
  IWBaseHTML40Component, IWColorComboBox, IWTMSImgCtrls, IWTMSCtrls,
  IWCompTabControl, IWHelpTip, IWTMSPopup, IWTMSMenus, GanttCh, IWCompMemo,
  IWCompCheckbox, ZAbstractConnection, ZConnection, DB, ZAbstractRODataset,
  ZAbstractDataset, ZDataset;
type
  TIWUserSession = class(TIWUserSessionBase)
    ZQuery1: TZQuery;
    ZConnection1: TZConnection;
  private
    { Private declarations }
  public
    { Public declarations }
    edtNomeCrit : array of TiwEdit;
    edtNomeAlt : array of TiwEdit;
    cmbbxTipoCrit : array of TiwComboBox;
    cmbbxNiveisCrit : array of TiwComboBox;
    edtMatrizConseq : array of array of TiwEdit;

    NomeCrit, NomeAlt : array of string;
    NumCrit, Numalt : integer;
    Tipocrit : array of integer;
    Niveis : array of integer;
    MatrizConseq, MatrizConseqNorm : array of array of double;
    PesoCrit : array of double;
    CasesOrdemCrit : array of array of integer;
    //OrdemCrit : array of integer;
    Max, Min : array of double;
    MatrizPOA, MatrizPOAPromethee : array of array of integer;
    ResultadoROC, ResultadoPromethee : array of array of double;
    MatrizSol, MatrizSolPromethee : array of array of double;
    ResultAlt, ResultAltPromethee : array of double;
    ResultSol, ResultSolPromethee : array of double;
    coresalt, coresalt1: array of Tcolor;
    CaseSol, CaseSolPromethee : array of integer;
    excelverif : boolean;
    Total: double;
    Crit10, Crit12, Crit13,  Crit14 : array of real;
    Crit20, Crit21, Crit23, Crit24 : array of real;
    MatrizDifVG, MatrizDifVGPromethee : array of array of double;
    MediaDifSol, MaxDifSol, DesvioPadraoDifSol, MinDifSol : array of array of double;
    MediaDifSolPromethee, MaxDifSolPromethee, DesvioPadraoDifSolPromethee, MinDifSolPromethee : array of array of double;
    SobClassMatrix: array of array of real;
    PositiveFlow, NegativeFlow, NetFlow : array of real;
    matrizParaPar: array of array of array of real;
    comp1, comp2, ind, pref, gauss: real;
    OrdemSol, OrdemSolPromethee, ordemalt, ordemaltpromethee : array of integer;
    MediaGeral, DesvioPadraoGeral, MaximoGeral, MinimoGeral, MaxMin, MinMax : array of real;
    MediaGeralNaosol, DesvioPadraoGeralNaosol, MaximoGeralNaosol, MinimoGeralNaosol, MaxMinNaosol, MinMaxNaosol : array of real;
    MediaGeralPromethee, DesvioPadraoGeralPromethee, MaximoGeralPromethee, MinimoGeralPromethee, MaxMinPromethee, MinMaxPromethee : array of real;
    MediaGeralNaosolPromethee, DesvioPadraoGeralNaosolPromethee, MaximoGeralNaosolPromethee, MinimoGeralNaosolPromethee, MaxMinNaosolPromethee, MinMaxNaosolPromethee : array of real;
    SimilaridadeGeral : real;
    SimilaridadeAlt, SomenteAditivo, SomenteSobreclassificacao : array of real;
    MinMaxMaxNaoSol,MinMaxMaxNaoSolPromethee : real;
    primeiroacesso : boolean;
    email, ID_user, senha : string;
    analyst : integer;
    CritName, RestName, AltName, CritSigla, AltSigla: array of string;
    critcode, altcode, PrefDir: array of Integer;
    CurrentID_problem : integer;
    NameProblem : string;
    problems, id_problem, date_problem : array of string;
    novoproblema, retomada : boolean;
    Naltsel, naltselpromethee: Integer;
    combinacaocrit : array of array of double;
    combinacaoalt : array of array of double;
    nperg : integer;
    VetorPerguntaElicitacao : array of array of double;
    VetorRespostaElicitacao : array of array of integer;
    novapopulacaoOrdemCrit : array of array of integer;
    NovaPopulacaoResultadoRoc : array of array of double;
    NovaPopulacaoMatrizPOA : array of array of integer;
    TamanhoNovaPopulacao : array of integer;
    problemafinalizado : boolean;
    pularcomparacao : array of array of boolean;
    combCritAltX, combcritaltZ, combcritaltO : array of array of double;
    combAltAltX, combAltaltZ, combAltaltO : array of array of double;
    AltZ, AltX : integer;

    combinacaocritPromethee : array of array of double;
    combinacaoaltPromethee : array of array of double;
    npergPromethee : integer;
    VetorPerguntaElicitacaoPromethee : array of array of double;
    VetorRespostaElicitacaoPromethee : array of array of integer;
    novapopulacaoOrdemCritPromethee : array of array of integer;
    NovaPopulacaoResultadoRocPromethee : array of array of double;
    NovaPopulacaoMatrizPOAPromethee : array of array of integer;
    TamanhoNovaPopulacaoPromethee : array of integer;
    problemafinalizadoPromethee : boolean;
    pularcomparacaoPromethee : array of array of boolean;
    combCritAltXPromethee, combcritaltZPromethee, combcritaltOPromethee : array of array of double;
    AltZPromethee, AltXPromethee : integer;
    MaxAltVG, minAltVG, maxaltsolVG, minaltsolVG : array of double;
    MinVGSol, MaxVGSol : double;



  end;

implementation

uses UntEntrada;

{$R *.dfm}

end.