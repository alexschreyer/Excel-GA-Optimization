unit excel_optim_main;
{ GA Optimization for Excel }


{ ============================================================================ }


interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, OleServer, GA, Math, StdCtrls, Spin, XPMan, ComObj, ActiveX,
  ComCtrls, Menus, ExcelXP, about;

type
  TForm1 = class(TForm)
    pcMainPage: TPageControl;
    TabSheet1: TTabSheet;
    gbPopulation: TGroupBox;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    edNumChrom: TSpinEdit;
    edCrossProb: TEdit;
    edRandSelProb: TEdit;
    edChromMutProb: TEdit;
    cbCrossType: TComboBox;
    GroupBox4: TGroupBox;
    Label7: TLabel;
    edMaxGen: TSpinEdit;
    TabSheet2: TTabSheet;
    GroupBox3: TGroupBox;
    btnRun: TButton;
    StatusBar1: TStatusBar;
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    Exit1: TMenuItem;
    Help1: TMenuItem;
    About1: TMenuItem;
    TabSheet3: TTabSheet;
    GroupBox1: TGroupBox;
    edF_R: TEdit;
    edF_C: TEdit;
    cbFuncTyp: TComboBox;
    Label1: TLabel;
    GroupBox2: TGroupBox;
    edV1_R: TEdit;
    edV1_C: TEdit;
    edVar1: TEdit;
    edVar2: TEdit;
    edV2_R: TEdit;
    edV2_C: TEdit;
    edVar3: TEdit;
    edV3_R: TEdit;
    edV3_C: TEdit;
    LoadDefinition1: TMenuItem;
    SaveDefinition1: TMenuItem;
    GroupBox5: TGroupBox;
    edC1_R: TEdit;
    edC1_C: TEdit;
    cbConTyp1: TComboBox;
    edConV1: TEdit;
    cbConTyp2: TComboBox;
    edConV2: TEdit;
    edC2_R: TEdit;
    edC2_C: TEdit;
    cbConTyp3: TComboBox;
    edConV3: TEdit;
    edC3_R: TEdit;
    edC3_C: TEdit;
    cbConTyp4: TComboBox;
    edConV4: TEdit;
    edC4_R: TEdit;
    edC4_C: TEdit;
    GroupBox6: TGroupBox;
    edExcelFilename: TEdit;
    btFileSelect: TButton;
    OpenDialog1: TOpenDialog;
    Label16: TLabel;
    edV1L: TEdit;
    Label17: TLabel;
    edV1H: TEdit;
    edV2L: TEdit;
    edV2H: TEdit;
    edV3L: TEdit;
    edV3H: TEdit;
    edCon1: TEdit;
    edCon2: TEdit;
    edCon3: TEdit;
    edCon4: TEdit;
    edCon5: TEdit;
    edC5_R: TEdit;
    edC5_C: TEdit;
    cbConTyp5: TComboBox;
    edConV5: TEdit;
    Label18: TLabel;
    Label2: TLabel;
    Label19: TLabel;
    Label3: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    Label23: TLabel;
    edFV: TEdit;
    Label24: TLabel;
    SaveDialog1: TSaveDialog;
    Analysis1: TMenuItem;
    Run1: TMenuItem;
    SaveOpenModel1: TMenuItem;
    cbPlotAll: TCheckBox;
    Memo1: TMemo;
    Label15: TLabel;
    edConvTol: TEdit;
    Label25: TLabel;
    edNumPrec: TSpinEdit;
    Label26: TLabel;
    cbIntVar2: TCheckBox;
    cbIntVar3: TCheckBox;
    cbIntVar1: TCheckBox;
    edVar4: TEdit;
    cbIntVar4: TCheckBox;
    edV4_R: TEdit;
    edV4_C: TEdit;
    edV4L: TEdit;
    edV4H: TEdit;
    edVar5: TEdit;
    cbIntVar5: TCheckBox;
    edV5_R: TEdit;
    edV5_C: TEdit;
    edV5L: TEdit;
    edV5H: TEdit;
    GroupBox7: TGroupBox;
    Label13: TLabel;
    edPrelimRuns: TSpinEdit;
    Label14: TLabel;
    edMaxPrelimGenes: TSpinEdit;
    Label27: TLabel;
    cbWorkSheets: TComboBox;
    Label28: TLabel;
    GroupBox8: TGroupBox;
    Label22: TLabel;
    edConstTol: TEdit;
    Label29: TLabel;
    edConstPen: TEdit;
    cbConstPenType: TComboBox;
    edFunc1: TEdit;
    Button1: TButton;
    Label30: TLabel;
    rbGATypeAd: TRadioButton;
    rbGATypeSw: TRadioButton;
    procedure btnRunClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure About1Click(Sender: TObject);
    procedure btFileSelectClick(Sender: TObject);
    procedure LoadDefinition1Click(Sender: TObject);
    procedure SaveDefinition1Click(Sender: TObject);
    procedure cbFuncTypChange(Sender: TObject);
    procedure SaveOpenModel1Click(Sender: TObject);
    procedure cbWorkSheetsChange(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure rbGATypeSwClick(Sender: TObject);
  private
    iNumVar, iNumConstr : integer;
    rPenalty : double;
    openFile : string;
    function CheckLoadExcel(fname : string; cbSheetList : TComboBox) : boolean ;
    { Private declarations }
  public
    { Public declarations }
  end;


type TGAFloatFunc = Class(TGAFloat)
  { Class type for Float-based optimization }
	private
  protected
	  function GetFitness(iChromIndex : Integer) : Double; override;
  public
  published
	end;

type TGACellSequence = Class(TGASequenceList)
  { Class type for Sequence-based optimization }
	private
  protected
	  function GetFitness(iChromIndex : Integer) : Double; override;
  public
		procedure SetInitialSequence;
  published
	end;

var
  Form1: TForm1;
  XLApp, XLWSheet, XLSheet : Variant;
  F_C : array[0..1] of integer;
  DV_C : array of array of integer;
  CN_C : array of array of integer;
  CN_Lim : array of double;
  rConstTol, rFTargetVal : double;


{ ============================================================================ }


implementation

{$R *.dfm}


{ ======== HELPER FUNCTIONS ================================================== }


function CRound(x: Extended): LongInt;
{ Do correct rounding, not bank rounding }
begin
  Result := Trunc(x);
  if (Frac(x) >= 0.5) then
    Result := Result + 1;
end;


function CRoundTo(const AValue: Double; const ADigit: TRoundToRange): Double;
{ Correct rounding RoundTo }
var
  LFactor: Double;
begin
  LFactor := IntPower(10, ADigit);
  Result := Trunc(AValue) +
            CRound(Frac(AValue) * LFactor) / LFactor;
end;


{ ========== OPTIMIZATION FUNCTIONS ========================================== }


function TGAFloatFunc.GetFitness(iChromIndex: Integer): Double;
{ The fitness function for GA type Adjust Values in Cells }
var
	i, nd : Integer;
  rval : double;
begin
  // Send last genes to Excel
  for i := 0 to FChromosomeDim-1 do
    XLSheet.Cells[DV_C[i,0],DV_C[i,1]] := (FChromosomes[iChromIndex] As TChromFloat).GetGene(i);

  // Read target cell and modify for type of optimization
  Result := XLSheet.Cells[F_C[0],F_C[1]];
  if Form1.cbFuncTyp.ItemIndex = 0 then
    Result := -1 * Result
  else if Form1.cbFuncTyp.ItemIndex = 2 then
    Result := -1 * Abs(Result - rFTargetVal);

  // Apply constraint penalties with ABSOLUTE constraint tolerances
  // Round to number of digits specified for comparison
  nd := Form1.edNumPrec.Value;
  for i := 0 to Form1.iNumConstr-1 do begin
    rval := XLSheet.Cells[CN_C[i,0],CN_C[i,1]];
    rval := CRoundTo(rval, nd);
    case CN_C[i,2] of  // Check type of constraint
     0 : if (rval > CRoundTo((CN_Lim[i] + rConstTol), nd)) then   // "<="
            case Form1.cbConstPenType.ItemIndex of
            0 : begin   // Fixed (absolute) penalty
                  Result := Result + Form1.rPenalty;
                end;
            1 : begin   // Scaled (linear) penalty;
                  Result := Result + (Form1.rPenalty * (Abs(rval - CN_Lim[i])/100));
                end;
            2 : begin   // Scaled (exponential 1) penalty;
                  Result := Result + (Form1.rPenalty * (Exp(Abs(rval - CN_Lim[i])/100) - 1));
                end;
            3 : begin   // Scaled (exponential 2) penalty;
                  Result := Result + (Form1.rPenalty * (Exp(Abs(rval - CN_Lim[i])/50) - 1));
                end;
           end;
     1 : if (rval < CRoundTo((CN_Lim[i] - rConstTol), nd)) then   // ">="
            case Form1.cbConstPenType.ItemIndex of
            0 : begin   // Fixed (absolute) penalty
                  Result := Result + Form1.rPenalty;
                end;
            1 : begin   // Scaled (linear) penalty;
                  Result := Result + (Form1.rPenalty * (Abs(rval - CN_Lim[i])/100));
                end;
            2 : begin   // Scaled (exponential 1) penalty;
                  Result := Result + (Form1.rPenalty * (Exp(Abs(rval - CN_Lim[i])/100) - 1));
                end;
            3 : begin   // Scaled (exponential 2) penalty;
                  Result := Result + (Form1.rPenalty * (Exp(Abs(rval - CN_Lim[i])/50) - 1));
                end;
           end;
     2 : if (rval < CRoundTo((CN_Lim[i] - rConstTol), nd)) or
            (rval > CRoundTo((CN_Lim[i] + rConstTol), nd)) then   // "="
            case Form1.cbConstPenType.ItemIndex of
            0 : begin   // Fixed (absolute) penalty
                  Result := Result + Form1.rPenalty;
                end;
            1 : begin   // Scaled (linear) penalty;
                  Result := Result + (Form1.rPenalty * (Abs(rval - CN_Lim[i])/100));
                end;
            2 : begin   // Scaled (exponential 1) penalty;
                  Result := Result + (Form1.rPenalty * (Exp(Abs(rval - CN_Lim[i])/100) - 1));
                end;
            3 : begin   // Scaled (exponential 2) penalty;
                  Result := Result + (Form1.rPenalty * (Exp(Abs(rval - CN_Lim[i])/50) - 1));
                end;
           end;
    end;
  end;

  Application.ProcessMessages;

  {  OLD ATTEMPTS:
  
  // Apply fixed constraint penalties without constraint tolerances
  // Runs fine
  for i := 0 to Form1.iNumConstr-1 do begin
    rval := XLSheet.Cells[CN_C[i,0],CN_C[i,1]];
    case CN_C[i,2] of  // Check type of constraint
     0 : if (rval > CN_Lim[i]) then
             Result := penalty;
     1 : if (rval < CN_Lim[i]) then
             Result := penalty;
     2 : if (rval < CN_Lim[i]) or (rval > CN_Lim[i]) then
             Result := penalty;
    end;
  end;

  // Apply constraint penalties with absolute constraint tolerances
  // Seems to work fine now.
  for i := 0 to Form1.iNumConstr-1 do begin
    rval := XLSheet.Cells[CN_C[i,0],CN_C[i,1]];
    case CN_C[i,2] of  // Check type of constraint
     0 : if (rval > (CN_Lim[i] + rConstTol)) then
             Result := penalty;
     1 : if (rval < (CN_Lim[i] - rConstTol)) then
             Result := penalty;
     2 : if (rval < (CN_Lim[i] - rConstTol)) or (rval > (CN_Lim[i] + rConstTol)) then
             Result := penalty;
    end;
  end;

  // Apply constraint penalties with exterior penalty method without tolerances
  for i := 0 to Form1.iNumConstr-1 do begin
    rval := XLSheet.Cells[CN_C[i,0],CN_C[i,1]];
    case CN_C[i,2] of
     0 : if (rval > (CN_Lim[i])) then
             Result := Result + penalty * Sqr( Max(0, rval) );
     1 : if (rval < (CN_Lim[i])) then
             Result := Result + penalty * Sqr( Min(0, rval) );
     2 : if (rval < (CN_Lim[i])) or (rval > (CN_Lim[i])) then
             Result := Result + penalty * Sqr( Max(0, rval) );
    end;
  end;

  // Apply constraint penalties with absolute/rel. constraint tolerances  FIX
  for i := 0 to Form1.iNumConstr-1 do begin
    rval := XLSheet.Cells[CN_C[i,0],CN_C[i,1]];
    case CN_C[i,2] of
     0 : if (rval > (CN_Lim[i] - rConstTol)) then
             Result := penalty;
     1 : if CN_Lim[i] <> 0 then
           if rval < (CN_Lim[i] * (1 + (rConstTol * CN_Lim[i] / Abs(CN_Lim[i])))) then
             Result := penalty
         else
           if (rval < 0) then
             Result := penalty;
     2 : if CN_Lim[i] <> 0 then
           if (rval < (CN_Lim[i] * (1 + (rConstTol * CN_Lim[i] / Abs(CN_Lim[i])))))
             or (rval > (CN_Lim[i] * (1 + (rConstTol * CN_Lim[i] / Abs(CN_Lim[i]))))) then
           Result := penalty
         else
           if (rval < 0) or (rval > 0) then
             Result := penalty;
    end;
  end;
  }
end;


procedure TGACellSequence.SetInitialSequence;
{ can go to run button... }
var
	i : Integer;
begin
    	//initialize one dimensional locations (e.g. traveling salesman's cities)
  for i := 0 to FChromosomeDim-1 do
		FSequence[i] := i + 1;
end;


function TGACellSequence.GetFitness(iChromIndex: Integer): Double;
{ The fitness function for GA type Swap Cells }
var
	rDist, rLocCity1, rLocCity2 : Double;
  i, index1, index2 : Integer;
  sChromosome, sGene1, sGene2 : String;
begin
	sChromosome := FChromosomes[iChromIndex].GetGenesAsStr;

  rDist := 0;
	for i := 1 to Length(sChromosome)-1 do
  begin
    sGene1 := sChromosome[i];
  	index1 := Pos(sGene1, FPossGeneValues) - 1;

    sGene2 := sChromosome[i+1];
  	index2 := Pos(sGene2, FPossGeneValues) - 1;

  	rLocCity1 := FSequence[index1];
    rLocCity2 := FSequence[index2];
  	rDist := rDist + abs(rLocCity1 - rLocCity2);
  end;

//  Result := rDist;  this would maximize the dist (find greatest dist)
	if abs(rDist) > 1e-12 then
	  Result := 1 / rDist  //this minimizes dist (find shortest dist)
  else
  	Result := 1 / 1e-12;
end;


procedure TForm1.btnRunClick(Sender: TObject);
{ Start the calculations when user clicks on button }
var
	pop : TGAFloatFunc;  // pop = population
	iNumGenerations, i : Integer;
  LLim, ULim : array of Double;
  VarType : array of Boolean;
  CrossType : TCrossover;
begin

  // ADD INPUT CHECKING HERE!!!

  pcMainPage.ActivePageIndex := 1;

  // Start Excel application object
  i := cbWorkSheets.ItemIndex;
  if CheckLoadExcel(edExcelFilename.Text, cbWorkSheets) = false then
    Exit;
  cbWorkSheets.ItemIndex := i;
  XLSheet := XLWSheet.Worksheets[cbWorkSheets.ItemIndex + 1];
  XLSheet.Activate;
  XLSheet.Visible := true;

  iNumVar := 5;                                     // Number of variables
  iNumConstr := 5;                                  // Number of onstraints
  rConstTol := StrtoFloat(edConstTol.Text);         // Constraint tolerance (absolute)
  rPenalty := - StrtoFloat(Form1.edConstPen.Text);  // Neg. for maximization

  // Set array lengths
  Setlength(LLim, iNumVar);
  Setlength(ULim, iNumVar);
  Setlength(VarType, iNumVar);
  Setlength(DV_C, iNumVar, 2);
  Setlength(CN_C, iNumConstr, 3);
  Setlength(CN_Lim, iNumConstr);

  // Set limits on values:
  LLim[0] := StrtoFloat(edV1L.Text);
  ULim[0] := StrtoFloat(edV1H.Text);
  LLim[1] := StrtoFloat(edV2L.Text);
  ULim[1] := StrtoFloat(edV2H.Text);
  LLim[2] := StrtoFloat(edV3L.Text);
  ULim[2] := StrtoFloat(edV3H.Text);
  LLim[3] := StrtoFloat(edV4L.Text);
  ULim[3] := StrtoFloat(edV4H.Text);
  LLim[4] := StrtoFloat(edV5L.Text);
  ULim[4] := StrtoFloat(edV5H.Text);

  // Set integer switch (0 - real, 1 - integer)
  VarType[0] := cbIntVar1.Checked;
  VarType[1] := cbIntVar2.Checked;
  VarType[2] := cbIntVar3.Checked;
  VarType[3] := cbIntVar4.Checked;
  VarType[4] := cbIntVar5.Checked;

  // Assign cell locations
  F_C[0] := StrtoInt(edF_R.Text);
  F_C[1] := StrtoInt(edF_C.Text);
  if edFV.Text <> '' then
    rFTargetVal := StrtoFloat(edFV.Text)
  else
    rFTargetVal := 0;

  DV_C[0,0] := StrtoInt(edV1_R.Text);
  DV_C[0,1] := StrtoInt(edV1_C.Text);
  DV_C[1,0] := StrtoInt(edV2_R.Text);
  DV_C[1,1] := StrtoInt(edV2_C.Text);
  DV_C[2,0] := StrtoInt(edV3_R.Text);
  DV_C[2,1] := StrtoInt(edV3_C.Text);
  DV_C[3,0] := StrtoInt(edV4_R.Text);
  DV_C[3,1] := StrtoInt(edV4_C.Text);
  DV_C[4,0] := StrtoInt(edV5_R.Text);
  DV_C[4,1] := StrtoInt(edV5_C.Text);

  CN_C[0,0] := StrtoInt(edC1_R.Text);
  CN_C[0,1] := StrtoInt(edC1_C.Text);
  CN_C[0,2] := cbConTyp1.ItemIndex;
  CN_Lim[0] := StrtoFloat(edConV1.Text);
  CN_C[1,0] := StrtoInt(edC2_R.Text);
  CN_C[1,1] := StrtoInt(edC2_C.Text);
  CN_C[1,2] := cbConTyp2.ItemIndex;
  CN_Lim[1] := StrtoFloat(edConV2.Text);
  CN_C[2,0] := StrtoInt(edC3_R.Text);
  CN_C[2,1] := StrtoInt(edC3_C.Text);
  CN_C[2,2] := cbConTyp3.ItemIndex;
  CN_Lim[2] := StrtoFloat(edConV3.Text);
  CN_C[3,0] := StrtoInt(edC4_R.Text);
  CN_C[3,1] := StrtoInt(edC4_C.Text);
  CN_C[3,2] := cbConTyp4.ItemIndex;
  CN_Lim[3] := StrtoFloat(edConV4.Text);
  CN_C[4,0] := StrtoInt(edC5_R.Text);
  CN_C[4,1] := StrtoInt(edC5_C.Text);
  CN_C[4,2] := cbConTyp5.ItemIndex;
  CN_Lim[4] := StrtoFloat(edConV5.Text);

  {
  for i := 0 to (iNumVar)-1 do
  begin
    LLim[i] := 0;
    ULim[i] := 10;
  end;
  }

  CrossType := ctOnePoint;
  case cbCrossType.ItemIndex of
   0 : CrossType := ctOnePoint;
   1 : CrossType := ctTwoPoint;
   2 : CrossType := ctUniform;
   3 : CrossType := ctRoulette;
  end;

  // Run optimization
	pop := TGAFloatFunc.Create(iNumVar,	//chromosome dim (number of genes)
  													edNumChrom.Value, 				//population of chromosomes
                            StrtoFloat(edConvTol.Text),     // Convergence tolerance
                            StrtoFloat(edCrossProb.Text),       	//crossover probability
                          	CRound(StrtoFloat(edRandSelProb.Text)*100),          //random selection chance % (regardless of fitness)
                            edMaxGen.Value,       	//stop after this many generations
                            edPrelimRuns.Value,         //num prelim runs (to build good breeding stock for final--full run)
                          	edMaxPrelimGenes.Value,         //max prelim generations
                            StrtoFloat(edChromMutProb.Text),     	  //chromosome mutation prob.
                            Memo1, 	// memo to display results in
                            StatusBar1, 	//status bar to display results in
                            CrossType, //crossover type
                            edNumPrec.Value,          //num decimal pts of precision
                            LLim, ULim,       // Limits on vars
                            cbPlotAll.Checked,    // Plot all genes into listbox
                            VarType);   // Variable type: real or integer
  try
    Application.ProcessMessages;
  	Screen.Cursor := crHourglass;
    btnRun.Enabled := false;
    Run1.Enabled := false;

  	iNumGenerations := pop.Evolve;
    // Send best genes to Excel
    for i := 0 to iNumVar-1 do
      XLSheet.Cells[DV_C[i,0],DV_C[i,1]] := (pop.FChromosomes[pop.FBestFitnessChromIndex] As TChromFloat).GetGene(i);;
    // GraphGAStats(pop, iNumGenerations);

  finally
    Application.ProcessMessages;
  	Screen.Cursor := crDefault;
    btnRun.Enabled := true;
    Run1.Enabled := true;
  	pop.Free;
  end;
end;


{ ========= GUI FUNCTIONS ==================================================== }


procedure TForm1.FormCreate(Sender: TObject);
{ Arrange form }
begin
  // Put it just lower than the open window's caption bar
  Top := GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYSIZEFRAME);
  Left := Screen.Width - Width;
  pcMainPage.ActivePageIndex := 0;
  openFile := '';
  XLApp := Null;
end;


procedure TForm1.Exit1Click(Sender: TObject);
{ Exit program }
begin
  Close;
end;


procedure TForm1.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
{ Close Excel when closing? }
begin
  if VarisNull(XLApp) = false then
    if MessageDlg('Do you also want to close Excel?',
      mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        XLApp.Quit;
  VarClear(XLApp);
end;


procedure TForm1.About1Click(Sender: TObject);
{ Show about box }
begin
  Application.ProcessMessages;
  with TAboutBox.Create(Application) do
    try
      ShowModal;
    finally
      Free;
    end;
end;


procedure TForm1.cbWorkSheetsChange(Sender: TObject);
{ Switch sheets in Excel when dropdown is changed }
begin
  XLWSheet.Worksheets[cbWorkSheets.ItemIndex+1].Activate;
end;


procedure TForm1.cbFuncTypChange(Sender: TObject);
{ Enabled toggle for function target value }
begin
  if cbFuncTyp.ItemIndex = 2 then
  begin
    edFV.Enabled := true;
    Label23.Enabled := true;
  end
  else
  begin
    edFV.Enabled := false;
    Label23.Enabled := false;
  end;
end;


procedure TForm1.rbGATypeSwClick(Sender: TObject);
{ Toggle display for GA types }
var
  state : boolean;
begin
  state := rbGATypeAd.Checked;

  Label16.Enabled := state;
  Label17.Enabled := state;
  edV1L.Enabled := state;
  edV2L.Enabled := state;
  edV3L.Enabled := state;
  edV4L.Enabled := state;
  edV5L.Enabled := state;
  edV1H.Enabled := state;
  edV2H.Enabled := state;
  edV3H.Enabled := state;
  edV4H.Enabled := state;
  edV5H.Enabled := state;
  cbIntVar1.Enabled := state;
  cbIntVar2.Enabled := state;
  cbIntVar3.Enabled := state;
  cbIntVar4.Enabled := state;
  cbIntVar5.Enabled := state;
end;


function TForm1.CheckLoadExcel(fname : string; cbSheetList : TComboBox) : boolean ;
{ Check for valid Excel file and show sheets }
var
  i : integer;
begin
  Result := false;
  try
    XLApp := GetActiveOleObject('Excel.Application');
  except
    XLApp := CreateOleObject('Excel.Application');
  end;
  if fname <> '' then
    if (XLApp.Workbooks.Count = 0) or
      (XLApp.ActiveWorkbook.Name <> ExtractFileName(fname)) then
      try // open new if it isn't already open and active
        XLWSheet := XLApp.Workbooks.Open(fname);
      except begin
        MessageDlg('No valid Excel file found.', mtWarning, [mbOK], 0);
        cbSheetList.Clear;
        fname := '';
        Exit;
        end;
      end
    else
      XLWSheet := XLApp.ActiveWorkbook
  else begin
    MessageDlg('No Excel file selected. Please select file first.',
      mtWarning, [mbOK], 0);
    cbSheetList.Clear;
    fname := '';
    Exit;
    end;
  cbSheetList.Items.Clear;
  for i := 1 to XLWSheet.WorkSheets.Count do
    cbSheetList.Items.Add(XLWSheet.WorkSheets[i].Name);
  cbSheetList.ItemIndex := 0;
  XLApp.Visible := true;
  XLWSheet.Activate;
  Result := true;
end;


procedure TForm1.btFileSelectClick(Sender: TObject);
{ Show dialog - select Excel file }
begin
  OpenDialog1.Title := 'Select EXCEL file with model calculations';
  OpenDialog1.FilterIndex := 1;
  if OpenDialog1.Execute then
    edExcelFilename.Text := OpenDialog1.FileName
  else
    Exit;
  // Check Excel file and load sheets
  CheckLoadExcel(edExcelFilename.Text, cbWorkSheets);
end;


procedure TForm1.LoadDefinition1Click(Sender: TObject);
{ Show dialog - load model from file }
var
  txtfile : TextFile;
  fname, sbuf : string;
begin
  Application.ProcessMessages;
  OpenDialog1.Title := 'Load problem definition from TXT file';
  OpenDialog1.FilterIndex := 2;
  if OpenDialog1.Execute then
    fname := OpenDialog1.FileName
  else
    Exit;
  AssignFile(txtfile, fname);
  try
    Reset(txtfile);
    ReadLn(txtfile, sbuf);
      edExcelFileName.Text := Copy(sbuf,2,Length(sbuf)-2);
    ReadLn(txtfile, sbuf);
    if CheckLoadExcel(edExcelFilename.Text, cbWorkSheets) then
      cbWorkSheets.ItemIndex := StrtoInt(Trim(sbuf));
    ReadLn(txtfile, sbuf);
      edFunc1.Text := Trim(Copy(sbuf,1,20));
      cbFuncTyp.ItemIndex := StrtoInt(Trim(Copy(sbuf,21,10)));
      edF_R.Text := Trim(Copy(sbuf,31,10));
      edF_C.Text := Trim(Copy(sbuf,41,10));
      edFV.Text := Trim(Copy(sbuf,51,10));
      cbFuncTypChange(Self);
    ReadLn(txtfile); // Number of design variables goes here
    ReadLn(txtfile, sbuf);
      edVar1.Text := Trim(Copy(sbuf,1,20));
      edV1_R.Text := Trim(Copy(sbuf,21,10));
      edV1_C.Text := Trim(Copy(sbuf,31,10));
      edV1L.Text := Trim(Copy(sbuf,41,10));
      edV1H.Text := Trim(Copy(sbuf,51,10));
    ReadLn(txtfile, sbuf);
      edVar2.Text := Trim(Copy(sbuf,1,20));
      edV2_R.Text := Trim(Copy(sbuf,21,10));
      edV2_C.Text := Trim(Copy(sbuf,31,10));
      edV2L.Text := Trim(Copy(sbuf,41,10));
      edV2H.Text := Trim(Copy(sbuf,51,10));
    ReadLn(txtfile, sbuf);
      edVar3.Text := Trim(Copy(sbuf,1,20));
      edV3_R.Text := Trim(Copy(sbuf,21,10));
      edV3_C.Text := Trim(Copy(sbuf,31,10));
      edV3L.Text := Trim(Copy(sbuf,41,10));
      edV3H.Text := Trim(Copy(sbuf,51,10));
    ReadLn(txtfile, sbuf);
      edVar4.Text := Trim(Copy(sbuf,1,20));
      edV4_R.Text := Trim(Copy(sbuf,21,10));
      edV4_C.Text := Trim(Copy(sbuf,31,10));
      edV4L.Text := Trim(Copy(sbuf,41,10));
      edV4H.Text := Trim(Copy(sbuf,51,10));
    ReadLn(txtfile, sbuf);
      edVar5.Text := Trim(Copy(sbuf,1,20));
      edV5_R.Text := Trim(Copy(sbuf,21,10));
      edV5_C.Text := Trim(Copy(sbuf,31,10));
      edV5L.Text := Trim(Copy(sbuf,41,10));
      edV5H.Text := Trim(Copy(sbuf,51,10));
    ReadLn(txtfile, sbuf);
      cbIntVar1.Checked := StrtoBool(sbuf);
    ReadLn(txtfile, sbuf);
      cbIntVar2.Checked := StrtoBool(sbuf);
    ReadLn(txtfile, sbuf);
      cbIntVar3.Checked := StrtoBool(sbuf);
    ReadLn(txtfile, sbuf);
      cbIntVar4.Checked := StrtoBool(sbuf);
    ReadLn(txtfile, sbuf);
      cbIntVar5.Checked := StrtoBool(sbuf);
    ReadLn(txtfile); // Number of constraints goes here
    ReadLn(txtfile, sbuf);
      edCon1.Text := Trim(Copy(sbuf,1,20));
      edC1_R.Text := Trim(Copy(sbuf,21,10));
      edC1_C.Text := Trim(Copy(sbuf,31,10));
      cbConTyp1.ItemIndex := StrtoInt(Trim(Copy(sbuf,41,10)));
      edConV1.Text := Trim(Copy(sbuf,51,10));
    ReadLn(txtfile, sbuf);
      edCon2.Text := Trim(Copy(sbuf,1,20));
      edC2_R.Text := Trim(Copy(sbuf,21,10));
      edC2_C.Text := Trim(Copy(sbuf,31,10));
      cbConTyp2.ItemIndex := StrtoInt(Trim(Copy(sbuf,41,10)));
      edConV2.Text := Trim(Copy(sbuf,51,10));
    ReadLn(txtfile, sbuf);
      edCon3.Text := Trim(Copy(sbuf,1,20));
      edC3_R.Text := Trim(Copy(sbuf,21,10));
      edC3_C.Text := Trim(Copy(sbuf,31,10));
      cbConTyp3.ItemIndex := StrtoInt(Trim(Copy(sbuf,41,10)));
      edConV3.Text := Trim(Copy(sbuf,51,10));
    ReadLn(txtfile, sbuf);
      edCon4.Text := Trim(Copy(sbuf,1,20));
      edC4_R.Text := Trim(Copy(sbuf,21,10));
      edC4_C.Text := Trim(Copy(sbuf,31,10));
      cbConTyp4.ItemIndex := StrtoInt(Trim(Copy(sbuf,41,10)));
      edConV4.Text := Trim(Copy(sbuf,51,10));
    ReadLn(txtfile, sbuf);
      edCon5.Text := Trim(Copy(sbuf,1,20));
      edC5_R.Text := Trim(Copy(sbuf,21,10));
      edC5_C.Text := Trim(Copy(sbuf,31,10));
      cbConTyp5.ItemIndex := StrtoInt(Trim(Copy(sbuf,41,10)));
      edConV5.Text := Trim(Copy(sbuf,51,10));
    ReadLn(txtfile, sbuf);
      edNumChrom.Value := StrtoInt(Trim(Copy(sbuf,1,10)));
      edCrossProb.Text := Trim(Copy(sbuf,11,10));
      edRandSelProb.Text := Trim(Copy(sbuf,21,10));
      edChromMutProb.Text := Trim(Copy(sbuf,31,10));
      cbCrossType.ItemIndex := StrtoInt(Trim(Copy(sbuf,41,10)));
    ReadLn(txtfile, sbuf);
      edConvTol.Text := Trim(Copy(sbuf,1,10));
      edConstTol.Text := Trim(Copy(sbuf,11,10));
      edMaxGen.Value := StrtoInt(Trim(Copy(sbuf,21,10)));
      edPrelimRuns.Text := Trim(Copy(sbuf,31,10));
      edMaxPrelimGenes.Text := Trim(Copy(sbuf,41,10));
    ReadLn(txtfile, sbuf);
      edNumPrec.Value := StrtoInt(Trim(Copy(sbuf,1,10)));
    ReadLn(txtfile, sbuf);
      cbConstPenType.ItemIndex := StrtoInt(Trim(Copy(sbuf,1,5)));
      edConstPen.Text := Trim(Copy(sbuf,6,15));
    ReadLn(txtfile, sbuf);
      rbGATypeAd.Checked := StrtoBool(sbuf);
    ReadLn(txtfile, sbuf);
      rbGATypeSw.Checked := StrtoBool(sbuf);
    openFile := fname;
    SaveOpenModel1.Enabled := true;
    Caption := 'GA Optimization for Excel ['+ExtractFileName(openFile)+']';
    Memo1.Clear;
  finally
    CloseFile(txtfile);
  end;
end;


procedure TForm1.SaveDefinition1Click(Sender: TObject);
{ Show dialog - save model to file }
var
  txtfile : TextFile;
  fname : string;
begin
  Application.ProcessMessages;
  SaveDialog1.Title := 'Save problem definition to TXT file';
  SaveDialog1.FilterIndex := 2;
  if SaveDialog1.Execute then
    fname := SaveDialog1.FileName
  else
    Exit;
  if FileExists(fname) then
    if (MessageDlg('File '+fname+' exists. Overwrite?', mtConfirmation,
      [mbOk, mbCancel], 0) = mrCancel) then
      Exit;
  AssignFile(txtfile, fname);
  try
    Rewrite(txtfile);
    WriteLn(txtfile, QuotedStr(edExcelFileName.Text));
    WriteLn(txtfile, Format('%5d', [cbWorkSheets.ItemIndex]));
    WriteLn(txtfile, Format('%20s%10d%10s%10s%10s', [edFunc1.Text, cbFuncTyp.ItemIndex, edF_R.Text, edF_C.Text, edFV.Text]));
    WriteLn(txtfile, Format('%10d', [5]));
    WriteLn(txtfile, Format('%20s%10s%10s%10s%10s', [edVar1.Text, edV1_R.Text, edV1_C.Text,
      edV1L.Text, edV1H.Text]));
    WriteLn(txtfile, Format('%20s%10s%10s%10s%10s', [edVar2.Text, edV2_R.Text, edV2_C.Text,
      edV2L.Text, edV2H.Text]));
    WriteLn(txtfile, Format('%20s%10s%10s%10s%10s', [edVar3.Text, edV3_R.Text, edV3_C.Text,
      edV3L.Text, edV3H.Text]));
    WriteLn(txtfile, Format('%20s%10s%10s%10s%10s', [edVar4.Text, edV4_R.Text, edV4_C.Text,
      edV4L.Text, edV4H.Text]));
    WriteLn(txtfile, Format('%20s%10s%10s%10s%10s', [edVar5.Text, edV5_R.Text, edV5_C.Text,
      edV5L.Text, edV5H.Text]));
    WriteLn(txtfile, cbIntVar1.Checked);
    WriteLn(txtfile, cbIntVar2.Checked);
    WriteLn(txtfile, cbIntVar3.Checked);
    WriteLn(txtfile, cbIntVar4.Checked);
    WriteLn(txtfile, cbIntVar5.Checked);
    WriteLn(txtfile, Format('%10d', [5]));
    WriteLn(txtfile, Format('%20s%10s%10s%10d%10s', [edCon1.Text, edC1_R.Text, edC1_C.Text,
      cbConTyp1.ItemIndex, edConV1.Text]));
    WriteLn(txtfile, Format('%20s%10s%10s%10d%10s', [edCon2.Text, edC2_R.Text, edC2_C.Text,
      cbConTyp2.ItemIndex, edConV2.Text]));
    WriteLn(txtfile, Format('%20s%10s%10s%10d%10s', [edCon3.Text, edC3_R.Text, edC3_C.Text,
      cbConTyp3.ItemIndex, edConV3.Text]));
    WriteLn(txtfile, Format('%20s%10s%10s%10d%10s', [edCon4.Text, edC4_R.Text, edC4_C.Text,
      cbConTyp4.ItemIndex, edConV4.Text]));
    WriteLn(txtfile, Format('%20s%10s%10s%10d%10s', [edCon5.Text, edC5_R.Text, edC5_C.Text,
      cbConTyp5.ItemIndex, edConV5.Text]));
    WriteLn(txtfile, Format('%10d%10s%10s%10s%10d', [edNumChrom.Value, edCrossProb.Text, edRandSelProb.Text,
      edChromMutProb.Text, cbCrossType.ItemIndex]));
    WriteLn(txtfile, Format('%10s%10s%10d%10s%10s', [edConvTol.Text, edConstTol.Text, edMaxGen.Value,
      edPrelimRuns.Text, edMaxPrelimGenes.Text]));
    WriteLn(txtfile, Format('%10d', [edNumPrec.Value]));
    WriteLn(txtfile, Format('%5d%10s', [cbConstPenType.ItemIndex, edConstPen.Text]));
    WriteLn(txtfile, rbGATypeAd.Checked);
    WriteLn(txtfile, rbGATypeSw.Checked);
    openFile := fname;
    SaveOpenModel1.Enabled := true;
    Caption := 'GA Optimization for Excel ['+ExtractFileName(openFile)+']';
  finally
    CloseFile(txtfile);
  end;
end;


procedure TForm1.SaveOpenModel1Click(Sender: TObject);
{ Save current model to file }
var
  txtfile : TextFile;
begin
  if openFile <> '' then begin
  AssignFile(txtfile, openFile);
  try
    Rewrite(txtfile);
    WriteLn(txtfile, QuotedStr(edExcelFileName.Text));
    WriteLn(txtfile, Format('%5d', [cbWorkSheets.ItemIndex]));
    WriteLn(txtfile, Format('%20s%10d%10s%10s%10s', [edFunc1.Text, cbFuncTyp.ItemIndex, edF_R.Text, edF_C.Text, edFV.Text]));
    WriteLn(txtfile, Format('%10d', [5]));
    WriteLn(txtfile, Format('%20s%10s%10s%10s%10s', [edVar1.Text, edV1_R.Text, edV1_C.Text,
      edV1L.Text, edV1H.Text]));
    WriteLn(txtfile, Format('%20s%10s%10s%10s%10s', [edVar2.Text, edV2_R.Text, edV2_C.Text,
      edV2L.Text, edV2H.Text]));
    WriteLn(txtfile, Format('%20s%10s%10s%10s%10s', [edVar3.Text, edV3_R.Text, edV3_C.Text,
      edV3L.Text, edV3H.Text]));
    WriteLn(txtfile, Format('%20s%10s%10s%10s%10s', [edVar4.Text, edV4_R.Text, edV4_C.Text,
      edV4L.Text, edV4H.Text]));
    WriteLn(txtfile, Format('%20s%10s%10s%10s%10s', [edVar5.Text, edV5_R.Text, edV5_C.Text,
      edV5L.Text, edV5H.Text]));
    WriteLn(txtfile, cbIntVar1.Checked);
    WriteLn(txtfile, cbIntVar2.Checked);
    WriteLn(txtfile, cbIntVar3.Checked);
    WriteLn(txtfile, cbIntVar4.Checked);
    WriteLn(txtfile, cbIntVar5.Checked);
    WriteLn(txtfile, Format('%10d', [5]));
    WriteLn(txtfile, Format('%20s%10s%10s%10d%10s', [edCon1.Text, edC1_R.Text, edC1_C.Text,
      cbConTyp1.ItemIndex, edConV1.Text]));
    WriteLn(txtfile, Format('%20s%10s%10s%10d%10s', [edCon2.Text, edC2_R.Text, edC2_C.Text,
      cbConTyp2.ItemIndex, edConV2.Text]));
    WriteLn(txtfile, Format('%20s%10s%10s%10d%10s', [edCon3.Text, edC3_R.Text, edC3_C.Text,
      cbConTyp3.ItemIndex, edConV3.Text]));
    WriteLn(txtfile, Format('%20s%10s%10s%10d%10s', [edCon4.Text, edC4_R.Text, edC4_C.Text,
      cbConTyp4.ItemIndex, edConV4.Text]));
    WriteLn(txtfile, Format('%20s%10s%10s%10d%10s', [edCon5.Text, edC5_R.Text, edC5_C.Text,
      cbConTyp5.ItemIndex, edConV5.Text]));
    WriteLn(txtfile, Format('%10d%10s%10s%10s%10d', [edNumChrom.Value, edCrossProb.Text, edRandSelProb.Text,
      edChromMutProb.Text, cbCrossType.ItemIndex]));
    WriteLn(txtfile, Format('%10s%10s%10d%10s%10s', [edConvTol.Text, edConstTol.Text, edMaxGen.Value,
      edPrelimRuns.Text, edMaxPrelimGenes.Text]));
    WriteLn(txtfile, Format('%10d', [edNumPrec.Value]));
    WriteLn(txtfile, Format('%5d%10s', [cbConstPenType.ItemIndex, edConstPen.Text]));
    WriteLn(txtfile, rbGATypeAd.Checked);
    WriteLn(txtfile, rbGATypeSw.Checked);
  finally
    CloseFile(txtfile);
  end;
  end;
end;


{ ============================================================================ }




end.
