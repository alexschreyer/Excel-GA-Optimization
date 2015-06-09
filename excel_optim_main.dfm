object Form1: TForm1
  Left = 333
  Top = 135
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'GA Optimization for EXCEL'
  ClientHeight = 557
  ClientWidth = 484
  Color = clBtnFace
  Constraints.MaxHeight = 999
  Constraints.MaxWidth = 492
  Constraints.MinHeight = 435
  Constraints.MinWidth = 401
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsStayOnTop
  Menu = MainMenu1
  OldCreateOrder = False
  ScreenSnap = True
  ShowHint = True
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  PixelsPerInch = 120
  TextHeight = 16
  object pcMainPage: TPageControl
    Left = 0
    Top = 0
    Width = 484
    Height = 557
    ActivePage = TabSheet3
    Align = alClient
    TabOrder = 0
    object TabSheet3: TTabSheet
      Caption = 'Model Formulation'
      ImageIndex = 2
      object GroupBox1: TGroupBox
        Left = 0
        Top = 97
        Width = 476
        Height = 96
        Align = alTop
        Caption = 'Function'
        TabOrder = 1
        object Label1: TLabel
          Left = 256
          Top = 16
          Width = 30
          Height = 16
          Caption = 'Row:'
        end
        object Label18: TLabel
          Left = 297
          Top = 16
          Width = 23
          Height = 16
          Caption = 'Col:'
        end
        object Label23: TLabel
          Left = 352
          Top = 16
          Width = 43
          Height = 16
          Caption = 'Target:'
          Enabled = False
        end
        object Label30: TLabel
          Left = 8
          Top = 66
          Width = 18
          Height = 16
          Caption = 'by:'
        end
        object edF_R: TEdit
          Left = 256
          Top = 32
          Width = 41
          Height = 24
          TabOrder = 1
          Text = '12'
        end
        object edF_C: TEdit
          Left = 297
          Top = 32
          Width = 40
          Height = 24
          TabOrder = 2
          Text = '4'
        end
        object cbFuncTyp: TComboBox
          Left = 9
          Top = 32
          Width = 96
          Height = 24
          Hint = 'Type of optimization'
          Style = csDropDownList
          ItemHeight = 16
          ItemIndex = 0
          TabOrder = 0
          Text = 'Minimize'
          OnChange = cbFuncTypChange
          Items.Strings = (
            'Minimize'
            'Maximize'
            'Target')
        end
        object edFV: TEdit
          Left = 352
          Top = 32
          Width = 57
          Height = 24
          Hint = 'Target value for optimization'
          Enabled = False
          TabOrder = 3
          Text = '0'
        end
        object edFunc1: TEdit
          Left = 112
          Top = 32
          Width = 129
          Height = 24
          Hint = 'Function name'
          MaxLength = 20
          TabOrder = 4
          Text = 'MinWhat'
        end
        object rbGATypeAd: TRadioButton
          Left = 40
          Top = 66
          Width = 209
          Height = 17
          Caption = 'adjusting values in cells'
          Checked = True
          TabOrder = 5
          TabStop = True
          OnClick = rbGATypeSwClick
        end
        object rbGATypeSw: TRadioButton
          Left = 256
          Top = 66
          Width = 201
          Height = 17
          Caption = 'swapping cells'
          TabOrder = 6
          OnClick = rbGATypeSwClick
        end
      end
      object GroupBox2: TGroupBox
        Left = 0
        Top = 193
        Width = 476
        Height = 168
        Align = alTop
        Caption = 'Design Variables'
        TabOrder = 2
        object Label16: TLabel
          Left = 352
          Top = 16
          Width = 27
          Height = 16
          Caption = 'Low:'
        end
        object Label17: TLabel
          Left = 409
          Top = 16
          Width = 31
          Height = 16
          Caption = 'High:'
        end
        object Label2: TLabel
          Left = 256
          Top = 16
          Width = 30
          Height = 16
          Caption = 'Row:'
        end
        object Label19: TLabel
          Left = 297
          Top = 16
          Width = 23
          Height = 16
          Caption = 'Col:'
        end
        object edV1_R: TEdit
          Left = 256
          Top = 32
          Width = 41
          Height = 24
          TabOrder = 2
          Text = '2'
        end
        object edV1_C: TEdit
          Left = 297
          Top = 32
          Width = 40
          Height = 24
          TabOrder = 3
          Text = '4'
        end
        object edVar1: TEdit
          Left = 9
          Top = 32
          Width = 192
          Height = 24
          Hint = 'Variable1 name'
          MaxLength = 20
          TabOrder = 0
          Text = 'Var1'
        end
        object edVar2: TEdit
          Left = 9
          Top = 56
          Width = 192
          Height = 24
          Hint = 'Variable2 name'
          MaxLength = 20
          TabOrder = 6
          Text = 'Var2'
        end
        object edV2_R: TEdit
          Left = 256
          Top = 56
          Width = 41
          Height = 24
          TabOrder = 8
          Text = '3'
        end
        object edV2_C: TEdit
          Left = 297
          Top = 56
          Width = 40
          Height = 24
          TabOrder = 9
          Text = '4'
        end
        object edVar3: TEdit
          Left = 9
          Top = 80
          Width = 192
          Height = 24
          Hint = 'Variable3 name'
          MaxLength = 20
          TabOrder = 12
          Text = 'Var3'
        end
        object edV3_R: TEdit
          Left = 256
          Top = 80
          Width = 41
          Height = 24
          TabOrder = 14
          Text = '4'
        end
        object edV3_C: TEdit
          Left = 297
          Top = 80
          Width = 40
          Height = 24
          TabOrder = 15
          Text = '4'
        end
        object edV1L: TEdit
          Left = 352
          Top = 32
          Width = 57
          Height = 24
          TabOrder = 4
          Text = '0.01'
        end
        object edV1H: TEdit
          Left = 409
          Top = 32
          Width = 56
          Height = 24
          TabOrder = 5
          Text = '1000'
        end
        object edV2L: TEdit
          Left = 352
          Top = 56
          Width = 57
          Height = 24
          TabOrder = 10
          Text = '0.01'
        end
        object edV2H: TEdit
          Left = 409
          Top = 56
          Width = 56
          Height = 24
          TabOrder = 11
          Text = '1000'
        end
        object edV3L: TEdit
          Left = 352
          Top = 80
          Width = 57
          Height = 24
          TabOrder = 16
          Text = '0.01'
        end
        object edV3H: TEdit
          Left = 409
          Top = 80
          Width = 56
          Height = 24
          TabOrder = 17
          Text = '1000'
        end
        object cbIntVar2: TCheckBox
          Left = 208
          Top = 60
          Width = 41
          Height = 17
          Hint = 'Integer values only'
          Caption = 'Int.'
          TabOrder = 7
        end
        object cbIntVar3: TCheckBox
          Left = 208
          Top = 84
          Width = 41
          Height = 17
          Hint = 'Integer values only'
          Caption = 'Int.'
          TabOrder = 13
        end
        object cbIntVar1: TCheckBox
          Left = 208
          Top = 36
          Width = 41
          Height = 17
          Hint = 'Integer values only'
          Caption = 'Int.'
          TabOrder = 1
        end
        object edVar4: TEdit
          Left = 9
          Top = 104
          Width = 192
          Height = 24
          Hint = 'Variable4 name'
          MaxLength = 20
          TabOrder = 18
          Text = 'Var4'
        end
        object cbIntVar4: TCheckBox
          Left = 208
          Top = 108
          Width = 41
          Height = 17
          Hint = 'Integer values only'
          Caption = 'Int.'
          TabOrder = 19
        end
        object edV4_R: TEdit
          Left = 256
          Top = 104
          Width = 41
          Height = 24
          TabOrder = 20
          Text = '4'
        end
        object edV4_C: TEdit
          Left = 297
          Top = 104
          Width = 40
          Height = 24
          TabOrder = 21
          Text = '4'
        end
        object edV4L: TEdit
          Left = 352
          Top = 104
          Width = 57
          Height = 24
          TabOrder = 22
          Text = '0.01'
        end
        object edV4H: TEdit
          Left = 409
          Top = 104
          Width = 56
          Height = 24
          TabOrder = 23
          Text = '1000'
        end
        object edVar5: TEdit
          Left = 9
          Top = 128
          Width = 192
          Height = 24
          Hint = 'Variable5 name'
          MaxLength = 20
          TabOrder = 24
          Text = 'Var5'
        end
        object cbIntVar5: TCheckBox
          Left = 208
          Top = 132
          Width = 41
          Height = 17
          Hint = 'Integer values only'
          Caption = 'Int.'
          TabOrder = 25
        end
        object edV5_R: TEdit
          Left = 256
          Top = 128
          Width = 41
          Height = 24
          TabOrder = 26
          Text = '4'
        end
        object edV5_C: TEdit
          Left = 297
          Top = 128
          Width = 40
          Height = 24
          TabOrder = 27
          Text = '4'
        end
        object edV5L: TEdit
          Left = 352
          Top = 128
          Width = 57
          Height = 24
          TabOrder = 28
          Text = '0.01'
        end
        object edV5H: TEdit
          Left = 409
          Top = 128
          Width = 56
          Height = 24
          TabOrder = 29
          Text = '1000'
        end
      end
      object GroupBox5: TGroupBox
        Left = 0
        Top = 361
        Width = 476
        Height = 165
        Align = alClient
        Caption = 'Constraints'
        TabOrder = 3
        object Label3: TLabel
          Left = 256
          Top = 16
          Width = 30
          Height = 16
          Caption = 'Row:'
        end
        object Label20: TLabel
          Left = 297
          Top = 16
          Width = 23
          Height = 16
          Caption = 'Col:'
        end
        object Label21: TLabel
          Left = 409
          Top = 16
          Width = 38
          Height = 16
          Caption = 'Value:'
        end
        object edC1_R: TEdit
          Left = 256
          Top = 32
          Width = 41
          Height = 24
          TabOrder = 1
          Text = '6'
        end
        object edC1_C: TEdit
          Left = 297
          Top = 32
          Width = 40
          Height = 24
          TabOrder = 2
          Text = '4'
        end
        object cbConTyp1: TComboBox
          Left = 352
          Top = 32
          Width = 57
          Height = 24
          Hint = 'Type of constraint'
          Style = csDropDownList
          ItemHeight = 16
          ItemIndex = 0
          TabOrder = 3
          Text = '<='
          Items.Strings = (
            '<='
            '>='
            '=')
        end
        object edConV1: TEdit
          Left = 409
          Top = 32
          Width = 56
          Height = 24
          TabOrder = 4
          Text = '0'
        end
        object cbConTyp2: TComboBox
          Left = 352
          Top = 56
          Width = 57
          Height = 24
          Hint = 'Type of constraint'
          Style = csDropDownList
          ItemHeight = 16
          ItemIndex = 0
          TabOrder = 8
          Text = '<='
          Items.Strings = (
            '<='
            '>='
            '=')
        end
        object edConV2: TEdit
          Left = 409
          Top = 56
          Width = 56
          Height = 24
          TabOrder = 9
          Text = '0'
        end
        object edC2_R: TEdit
          Left = 256
          Top = 56
          Width = 41
          Height = 24
          TabOrder = 6
          Text = '7'
        end
        object edC2_C: TEdit
          Left = 297
          Top = 56
          Width = 40
          Height = 24
          TabOrder = 7
          Text = '4'
        end
        object cbConTyp3: TComboBox
          Left = 352
          Top = 80
          Width = 57
          Height = 24
          Hint = 'Type of constraint'
          Style = csDropDownList
          ItemHeight = 16
          ItemIndex = 0
          TabOrder = 13
          Text = '<='
          Items.Strings = (
            '<='
            '>='
            '=')
        end
        object edConV3: TEdit
          Left = 409
          Top = 80
          Width = 56
          Height = 24
          TabOrder = 14
          Text = '0'
        end
        object edC3_R: TEdit
          Left = 256
          Top = 80
          Width = 41
          Height = 24
          TabOrder = 11
          Text = '8'
        end
        object edC3_C: TEdit
          Left = 297
          Top = 80
          Width = 40
          Height = 24
          TabOrder = 12
          Text = '4'
        end
        object cbConTyp4: TComboBox
          Left = 352
          Top = 104
          Width = 57
          Height = 24
          Hint = 'Type of constraint'
          Style = csDropDownList
          ItemHeight = 16
          ItemIndex = 0
          TabOrder = 18
          Text = '<='
          Items.Strings = (
            '<='
            '>='
            '=')
        end
        object edConV4: TEdit
          Left = 409
          Top = 104
          Width = 56
          Height = 24
          TabOrder = 19
          Text = '0'
        end
        object edC4_R: TEdit
          Left = 256
          Top = 104
          Width = 41
          Height = 24
          TabOrder = 16
          Text = '9'
        end
        object edC4_C: TEdit
          Left = 297
          Top = 104
          Width = 40
          Height = 24
          TabOrder = 17
          Text = '4'
        end
        object edCon1: TEdit
          Left = 9
          Top = 32
          Width = 232
          Height = 24
          Hint = 'Constraint1 name'
          MaxLength = 20
          TabOrder = 0
          Text = 'Con1'
        end
        object edCon2: TEdit
          Left = 9
          Top = 56
          Width = 232
          Height = 24
          Hint = 'Constraint2 name'
          MaxLength = 20
          TabOrder = 5
          Text = 'Con2'
        end
        object edCon3: TEdit
          Left = 9
          Top = 80
          Width = 232
          Height = 24
          Hint = 'Constraint3 name'
          MaxLength = 20
          TabOrder = 10
          Text = 'Con3'
        end
        object edCon4: TEdit
          Left = 9
          Top = 104
          Width = 232
          Height = 24
          Hint = 'Constraint4 name'
          MaxLength = 20
          TabOrder = 15
          Text = 'Con4'
        end
        object edCon5: TEdit
          Left = 9
          Top = 128
          Width = 232
          Height = 24
          Hint = 'Constraint5 name'
          MaxLength = 20
          TabOrder = 20
          Text = 'Con5'
        end
        object edC5_R: TEdit
          Left = 256
          Top = 128
          Width = 41
          Height = 24
          TabOrder = 21
          Text = '10'
        end
        object edC5_C: TEdit
          Left = 297
          Top = 128
          Width = 40
          Height = 24
          TabOrder = 22
          Text = '4'
        end
        object cbConTyp5: TComboBox
          Left = 352
          Top = 128
          Width = 57
          Height = 24
          Hint = 'Type of constraint'
          Style = csDropDownList
          ItemHeight = 16
          ItemIndex = 0
          TabOrder = 23
          Text = '<='
          Items.Strings = (
            '<='
            '>='
            '=')
        end
        object edConV5: TEdit
          Left = 409
          Top = 128
          Width = 56
          Height = 24
          TabOrder = 24
          Text = '0'
        end
      end
      object GroupBox6: TGroupBox
        Left = 0
        Top = 0
        Width = 476
        Height = 97
        Align = alTop
        Caption = 'EXCEL File'
        TabOrder = 0
        object Label27: TLabel
          Left = 8
          Top = 60
          Width = 38
          Height = 16
          Caption = 'Sheet:'
        end
        object edExcelFilename: TEdit
          Left = 64
          Top = 25
          Width = 377
          Height = 24
          Hint = 'EXCEL file name'
          TabStop = False
          ReadOnly = True
          TabOrder = 2
          Text = 'Select file...'
        end
        object btFileSelect: TButton
          Left = 441
          Top = 25
          Width = 24
          Height = 24
          Hint = 'Select EXCEL file...'
          Caption = '...'
          TabOrder = 0
          OnClick = btFileSelectClick
        end
        object cbWorkSheets: TComboBox
          Left = 64
          Top = 56
          Width = 377
          Height = 24
          Hint = 'Worksheet with calculation cells'
          Style = csDropDownList
          ItemHeight = 16
          TabOrder = 1
          OnChange = cbWorkSheetsChange
        end
        object Button1: TButton
          Left = 8
          Top = 24
          Width = 49
          Height = 25
          Hint = 'Load Model...'
          Caption = 'Model'
          TabOrder = 3
          OnClick = LoadDefinition1Click
        end
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'Run GA'
      ImageIndex = 1
      object GroupBox3: TGroupBox
        Left = 0
        Top = 0
        Width = 476
        Height = 526
        Align = alClient
        Caption = 'Optimize in EXCEL'
        TabOrder = 0
        object btnRun: TButton
          Left = 9
          Top = 25
          Width = 80
          Height = 32
          Caption = 'Run'
          TabOrder = 0
          OnClick = btnRunClick
        end
        object StatusBar1: TStatusBar
          Left = 2
          Top = 505
          Width = 472
          Height = 19
          Panels = <>
          SimplePanel = True
        end
        object cbPlotAll: TCheckBox
          Left = 112
          Top = 32
          Width = 289
          Height = 17
          Caption = 'Plot each chromosome in every generation'
          TabOrder = 1
        end
        object Memo1: TMemo
          Left = 2
          Top = 67
          Width = 472
          Height = 438
          Align = alBottom
          Anchors = [akLeft, akTop, akRight, akBottom]
          ScrollBars = ssVertical
          TabOrder = 2
        end
      end
    end
    object TabSheet1: TTabSheet
      Caption = 'GA Settings'
      object gbPopulation: TGroupBox
        Left = 0
        Top = 0
        Width = 476
        Height = 185
        Align = alTop
        Caption = 'Population Parameters'
        TabOrder = 0
        object Label4: TLabel
          Left = 25
          Top = 22
          Width = 230
          Height = 16
          Caption = 'Number of chromosomes in population'
        end
        object Label5: TLabel
          Left = 25
          Top = 52
          Width = 132
          Height = 16
          Caption = 'Cross-over probability'
        end
        object Label6: TLabel
          Left = 25
          Top = 148
          Width = 175
          Height = 16
          Caption = 'Random selection probability'
        end
        object Label8: TLabel
          Left = 25
          Top = 116
          Width = 200
          Height = 16
          Caption = 'Chromosome mutation probability'
        end
        object Label9: TLabel
          Left = 25
          Top = 84
          Width = 95
          Height = 16
          Caption = 'Cross-over type'
        end
        object Label10: TLabel
          Left = 368
          Top = 52
          Width = 44
          Height = 16
          Caption = '0.6 - 0.8'
        end
        object Label11: TLabel
          Left = 368
          Top = 116
          Width = 58
          Height = 16
          Caption = '0.01 - 0.02'
        end
        object Label12: TLabel
          Left = 368
          Top = 84
          Width = 58
          Height = 16
          Caption = 'One Point'
        end
        object Label24: TLabel
          Left = 368
          Top = 20
          Width = 80
          Height = 16
          Caption = '2N - 4N, even'
        end
        object Label26: TLabel
          Left = 368
          Top = 148
          Width = 17
          Height = 16
          Caption = '0.1'
        end
        object edNumChrom: TSpinEdit
          Left = 288
          Top = 17
          Width = 73
          Height = 26
          MaxValue = 999
          MinValue = 0
          TabOrder = 0
          Value = 16
        end
        object edCrossProb: TEdit
          Left = 288
          Top = 49
          Width = 73
          Height = 24
          TabOrder = 1
          Text = '0.9'
        end
        object edRandSelProb: TEdit
          Left = 288
          Top = 145
          Width = 73
          Height = 24
          TabOrder = 4
          Text = '0.1'
        end
        object edChromMutProb: TEdit
          Left = 288
          Top = 113
          Width = 73
          Height = 24
          TabOrder = 3
          Text = '0.1'
        end
        object cbCrossType: TComboBox
          Left = 240
          Top = 81
          Width = 121
          Height = 24
          Style = csDropDownList
          ItemHeight = 16
          ItemIndex = 0
          TabOrder = 2
          Text = 'One Point'
          Items.Strings = (
            'One Point'
            'Two Points'
            'Uniform'
            'Random Method')
        end
      end
      object GroupBox4: TGroupBox
        Left = 0
        Top = 276
        Width = 476
        Height = 133
        Align = alTop
        Caption = 'Main Run Parameters'
        TabOrder = 2
        object Label7: TLabel
          Left = 25
          Top = 22
          Width = 164
          Height = 16
          Caption = 'Max. number of generations'
        end
        object Label15: TLabel
          Left = 25
          Top = 53
          Width = 227
          Height = 16
          Caption = 'Convergence tolerance (on Avg. Dev.)'
        end
        object Label25: TLabel
          Left = 25
          Top = 83
          Width = 151
          Height = 16
          Caption = 'Numeric precision (digits)'
        end
        object Label28: TLabel
          Left = 40
          Top = 104
          Width = 296
          Height = 16
          Caption = '(rounds variables and constraints to this precision)'
        end
        object edMaxGen: TSpinEdit
          Left = 288
          Top = 17
          Width = 73
          Height = 26
          MaxValue = 999
          MinValue = 0
          TabOrder = 0
          Value = 100
        end
        object edConvTol: TEdit
          Left = 288
          Top = 49
          Width = 73
          Height = 24
          TabOrder = 1
          Text = '1e-5'
        end
        object edNumPrec: TSpinEdit
          Left = 288
          Top = 78
          Width = 73
          Height = 26
          MaxValue = 16
          MinValue = 1
          TabOrder = 2
          Value = 12
        end
      end
      object GroupBox7: TGroupBox
        Left = 0
        Top = 409
        Width = 476
        Height = 117
        Align = alClient
        Caption = 'Preliminary Run (optional) Parameters'
        TabOrder = 3
        object Label13: TLabel
          Left = 25
          Top = 22
          Width = 159
          Height = 16
          Caption = 'Number of preliminary runs'
        end
        object Label14: TLabel
          Left = 25
          Top = 54
          Width = 208
          Height = 16
          Caption = 'Max. number of generations per run'
        end
        object edPrelimRuns: TSpinEdit
          Left = 288
          Top = 17
          Width = 73
          Height = 26
          MaxValue = 999
          MinValue = 0
          TabOrder = 0
          Value = 4
        end
        object edMaxPrelimGenes: TSpinEdit
          Left = 288
          Top = 49
          Width = 73
          Height = 26
          MaxValue = 999
          MinValue = 0
          TabOrder = 1
          Value = 10
        end
      end
      object GroupBox8: TGroupBox
        Left = 0
        Top = 185
        Width = 476
        Height = 91
        Align = alTop
        Caption = 'Constraint Parameters'
        TabOrder = 1
        object Label22: TLabel
          Left = 25
          Top = 53
          Width = 172
          Height = 16
          Caption = 'Absolute constraint tolerance'
        end
        object Label29: TLabel
          Left = 24
          Top = 21
          Width = 106
          Height = 16
          Caption = 'Constraint penalty'
        end
        object edConstTol: TEdit
          Left = 288
          Top = 49
          Width = 73
          Height = 24
          TabOrder = 2
          Text = '0'
        end
        object edConstPen: TEdit
          Left = 288
          Top = 17
          Width = 73
          Height = 24
          TabOrder = 1
          Text = '1e12'
        end
        object cbConstPenType: TComboBox
          Left = 192
          Top = 17
          Width = 89
          Height = 24
          Style = csDropDownList
          ItemHeight = 16
          ItemIndex = 0
          TabOrder = 0
          Text = 'Fixed'
          Items.Strings = (
            'Fixed'
            'Linear'
            'Exp. 1'
            'Exp. 2')
        end
      end
    end
  end
  object MainMenu1: TMainMenu
    Left = 372
    Top = 3
    object File1: TMenuItem
      Caption = '&File'
      object LoadDefinition1: TMenuItem
        Caption = '&Load Model...'
        ShortCut = 16460
        OnClick = LoadDefinition1Click
      end
      object SaveOpenModel1: TMenuItem
        Caption = '&Save Model'
        Enabled = False
        ShortCut = 16467
        OnClick = SaveOpenModel1Click
      end
      object SaveDefinition1: TMenuItem
        Caption = 'Save Model &As...'
        OnClick = SaveDefinition1Click
      end
      object Exit1: TMenuItem
        Caption = 'E&xit'
        OnClick = Exit1Click
      end
    end
    object Analysis1: TMenuItem
      Caption = '&Analysis'
      object Run1: TMenuItem
        Caption = '&Run'
        ShortCut = 113
        OnClick = btnRunClick
      end
    end
    object Help1: TMenuItem
      Caption = '&Help'
      object About1: TMenuItem
        Caption = 'A&bout'
        ShortCut = 112
        OnClick = About1Click
      end
    end
  end
  object OpenDialog1: TOpenDialog
    Filter = 'Excel file (*.xls)|*.xls|Text File (*.txt)|*.txt'
    Options = [ofHideReadOnly, ofNoChangeDir, ofEnableSizing]
    Left = 400
  end
  object SaveDialog1: TSaveDialog
    DefaultExt = '*.txt'
    Filter = 'Excel file (*.xls)|*.xls|Text File (*.txt)|*.txt'
    Options = [ofHideReadOnly, ofNoChangeDir, ofEnableSizing]
    Left = 432
  end
end
