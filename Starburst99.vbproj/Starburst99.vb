Imports System.IO


Public Class Starburst99
  Dim ps As System.Diagnostics.Process
  Dim startupDirName = Application.StartupPath
  Dim programDirName = startupDirName
  Dim programDir As New DirectoryInfo(programDirName)
  Dim programName = "galaxy.exe"

  Dim modelsDirName = startupDirName + "\models"
  Dim modelDir As DirectoryInfo

  Dim inputDirName = "Input"
  Dim inputDir As DirectoryInfo

  Dim inputFileName = "fort.1"
  Dim inputFile As FileInfo

  Dim outputDirName = "Output"
  Dim outputDir As DirectoryInfo

  Dim helpDirName = startupDirName + "\help"
  Dim helpFileName = "help.htm"
  Dim registerFileName = "Registration.url"


  Private Sub btnRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRun.Click

    'MessageBox.Show(Application.StartupPath)

    'Validate inputs
    If ValidateInputs() Then
      'Create inputfile under program directory

      ParaTabs.SelectTab(4)
      SystemMessage.Text = "Starburst99 simulation is running..."

      'System.Threading.Thread.Sleep(1000)

      Me.WindowState = System.Windows.Forms.FormWindowState.Minimized

      'First, create input/output folders as well as input file
      CreateFolders()

      'Dim i As Integer
      'i = Shell("c:\huang2\test.bat", AppWinStyle.Hide, True, )
      'MessageBox.Show(i)

      'ps = Process.Start("c:\Huang2\test.bat")
      'MessageBox.Show(ps.Id)

      Dim myProcess As System.Diagnostics.Process = New System.Diagnostics.Process()
            myProcess.StartInfo.FileName = programDir.FullName + "\" + programName
            myProcess.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
            myProcess.Start()
            myProcess.WaitForExit()
            myProcess.Close()

      CopyOutputfiles(programDir, outputDir)

      Me.WindowState = System.Windows.Forms.FormWindowState.Normal

      SystemMessage.Text = "Starburst99 simulation completed." & ControlChars.Lf & _
           "Please click ""Browse the Files"" Button to view results."
      ParaTabs.SelectTab(4)

    End If



  End Sub

  Private Sub Starburst99_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    'MessageBox.Show("before killing")
    Dim plist As Process() = Process.GetProcesses()
    For Each p As Process In plist
      Try
        If p.MainModule.ModuleName.ToUpper() = "GALAXY.EXE" Then p.Kill()
      Catch
        'seems listing modules for some processes fails, so better ignore any exceptions here
      End Try
    Next p
  End Sub

  Private Sub Starburst99_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    SystemMessage.Text = "Ready to run..."
  End Sub

  Function ValidateInputs() As Boolean
    Dim errorMessage As New String("")

    If ModelName.Text = "" Then
      errorMessage = AppendMessage(errorMessage, "Model Name can NOT be empty. ")
    End If

    errorMessage = RangeValidation(errorMessage, SFR.Text, 0, 10000, _
     "Input for Star Formation Rate is outside the supported range!")

    errorMessage = RangeValidation(errorMessage, TOMA.Text, 0, 100000000.0, _
     "Input for Total Stellar Mass is outside the supported range!")

    errorMessage = RangeValidation(errorMessage, NINTERV.Text, 0, 11, _
     "Input for Number of IMF Intervals is outside the supported range!")

    errorMessage = ArrayRangeValidation(errorMessage, XPONENT.Text, 0, 11, _
     "Input for IMF Exponents is outside the supported range!")

    errorMessage = ArrayRangeValidation(errorMessage, XMASLIM.Text, 0, 999, _
      "Input for Mass Boundaries for IMF is outside the supported range!")

    errorMessage = RangeValidation(errorMessage, SnCut.Text, 0, 999, _
      "Input for Supernova Cut-Off Mass is outside the supported range!")

    errorMessage = RangeValidation(errorMessage, BhCut.Text, 0, 999, _
      "Input for Black Hole Cut-Off Mass is outside the supported range!")

    errorMessage = RangeValidation(errorMessage, Time1.Text, 0, 999, _
      "Input for Initial Time is outside the supported range!")

    errorMessage = RangeValidation(errorMessage, TBIV.Text, 0, 999, _
      "Input for Time Step (if linear) is outside the supported range!")

    errorMessage = RangeValidation(errorMessage, ITBIV.Text, 0, 999, _
      "Input for Number of Steps is outside the supported range!")

    errorMessage = RangeValidation(errorMessage, TMax.Text, 0, 999, _
      "Input for Last Grid Point is outside the supported range!")

    If (LorAll1.Checked) Then
      errorMessage = RangeValidation(errorMessage, Lmin.Text, 0, 999, _
        "Input for Selected Tracks is outside the supported range!")
      errorMessage = RangeValidation(errorMessage, Lmax.Text, 0, 999, _
        "Input for Selected Tracks is outside the supported range!")
    End If

    errorMessage = RangeValidation(errorMessage, TDel.Text, 0, 999, _
      "Input for Time Step to Print Spectra is outside the supported range!")


    If errorMessage = "" Then
      Return True
    Else
      MessageBox.Show("Validation error. Please check Runtime Message tab for more details.")
      SystemMessage.Text = errorMessage
      ParaTabs.SelectTab(4)
      Return False
    End If

  End Function

  Private Sub PrepareInputFile(ByVal inFileName As String)

    'Dim oFile As System.IO.File
    Dim oWrite As System.IO.StreamWriter
    oWrite = File.CreateText(inFileName)
    'oWrite = oFile.CreateText("C:\sample.txt")

    oWrite.WriteLine("MODEL DESIGNATION:                                           [NAME]")
    oWrite.WriteLine(ModelName.Text)

    oWrite.WriteLine("CONTINUOUS STAR FORMATION (>0) OR FIXED MASS (<=0):          [ISF]")
    If ISF1.Checked Then
      oWrite.WriteLine(1)
    ElseIf ISF2.Checked Then
      oWrite.WriteLine(-1)
    End If

    oWrite.WriteLine("TOTAL STELLAR MASS [10e6 SOLAR MASSES] IF 'FIXED MASS' IS CHOSEN: [TOMA]")
    If TOMA.Text = "" Then
      TOMA.Text = "1.0"
    End If
        oWrite.WriteLine(ConvertTextIntoReal(TOMA.Text))
        'oWrite.WriteLine(CDec(TOMA.Text).ToString("N3"))

    oWrite.WriteLine("SFR [SOLAR MASSES PER YEAR] IF 'CONT. SF' IS CHOSEN:         [SFR]")
    If SFR.Text = "" Then
      SFR.Text = "1.0"
    End If
        oWrite.WriteLine(ConvertTextIntoReal(SFR.Text))
        'oWrite.WriteLine(CDec(SFR.Text).ToString("N3"))

    oWrite.WriteLine("NUMBER OF INTERVALS FOR THE IMF (KROUPA=2):                  [NINTERV]")
    If NINTERV.Text = "" Then
      NINTERV.Text = "2"
    End If
    oWrite.WriteLine(NINTERV.Text)

    oWrite.WriteLine("IMF EXPONENTS (KROUPA=1.3,2.3):                              [XPONENT]")
    If XPONENT.Text = "" Then
      XPONENT.Text = "1.3,2.3"
    End If
    'oWrite.WriteLine(XPONENT.Text)
    oWrite.WriteLine(SplitAndFormat(XPONENT.Text, 3))

    oWrite.WriteLine("MASS BOUNDARIES FOR IMF (KROUPA=0.1,0.5,100) [SOLAR MASSES]: [XMASLIM]")
    If XMASLIM.Text = "" Then
      XMASLIM.Text = "0.1,0.5,100."
    End If
    'Might need to add code here to loop through the values and reformat with the right number of decimal points.
    'oWrite.WriteLine(XMASLIM.Text)
    oWrite.WriteLine(SplitAndFormat(XMASLIM.Text, 3))


    oWrite.WriteLine("SUPERNOVA CUT-OFF MASS [SOLAR MASSES]:                       [SNCUT]")
    If SnCut.Text = "" Then
      SnCut.Text = "8."
    End If
        oWrite.WriteLine(ConvertTextIntoReal(SnCut.Text))
        'oWrite.WriteLine(CDec(SnCut.Text).ToString("N3"))

    oWrite.WriteLine("BLACK HOLE CUT-OFF MASS [SOLAR MASSES]:                      [BHCUT]")
    If BhCut.Text = "" Then
      BhCut.Text = "120."
    End If
        oWrite.WriteLine(ConvertTextIntoReal(BhCut.Text))
        'oWrite.WriteLine(CDec(BhCut.Text).ToString("N3"))

    oWrite.WriteLine("METALLICITY + TRACKS:                                        [IZ]")
    oWrite.WriteLine("GENEVA STD: 11=0.001;  12=0.004; 13=0.008; 14=0.020; 15=0.040")
    oWrite.WriteLine("GENEVA HIGH:21=0.001;  22=0.004; 23=0.008; 24=0.020; 25=0.040")
    oWrite.WriteLine("PADOVA STD: 31=0.0004; 32=0.004; 33=0.008; 34=0.020; 35=0.050")
    oWrite.WriteLine("PADOVA AGB: 41=0.0004; 42=0.004; 43=0.008; 44=0.020; 45=0.050")
    Dim tempIZ As Double
    If MLType1.Checked Then
      tempIZ = 10 + MLoss1.SelectedIndex + 1
    ElseIf MLType2.Checked Then
      tempIZ = 20 + MLoss2.SelectedIndex + 1
    ElseIf MLType3.Checked Then
      tempIZ = 30 + MLoss3.SelectedIndex + 1
    ElseIf MLType4.Checked Then
      tempIZ = 40 + MLoss4.SelectedIndex + 1
    End If
    oWrite.WriteLine(tempIZ)

    oWrite.WriteLine("WIND MODEL (0: MAEDER; 1: EMP.; 2: THEOR.; 3: ELSON):        [IWIND]")
    oWrite.WriteLine(IWind.SelectedIndex)

    oWrite.WriteLine("INITIAL TIME [1.E6 YEARS]:                                   [TIME1]")
    If Time1.Text = "" Then
      Time1.Text = "0.01"
    End If
        oWrite.WriteLine(ConvertTextIntoReal(Time1.Text))
        'MessageBox.Show("F: " + CDec(Time1.Text).ToString("F"))
        'MessageBox.Show("F2: " + CDec(Time1.Text).ToString("F2"))
        'MessageBox.Show("F3: " + CDec(Time1.Text).ToString("F3"))
        'MessageBox.Show("F4: " + CDec(Time1.Text).ToString("F4"))
        'MessageBox.Show("Empty: " + CDec(Time1.Text).ToString())
        'MessageBox.Show("Custom: " + ConvertTextIntoReal(Time1.Text))
        'oWrite.WriteLine(CDec(Time1.Text).ToString("N3"))

    oWrite.WriteLine("TIME SCALE: LINEAR (=0) OR LOGARITHMIC (=1)                  [JTIME]")
    oWrite.WriteLine(JTIME.SelectedIndex)

    oWrite.WriteLine("TIME STEP [1.e6 YEARS] (ONLY USED IF JTIME=0):               [TBIV]")
    If TBIV.Text = "" Then
      TBIV.Text = "0.1"
    End If
        oWrite.WriteLine(ConvertTextIntoReal(TBIV.Text))
        'oWrite.WriteLine(CDec(TBIV.Text).ToString("N3"))

    oWrite.WriteLine("NUMBER OF STEPS        (ONLY USED IF JTIME=1):               [ITBIV]")
    If ITBIV.Text = "" Then
      ITBIV.Text = "1000"
    End If
    oWrite.WriteLine(ITBIV.Text)

    oWrite.WriteLine("LAST GRID POINT [1.e6 YEARS]:                                [TMAX]")
    If TMax.Text = "" Then
      TMax.Text = "100"
    End If
        oWrite.WriteLine(ConvertTextIntoReal(TMax.Text))
        'oWrite.WriteLine(CDec(TMax.Text).ToString("N3"))

    oWrite.WriteLine("SMALL (=0) OR LARGE (=1) MASS GRID;")
    oWrite.WriteLine("ISOCHRONE ON  LARGE GRID (=2) OR FULL ISOCHRONE (=3):        [JMG]")
    oWrite.WriteLine(JMG.SelectedIndex)

    oWrite.WriteLine("LMIN, LMAX (ALL=0):                                          [LMIN,LMAX]")
    'Add code to process LminMax
    Dim LminMax As String
    If LorAll2.Checked Then
      LminMax = "0"
    Else
      LminMax = Lmin.Text + "," + Lmax.Text
    End If
    oWrite.WriteLine(LminMax)

    oWrite.WriteLine("TIME STEP FOR PRINTING OUT THE SYNTHETIC SPECTRA [1.e6YR]:   [TDEL]")
    If TDel.Text = "" Then
      TDel.Text = "2"
    End If
        oWrite.WriteLine(ConvertTextIntoReal(TDel.Text))
        'oWrite.WriteLine(CDec(TDel.Text).ToString("N3"))

    oWrite.WriteLine("ATMOSPHERE: 1=PLA, 2=LEJ, 3=LEJ+SCH, 4=LEJ+SMI, 5=PAU+SMI:   [IATMOS]")
    oWrite.WriteLine(IAtmos.SelectedIndex + 1)

    oWrite.WriteLine("METALLICITY OF THE HIGH RESOLUTION MODELS                    [ILIB]")
    oWrite.WriteLine("(1=0.001, 2= 0.008, 3=0.020, 4=0.040):")
    oWrite.WriteLine(ILIB.SelectedIndex + 1)

    oWrite.WriteLine("LIBRARY FOR THE UV LINE SPECTRUM: (1=SOLAR, 2=LMC/SMC)       [ILINE]")
    oWrite.WriteLine(ILINE.SelectedIndex + 1)

    oWrite.WriteLine("RSG FEATURE: MICROTURB. VEL (1-6), SOL/NON-SOL ABUND (0,1)   [IVT,IRSG]")
    oWrite.WriteLine(IVT.Text + "," + IRSG.SelectedIndex.ToString())

    oWrite.WriteLine("OUTPUT FILES (NO<0, YES>=0)                                  [IO1,...]")
    If OutFile1.Checked Then
      oWrite.Write("+1,")
    Else
      oWrite.Write("-1,")
    End If
    If Outfile2.Checked Then
      oWrite.Write("+1,")
    Else
      oWrite.Write("-1,")
    End If
    If Outfile3.Checked Then
      oWrite.Write("+1,")
    Else
      oWrite.Write("-1,")
    End If
    If Outfile4.Checked Then
      oWrite.Write("+1,")
    Else
      oWrite.Write("-1,")
    End If
    If Outfile5.Checked Then
      oWrite.Write("+1,")
    Else
      oWrite.Write("-1,")
    End If
    If Outfile6.Checked Then
      oWrite.Write("+1,")
    Else
      oWrite.Write("-1,")
    End If
    If Outfile7.Checked Then
      oWrite.Write("+1,")
    Else
      oWrite.Write("-1,")
    End If
    If Outfile8.Checked Then
      oWrite.Write("+1,")
    Else
      oWrite.Write("-1,")
    End If
    If Outfile9.Checked Then
      oWrite.Write("+1,")
    Else
      oWrite.Write("-1,")
    End If
    If Outfile10.Checked Then
      oWrite.Write("+1,")
    Else
      oWrite.Write("-1,")
    End If
    If Outfile11.Checked Then
      oWrite.Write("+1,")
    Else
      oWrite.Write("-1,")
    End If
    If Outfile12.Checked Then
      oWrite.Write("+1,")
    Else
      oWrite.Write("-1,")
    End If
    If Outfile13.Checked Then
      oWrite.Write("+1,")
    Else
      oWrite.Write("-1,")
    End If
    If Outfile14.Checked Then
      oWrite.Write("+1,")
    Else
      oWrite.Write("-1,")
    End If
    If Outfile15.Checked Then
      oWrite.WriteLine("+1")
    Else
      oWrite.WriteLine("-1")
    End If

    oWrite.WriteLine("******************************************************************")
    oWrite.WriteLine("  OUTPUT FILES:         1    SYNTHESIS.QUANTA")
    oWrite.WriteLine("                        2    SYNTHESIS.SNR")
    oWrite.WriteLine("                        3    SYNTHESIS.HRD")
    oWrite.WriteLine("                        4    SYNTHESIS.POWER")
    oWrite.WriteLine("                        5    SYNTHESIS.SP")
    oWrite.WriteLine("                        6    SYNTHESIS.YIELDS")
    oWrite.WriteLine("                        7    SYNTHESIS.SPECTRUM")
    oWrite.WriteLine("                        8    SYNTHESIS.LINE")
    oWrite.WriteLine("                        9    SYNTHESIS.COLOR")
    oWrite.WriteLine("                       10    SYNTHESIS.WIDTH")
    oWrite.WriteLine("                       11    SYNTHESIS.FEATURES")
    oWrite.WriteLine("                       12    SYNTHESIS.OVI")
    oWrite.WriteLine("                       13    SYNTHESIS.HIRES")
    oWrite.WriteLine("                       14    SYNTHESIS.WRLINES")
    oWrite.WriteLine("                       15    SYNTHESIS.IFASPEC")
    oWrite.WriteLine("")
    oWrite.WriteLine("  CORRESPONDENCE I VS. MASS:")
    oWrite.WriteLine("  M   120 115 110 105 100  95  90  85  80  75  70  65 ")
    oWrite.WriteLine("  I     1   2   3   4   5   6   7   8   9  10  11  12")
    oWrite.WriteLine("")
    oWrite.WriteLine("  M    60  55  50  45  40  35  30  25  20  15  10   5 ")
    oWrite.WriteLine("  I    13  14  15  16  17  18  19  20  21  22  23  24")
    oWrite.WriteLine("")
    oWrite.WriteLine("")
    oWrite.WriteLine("  M   120 119 118 117 116 115 114 113 112 111 110 109 108 107 106 ")
    oWrite.WriteLine("  I     1   2   3   4   5   6   7   8   9  10  11  12  13  14  15")
    oWrite.WriteLine("")
    oWrite.WriteLine("  M   105 104 103 102 101 100  99  98  97  96  95  94  93  92  91")
    oWrite.WriteLine("  I    16  17  18  19  20  21  22  23  24  25  26  27  28  29  30")
    oWrite.WriteLine("")
    oWrite.WriteLine("  M    90  89  88  87  86  85  84  83  82  81  80  79  78  77  76")
    oWrite.WriteLine("  I    31  32  33  34  35  36  37  38  39  40  41  42  43  44  45")
    oWrite.WriteLine("")
    oWrite.WriteLine("  M    75  74  73  72  71  70  69  68  67  66  65  64  63  62  61")
    oWrite.WriteLine("  I    46  47  48  49  50  51  52  53  54  55  56  57  58  59  60")
    oWrite.WriteLine("")
    oWrite.WriteLine("  M    60  59  58  57  56  55  54  53  52  51  50  49  48  47  46")
    oWrite.WriteLine("  I    61  62  63  64  65  66  67  68  69  70  71  72  73  74  75")
    oWrite.WriteLine("")
    oWrite.WriteLine("  M    45  44  43  42  41  40  39  38  37  36  35  34  33  32  31")
    oWrite.WriteLine("  I    76  77  78  79  80  81  82  83  84  85  86  87  88  89  90")
    oWrite.WriteLine("")
    oWrite.WriteLine("  M    30  29  28  27  26  25  24  23  22  21  20  19  18  17  16")
    oWrite.WriteLine("  I    91  92  93  94  95  96  97  98  99 100 101 102 103 104 105")
    oWrite.WriteLine("")
    oWrite.WriteLine("  M    15  14  13  12  11  10   9   8   7   6   5   4   3   2   1")
    oWrite.WriteLine("  I   106 107 108 109 110 111 112 113 114 115 116 117 118 119 120  ")
    oWrite.WriteLine("******************************************************************** ")


    oWrite.Close()
    Copyfile(inFileName, programDir.FullName + "\" + inputFileName)


  End Sub

  Private Sub preSelectOutFiles()
    Dim x As Integer
    For x = 0 To OutFiles.Items.Count - 1
      If x <> 2 And x <> 14 Then
        OutFiles.SelectedIndex() = x
      End If
    Next x
  End Sub

  Private Sub CreateFolders()

    modelDir = New DirectoryInfo(modelsDirName + "\" + ModelName.Text)
    modelDir.Create()
    inputDir = New DirectoryInfo(modelDir.FullName + "\" + inputDirName)
    inputDir.Create()
    'inputFile = New FileInfo(inputDir.FullName + "\" + inputFileName)
    'inputFile.Create()
    PrepareInputFile(inputDir.FullName + "\" + inputFileName)
    outputDir = New DirectoryInfo(modelDir.FullName + "\" + outputDirName)
    outputDir.Create()

  End Sub

  Private Sub CopyOutputfiles(ByVal source As DirectoryInfo, ByVal target As DirectoryInfo)
        Movefile(source.FullName + "\" + "fort.99", target.FullName + "\" + ModelName.Text + "." + "output")
        Movefile(source.FullName + "\" + "fort.99", target.FullName + "\" + ModelName.Text + "." + "output")
        Movefile(source.FullName + "\" + "fort.98", target.FullName + "\" + ModelName.Text + "." + "quanta")
        Movefile(source.FullName + "\" + "fort.97", target.FullName + "\" + ModelName.Text + "." + "snr")
        Movefile(source.FullName + "\" + "fort.96", target.FullName + "\" + ModelName.Text + "." + "hrd")
        Movefile(source.FullName + "\" + "fort.95", target.FullName + "\" + ModelName.Text + "." + "power")
        Movefile(source.FullName + "\" + "fort.94", target.FullName + "\" + ModelName.Text + "." + "sptyp1")
        Movefile(source.FullName + "\" + "fort.90", target.FullName + "\" + ModelName.Text + "." + "sptyp2")
        Movefile(source.FullName + "\" + "fort.93", target.FullName + "\" + ModelName.Text + "." + "yield")
        Movefile(source.FullName + "\" + "fort.92", target.FullName + "\" + ModelName.Text + "." + "spectrum")
        Movefile(source.FullName + "\" + "fort.91", target.FullName + "\" + ModelName.Text + "." + "uvline")
        Movefile(source.FullName + "\" + "fort.89", target.FullName + "\" + ModelName.Text + "." + "color")
        Movefile(source.FullName + "\" + "fort.88", target.FullName + "\" + ModelName.Text + "." + "ewidth")
        Movefile(source.FullName + "\" + "fort.87", target.FullName + "\" + ModelName.Text + "." + "irfeature")
        Movefile(source.FullName + "\" + "fort.86", target.FullName + "\" + ModelName.Text + "." + "ovi")
        Movefile(source.FullName + "\" + "fort.82", target.FullName + "\" + ModelName.Text + "." + "hires")
        Movefile(source.FullName + "\" + "fort.84", target.FullName + "\" + ModelName.Text + "." + "wrlines")
        Movefile(source.FullName + "\" + "fort.83", target.FullName + "\" + ModelName.Text + "." + "ifaspec")
        Movefile(source.FullName + "\" + "fort.50", target.FullName + "\" + ModelName.Text + "." + "mapspec")
        Movefile(source.FullName + "\" + "fort.66", target.FullName + "\" + ModelName.Text + "." + "log")


        'Movefile(source.FullName + "\" + "fort.99", target.FullName + "\" + ModelName.Text + "." + "output")
        Movefile(source.FullName + "\" + "fort.980", target.FullName + "\" + ModelName.Text + "_" + "quanta.xml")
        Movefile(source.FullName + "\" + "fort.970", target.FullName + "\" + ModelName.Text + "_" + "snr.xml")
        'Movefile(source.FullName + "\" + "fort.96", target.FullName + "\" + ModelName.Text + "." + "hrd.xml")
        Movefile(source.FullName + "\" + "fort.950", target.FullName + "\" + ModelName.Text + "_" + "power.xml")
        Movefile(source.FullName + "\" + "fort.940", target.FullName + "\" + ModelName.Text + "_" + "sptyp1.xml")
        'Movefile(source.FullName + "\" + "fort.90", target.FullName + "\" + ModelName.Text + "." + "sptyp2.xml")
        Movefile(source.FullName + "\" + "fort.930", target.FullName + "\" + ModelName.Text + "_" + "yield.xml")
        Movefile(source.FullName + "\" + "fort.920", target.FullName + "\" + ModelName.Text + "_" + "spectrum.xml")
        Movefile(source.FullName + "\" + "fort.910", target.FullName + "\" + ModelName.Text + "_" + "uvline.xml")
        Movefile(source.FullName + "\" + "fort.890", target.FullName + "\" + ModelName.Text + "_" + "color.xml")
        Movefile(source.FullName + "\" + "fort.880", target.FullName + "\" + ModelName.Text + "_" + "ewidth.xml")
        Movefile(source.FullName + "\" + "fort.870", target.FullName + "\" + ModelName.Text + "_" + "irfeature.xml")
        Movefile(source.FullName + "\" + "fort.860", target.FullName + "\" + ModelName.Text + "_" + "ovi.xml")
        Movefile(source.FullName + "\" + "fort.820", target.FullName + "\" + ModelName.Text + "_" + "hires.xml")
        Movefile(source.FullName + "\" + "fort.840", target.FullName + "\" + ModelName.Text + "_" + "wrlines.xml")
        Movefile(source.FullName + "\" + "fort.830", target.FullName + "\" + ModelName.Text + "_" + "ifaspec.xml")
        'Movefile(source.FullName + "\" + "fort.50", target.FullName + "\" + ModelName.Text + "." + "mapspec.xml")
        'Movefile(source.FullName + "\" + "fort.66", target.FullName + "\" + ModelName.Text + "." + "log")







    '/bin/cp fort.1 	$1.input$2
    '/bin/mv fort.99 $1.output$2
    '/bin/mv fort.98 $1.quanta$2
    '/bin/mv fort.97 $1.snr$2
    '/bin/mv fort.96 $1.hrd$2
    '/bin/mv fort.95 $1.power$2
    '/bin/mv fort.94 $1.sptyp1$2
    '/bin/mv fort.90 $1.sptyp2$2
    '/bin/mv fort.93 $1.yield$2
    '/bin/mv fort.92 $1.spectrum$2
    '/bin/mv fort.91 $1.uvline$2
    '/bin/mv fort.89 $1.color$2
    '/bin/mv fort.88 $1.ewidth$2
    '/bin/mv fort.87 $1.irfeature$2
    '/bin/mv fort.86 $1.ovi$2
    '/bin/mv fort.82 $1.hires$2
    '/bin/mv fort.84 $1.wrlines$2
    '/bin/mv fort.83 $1.ifaspec$2
    '/bin/mv fort.50 $1.mapspec$2

  End Sub

  Private Sub Copyfile(ByVal sourceFilePath As String, ByVal targetFilePath As String)
    Dim fFile1 As New FileInfo(sourceFilePath)
    If fFile1.Exists Then
            fFile1.CopyTo(targetFilePath, True)
        End If
    End Sub

    Private Sub Movefile(ByVal sourceFilePath As String, ByVal targetFilePath As String)
        Dim fFile1 As New FileInfo(sourceFilePath)
        If fFile1.Exists Then
            fFile1.CopyTo(targetFilePath, True)
            fFile1.Delete()
        End If
    End Sub

  Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
    Dim openFileDialog1 As New OpenFileDialog()
    openFileDialog1.InitialDirectory = modelsDirName

    'By default *.* file 
    openFileDialog1.Filter = "All Files (*.*)|*.*"
    openFileDialog1.Title = "Open a simulation file---> *.*"

    'if file is selected 
    While openFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.OK
      Shell("notepad.exe " & openFileDialog1.FileName, AppWinStyle.NormalFocus)
    End While

  End Sub

  Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
    'Reset model name
    ModelName.Text = ""

    'Reset tab "Star Formation Law"
    ISF2.Checked = True
    SFR.Text = "1.0"
    TOMA.Text = "1.0"
    NINTERV.Text = "2"
    XPONENT.Text = "1.3,2.3"
    XMASLIM.Text = "0.1,0.5,100."
    SnCut.Text = "8."
    BhCut.Text = "120."

    'Reset tab "Metallicity and Tracks"
    MLType4.Checked = True
    MLoss1.SelectedItem = "0.020"
    MLoss2.SelectedItem = "0.020"
    MLoss3.SelectedItem = "0.020"
    MLoss4.SelectedItem = "0.020"
    IWind.SelectedItem = "Evolution"
    Time1.Text = "0.01"
    JTIME.SelectedItem = "Linear"
    TBIV.Text = "0.1"
    ITBIV.Text = "1000"
        TMax.Text = "50"
    JMG.SelectedItem = "Full Isochrone"
    LorAll2.Checked = True
    Lmin.Text = "100"
    Lmax.Text = "100"

    'Reset tab "Atmospheres and Spectra"
    TDel.Text = "2"
    IAtmos.SelectedItem = "Pauldrach/Hillier"
    ILIB.SelectedItem = "0.020"
    ILINE.SelectedItem = "Solar"
    IVT.SelectedItem = "3"
    IRSG.SelectedItem = "Solar"

    'Reset tab "Output Files"
    OutFile1.Checked = True
    Outfile2.Checked = True
    Outfile3.Checked = False
    Outfile4.Checked = True
    Outfile5.Checked = True
    Outfile6.Checked = True
    Outfile7.Checked = True
    Outfile8.Checked = True
    Outfile9.Checked = True
    Outfile10.Checked = True
    Outfile11.Checked = True
    Outfile12.Checked = True
    Outfile13.Checked = True
    Outfile14.Checked = True
        Outfile15.Checked = True
  End Sub

  Function AppendMessage(ByVal message As String, ByVal addition As String) As String
    Dim newMessage As String
    newMessage = message + addition + ControlChars.Lf
    Return newMessage
  End Function

  Function RangeValidation(ByVal errorMessage As String, ByVal value As String, ByVal lowerRange As Integer, ByVal upperRange As Integer, ByVal newMessage As String) As String

    If (value = "") OrElse (Not IsNumeric(value)) OrElse _
   (Not (lowerRange = 999) And value <= lowerRange) OrElse _
   (Not (upperRange = 999) And value >= upperRange) Then
      errorMessage = AppendMessage(errorMessage, newMessage)
    End If
    Return errorMessage

  End Function

  Function ArrayRangeValidation(ByVal errorMessage As String, ByVal values As String, ByVal lowerRange As Integer, ByVal upperRange As Integer, ByVal newMessage As String) As String


    Dim valueArray As String() = Nothing
    valueArray = values.Split(",")
    Dim s As String
    For Each s In valueArray
      errorMessage = RangeValidation(errorMessage, s, lowerRange, upperRange, newMessage)
    Next s

    Return errorMessage

  End Function

  Function SplitAndFormat(ByVal values As String, ByVal precision As Integer) As String
    Dim result As String
    Dim Sprecision As String
    Sprecision = "N" + precision.ToString
    result = ""
    Dim valueArray As String() = Nothing
    valueArray = values.Split(",")
    Dim s As String
    For Each s In valueArray
            result = result + "," + CDec(s).ToString(Sprecision)
    Next
    result = result.Remove(0, 1)

    Return result
  End Function

    Function ConvertTextIntoReal(ByVal inText As String) As String
        If inText.IndexOf(".") = -1 Then
            inText = inText + ".00"
        End If
        Return inText
    End Function

  Private Sub ViewHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles viewHelp.Click
    Process.Start(helpDirName + "\" + helpFileName)
  End Sub

  Private Sub register_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles register.Click
    Process.Start(helpDirName + "\" + registerFileName)
  End Sub


    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub
End Class
