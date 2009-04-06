PRO GETIS, event
  
  ; Use the ENVI dialog box to select a file
  ENVI_SELECT, fid=file,dims=dims,pos=pos
  
  ; If the dialog box was cancelled then stop the procedure
  IF file[0] EQ -1 THEN RETURN
  
  ; Create dialog box window
  TLB = WIDGET_AUTO_BASE(title="Create Getis Image")
  
  ; Create dropdown list to select distance value
  list = ['d = 1 (3x3 square)', 'd = 2 (5x5 square)', 'd = 3 (7x7 square)']
  distance = WIDGET_PMENU(TLB, /AUTO_MANAGE, list=list, uvalue='d')
  
  ; Start the automatic management of the window
  result = AUTO_WID_MNG(TLB) 
  
  ; If the OK button was pressed
  IF result.accept EQ 0 THEN RETURN
  
  ; Get the filename of the selected file
  ENVI_FILE_QUERY, file, fname=fname
  
  ; Initialise the progress bar window
  ENVI_REPORT_INIT, ['Input File: ' + fname, 'Output File: filename'], title='Getis status', base=base, /INTERRUPT
  
  ; Call the function to create the Getis image
  CREATE_GETIS_IMAGE, file, dims, result.d + 1, base
END

; Creates a Getis image given a FID, the dimensions of the file, a distance to use for the getis routine
; and a base window to send progress updates to
PRO CREATE_GETIS_IMAGE, file, dims, distance, report_base

  ; Print debugging info
  print, "Distance", distance

  ; TODO: Get this to loop through bands
  ; Get the data for the first band of the file (ignores pos from earlier)
  WholeBand = ENVI_GET_DATA(fid=file, dims=dims, pos=0)
  
  ; Calculate the dimensions of WholeBand
  SizeInfo = SIZE(WholeBand, /DIMENSIONS)
  NumRows = SizeInfo[0]
  NumCols = SizeInfo[1]
  
  ; Let the progress bar know what the denominator of the fraction is - the max
  ENVI_REPORT_INC, report_base, NumRows
  
  ; --- Calculate variable values for the WholeBand
  
  ; Get the global mean
  GlobMean = MEAN(WholeBand)
  
  ; Get the global variance
  GlobVariance = VARIANCE(WholeBand)
  
  ; Get the number of values in the whole image
  GlobNumber = NumRows * NumCols
  
  ; Create the output array to store the Getis values in - NB: Must be an array of floats
  OutputArray = FLTARR(NumRows, NumCols)
  
  ; For each pixel in the image
  FOR Rows = 0, NumRows - 1 DO BEGIN
    
    ; Send an update to the progress window telling it to let us know if cancel has been pressed
    ENVI_REPORT_STAT, report_base, Rows, NumRows - 1, cancel=cancelled
    
    ; If cancel has been pressed then...
    IF cancelled EQ 1 THEN BEGIN
      ; Close the progress window
      ENVI_REPORT_INIT,base=report_base, /FINISH
      ; Exit the function
      RETURN
    ENDIF
    FOR Cols = 0, NumCols - 1 DO BEGIN
      ; Make sure RowBottom doesn't go below 0
      RowBottom = Rows - Distance
      IF RowBottom LT 0 THEN RowBottom = 0
      
      ; Make sure RowTop doesn't go above NumRows
      RowTop = Rows + Distance
      IF RowTop GE NumRows THEN RowTop = NumRows - 1
      
      ColBottom = Cols - Distance
      IF ColBottom LT 0 THEN ColBottom = 0
      
      ColTop = Cols + Distance
      IF ColTop GE NumCols THEN ColTop = (NumCols - 1)
      
      ; Get the subset of the image corresponding to the Area of Interest (AOI)
      AOI = WholeBand[RowBottom:RowTop, ColBottom:ColTop]
      
      ; Calculate the getis value for this AOI
      getis = CALCULATE_GETIS(GlobMean, GlobVariance, GlobNumber, AOI)
           
      ; Set the pixel in the output image equal to the getis value
      OutputArray[Rows, Cols] = getis
    ENDFOR
  ENDFOR
  
  ; Code to scale 0-255 - not used at the moment
  
  ;MaxOutputArray = MAX(OutputArray)
  ;MinOutputArray = MIN(OutputArray)
  ;RangeOutputArray = MaxOutputArray - MinOutputArray
  
  ;print, MaxOutputArray
  ;print, MinOutputArray
  ;print, RangeOutputArray
  
  ;ScaledArray = OutputArray - MinOutputArray
  ;ScaledArray = ScaledArray * (255 / MaxOutputArray - MinOutputArray)
  
  ; Write the data to an image in ENVI memory
  ENVI_ENTER_DATA, OutputArray
  
  ; Close the progress window
  ENVI_REPORT_INIT,base=report_base, /FINISH
END

; Calculates the getis value for an AOI given the AOI as an array, and various
; values of global constants - Mean, Variance and Number of pixels
FUNCTION CALCULATE_GETIS, GlobMean, GlobVariance, GlobNumber, AOI
  ; --- Calculate variable values for the AOI
  
  ; Get the Sum of the values in the AOI
  AOISum = TOTAL(aoi)
  
  ; Get number of values in AOI
  SizeInfo = SIZE(aoi, /DIMENSIONS)
  SizeOfSize = SIZE(SizeInfo, /DIMENSIONS)
  IF SizeOfSize EQ 2 THEN AOINumber = SizeInfo[0] * SizeInfo[1]
  IF SizeOfSize EQ 1 THEN AOINumber = SizeInfo[0]
  
  ; --- Start Calculating Getis Statistic
  
  ; Calculate the top of the fraction
  TopFraction = AOISum - (FLOAT(AOINumber) * GlobMean)
  
  ; Calculate the square root
  SquareRootAnswer = SQRT((FLOAT(AOINumber) * (GlobNumber - AOINumber))/(GlobNumber - 1))
  
  ; Calculate bottom of fraction
  BottomFraction = GlobVariance * SquareRootAnswer
  
  ; Calculate Getis Statistic
  Getis = TopFraction / BottomFraction
  
  RETURN, Getis
END