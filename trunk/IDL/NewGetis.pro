PRO NEWGETIS, event
  
  ; Use the ENVI dialog box to select a file
  ENVI_SELECT, fid=file,dims=dims,pos=pos
  
  ; If the dialog box was cancelled then stop the procedure
  IF file[0] EQ -1 THEN RETURN
  
  ; Create dialog box window
  TLB = WIDGET_AUTO_BASE(title="Create Getis Image")
  
  ; Create dropdown list to select distance value
  list = ['d = 1 (3x3 square)', 'd = 2 (5x5 square)', 'd = 3 (7x7 square)']
  distance = WIDGET_PMENU(TLB, /AUTO_MANAGE, list=list, uvalue='d')
  
  ; Create the widget to let the user select file or memory output
  FileOrMem = WIDGET_OUTFM(TLB, /AUTO_MANAGE, uvalue='fm')
  
  ; Start the automatic management of the window
  result = AUTO_WID_MNG(TLB) 
  
  ; If the OK button was pressed
  IF result.accept EQ 0 THEN RETURN
  
  ; Get the details of the file, ready to write to the disk if needed
  ENVI_FILE_QUERY, file, fname=fname, data_type=data_type, xstart=xstart, $
    ystart=ystart, INTERLEAVE=interleave
    
  ; Get the map info of the file so that we can output it to the new file
  map_info = ENVI_GET_MAP_INFO(FID=file)
  
  ; Initialise the progress bar window - differently depending if the output is
  ; to memory or to file
  IF result.fm.in_memory EQ 1 THEN BEGIN
    ENVI_REPORT_INIT, ['Input File: ' + fname, 'Output to memory'], title='Getis status', base=base, /INTERRUPT
  ENDIF ELSE BEGIN
    ENVI_REPORT_INIT, ['Input File: ' + fname, 'Output File: ' + result.fm.name], title='Getis status', base=base, /INTERRUPT
  ENDELSE
  
  ; Call the function to create the Getis image
  GetisImage = CREATE_GETIS_IMAGE(file, dims, pos, result.d + 1, base)
  
  ; Create classification image for top 20% of data
   
  NumRows = (dims[2] - dims[1]) + 1
  NumCols = (dims[4] - dims[3]) + 1
  
  ClassImage = INTARR(NumRows, NumCols)
  Range = MAX(GetisImage) - MIN(GetisImage)
  
  MaxValue = MIN(GetisImage) + (0.7 * Range)
  indices = WHERE(GetisImage GT MaxValue)
  ClassImage[indices] = 1
  MaxValue = MIN(GetisImage) + (0.8 * Range)
  indices = WHERE(GetisImage GT MaxValue)
  ClassImage[indices] = 2
  MaxValue = MIN(GetisImage) + (0.9 * Range)
  indices = WHERE(GetisImage GT MaxValue)
  ClassImage[indices] = 3
  
  LookupArray = [ [255, 255, 255], [255, 0, 0], [0, 255, 0], [0, 0, 255] ]
  
  FileType = ENVI_FILE_TYPE("ENVI Classification")
  
  ENVI_ENTER_DATA, ClassImage, NUM_CLASSES=4, FILE_TYPE=FileType, LOOKUP=LookupArray, CLASS_NAMES=["Unclassified", '70%', '80%', '90%']
  
  
  IF result.fm.in_memory EQ 1 THEN BEGIN
    ; If the user wanted the result to go to memory then just output it there
    ENVI_ENTER_DATA, GetisImage
  ENDIF ELSE BEGIN
    ; If the output is to file then open the file, write the binary data
    ; and close the file
    OpenW, unit, result.fm.name, /GET_LUN
    WriteU, unit, GetisImage
    FREE_LUN, unit
    
    ; Then calculate the values needed to create the header file, and create it
    NSamples = dims[2] - dims[1] + 1
    NLines = dims[4] - dims[3] + 1
    NBands = N_ELEMENTS(pos)
    ENVI_SETUP_HEAD, FNAME=result.fm.name, NS=NSamples, NL=NLines, NB=NBands, $
      DATA_TYPE=4, offset=0, INTERLEAVE=interleave, $
      XSTART=xstart+dims[1], YSTART=ystart+dims[3], $
      DESCRIP="Getis Image Output", MAP_INFO=map_info, /OPEN, /WRITE
  ENDELSE
  
 
END

; Creates a Getis image given a FID, the dimensions of the file, a distance to use for the getis routine
; and a base window to send progress updates to
FUNCTION CREATE_GETIS_IMAGE, file, dims, pos, distance, report_base
  NumRows = dims[2] - dims[1]
  NumCols = dims[4] - dims[3]
  
  NumPos = N_ELEMENTS(pos)
  
  ; Let the progress bar know how many bands we're dealing with (denom. of fraction)
  ENVI_REPORT_INC, report_base, NumPos
  
  FOR CurrPos = 0, NumPos - 1 DO BEGIN
    ; Send an update to the progress window telling it to let us know if cancel has been pressed
    ENVI_REPORT_STAT, report_base, CurrPos, NumPos
    
    ; Get the data for the current band
    WholeBand = ENVI_GET_DATA(fid=file, dims=dims, pos=pos[CurrPos])
    
    ; Get the global mean
    GlobMean = MEAN(WholeBand)
    
    ; Get the global variance
    GlobVariance = VARIANCE(WholeBand)
    
    ; Get the number of values in the whole image
    GlobNumber = NumRows * NumCols
    
    ; Converts a distance to the length of each side of the square
    ; Eg. A distance of 1 to a length of 3
    DimOfArray = (distance * 2) + 1
    
    NumOfElements = DimOfArray * DimOfArray
    
    ; Create the kernel for the summing CONVOL operation
    Kernel = FLTARR(DimOfArray, DimOfArray)
    Kernel = Kernel + 1
    
    ; Create an image where each element is the sum of the elements within
    ; d around it
    SummedImage = CONVOL(FLOAT(WholeBand), Kernel, /CENTER, /EDGE_ZERO)
    
    ; Create an image where each element is the result of the top fraction part
    ; of the getis formula
    TopFraction = SummedImage - (FLOAT(NumOfElements) * GlobMean)
    
    ; Calculate the square root bit of the formula and then create a single variable
    ; with the bottom fraction part of the formula (this is constant for all pixels)
    SquareRootAnswer = SQRT((FLOAT(NumOfElements) * (GlobNumber - NumOfElements))/(GlobNumber - 1))
    BottomFraction = GlobVariance * SquareRootAnswer
    
    ; Create an image with the getis values in it
    Getis = FLOAT(TopFraction) / BottomFraction
    
    ; If it's the first time then copy the Getis result to OutputArray,
    ; if not then append it to the end of OutputArray
    IF (CurrPos EQ 0) THEN OutputArray = Getis ELSE OutputArray = [ [[OutputArray]], [[Getis]] ]
  ENDFOR
  
  ; Close the progress window
  ENVI_REPORT_INIT,base=report_base, /FINISH
  
  RETURN, OutputArray
END