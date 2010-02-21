PRO ADD_ZEROES_BAND
  ENVI_SELECT, fid=fid, dims=dims
  
  ENVI_FILE_QUERY, fid, nb=nb, ns=ns, nl=nl
  
  FOR i = 0, nb-1 DO BEGIN
    band = ENVI_GET_DATA(fid=fid, pos=i, dims=dims)
    
    if N_ELEMENTS(WholeImage) EQ 0 THEN WholeImage = band ELSE WholeImage = [ [[WholeImage]], [[band]] ]
  ENDFOR
  
  Zeroes = fltarr(nl, ns)
  
  ENVI_ENTER_DATA, [ [[WholeImage]], [[Zeroes]] ]
END