FUNCTION GET_DATA_IMAGE, fid, dims, pos
  FOR i = 1, N_ELEMENTS(pos) - 1 DO BEGIN
    new_band = ENVI_GET_DATA(fid=fid, dims=dims, pos=pos[i])
    IF N_ELEMENTS(data_image) EQ 0 THEN data_image = new_band ELSE data_image = [ [[data_image]], [[new_band]] ]
  ENDFOR
  
  return, data_image
END