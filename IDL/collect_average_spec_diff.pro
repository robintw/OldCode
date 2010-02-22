; Give this function the x and y co-ords got from the display group
FUNCTION GET_Z_PROFILE, fid, x, y
  ENVI_FILE_QUERY, fid, ns=ns, nl=nl, nb=nb, dims=dims, xstart=xstart, ystart=ystart
  
  new_x = x - xstart - 1
  new_y = y - ystart - 1
  
  data = ENVI_GET_SLICE(fid=fid, /bip, pos=indgen(nb), line=new_y, xs=new_x, xe=new_x)
  
  return, data
END

FUNCTION COLLECT_SPECTRA
  
END

PRO COLLECT_AVERAGE_SPEC_DIFF
  ;COMMON, SPECTRA_ARRAYS, CASI_spectra_array, SPOT_spectra_array, CASI_i, SPOT_i
  ; Select CASI file
  ENVI_SELECT, fid=CASI_fid, title="Select CASI file"
  
  ENVI_SELECT, fid=SPOT_fid, title="Select SPOT file"
  
  display = envi_get_display_numbers()
  if (display[0] eq -1) then return
  
  
  
  disp_get_location, display[0], xloc, yloc 
  CASI_spectra = GET_Z_PROFILE(CASI_fid, xloc, yloc)
  
  disp_get_location, display[1], xloc, yloc
  SPOT_spectra = GET_Z_PROFILE(SPOT_fid, xloc, yloc)
END