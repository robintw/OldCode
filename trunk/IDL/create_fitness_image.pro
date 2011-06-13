FUNCTION GET_PERCENTILE_THRESHOLD, data, percentage
  sorted_indices = SORT(data)

  len = N_ELEMENTS(data)
  
  percentage = 1 - (float(percentage)/100)
  
  threshold =  data[sorted_indices[ceil(percentage * len)]]
  
  print, threshold
  
  return, threshold
END

FUNCTION PROCESS_SMACC, fid, dims, pos
  ; Get all of the bands of the SMACC image into one array
  FOR i = 0, N_ELEMENTS(pos) - 1 DO BEGIN
    band = ENVI_GET_DATA(fid=fid, dims=dims, pos=i)
    IF N_ELEMENTS(smacc) EQ 0 THEN smacc = band ELSE smacc = [ [[smacc]], [[band]] ]
  ENDFOR
  
  ; Remove the first band which is the Shadow endmember
  smacc = smacc[*, *, 1:pos[N_ELEMENTS(pos)-1]]
  
  ; Get the pixel-wise maximum (ie. for each pixel the maximum across all bands)
  smacc = MAX(smacc, dimension=3)
  
  return, smacc
END

FUNCTION PROCESS_GETIS_BAND, getis, getis_percentage_thresh
  ; We split it into two bits: where Getis >= 0 (bright uniform) and where Getis < 0
  ; (dark uniform)
  
  bright_getis = getis[WHERE(getis GE 0)]
  dark_getis = getis[WHERE(getis LT 0)]

  ; Calculate the threshold below which we should ignore everthing
  bright_thresh = GET_PERCENTILE_THRESHOLD(bright_getis, getis_percentage_thresh)
  
  ; Set everything below this threshold to zero
  getis[WHERE(getis GT 0 AND getis LT bright_thresh)] = 0
  
  ; For everything above this threshold we want to scale the data
  
  ; So get the data out
  to_scale_indices = WHERE(getis GE bright_thresh)
  to_scale = getis[to_scale_indices]
   
  ; Scale this data
  scaled = (to_scale - MIN(to_scale)) * (1 / (MAX(to_scale) - MIN(to_scale)))
    
  getis[to_scale_indices] = scaled
  
  ; Calculate the threshold ABOVE which we should ignore everything (as this will be a negative number)
  dark_thresh = GET_PERCENTILE_THRESHOLD(dark_getis, 100-getis_percentage_thresh)
  
  ; Set everything above this threshold to zero
  getis[WHERE(getis LT 0 AND getis GT dark_thresh)] = 0
  
  ; For everything below the threshold we want to scale the data
  to_scale_indices = WHERE(getis LE dark_thresh)
  to_scale = getis[to_scale_indices]
  
  ; Scale this data
  scaled = (to_scale - MIN(to_scale)) * (1 / (MAX(to_scale) - MIN(to_scale)))
  
  ; Flip the data so that the good stuff (very negative values) is at the top
  scaled = 1 - scaled
  getis[to_scale_indices] = scaled
  
  return, getis
END

PRO CREATE_FITNESS_IMAGE
  getis_coef = 0.3
  smacc_coef = 0.5
  ncp_coef = 0.2
  getis_percentage_thresh = 5
  
  ENVI_SELECT, fid=getis_fid, dims=getis_dims, pos=getis_pos, title="Select Getis image"
    
  ENVI_SELECT, fid=smacc_fid, dims=smacc_dims, pos=smacc_pos, title="Select SMACC image"  
    
  ENVI_SELECT, fid=ncp_fid, dims=ncp_dims, pos=ncp_pos, title="Select MAD NCP image", /BAND_ONLY
  
  ; Get the NCP band - it's just one band so that's nice and simple
  ncp = ENVI_GET_DATA(fid=ncp_fid, dims=ncp_dims, pos=ncp_pos)
  
  ; Get the SMACC image and process it down to one band
  smacc = PROCESS_SMACC(smacc_fid, smacc_dims, smacc_pos)
  
  ; It's easy to scale SMACC and the NCP as they are both given as fractions anyway
  new_smacc = smacc_coef * smacc
  new_ncp = ncp_coef * ncp
  
  getis_band_scale = 1 / FLOAT(N_ELEMENTS(getis_pos))
  
  ; Getis is a bit harder!
  ; Loop through all of the bands, process them to convert to sensible numbers, scale each band and add
  FOR i = 0, N_ELEMENTS(getis_pos) - 1 DO BEGIN
    getis = ENVI_GET_DATA(fid=getis_fid, dims=getis_dims, pos=i)
    processed = PROCESS_GETIS_BAND(getis, getis_percentage_thresh)
    IF N_ELEMENTS(new_getis) EQ 0 THEN new_getis = getis_band_scale * processed ELSE new_getis = new_getis + (getis_band_scale * processed)
  ENDFOR
  
  ; Calculate overall fitness image
  fitness = new_getis + new_smacc + new_ncp
  
  ENVI_ENTER_DATA, fitness
END