FUNCTION CREATE_SMACC_ROI_CLASS_IMAGE, percentage, fid, dims, pos  
  ; If the dialog box was cancelled then stop the procedure
  IF fid[0] EQ -1 THEN RETURN, -1
  
  ENVI_FILE_QUERY, fid, fname=fname
  
  ; Create an array to hold the roi_ids returned by the Percentile Threshold function
  roi_ids=lonarr(n_elements(pos))
  
  help, pos
  
  ; For each band...
  FOR i=0, N_ELEMENTS(pos)-1 DO BEGIN
    ; Create a name for the ROI
    name = "Band " + STRCOMPRESS(STRING(i)) + " " + STRCOMPRESS(STRING(percentage, FORMAT="(f5.3)")) + "%"
    
    print, name
    
    ;Create the ROI using a percentile threshold
    roi_id = ROI_PERCENTILE_THRESHOLD(percentage, name, 2+i, fid=fid, dims=dims, pos=pos[i], /ensure_above_zero)
    
    ; Put the ROI ID into the array
    roi_ids[i] = roi_id
  ENDFOR
  
  ; Convert the ROIs to a classification image where every pixel that is in any of the ROIs gets given a value of 1
  ENVI_DOIT, 'ENVI_ROI_TO_IMAGE_DOIT', class_values=replicate(long(1), N_ELEMENTS(pos)), FID=fid, ROI_IDS=roi_ids, out_name=fname+"_SMACC_ClassImage.bsq", r_fid=r_fid
  
  return, r_fid
END