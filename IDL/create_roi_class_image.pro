PRO CREATE_ROI_CLASS_IMAGE
  percentage = 0.2

  ; Use the ENVI dialog box to select a file
  ENVI_SELECT, fid=fid,dims=dims,pos=pos
  
  ; If the dialog box was cancelled then stop the procedure
  IF fid[0] EQ -1 THEN RETURN
  
  ; Create an array to hold the roi_ids returned by the Percentile Threshold function
  roi_ids=lonarr(n_elements(pos))
  
  ; For each band...
  FOR i=0, N_ELEMENTS(pos)-1 DO BEGIN
    ; Create a name for the ROI
    name = "Band " + STRCOMPRESS(STRING(i)) + " " + STRCOMPRESS(STRING(percentage, FORMAT="(f5.3)")) + "%"
  
    ;Create the ROI using a percentile threshold
    roi_id = ROI_PERCENTILE_THRESHOLD(percentage, name, 2+i, fid=fid, dims=dims, pos=pos[i])
    
    ; Put the ROI ID into the array
    roi_ids[i] = roi_id
  ENDFOR
  
  ; Convert the ROIs to a classification image where every pixel that is in any of the ROIs gets given a value of 1
  ENVI_DOIT, 'ENVI_ROI_TO_IMAGE_DOIT', class_values=replicate(long(1), N_ELEMENTS(pos)), FID=fid, ROI_IDS=roi_ids, /IN_MEMORY
END