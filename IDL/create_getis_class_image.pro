@NewGetis

PRO CREATE_GETIS_CLASS_IMAGE, fid, dims, pos, m_fid, m_pos
  ; Run this first
  ; ENVI_SELECT, fid=fid,dims=dims,pos=pos, /mask, m_fid=m_fid, m_pos=m_pos
  
  ;ENVI_FILE_QUERY, fid, fname=fname
  
  percentage = 0.1
  
  getis_fid = NEWGETIS_NOGUI(fid, dims, pos, m_fid, m_pos)
  
  color = 3
  
  print, pos
  
  FOR i=0, N_ELEMENTS(pos)-1 DO BEGIN
    print, "Doing band " + string(i)
    ; Get the high Getis values (bright uniform areas)    
    top_roi = ROI_PERCENTILE_THRESHOLD(percentage, "Band " + strcompress(string(pos[i])) + " top", color, fid=getis_fid, dims=dims, pos=pos[i], /ensure_above_zero)
    color = color + 1
   
    ; Get the low Getis values (dark uniform areas)
    bottom_roi = ROI_PERCENTILE_THRESHOLD(percentage, "Band " + strcompress(string(pos[i])) + " bottom", color, fid=getis_fid, dims=dims, pos=pos[i], /bottom, /ensure_below_zero)
    color = color + 1
    
    ENVI_DOIT, 'ENVI_ROI_TO_IMAGE_DOIT', class_values=[1, 2], FID=getis_fid, ROI_IDS=[top_roi, bottom_roi], /IN_MEMORY
  ENDFOR
  
  
  ;CREATE_ROI_CLASS_IMAGE, 0.2, abund_fid, r_dims, pos
  
  ; Do something like the following
  ;roi_id = ROI_PERCENTILE_THRESHOLD(0.1, "test 0.1 bottom <0", 3, fid=fid, dims=dims, pos=pos, /bottom, /below_zero)
END