@NewGetis

FUNCTION CREATE_GETIS_CLASS_IMAGE, fid, dims, pos, m_fid, m_pos, percentage, distance
  ; Run this first
  ; ENVI_SELECT, fid=fid,dims=dims,pos=pos, /mask, m_fid=m_fid, m_pos=m_pos
  
  ENVI_FILE_QUERY, fid, fname=fname
  
  getis_fid = NEWGETIS_NOGUI(fid, dims, pos, m_fid, m_pos, distance)
  
  color = 3
  
  print, pos
  
  top_roi_ids = lonarr(N_ELEMENTS(pos))
  bottom_roi_ids = lonarr(N_ELEMENTS(pos))
  
  FOR i=0, N_ELEMENTS(pos)-1 DO BEGIN
    print, "Doing band " + string(i)
    ; Get the high Getis values (bright uniform areas)    
    top_roi = ROI_PERCENTILE_THRESHOLD(percentage, "Band " + strcompress(string(pos[i])) + " top", color, fid=getis_fid, dims=dims, pos=pos[i], /ensure_above_zero)
    color = color + 1
    
    top_roi_ids[i] = top_roi
    
    ; Get the low Getis values (dark uniform areas)
    bottom_roi = ROI_PERCENTILE_THRESHOLD(percentage, "Band " + strcompress(string(pos[i])) + " bottom", color, fid=getis_fid, dims=dims, pos=pos[i], /bottom, /ensure_below_zero)
    color = color + 1
    
    bottom_roi_ids = bottom_roi
    
  ENDFOR
  
  
  fids = lonarr(2)
  
  ; Export the bottom ROIs
  ENVI_DOIT, 'ENVI_ROI_TO_IMAGE_DOIT', class_values=replicate(1, N_ELEMENTS(pos)), FID=getis_fid, ROI_IDS=bottom_roi_ids, out_name=fname+"BottomGetisClass.bsq", r_fid=r_fid
  fids[0] = r_fid
  
  ; Export the top ROIs
  ENVI_DOIT, 'ENVI_ROI_TO_IMAGE_DOIT', class_values=replicate(1, N_ELEMENTS(pos)), FID=getis_fid, ROI_IDS=top_roi_ids, out_name=fname+"TopGetisClass.bsq", r_fid=r_fid
  fids[1] = r_fid
  
  
  ; Do the layerstacking of the classification images
  
  poss = lonarr(N_ELEMENTS(fids))
  dims = lonarr(5, N_ELEMENTS(fids))
  
  FOR i=0, N_ELEMENTS(fids)-1 DO BEGIN
    ENVI_FILE_QUERY, fids[i], nb=nb, dims=r_dims
    poss[i] = nb - 1
    print, nb -1
    dims[*, i] = r_dims
  ENDFOR
  
  projection = ENVI_GET_PROJECTION(fid=fids[0], pixel_size=pixel_size)
  
  ENVI_DOIT, 'ENVI_LAYER_STACKING_DOIT', fid=fids, pos=poss, dims=dims, $
    out_dt=1, out_name=fname+"_GetisClassStacked.bsq", out_ps=pixel_size, $
    out_proj=projection, r_fid=r_fid

  return, r_fid
  
END