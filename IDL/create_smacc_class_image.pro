FUNCTION CREATE_SMACC_CLASS_IMAGE, fid, dims, pos, m_fid, m_pos, percentage
  ; Run this first
  ; ENVI_SELECT, fid=fid,dims=dims,pos=pos, /mask, m_fid=m_fid, m_pos=m_pos
  
  ENVI_FILE_QUERY, fid, fname=fname
  
  ENVI_DOIT, "ENVI_SMACC_DOIT", m_fid=m_fid, m_pos=m_pos, fid=fid, dims=dims, pos=pos, n_endmembers=5, abund_name=fname+"_abund.bsq", abund_r_fid=abund_fid, method=2, out_name=fname+"_speclib.sli", r_fid=r_fid
  
  ENVI_FILE_QUERY, abund_fid, dims=r_dims, nb=nb
  
  ; Do all the bands of the image
  smacc_pos = lindgen(nb)
  
  print, abund_fid
  
  r_fid = CREATE_SMACC_ROI_CLASS_IMAGE(percentage, abund_fid, r_dims, smacc_pos)
  
  return, r_fid
END