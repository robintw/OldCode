PRO CREATE_GETIS_CLASS_IMAGE, fid, dims, pos, m_fid, m_pos
  ; Run this first
  ; ENVI_SELECT, fid=fid,dims=dims,pos=pos, /mask, m_fid=m_fid, m_pos=m_pos
  
  ENVI_FILE_QUERY, fid, fname=fname
  
  ENVI_DOIT, "ENVI_SMACC_DOIT", m_fid=m_fid, m_pos=m_pos, fid=fid, dims=dims, pos=pos, n_endmembers=5, abund_name=fname+"_abund.bsq", abund_r_fid=abund_fid, method=2, out_name=fname+"_speclib.sli", r_fid=r_fid
  
  ENVI_FILE_QUERY, abund_fid, dims=r_dims
  
  print, abund_fid
  
  CREATE_ROI_CLASS_IMAGE, 0.2, abund_fid, r_dims, pos
  
  ; Do something like the following
  roi_id = ROI_PERCENTILE_THRESHOLD(0.1, "test 0.1 bottom <0", 3, fid=fid, dims=dims, pos=pos, /bottom, /below_zero)
END