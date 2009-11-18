@CREATE_SMACC_ROI_CLASS_IMAGE
FUNCTION CREATE_SMACC_CLASS_IMAGE, fid, dims, pos, m_fid, m_pos, percentage
  ; If this file is being run manually then the following must be run first to set up the variables correctly
  ; ENVI_SELECT, fid=fid,dims=dims,pos=pos, /mask, m_fid=m_fid, m_pos=m_pos
  
  ; Get the filename of the given file
  ENVI_FILE_QUERY, fid, fname=fname
  
  ; Perform the SMACC endmember extraction asking for 4 endmembers, with the constraint of
  ; summing to unity
  ENVI_DOIT, "ENVI_SMACC_DOIT", m_fid=m_fid, m_pos=m_pos, fid=fid, dims=dims, pos=pos, $
    n_endmembers=4, abund_name=fname+"_abund.bsq", abund_r_fid=abund_fid, method=2, $
    out_name=fname+"_speclib.sli", r_fid=r_fid
  
  ; Find out how many bands the image has (in case the number of endmembers above has been changed)
  ENVI_FILE_QUERY, abund_fid, dims=r_dims, nb=nb
  
  ; Create a list of all the bands in the image so that all of them can be processed
  smacc_pos = lindgen(nb)
  
  ; Create the classification image from the SMACC image
  r_fid = CREATE_SMACC_ROI_CLASS_IMAGE(percentage, abund_fid, r_dims, smacc_pos)
  
  ; Return the FID of the classification image
  return, r_fid
END