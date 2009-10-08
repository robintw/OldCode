PRO CREATE_GETIS_AND_SMACC_FILES, smacc_percentage, getis_percentage, getis_distance
  ; Open a dialog box to allow the input IKONOS and DEM file to be selected - along with a mask if needed
  ENVI_SELECT, fid=fid, dims=dims, pos=pos, /mask, m_fid=m_fid, m_pos=m_pos
  
  ; Create the SMACC classification image
  CREATE_SMACC_CLASS_IMAGE, fid, dims, pos, m_fid, m_pos, smacc_percentage
  
  CREATE_GETIS_CLASS_IMAGE, fid, dims, pos, m_fid, m_pos, getis_percentage, getis_distance
END