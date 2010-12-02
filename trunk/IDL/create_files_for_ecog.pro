; Run with CREATE_GETIS_AND_SMACC_FILES, 0.3, 0.3, 0.3, 1, "F:\Dissertation\Data\Final Processing Outputs\Ki-Default.bsq"
PRO CREATE_FILES_FOR_ECOG, smacc_percentage, getis_top_percentage, getis_bottom_percentage, getis_distance, out_file
  ; Open a dialog box to allow the input IKONOS file to be selected, along with a mask if needed
  ENVI_SELECT, fid=fid, dims=dims, pos=pos, /mask, m_fid=m_fid, m_pos=m_pos, title="Select the image, excluding the DEM band"
  
  ; If the dialog was cancelled then just exit
  IF fid EQ -1 THEN RETURN
  
  ; Open a dialog box to allow the input DEM band to be selected
  ENVI_SELECT, fid=dem_fid, dims=dem_dims, pos=dem_pos, title="Select the DEM band", /band_only
  
  ; If the dialog was cancelled then just exit
  IF dem_fid EQ -1 THEN RETURN
  
  ; Open a dialog box to allow the input ChiSq band to be selected
  ENVI_SELECT, fid=chi_sq_fid, dims=chi_sq_dims, pos=chi_sq_pos, title="Select the Chi-Sq band", /band_only
  
  ; If the dialog was cancelled then just exit
  IF chi_sq_fid EQ -1 THEN RETURN
  
  ; Create the SMACC classification image
  smacc_class_fid = CREATE_SMACC_CLASS_IMAGE(fid, dims, pos, m_fid, m_pos, smacc_percentage)
 
  ; Create the Getis classification image
  getis_class_fid = CREATE_GETIS_CLASS_IMAGE(fid, dims, pos, m_fid, m_pos, getis_top_percentage, getis_bottom_percentage, getis_distance)
  
  ; Create the slope and aspect images
  slope_aspect_fid = CREATE_SLOPE_ASPECT_IMAGES(dem_fid, dem_pos)
  
  ncp_fid = CREATE_NCP_IMAGE(chi_sq_fid, chi_sq_dims, chi_sq_pos, m_fid, m_pos)
  
  ; Layerstack the above three outputs together, appended on to the end of the original IKONOS+DEM file
  LAYERSTACK_FILES, [fid, smacc_class_fid, getis_class_fid, slope_aspect_fid, ncp_fid], out_file
END