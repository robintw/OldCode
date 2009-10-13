PRO CREATE_GETIS_AND_SMACC_FILES, smacc_percentage, getis_percentage, getis_distance, out_file
  ; Open a dialog box to allow the input IKONOS and DEM file to be selected - along with a mask if needed
  ENVI_SELECT, fid=fid, dims=dims, pos=pos, /mask, m_fid=m_fid, m_pos=m_pos, title="Select the image, excluding the DEM band"
  
  ; Get the DEM band
  ENVI_SELECT, fid=dem_fid, dims=dem_dims, pos=dem_pos, title="Select the DEM band", /band_only
  
  ; If the dialog was cancelled then just exit
  IF fid EQ -1 THEN RETURN
  
  ; Create the SMACC classification image
  smacc_class_fid = CREATE_SMACC_CLASS_IMAGE(fid, dims, pos, m_fid, m_pos, smacc_percentage)
  
  getis_class_fid = CREATE_GETIS_CLASS_IMAGE(fid, dims, pos, m_fid, m_pos, getis_percentage, getis_distance)
  
  help, dem_dims
  print, dem_dims
  
  slope_aspect_fid = CREATE_SLOPE_ASPECT_IMAGES(dem_fid, dem_pos)
  
  LAYERSTACK_FILES, [fid, smacc_class_fid, getis_class_fid, slope_aspect_fid], out_file
END