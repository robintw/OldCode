; Takes the fid, pos and dims of a DEM band and creates slope and aspect images from it
FUNCTION CREATE_SLOPE_ASPECT_IMAGES, dem_fid, dem_pos
  ENVI_FILE_QUERY, dem_fid, fname=fname, dims=dem_dims
  
  print, dem_fid
  print, dem_pos
  print, dem_dims
  
  ; Get the pixel size of the DEM
  projection = ENVI_GET_PROJECTION(fid=dem_fid, pixel_size=pixel_size)
  
  ENVI_DOIT, 'TOPO_DOIT', BPTR=[0,1], fid=dem_fid, pos=dem_pos, $
    out_name=fname+"_slope_aspect.bsq", dims=dem_dims, r_fid=r_fid, $
    kernel_size=3, pixel_size=pixel_size
  
  return, r_fid 
END