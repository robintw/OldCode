FUNCTION REMOVE_SMALL_REGIONS, regions, min_size
  max_region_id = MAX(regions)
  
  print, MAX(regions)
  
  FOR id = 1, max_region_id DO BEGIN
    indices = WHERE(regions EQ id, count)
    ; If this is a region with only one pixel in it
    ; then remove it
    IF count LE min_size THEN regions[indices] = 0
  ENDFOR
  
  return, regions
END