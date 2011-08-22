PRO REGIONS_TO_ROIS, regions
  max_region_id = MAX(regions)
  
  dims = SIZE(regions, /dimensions)
  ns = dims[0]
  nl = dims[1]
  
  
  FOR id = 1, max_region_id DO BEGIN
    indices = WHERE(regions EQ id, count)
    
    IF count LT 1 THEN CONTINUE
    print, count
    name = "Region " + STRTRIM(STRING(id),2)
    roi_id = ENVI_CREATE_ROI(ns=ns, nl=nl, color=5, name=name)
    separated_indices = ARRAY_INDICES(regions, indices)
    separated_indices = transpose(separated_indices)
    ENVI_DEFINE_ROI, roi_id, /POINT, xpts=separated_indices[*, 0], ypts=separated_indices[*, 1]
  ENDFOR
END