FUNCTION RUN_SEGMENTATION, fitness_image, data_image, seed_threshold, data_delta, min_fitness, min_region_size, do_morph_close
  ; Run region-growing algorithm
  regions = segment(fitness_image, data_image, seed_threshold, data_delta, min_fitness)
  print, "Finished segmentation"
  
  ;max_region_id = MAX(regions)
  
  ; Remove very small regions (either delete or merge?)
  regions = REMOVE_SMALL_REGIONS(regions, min_region_size-1)
  
  ; Run MORPH_CLOSE to remove small holes in regions and tidy them up a bit
  if do_morph_close EQ 1 THEN regions = MORPH_CLOSE(regions, intarr(3, 3)+1)
  
  
  ; Check other attributes and remove if necessary
  
  ; Vectorise
  
  regions = LABEL_REGION(regions, /ALL_NEIGHBORS)
  
  REGIONS_TO_ROIS, regions
  
  
  return, regions
END