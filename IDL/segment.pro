PRO REGION_GROW, seed_x, seed_y, image, region_id, regions, region_seed_values, data_delta, fitness_image, fitness_threshold
  ; Mark this pixel as in the region
  regions[seed_x, seed_y] = region_id
  
  ; Get surrounding pixels (cardinal directions only)
  xs = [seed_x + 1, seed_x - 1,   seed_x,     seed_x]
  ys = [seed_y,   seed_y,     seed_y + 1,   seed_y - 1]
  
  ; Get image size
  dim = SIZE(image, /DIMENSIONS)
  
  ; Sort out xs and ys for boundary problems
  FOR i = 0, N_ELEMENTS(xs) - 1 DO BEGIN
    xs[i] = 0 > xs[i] < (dim[0]-1)
    ys[i] = 0 > ys[i] < (dim[1]-1)
  ENDFOR
  
  ; Check each pixel to see if it fulfills H(x)
  FOR i = 0, 3 DO BEGIN
    ; If value is already in a region then skip to the next one
    IF regions[xs[i], ys[i]] NE 0 THEN CONTINUE
    
    ; Get the values (for all bands) of the image at this location
    values = image[xs[i], ys[i], *]
   
    region_indices = WHERE(regions EQ region_id, count)
    ;print, count
    
    IF count LT 2 THEN BEGIN
      ; This is the first pixel in this region, so just add it
      regions[xs[i], ys[i]] = region_id
      
      ; And set the key value for this region to the value of this pixel
      region_seed_values[region_id, *] = values
      
      ;print, "First pixel in region: Region ID = ", region_id, "values = ", values
      
      image[xs[i], ys[i]] = 0
    ENDIF ELSE BEGIN
      ; There are already pixels in the region, so check this isn't significantly different
      ; and if it isn't then add it to the region
      diff = region_seed_values[region_id, *] - values
      result = MAX(abs(diff))
      
      ;print, diff, "Result = ", result, "Data D = ", data_delta
      
      ; If the difference is less than the difference allowed in the image data
      IF result LE data_delta THEN BEGIN
        ; Also check the fitness image - we don't want to add any pixels with a very low fitness
        fitness = fitness_image[xs[i], ys[i]]
        IF fitness GT fitness_threshold THEN BEGIN
          ; Mark it as belonging to this region in the regions image
          regions[xs[i], ys[i]] = region_id
          ;image[xs[i], ys[i]] = 0
          ; Recurse to check all of the pixels around this one
          REGION_GROW, xs[i], ys[i], image, region_id, regions, region_seed_values, data_delta, fitness_image, fitness_threshold
        ENDIF
      ENDIF
    ENDELSE
  ENDFOR
  
END

FUNCTION segment, fitness_image, data_image, fitness_thresh, data_delta, fitness_adding_threshold
  max_num = 500

  ; Get the size of the image
  image_size = SIZE(data_image,/dimensions)
  
  ; Create an output image for the labelled regions map
  ; the same size as input image
  regions = intarr(image_size[0], image_size[1])
  
  ; Create a large array to store all of the seed values in
  region_seed_values = intarr(10000, image_size[2])
  
  ; Initialise region_id
  region_id = 0
  
  WHILE 1 DO BEGIN  
    ; Get the maximum value in the fitness image
    image_max = MAX(fitness_image, subscript)
  
    ; If it's less than the threshold then there aren't any useful
    ; seeds left in the image, so exit
    IF image_max LT fitness_thresh OR region_id GT max_num THEN BREAK
  
    ; Increment region ID
    region_id += 1
    ;print, region_id
    
    ; Otherwise get the X and Y indices for the max value
    subs_xy = ARRAY_INDICES(fitness_image, subscript)
    seed_x = subs_xy[0]
    seed_y = subs_xy[1]
  
    ; Grow a region from that seed
    REGION_GROW, seed_x, seed_y, data_image, region_id, regions, region_seed_values, data_delta, fitness_image, fitness_adding_threshold
  
    n_pixels = WHERE(regions EQ region_id, count)
    ;print, "N pixels in new region: ", count
    ; Remove the maximum we just found so that next time the MAX function gives
    ; us the next value down
    fitness_image[subs_xy[0], subs_xy[1]] = 0
  ENDWHILE
  
  help, regions
  return, regions
END