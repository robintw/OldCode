; Shrink an individual ROI from the file specified with fid
PRO SHRINK_ROI, fid, roi_id
  ; Get the number of samples
  ENVI_FILE_QUERY, fid, ns=ns, nl=nl
  
  ; Get the array of 1D points then convert them to actual x and y co-ords
  points = ENVI_GET_ROI(roi_id)
  
  ; If there are no points in the ROI then exit
  if points[0] EQ -1 THEN RETURN
  
  ; Create the image array
  image_array = intarr(ns, nl)
 
  ; Extract the point indices to X and Y co-ordinates
  point_indices = ARRAY_INDICES(image_array, points)

  
  ; Set the area covered by the ROI to 1 in the image_array
  image_array[point_indices[0, *], point_indices[1, *]] = 1
  
  ; Create the kernel for the summing CONVOL operation - no diagonals
  Kernel = FLTARR(3, 3)
  Kernel[0, *] = [0, 1, 0]
  Kernel[1, *] = [1, 1, 1]
  Kernel[2, *] = [0, 1, 0]
  
  ; Create an image where each element is the sum of the elements within
  ; d around it
  summed_image = CONVOL(image_array, Kernel, /CENTER, /EDGE_TRUNCATE)
 
  ; Select the indices where the pixels are entirely surrounded by other pixels
  ; That is, all the pixels we want to keep in the shrunk ROI
  where_answer = WHERE(summed_image EQ 5, count)
  
  IF count EQ 0 THEN RETURN
  
  new_indices = ARRAY_INDICES(summed_image, where_answer)

  ; Extract the X and Y indices from the array
  new_x_indices = reform(new_indices[0, *])
  new_y_indices = reform(new_indices[1, *])
  
  ; Create the new ROI and associated the points with it
  new_roi_id = ENVI_CREATE_ROI(nl=nl, ns=ns, name="Shrunk ROI")
  ENVI_DEFINE_ROI, new_roi_id, /point, xpts=new_x_indices, ypts=new_y_indices
END

; Shrink all the ROIs associated with an image
PRO SHRINK_ALL_ROIS
  ; Select the file the ROIs are associated with
  ENVI_SELECT, fid=fid
  
  ; Get the list of ROIs and run through it shrinking all of them
  roi_ids = ENVI_GET_ROI_IDS(fid=fid)
  FOR i = 0, N_ELEMENTS(roi_ids) - 1 DO BEGIN
    print, "DOING ROI ID ", roi_ids[i]
    SHRINK_ROI, fid, roi_ids[i]
  ENDFOR
END