FUNCTION SCALE_ARRAY, array, from, to
  amax = max(array)
  amin = min(array)
  
  print, "Scaling: Min = ", amin, "Max = ", amax
  
  return, (array - amin) * (to - from) / FLOAT(amax - amin)
END

PRO SCALE_GETIS
  ENVI_SELECT, fid=fid, dims=dims, pos=pos
  
  getis = ENVI_GET_DATA(fid=fid, dims=dims, pos=pos)
  

  ;indices = WHERE(getis GT 0)
  
  ;getis[indices] = SCALE_ARRAY(getis[indices], 0, 1)
  
  ;print, max(getis[indices]), min(getis[indices])
  
  ;new_indices = WHERE(getis LT 0)
  
  ;getis[new_indices] = SCALE_ARRAY(getis[new_indices], -1, 0)
  
  ;print, max(getis[new_indices]), min(getis[new_indices])
  
  indices = WHERE(getis LT 0)
  
  getis[indices] = abs(getis[indices] * (max(getis) / min(getis)))
  
  ENVI_ENTER_DATA, SCALE_ARRAY(getis, 0, 1)
END