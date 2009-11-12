; Call this function with the size of the image you want in x and y, and a base filename for the output
; images in base_filename
PRO CREATE_GLT_IMAGE, x, y, base_filename, reverse_cols=reverse_cols, reverse_rows=reverse_rows
  ; Create one row for the column indices image - make it go from 1 rather than 0 by adding 1 to it
  row = indgen(x) + 1
  
  if keyword_set(reverse_cols) then row = reverse(row)
  
  ; Replicate this down the image
  col_indices_image = cmreplicate(row, y)
  
  ; Create one column for the row indices image - make it go from 1 rather than 0 by adding 1 to it
  column = indgen(y) + 1
  
  if keyword_set(reverse_rows) then column = reverse(column)
  ; Replicate this across the image
  row_indices_image = transpose(cmreplicate(column, x))
  
  ;print, col_indices_image
  
  ;print, "---------"
  
  ;print, row_indices_image
  
  ENVI_WRITE_ENVI_FILE, col_indices_image, data_type=2, interleave=0, nb=1, nl=y, ns=x, offset=0, out_name=base_filename+"_ColIndices.bsq"
  ENVI_WRITE_ENVI_FILE, row_indices_image, data_type=2, interleave=0, nb=1, nl=y, ns=x, offset=0, out_name=base_filename+"_RowIndices.bsq"
END