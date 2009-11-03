; Call this function with the size of the image you want in x and y, and a base filename for the output
; images in base_filename
PRO CREATE_GLT_IMAGE, x, y, base_filename
  ; Create one row for the column indices image - make it go from 1 rather than 0 by adding 1 to it
  row = indgen(x) + 1
  ; Replicate this down the image
  col_indices_image = cmreplicate(row, y)
  
  ; Create one column for the row indices image - make it go from 1 rather than 0 by adding 1 to it
  column = indgen(y) + 1
  ; Replicate this across the image
  row_indices_image = transpose(cmreplicate(column, x))
  
  ENVI_WRITE_ENVI_FILE, col_indices_image, data_type=2, interleave=0, nb=1, nl=y, ns=x, offset=0, out_name=base_filename+"_ColIndices.bsq"
  ENVI_WRITE_ENVI_FILE, row_indices_image, data_type=2, interleave=0, nb=1, nl=y, ns=x, offset=0, out_name=base_filename+"_RowIndices.bsq"
END