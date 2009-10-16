; Layerstack every band of the array of fids (input_fids) together into a specified output file.
PRO LAYERSTACK_FILES, input_fids, out_name
  total_nb = 0
  
  FOR i=0, N_ELEMENTS(input_fids) - 1 DO BEGIN
    ENVI_FILE_QUERY, input_fids[i], nb=nb, ns=ns, nl=nl, dims=r_dims
    pos_to_concat = lindgen(nb)
    fids_to_concat = replicate(input_fids[i], nb)
    
    total_nb = total_nb + nb
    ; If it's the first time through then init the arrays
    IF i EQ 0 THEN BEGIN
      output_pos = pos_to_concat
      output_fids = fids_to_concat
    ENDIF ELSE BEGIN
      output_pos = [output_pos, pos_to_concat]
      output_fids = [output_fids, fids_to_concat]
    ENDELSE
  ENDFOR
  
  ; We're assuming that we want the output file to be the same dims as the last input file - as all the input
  ; files will be the same size.
  output_dims = cmreplicate(r_dims, total_nb)
  
  ; Take the projection from the first file
  projection = ENVI_GET_PROJECTION(fid=input_fids[0], pixel_size=pixel_size)
  
  ENVI_DOIT, 'ENVI_LAYER_STACKING_DOIT', fid=output_fids, pos=output_pos, dims=output_dims, $
    out_dt=2, out_name=out_name, out_ps=pixel_size, $
    out_proj=projection, r_fid=r_fid
END