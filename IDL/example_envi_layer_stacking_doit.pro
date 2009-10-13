pro example_envi_layer_stacking_doit
  ;
  ; Open the first input file. 
  ; We will also open the one band
  ; dem file to layer stack with 
  ; this file. 
  ;
  ;envi_open_file, ’bhtmref.img’, r_fid=t_fid
  ;if (t_fid eq -1) then begin
    ;envi_batch_exit
    ;return
  ;endif
  ;
  ; Open the second input file. 
  ;
  ;envi_open_file, ’bhdemsub.img’, r_fid=d_fid
  ;if (d_fid eq -1) then begin
    ;envi_batch_exit
    ;return
  ;endif
  ;
  ; Use all the bands from both files
  ; and all spatial pixels. First build the 
  ; array of FID, POS and DIMS for both 
  ; files.
  ; #
  

  ENVI_SELECT, fid=t_fid
  ENVI_SELECT, fid=d_fid
  
  envi_file_query, t_fid, $
    ns=t_ns, nl=t_nl, nb=t_nb
  envi_file_query, d_fid, $
    ns=d_ns, nl=d_nl, nb=d_nb
  ;
  nb = t_nb + d_nb
  fid = lonarr(nb)
  pos = lonarr(nb)
  dims = lonarr(5,nb)
  ;
  for i=0L,t_nb-1 do begin
    fid[i] = t_fid
    pos[i] = i
    dims[0,i] = [-1,0,t_ns-1,0,t_nl-1]
  endfor

  for i=t_nb,nb-1 do begin
    fid[i] = d_fid
    pos[i] = i-t_nb
    dims[0,i] = [-1,0,d_ns-1,0,d_nl-1]
  endfor
  ;
  ; Set the output projection and 
  ; pixel size from the TM file. Save
  ; the result to disk and use floating
  ; point output data.
  ; 
  
  help, fid
  help, pos
  help, dims
  
  
  print, dims
  ;
  ; Call the layer stacking routine. Do not
  ; set the exclusive keyword allow for an
  ; inclusive result. Use cubic convolution
  ; for the interpolation method.
  ;
  ;envi_doit, ’envi_layer_stacking_doit’, $
    ;fid=fid, pos=pos, dims=dims, $
    ;out_dt=out_dt, out_name=out_name, $
    ;interp=2, out_ps=out_ps, $
    ;out_proj=out_proj, r_fid=r_fid
  end