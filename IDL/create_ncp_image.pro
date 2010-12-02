FUNCTION CREATE_NCP_IMAGE_GUI
; Load the image
envi_select, title='Choose (spatial subset of) chi_square image', $
 fid=fid, dims=dims,pos=pos,/band_only,/mask,m_fid=m_fid

ENVI_FILE_QUERY, fid, nb=nb, ns = num_cols, nl=num_rows

; Load the mask image
if m_fid ne -1 then mask = envi_get_data(fid=m_fid,dims=dims1,pos=0) $
      else mask = bytarr(num_cols,num_rows)+1B



; Generate the NCP image
chi_sqr = envi_get_data(fid=fid,dims=dims,pos=pos)
ncp = 1.0 - chisqr_pdf(chi_sqr,nb-1)

; Apply the mask
idx = where(mask,count,complement=idxc,ncomplement=ncomplement)
if ncomplement gt 0 then ncp[idxc] = 0.0

tvscl, ncp

ENVI_ENTER_DATA, ncp
END











FUNCTION CREATE_NCP_IMAGE, fid, dims, pos, m_fid, m_pos
ENVI_FILE_QUERY, fid, nb=nb, ns = num_cols, nl=num_rows, fname=fname, data_type=data_type, xstart=xstart, $
    ystart=ystart, INTERLEAVE=interleave, dims=dims, pos=pos

; Load the mask image
if m_fid ne -1 then mask = envi_get_data(fid=m_fid,dims=dims1,pos=0) $
      else mask = bytarr(num_cols,num_rows)+1B

; Generate the NCP image
chi_sqr = envi_get_data(fid=fid,dims=dims,pos=pos)
ncp = 1.0 - chisqr_pdf(chi_sqr,nb-1)

; Apply the mask
idx = where(mask,count,complement=idxc,ncomplement=ncomplement)
if ncomplement gt 0 then ncp[idxc] = 0.0

output_file = fname + "_NCP.bsq"

; If the output is to file then open the file, write the binary data
  ; and close the file
  OpenW, unit, output_file, /GET_LUN
  WriteU, unit, ncp
  FREE_LUN, unit
  
  ; Get the map info of the file so that we can output it to the new file
  map_info = ENVI_GET_MAP_INFO(FID=fid)
    
  ; Then calculate the values needed to create the header file, and create it
  NSamples = dims[2] - dims[1] + 1
  NLines = dims[4] - dims[3] + 1
  NBands = N_ELEMENTS(pos)
  ENVI_SETUP_HEAD, FNAME=output_file, NS=num_cols, NL=num_rows, NB=1, $
    DATA_TYPE=4, offset=0, INTERLEAVE=0, $
    XSTART=xstart+dims[1], YSTART=ystart+dims[3], $
    DESCRIP="NCP Image Output", MAP_INFO=map_info, /OPEN, /WRITE, r_fid=r_fid
  
  return, r_fid
END



