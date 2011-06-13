PRO CREATE_NCP_IMAGE_GUI
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

ENVI_ENTER_DATA, ncp
END