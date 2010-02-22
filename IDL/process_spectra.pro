PRO PROCESS_SPECTRA, fid=fid, pos=pos, dims=dims, $ 
   spec=spec, snames=snames, scolors=scolors, _extra=extra
   help, spec 
   print, spec
   print, snames
   print, scolors
   
   ; Create dialog box window
  TLB = WIDGET_AUTO_BASE(title="Select output Spectral Library")
  
  ; Create the widget to let the user select file or memory output
  W_FileOrMem = WIDGET_OUTFM(TLB, /AUTO_MANAGE, uvalue='fm')
  
  ; Start the automatic management of the window
  result = AUTO_WID_MNG(TLB) 
  
  ; If the OK button was pressed
  IF result.accept EQ 0 THEN RETURN
   
   ; Take the collected spectra and export to a spectral library
   OpenW, unit, result.fm.name, /GET_LUN
   WriteU, unit, CVImage
   FREE_LUN, unit
END 
 
PRO COLLECT_SPECTRA, event
   ENVI_SELECT, fid=fid, pos=pos, dims=dims
   ENVI_COLLECT_SPECTRA, dims=dims, fid=fid, pos=pos,$ 
     title=title, procedure='PROCESS_SPECTRA',$ 
     h_info=info 
END 
