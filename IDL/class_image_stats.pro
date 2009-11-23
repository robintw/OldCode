;+
; <PRO class_image_stats>
; The CLASS_IMAGE_STATS routine is an ENVI batch procedure. It is used for calculating
; statistics on an ENVI classification image from polygons defined by a shapefile.
;
; The program will take each polygon from the shapefile and convert it to a region of interest.
; The number of points for each class inside the ROI are counted. A text file is created which
; records the class value, number of points per class, the running total, class_percentages
; and the accumulated percent.
;
; @Param
;   shape_file {in} {required} {type=string} {default=none}
;     A scalar string specifying the full pathname of the SHAPEFILE file to read.
;
; @Param
;   image_file {in} {required} {type=string} {default=none}
;     A scalar string specifying the full pathname of the IMAGE FILE file to read.
;
; @Param
;   stats_file {in} {required} {type=string} {default=none}
;     A scalar string specifying the full pathname of the STATISTICS FILE file to write.
;
; @Author
;	Andy Pursch - ITT-VIS
;	apursch\@ittvis.com
;
; @Copyright
;   ITT Visual Information Solutions
;
; @Categories
;   classification, batch, statistics, roi
;
; @History
;   Original, June 2006
;
; @Requires
;   ENVI 4.2, IDL 6.2
;
; @Restrictions
;   Requires the shapefile to have an associated projection file (.prj)
;
; @Version
;   1.0
;-
PRO class_image_stats, shape_file, image_file
  ; This program needs to use undocumented procedures so source code
  ; can not be provided. Restore them from a save file instead.
  RESTORE,'F:\Dissertation\Programming\IDL-SVN\get_shape_proj.sav'

  ; Set some compiler options
  COMPILE_OPT idl2

  ; Return to the calling program if there is an error
  ON_ERROR, 2

  ; If any files are not specified on command line ask the user to select them
  IF N_ELEMENTS(image_file) EQ 0 THEN $
  		image_file = DIALOG_PICKFILE(TITLE='Select the input ENVI image file')
  IF N_ELEMENTS(shape_file) EQ 0 THEN $
  		shape_file = DIALOG_PICKFILE(title='Select the input shapefile')

  ; Open (and check) the image file.
  ENVI_OPEN_FILE, image_file, r_fid=fid, /no_interactive_query, /no_realize
  IF fid EQ -1 THEN BEGIN
  		void = DIALOG_MESSAGE('Unable to open image file', /error)
    	RETURN
  ENDIF

  ; Retrieve information about the file
  ENVI_FILE_QUERY, fid, file_type=file_type, nl=num_lines, $
  						ns=num_samps, num_classes=num_classes

  ; Open (and check) the shape file. Use the IDL shapefile object since ENVI's
  ; shapefile API is not exposed to the user.
  shape_file_obj = OBJ_NEW('idlffshape', shape_file)
  IF NOT OBJ_VALID(shape_file_obj) THEN BEGIN
    	void = DIALOG_MESSAGE('Unable to access shape file', /error)
    	RETURN
  ENDIF

;  ; If we found the shapefile now look for the .prj file. This contains the projection information
;  basename = strmid(shape_file, 0, strpos(shape_file,'.'))
;  prj_file_name = basename + '.prj'
;  status = FILE_TEST( prj_file_name, /read)
;  if status EQ 1 then proj = READ_SHAPEFILE_PROJ(prj_file_name) $
;  else begin
;  		msg = dialog_message('Can not locate associated projection file (.prj) for shapefile',/error)
;  		return
;  endelse

  ; Get the number of entities and the entity type.
  shape_file_obj->GetProperty, N_ENTITIES=num_ent, $
   		ENTITY_TYPE=ent_type

 ; Get the info for all attributes.
 shape_file_obj->GetProperty, ATTRIBUTE_INFO=attr_info

  ; Use the information returned in the projection structure.
  ;
  ; PROJ should be a structure if the projection is defined and known. In IDL this means
  ; it has a type value of 8. If PROJ is not 8 then I assume it is an "unsupported" projection.
;  IF SIZE(proj, /type) NE 8 THEN BEGIN
;   		;msg = DIALOG_MESSAGE('The map projection appears to be unsupported',/error)
;  		;RETURN
;  ENDIF ELSE BEGIN
;  		; looks like we have a supported projection so retrieve the info
;  		name = proj.name
;  		params = proj.params
;  		datum = proj.datum
;  		type = proj.type
;  ENDELSE
  
  ;name = "GCS_OSGB_1936"
  ;params = ["Airy_1830",6377563.396,299.3249646]
  
;  
;  print, "-----------"
;  print, name
;  print, params
;  print, datum
;  print, type
;  
  ;
;  ; Create the projection for the vector layers
;  iproj = ENVI_PROJ_CREATE(type=type, $
;   		name=name, datum=datum, params=params)

  ; Loop through the entities, converting each to an ROI, getting/computing the stats,
  ; writing out the stats.
  icnt=0

  ; Setup for the progress bar that will be displayed
  DEVICE,DECOMPOSED=1
  ENVI_BATCH_STATUS_WINDOW,/ON
  repStr='Processing Classification Image'
  ENVI_REPORT_INIT,repStr,BASE=repBase,TITLE='Calculating ROI Statistics'
  ENVI_REPORT_INC,repBase,num_ent
  WIDGET_CONTROL,/HOURGLASS

  ; Start looping over all individual cells from the shapefile
  FOR i=0, num_ent-1 DO BEGIN

		; Update the processing status bar
        ENVI_REPORT_STAT,repBase,i, num_ent-1

    	; First, get the vertices for the entity.
    	this_entity = shape_file_obj -> GetEntity(i) ; Could add the attributes keyword here.
    	these_vertices = *(this_entity.vertices)
    	attr = shape_file_obj->getAttributes( i)
    	
    	
    	IF attr.attribute_2 EQ -1 THEN CONTINUE ELSE $
      IF attr.attribute_2 EQ 1 THEN name = "bright" ELSE $
      IF attr.attribute_2 EQ 2 THEN name = "dark"

      print, "these_vertices"
      print, these_vertices

    	; Convert the vertices' map cooordinates into x/y file
    	; coordinates.  Note we're assuming they're all in the
    	; same projection here.  If they aren't, this is where
    	; you'd put in some conversion code with envi_convert_projection_coordinates.
;    	oproj = ENVI_GET_PROJECTION(FID = fid)
		;ENVI_CONVERT_PROJECTION_COORDINATES, $
 					 ;reform(these_vertices[0,*]), reform(these_vertices[1,*]), iproj, $
  					;oxmap, oymap, oproj
    	;ENVI_CONVERT_FILE_COORDINATES, fid, x_file_coords, $
                                        ;y_file_coords, oxmap, oymap
                                        
      x_file_coords = reform(these_vertices[0,*])
      y_file_coords = reform(these_vertices[1,*])


      ;x_file_coords = num_samps - x_file_coords
      y_file_coords = num_lines - y_file_coords
      
    	; Do shape file polygon entities get returned as a closed
    	; set of coordinates or if it's just implicit, let's check.  ROIs require
    	; polygons to be explicitly closed.
    	last_element_index = N_ELEMENTS(x_file_coords)-1
    	IF x_file_coords[0] NE x_file_coords[last_element_index] THEN BEGIN
      			x_file_coords = [x_file_coords, x_file_coords[0]]
      			y_file_coords = [y_file_coords, y_file_coords[0]]
    	ENDIF

      print, "HELLO, ROBIN HERE!"
      print, x_file_coords

    	; Since the overlay grid covers a larger geographic extent than the image some of the
    	; polygons will lie off the image. We can skip them.
    	IF TOTAL(x_file_coords LT 0) EQ 0 AND TOTAL(y_file_coords LT 0) EQ 0 AND $
    		TOTAL(x_file_coords GT NUM_SAMPS) EQ 0 AND TOTAL(y_file_coords GT NUM_LINES) EQ 0 THEN BEGIN
			; define counter for the the number of valid ROIs
    		icnt++
    		
    		; Now that we're in file coordinates, make an roi.
    		this_roi_id = ENVI_CREATE_ROI(nl=num_lines, ns=num_samps, name=name+string(i))
    		ENVI_DEFINE_ROI, this_roi_id, /polygon, xpts=x_file_coords, ypts=y_file_coords
    		;ENVI_DEFINE_ROI, this_roi_id, /polygon, xpts=these_vertices[0,*], ypts=these_vertices[1,*]
    		print, "-------"
    		print, x_file_coords

  		ENDIF
  ENDFOR

  ENVI_REPORT_INIT,BASE=repBase,/FINISH
  ;ENVI_BATCH_STATUS_WINDOW,/OFF
  DEVICE,DECOMPOSED=0

  OBJ_DESTROY, shape_file_obj
  PRINT,'Total number of Valid ROIs found: ', icnt
END
