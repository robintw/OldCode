PRO RUN_GUI, event

; Parameters
; ----------
getis_coef = 0.3
smacc_coef = 0.3
ncp_coef = 0.4
getis_percentage_thresh = 5

seed_threshold = 0.75
data_delta = 5
min_fitness = 0.5

min_region_size = 3

do_morph_close = 0

ENVI_DELETE_ROIS, /ALL

ENVI_SELECT, fid=data_fid, dims=dims, pos=data_pos, title="Select data image"

data_image = GET_DATA_IMAGE(data_fid, dims, data_pos)

; Generate Fitness Image
fitness_image = CREATE_FITNESS_IMAGE(getis_coef, smacc_coef, ncp_coef, getis_percentage_thresh)

; Run Segmentation
result = RUN_SEGMENTATION(fitness_image, data_image, seed_threshold, data_delta, min_fitness, min_region_size, do_morph_close)

; Display image and ROIs
ENVI_DISPLAY_BANDS, [data_fid, data_fid, data_fid], data_pos, /new

END