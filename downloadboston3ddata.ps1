#mass download "http://www.bostonplans.org/3d-data-maps/3d-smart-model/3d-data-download"    3d model data  (sketchup version)

#the following are available:
# filetype: url example
# SKP: http://maps.bostonredevelopmentauthority.org/3d/bpda_3d_downloads/skp/BOS3D_A13.skp.zip
# DXF: http://maps.bostonredevelopmentauthority.org/3d/bpda_3d_downloads/dxf/BOS3D_A13_CAD.DXF.zip
# DAE(building): http://maps.bostonredevelopmentauthority.org/3d/bpda_3d_downloads/dae_buildings/BOS3D_A13_Buildings.zip
# DAE(terrain): http://maps.bostonredevelopmentauthority.org/3d/bpda_3d_downloads/dae_terrain/BOS3D_A13_Terrain.zip
# PNG(basemap): http://maps.bostonredevelopmentauthority.org/3d/bpda_3d_downloads/basemap_png/BOS3D_A13_GroundPlan.png
# PDF(basemap): http://maps.bostonredevelopmentauthority.org/3d/bpda_3d_downloads/pdf_groundplan/BOS3D_A13_PDF_Basemap.pdf
# JPG (aerial photo): http://maps.bostonredevelopmentauthority.org/3d/bpda_3d_downloads/aerial_2014_jpg/BOS3D_A13_2014_Aerial.jpg

# Lines 29 and 30 need to be edited to download from the correct url, and save to the correct filename  (for example the skp need to be changed to dxf, just look carefully at the URLs)
# I'll leave that as an exercise for the reader. 


$PATHTOSAVETO = "U:\mapsketchup\" #don't forget the \ at the end :^)
if (-not(test-path $PATHTOSAVETO)){
    New-Item -ItemType Directory -Path $PATHTOSAVETO | Out-Null
}


$x =  Invoke-RestMethod "http://maps.bostonredevelopmentauthority.org/3d/Bos3dTilesGeoJson.js"
$json = $x.substring($x.IndexOf('=')+1)|convertfrom-json

$json.features|foreach{
    $tileid = $_.properties.tile_id

    $url = "http://maps.bostonredevelopmentauthority.org/3d/bpda_3d_downloads/skp/BOS3D_$($tileid).skp.zip"
    Invoke-WebRequest -uri $url -OutFile "$($PATHTOSAVETO)\BOS3D_$($tileid).skp.zip"
}
