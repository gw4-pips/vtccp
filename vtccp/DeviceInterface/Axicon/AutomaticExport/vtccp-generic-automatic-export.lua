{%function acv_template_info()
return { name = "VTCCP Generic CSV Automatic Export", appendable = false }
end

function acv_report_headers()
local g = general
local p = strings.param_desc
cols = { g.msg.desc, g.symbology.desc, g.subsymb.desc, g.decode_time.desc,
         "Platform Host", "Software Product Code", "Software Version", "Calibration Date",
         "Reader Serial", "Matrix Size", "Data Structure", "GS1 Company Prefix",
         "Human Readable", "Print Growth", g.isograde.desc, p.xdim, p.gain, p.xgain, p.ydim, p.ygain,
         p.height, p.qr_version, p.qr_ecc_level, p.overall, p.rmin, p.rmax,
         p.symbol_contrast, p.edge_contrast, p.modulation, p.defects,
         p.decodability, p.decode, p.axial_nonuniformity, p.grid_nonuniformity,
         p.reflectance_margin, p.unused_error_correction, p.contrast_uniformity,
         p.fixed_pattern_damage, p.dmx_fpd_L_left, p.dmx_fpd_L_bottom,
         p.dmx_fpd_quiet_left, p.dmx_fpd_quiet_bottom, p.dmx_fpd_clocks,
         p.dmx_fpd_average, p.format_info, p.version_info }
return strings.csv_header_from_table(cols)
end%}
"{*strings.csv_escape(general.msg)*}","{*general.symbology*}","{*general.subsymb*}","{*general.decode_time*}","{*general.platform.hostname*}","{*general.version.code*}","{*general.version.current*}","{*reader.caltime_user*}","{*reader.serial*}","{*general.matrix_size*}","{*plugins.data_structure.status*}","{*plugins.gs1_company_prefix.status*}","{*plugins.human_readable.status*}","{*iso.dmx_fpd_average*}","{*general.isograde*}","{*general.xdim*}","{*general.gain*}","{*general.xgain*}","{*general.ydim*}","{*general.ygain*}","{*strings.csv_escape(general.height)*}","{*general.qr_version*}","{*general.qr_ecc_level*}","{*iso.overall*}","{*iso.rmin*}","{*iso.rmax*}","{*iso.symbol_contrast*}","{*iso.edge_contrast*}","{*iso.modulation*}","{*iso.defects*}","{*iso.decodability*}","{*iso.decode*}","{*iso.axial_nonuniformity*}","{*iso.grid_nonuniformity*}","{*iso.reflectance_margin*}","{*iso.unused_error_correction*}","{*iso.contrast_uniformity*}","{*iso.fixed_pattern_damage*}","{*iso.dmx_fpd_L_left*}","{*iso.dmx_fpd_L_bottom*}","{*iso.dmx_fpd_quiet_left*}","{*iso.dmx_fpd_quiet_bottom*}","{*iso.dmx_fpd_clocks*}","{*iso.dmx_fpd_average*}","{*iso.format_info*}","{*iso.version_info*}"