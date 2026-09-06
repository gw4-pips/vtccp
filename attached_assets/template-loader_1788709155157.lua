local u = require "utils"
require "gettext-utils"
--local dbg = require "dbg"

local strings = {
	verification_report = _("Verification Report"),
	indiv_linear_report = _("Verification Report (Linear Details)"),
	iso_parameter = _("ISO Parameter"),
	grade = _("Grade"),
	percentage = _("Percentage"),
	yes = _(u.bool_str.yes_no[true]),
	no = _(u.bool_str.yes_no[false]),
	tpl_name = {
		dmx_csv = _("Data Matrix CSV"),
		linear_csv = _("Linear CSV"),
		qr_csv = _("QR Code CSV"),
		std_csv = _("Standard CSV"),
		std_html = _("Standard HTML"),
		std_html_multi = _("Standard HTML (multiple codes)"),
		std_txt = _("Standard plain text"),
		gs1_html = _("GS1 format HTML"),
		plugins_html = _("Detailed plugins HTML"),
		indiv_html = _("Linear details HTML"),
	},
	av = setmetatable({}, {
		__index = function(t, k)
			return _(u.application_strs[k])
		end
	}),
	param_desc = setmetatable({}, {
		__index = function(t, k)
			return _(u.param_desc[k])
		end
	}),
	status = setmetatable({}, {
		__index = function(t, k)
			return _(u.plugin_status_str[k])
		end
	}),
}

function strings.csv_escape(s)
	-- tostring(nil) returns a literal string "nil", which we do not want.
	if not s then
		return nil
	end

	return string.gsub(tostring(s), '"', '""')
end

function strings.csv_header_from_table(t)
	for k, v in pairs(t) do
		t[k] = strings.csv_escape(v)
	end
	-- Appease MS Excel by starting with UTF-8 BOM
	return "\xEF\xBB\xBF" .. '"' .. table.concat(t, '","') .. '"' .. "\n"
end

local function get_unit_str(unit_id)
	if unit_id and u.unit_str[unit_id] then
		return _(u.unit_str[unit_id])
	end
	return ""
end

local reader_meta = {
	__index = function(t, k)
		if k == "desc" then
			return _(u.reader_desc[t.name])
		elseif k == "val" and t.unit_id then
			return t._v
		elseif k == "unit" and t.unit_id then
			return get_unit_str(t.unit_id)
		else
			return nil
		end
	end,

	__tostring = function(t)
		if t.name == "serial" then
			return string.format("%.5X", t._v)
		elseif t.unit_id then
			if t.name == "pixel_size" or t.name == "base_aperture" then
				return u.dim_str(u.mk_dim(t.unit_id, "double", t._v, 1))
			else
				return u.dim_str(u.mk_dim(t.unit_id, "int", t._v))
			end
		else
			return t._v
		end
	end
}

local verification_meta = {
	__index = function(t, k)
		if k == "desc" then
			return _(u.verification_desc[t.name])
		elseif k == "val" and t.unit_id then
			return t._v
		elseif k == "unit" and t.unit_id then
			return get_unit_str(t.unit_id)
		else
			return nil
		end
	end,

	__tostring = function(t)
		grade_params = { overall_grade = true, pass_grade = true }
		if grade_params[t.name] then
			return string.format("%.1f", u.round(t._v, 1))
		elseif t.name == "isograde" then
			if not t.parent.overall_grade then
				return ""
			end

			local pre = t.is_dpm and "DPM " or ""
			local suf = t.is_dpm and "/45Q" or ""
			return string.format("%s%.1f/%.2d/%d%s (%s)", pre,
			                     u.round(t.parent.overall_grade._v, 1),
			                     u.round(t.parent.aperture._v / 25.4),
			                     t.parent.wavelength._v, suf,
			                     u.get_ansi(t.parent.overall_grade._v))
		elseif t.unit_id then
			return u.dim_str(u.mk_dim(t.unit_id, "int", t._v))
		elseif t.name == "grade_passfail" then
			if not t.parent.overall_grade or t.parent.pass_grade._v < 0 then
				return ""
			end

			local overall_grade = u.round(t.parent.overall_grade._v, 1)
			local pass_grade = u.round(t.parent.pass_grade._v, 1)
			local passfail_str = strings.status.pass
			local comp_str = "≥"

			if overall_grade < pass_grade then
				passfail_str = strings.status.fail
				comp_str = "<"
			end

			return string.format("%s (%s%.1f)", passfail_str, comp_str, pass_grade)
		elseif t.name == "overall_status" then
			if not t.parent.overall_grade then
				return strings.status[t._v]
			end

			local summary = tostring(t.parent.isograde)

			if t.parent.pass_grade._v >= 0 then
				summary = summary .. " - " .. tostring(t.parent.grade_passfail)
			end

			return summary
		else
			return t._v
		end
	end
}

local date_meta = {
	__index = function(t, k)
		if k == "desc" then
			if t.name == "scan_time" or t.name == "decode_time" then
				return _(u.decode_desc[t.name])
			elseif t.name == "caltime_user" or t.name == "caltime_factory" then
				return _(u.reader_desc[t.name])
			else
				return nil
			end
		elseif k =="raw" then
			return t._v
		elseif t._v < 1 then
			if t.name == "caltime_user" or t.name == "caltime_factory" then
				return _("Not Calibrated")
			else
				return _("N/a")
			end
		elseif k == "date" then
			return os.date("%x", t._v)
		elseif k == "time" then
			return os.date("%X", t._v)
		elseif k == "timezone" then
			return u.get_tz_str()
		else
			return nil
		end
	end,

	__tostring = function(t)
		if t._v < 1 then
			return _("N/a")
		else
			return u.get_time_str(t._v)
		end
	end
}

local version_meta = {
	__index = function(t, k)
		if k == "desc" then
			return _(u.version_desc[t.name])
		elseif k == "build_time" then
			return u.get_time_str(t._v.build_time)
		else
			return nil
		end
	end,

	__tostring = function(t)
		return string.format("%s.%s.%s", t._v.major, t._v.minor,
		                     t._v.build_id)
	end
}

local platform_meta = {
	__index = function(t, k)
		if k == "desc" then
			return _(u.platform_desc[t.name])
		else
			return nil
		end
	end,

	__tostring = function(t)
		return t._v
	end
}

local decode_meta = {
	__index = function(t, k)
		if k == "desc" then
			return _(u.decode_desc[t.name])
		elseif t.name == "msg" and k == "html" then
			-- For EAN/UPC: don't escape spaces, make opts table empty
			local opt = u.is_ean_upc[t.parent.subsymb._v] and {}
			return function(span_class)
				return u.html_msg(t._v, span_class, opt)
			end
		elseif t.name == "msg" and k == "get_ai" then
			local aicu = require "aicheckutils"
			return function(ai)
				return aicu.get_ai(t._v, t.parent.subsymb._v, ai)
			end
		else
			return nil
		end
	end,

	__tostring = function(t)
		if t.name == "msg" then
			-- For EAN/UPC: don't escape spaces, make opts table empty
			local opt = u.is_ean_upc[t.parent.subsymb._v] and {}
			return u.parse_msg(t._v, opt)
		elseif t.name == "fullsymb" then
			return u.fullsymb_str(t.parent.symbology._v, t.parent.subsymb._v)
		else
			return _(u.symbology_desc[t._v])
		end
	end
}

local spec_meta = {
	__index = function(t, k)
		if k == "desc" then
			return _(u.spec_desc[t.name])
		else
			return nil
		end
	end,

	__tostring = function(t)
		if t.name == "print_spec" then
			return u.get_print_spec(t.parent.symbology._v, t.parent.isograde.is_dpm)
		elseif t.name == "symb_spec" then
			local qz_size = t.parent.qz_size and t.parent.qz_size.val
			return u.get_symb_spec(t.parent.symbology._v, t.parent.subsymb._v,
			                       qz_size)
		elseif t.name == "verif_spec" then
			return u.get_verif_spec(t.parent.symbology._v)
		end
	end
}

local general_meta = {
	__index = function(t, k)
		if k == "desc" then
			return _(u.param_desc[t._v.id])
		elseif k == "val" and t._v.param then
			if t._v.valid then
				return u.base_param_str(t._v.param)
			else
				return _("Invalid")
			end
		elseif k == "percent" and t._v.percent then
			if t._v.valid then
				return u.base_param_str(t._v.percent)
			else
				return _("Invalid")
			end
		elseif k == "unit" and t._v.unit then
			if t._v.valid then
				return get_unit_str(t._v.unit)
			else
				return _("Invalid")
			end
		else
			return nil
		end
	end,

	-- TODO: add options for formatting e.g. newlines between components,
	-- conditional for showing min/max values etc.
	__tostring = function(t)
		if t._v.valid then
			return u.param_str(t._v)
		else
			return _("Invalid")
		end
	end
}

local iso_meta = {
	__index = function(t, k)
		if k == "desc" then
			return _(u.param_desc[t._v.id])
		elseif k == "grade_ansi" and t._v.grade then
			if t._v.valid then
				return string.format("%s (%s)", u.base_param_str(t._v.grade),
				                     u.get_ansi(t._v.grade.val))
			else
				return _("Invalid")
			end
		elseif k == "grade" and t._v.grade then
			if t._v.valid then
				return u.base_param_str(t._v.grade)
			else
				return _("Invalid")
			end
		elseif k == "ansi" and t._v.grade then
			if t._v.valid then
				return u.get_ansi(t._v.grade.val)
			else
				return _("Invalid")
			end
		elseif k == "percent" and t._v.percent then
			if t._v.valid then
				return u.base_param_str(t._v.percent)
			else
				return _("Invalid")
			end
		else
			return nil
		end
	end,

	-- TODO: add options for formatting e.g. newlines between components,
	-- conditional for showing min/max values etc.
	__tostring = function(t)
		if not t._v.valid then
			return _("Invalid")
		elseif not t.grade and not t.percent then
			return _("N/a")
		end

		local s = ""
		if t.grade then
			s = string.format("%s (%s)", t.grade, t.ansi)
		end
		if t.percent then
			if s ~= "" then
				s = s .. ", "
			end
			s = s .. t.percent .. "%"
		end
		return s
	end
}

local plugins_meta = {
	__index = function(t, k)
		if k == "desc" then
			-- Return plugin name if description unavailable
			return t._v.desc or t._v.id
		elseif k == "status" then
			return _(u.plugin_status_str[t._v.status])
		elseif k == "output" then
			return t._v.output
		elseif k == "details" then
			return t._v.details or {}
		else
			return nil
		end
	end,

	__tostring = function(t)
		return string.format("%s (%s): %s", t._v.desc or t._v.id,
		                     _(u.plugin_status_str[t._v.status]), t._v.output)
	end
}

local userdata_meta = {
	__index = function(t, k)
		if k == "status" then
			return _(u.plugin_status_str[t._v.status])
		else
			return t._v[k]
		end
	end,

	__tostring = function(t)
		if t._v.status ~= "na" then
			return string.format("%s (%s): %s", t._v.title, t.status, t.value)
		else
			return string.format("%s: %s", t._v.title, t.value)
		end
	end
}

local param_unit_id = {
	pixel_size = "um",
	aperture = "um",
	base_aperture = "um",
	wavelength = "nm",
}

local function fill_reader_info(reader, ri)
	-- TODO: use proper dim tables for pixel size and base aperture
	local include = {
		["image_width"]   = true,
		["image_height"]  = true,
		["pixel_size"]    = true,  -- number (not integer)
		["base_aperture"] = true,  -- number (not integer)
		["wavelength"]    = true,
		["serial"]        = true,
	}

	local epochs = {
		["caltime_factory"] = true,
		["caltime_user"]    = true,
	}

	for k, v in pairs(ri) do
		if epochs[k] then
			reader[k] = { name = k, _v = v }
			setmetatable(reader[k], date_meta)
		elseif include[k] then
			reader[k] = { name = k, _v = v, unit_id = param_unit_id[k] }
			setmetatable(reader[k], reader_meta)
		end
	end
end

local function fill_decode_info(general, di)
--	local exclude = {
--		backend_type = true,
--		location = true,
--		module_width = true,
--		recommended_aperture = true,
--	}
	local include = {
		["symbology"] = true,
		["subsymb"]   = true,
		["msg"]       = true,
	}

	local epochs = { ["decode_time"] = true }

	for k, v in pairs(di) do
		if epochs[k] then
			general[k] = { name = k, _v = v }
			setmetatable(general[k], date_meta)
		elseif include[k] then
			general[k] = { name = k, _v = v, parent = general }
			setmetatable(general[k], decode_meta)
		end
	end
	-- Make the fullsymb string available under "general.fullsymb"
	general.fullsymb = { name = "fullsymb", parent = general }
	setmetatable(general.fullsymb, decode_meta)
end

local function fill_spec_info(general)
	local include = {
		["print_spec"] = true,
		["symb_spec"]  = true,
		["verif_spec"] = true,
	}

	for k, v in pairs(include) do
		general[k] = { name = k, parent = general }
		setmetatable(general[k], spec_meta)
	end
end

local function fill_verif_info(general, vi)
	local include = {
		["pass_grade"]    = true,
		["overall_grade"] = true,
		["aperture"]      = true,
		["wavelength"]    = true,
		["overall_status"] = true,
	}

	-- TODO: what do we do if "has_iso" is false?
	for k, v in pairs(vi) do
		if include[k] then
			general[k] = { name = k, _v = v, unit_id = param_unit_id[k],
			               parent = general }
			setmetatable(general[k], verification_meta)
		end
	end
	-- Make the formal grade available under "general.isograde"
	general.isograde = { name = "isograde",
	                     is_dpm = u.is_dpm({ verification_info = vi }),
	                     parent = general }
	setmetatable(general.isograde, verification_meta)

	general.grade_passfail = { name = "grade_passfail", parent = general }
	setmetatable(general.grade_passfail, verification_meta)
end

local function fill_version_info(version_dest, swvi)
	local include = {
		["code"]    = true,
		["current"] = true,
	}

	for k, v in pairs(swvi) do
		if include[k] then
			version_dest[k] = { name = k, _v = v }
			setmetatable(version_dest[k], version_meta)
		end
	end
end

local function fill_platform_info(platform_dest, pfi)
	local req = { "username", "hostname", "osversion" }

	if type(pfi) ~= "table" then
		pfi = {}
	end

	for k, v in ipairs(req) do
		if not pfi[v] then
			pfi[v] = u.platform_desc.unknown
		end
	end

	for k, v in pairs(pfi) do
		platform_dest[k] = { name = k, _v = v }
		setmetatable(platform_dest[k], platform_meta)
	end
end

local function fill_params(general, iso, pl)
	for k, v in ipairs(pl) do
		if v.dtype == "general" then
			local id = v.id
			if general[id] then
				-- duplicate top-level ID, needs a new tag
				id = v.id .. "2"
			end

			table.insert(general, id)
			general[id] = { _v = v }
			setmetatable(general[id], general_meta)
		elseif (v.dtype == "iso") and (v.type == "iso") then
			table.insert(iso, v.id)
			iso[v.id] = { _v = v }
			setmetatable(iso[v.id], iso_meta)

			-- Let dmx_fpd_* tags work for reports that explicitly use them
			if v.id == "fixed_pattern_damage" then
				for kk, vv in ipairs(v) do
					iso[vv.id] = { _v = vv }
					setmetatable(iso[vv.id], iso_meta)
				end
			end
		end
	end
end

local function fill_plugins(plugins, pil)
	local has_pass, has_warn, has_fail = false, false, false
	for k, v in ipairs(pil) do
		if v.status == "na" then
			-- Do nothing
		else
			table.insert(plugins, v.id)
			local set_wide_info
			if v.details then
				for i, d in ipairs(v.details) do
					v.details[i].status = _(u.plugin_status_str[d.status])
					-- Make info column wider if any details include notes
					if v.details[i].notes and #v.details[i].notes > 0 then
						set_wide_info = true
					end
				end
			end
			plugins[v.id] = { _v = v, wide_info = set_wide_info }
			setmetatable(plugins[v.id], plugins_meta)

			if v.status == "fail" then
				has_fail = true
			elseif v.status == "warn" then
				has_warn = true
			elseif v.status == "pass" then
				has_pass = true
			end
		end
	end
	local overall_status
	if has_fail then
		overall_status = "fail"
	elseif has_warn then
		overall_status = "warn"
	elseif has_pass then
		overall_status = "pass"
	else
		overall_status = "na"
	end
	plugins.overall_desc = _("Application Validation")
	plugins.overall_status = _(u.plugin_status_str[overall_status])
end

local function fill_userdata(userdata, u)
	for k, v in ipairs(u) do
		table.insert(userdata, v.id)
		userdata[v.id] = { _v = v }
		setmetatable(userdata[v.id], userdata_meta)
	end
end

local image_meta = {
	__index = function(t, k)
		if k == "base64" then
			return u.base64_enc(t._v)
		else
			return nil
		end
	end,

	__tostring = function(t)
		return t._v
	end
}

local function fill_images(img_dest, img)
	local include = {
		["cropped"] = true,
		["full"] = true
	}

	for k, v in pairs(img) do
		if include[k] then
			img_dest[k] = { _v = v }
			setmetatable(img_dest[k], image_meta)
		end
	end
end

local mod = {}

function mod.report_info(info)
	local t = require "template"
	local template_func = t.compile(info.template)

	-- This will probably fail due to there being no code, but if the info
	-- function is declared before it bombs out we should be able to call it.
	pcall(template_func, {
		info = info,
		strings = strings
	})

	if t.context.acv_template_info then
		return t.context.acv_template_info()
	else
		return nil
	end
end

function mod.create_report(info)
	local t = require "template"
	local f

	if info.template_string then
		-- Input and output are strings
		t.load = function(s)
			return s
		end
		t.print = function(s)
			return s
		end
	else
		local iomode = "w"
		if info.append then
			iomode = "a"
		end

		f = assert(io.open(info.outfile, iomode))
		t.print = function(...)
			f:write(...)
		end
	end

	u.set_metric(info.metric)

	local reader  = {}
	-- Reader info is per scan, grab it from the first code
	fill_reader_info(reader, info.code.reader_info)

	local userdata = {}
	fill_userdata(userdata, info.userdata)

	-- Add info.code as only array member if info.codes doesn't exist
	if not info.codes then
		info.codes = { info.code }
	end

	local codes = {}

	for i, c in ipairs(info.codes) do
		--print(i, c.decode_info.msg)

		codes[i] = {}
		codes[i].general = {
			version  = {},
			platform = {},
			image    = {},
		}

		codes[i].general.scan_time = { name = "scan_time", _v = c.scan_time }
		setmetatable(codes[i].general.scan_time, date_meta)

		fill_decode_info(codes[i].general, c.decode_info)
		fill_spec_info(codes[i].general)

		fill_verif_info(codes[i].general, c.verification_info)
		fill_version_info(codes[i].general.version, c.version)
		fill_platform_info(codes[i].general.platform, c.platform_info)
		fill_images(codes[i].general.image, c.image)

		codes[i].iso = {}
		fill_params(codes[i].general, codes[i].iso, c.param_list)

		codes[i].plugins = {}
		fill_plugins(codes[i].plugins, c.plugins)
	end

	local context = {
		info = info,
		reader = reader,
		general = codes[1].general,
		iso = codes[1].iso,
		plugins = codes[1].plugins,
		userdata = userdata,
		strings = strings,
		codes = codes,
	}

	-- Add full image only if it exists
	if info.image then
		context.image = {}
		fill_images(context.image, info.image)
	end

	if info.template_string then
		return {
			output = t.render(info.template_string, context)
		}
	else
		local template_func = t.compile(info.template)
		local template_output = template_func(context)

		if (u.filesize(f) == 0) and t.context.acv_report_headers then
			f:write(t.context.acv_report_headers())
		end
		f:write(template_output)
		f:close()
	end
end

return mod
