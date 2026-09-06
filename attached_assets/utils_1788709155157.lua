--require "strict"
require "gettext-utils"

local mod = { using_metric = true }

function mod.set_metric(is_metric)
	mod.using_metric = is_metric
end

mod.bool_str = {
	yes_no          = { [false] = N_("No"),    [true] = N_("Yes")  },
	true_false      = { [false] = N_("False"), [true] = N_("True") },
	pass_fail       = { [false] = N_("Fail"),  [true] = N_("Pass") },
	present_absent  = { [false] = N_("Absent"),  [true] = N_("Present") },
}

mod.plugin_status = {
	pass = 0,
	fail = 1,
	warn = 2,
	na   = 3,
}

local return_unknown_meta = {
	__index = function(t, k)
		return string.format(N_("Unknown (%s)"), k)
	end
}

mod.plugin_status_str = {
	pass = N_("Pass"),
	fail = N_("Fail"),
	warn = N_("Warn"),
	na   = N_("N/a"),
}
setmetatable(mod.plugin_status_str, return_unknown_meta)

-- Physical quantities, some may need to be converted to imperial units.
mod.unit_str = {
	um = N_("µm"),
	nm = N_("nm"),
	mm = N_("mm"),
}

mod.is_2d_code = {
	qr  = true,
	dmx = true,
}

local return_key_meta = {
	__index = function(t, k)
		return k
	end
}

mod.decode_desc = {
	symbology   = N_("Base Symbology"),
	subsymb     = N_("Sub-symbology"),
	fullsymb    = N_("Symbology"),
	msg         = N_("Decoded Message"),
	scan_time   = N_("Scan Time"),
	decode_time = N_("Decode Time"),
}
setmetatable(mod.decode_desc, return_key_meta)

mod.version_desc = {
	current = N_("Current Software Version"),
	code    = N_("Software Version"),
}
setmetatable(mod.version_desc, return_key_meta)

mod.platform_desc = {
	username  = N_("User Name"),
	hostname  = N_("Computer Name"),
	osversion = N_("OS Version"),
	unknown   = N_("Unknown"),
}
setmetatable(mod.platform_desc, return_key_meta)

mod.reader_desc = {
	--[[ reader_info ]]--
	serial          = N_("Serial Number"),
	base_aperture   = N_("Base Aperture"),
	wavelength      = N_("Wavelength"),
	caltime_factory	= N_("Factory Configuration"),
	caltime_user    = N_("User Calibration"),
	image_height    = N_("Image Height"),
	image_width     = N_("Image Width"),
	pixel_size      = N_("Pixel Size"),
}
setmetatable(mod.reader_desc, return_key_meta)

mod.verification_desc = {
	overall_grade = N_("Overall Grade"), -- Unused
	pass_grade    = N_("Pass grade"),
	isograde      = N_("ISO Grade"),
	aperture      = N_("Aperture"),
	wavelength    = N_("Wavelength"),
	overall_status = N_("Print Quality"),
}
setmetatable(mod.verification_desc, return_key_meta)

mod.spec_desc = {
	print_spec = N_("Print Quality Standard"),
	symb_spec  = N_("Symbology Specification"),
	verif_spec = N_("Verifier Conformance Standard"),
}
setmetatable(mod.spec_desc, return_key_meta)

-- Strings for plugins report
mod.application_strs = {
	title    = N_("Application Validation Report"),
	element  = N_("Element"),
	value    = N_("Value"),
	desc     = N_("Description"),
	status   = N_("Status"),
	info     = N_("Information"),
}
setmetatable(mod.application_strs, return_key_meta)

mod.symbology_desc = {
	--[[ Symbology ]]--
	qr         = N_("QR Code"),
	dmx        = N_("Data Matrix"),
	linear     = N_("Linear"),
	pdf417     = N_("PDF417"),
	laetus     = N_("Laetus"),
	------[[ Laetus ]]------
	standard  = N_("Standard"),
	minicode  = N_("Minicode"),
	boxcode   = N_("Box code"),
	microcode = N_("Microcode"),
	------[[ Linear ]]------
	nosymb      = N_("Not decoded"),
	ean8        = N_("EAN-8"),
	ean8_2      = N_("EAN-8+2"),
	ean8_5      = N_("EAN-8+5"),
	ean13       = N_("EAN-13"),
	ean13_2     = N_("EAN-13+2"),
	ean13_5     = N_("EAN-13+5"),
	upce        = N_("UPC-E"),
	upce_2      = N_("UPC-E+2"),
	upce_5      = N_("UPC-E+5"),
	upca        = N_("UPC-A"),
	upca_2      = N_("UPC-A+2"),
	upca_5      = N_("UPC-A+5"),
	ean_unknown = N_("Unknown EAN/UPC"),
	code39      = N_("Code 39"),
	code39_tri  = N_("Code 39 Trioptic"),
	itf         = N_("ITF"),
	itf14       = N_("ITF-14"),
	codabar     = N_("Codabar"),
	code128     = N_("Code 128"),
	gs1_128     = N_("GS1-128"),
	code128_kraft   = N_("Kraft Tassimo C128"),
	code128_kraft_a = N_("Kraft Tassimo C128 (A)"),
	code128_kraft_b = N_("Kraft Tassimo C128 (B)"),
	databar          = N_("GS1 Databar"),
	databar_expanded = N_("GS1 Databar Expanded"),
	databar_limited  = N_("GS1 Databar Limited"),
	databar_unknown  = N_("GS1 Databar (Unknown)"),
	msi_plessey = N_("MSI Plessey"),
	code93      = N_("Code 93"),
	unknown     = N_("Unknown"),
	------[[ QR: must quote key if it begins with digit ]]------
	model1              = N_("Model 1"),
	["2005"]            = N_("2005"),
	["2005_eci"]        = N_("2005, ECI"),
	["2005_fnc1_1"]     = N_("2005, FNC1 implied in 1st position"),
	["2005_eci_fnc1_1"] = N_("2005, FNC1 implied in 1st position, ECI"),
	["2005_fnc1_2"]     = N_("2005, FNC1 implied in 2nd position"),
	["2005_eci_fnc1_2"] = N_("2005, FNC1 implied in 2nd position, ECI"),
	------[[ DMX ]]------
	ecc_000_140          = N_("ECC 000-140"),
	ecc_200              = N_("ECC 200"),
	ecc_200_fnc1_1_5     = N_("ECC 200, FNC1 in 1st or 5th position"),
	ecc_200_fnc1_2_6     = N_("ECC 200, FNC1 in 2nd or 6th position"),
	ecc_200_eci          = N_("ECC 200, ECI"),
	ecc_200_eci_fnc1_1_5 = N_("ECC 200, FNC1 in 1st or 5th position, ECI"),
	ecc_200_eci_fnc1_2_6 = N_("ECC 200, FNC1 in 2nd or 6th position, ECI"),
	dmre                 = N_("DMRE"),
	dmre_fnc1_1_5        = N_("DMRE, FNC1 in 1st or 5th position"),
	dmre_fnc1_2_6        = N_("DMRE, FNC1 in 2nd or 6th position"),
	dmre_eci             = N_("DMRE, ECI"),
	dmre_eci_fnc1_1_5    = N_("DMRE, FNC1 in 1st or 5th position, ECI"),
	dmre_eci_fnc1_2_6    = N_("DMRE, FNC1 in 2nd or 6th position, ECI"),
	------[[ (Micro)PDF417 & Composites ]]------
	standard_pdf417 = N_("PDF417"),
	micro_pdf417    = N_("MicroPDF417"),
	cca             = N_("GS1 Composite Component (CC-A)"),
	ccb             = N_("GS1 Composite Component (CC-B)"),
	ccc             = N_("GS1 Composite Component (CC-C)"),
}
setmetatable(mod.symbology_desc, return_key_meta)

mod.is_ean_upc = {
	ean8    = true,
	ean8_2  = true,
	ean8_5  = true,
	ean13   = true,
	ean13_2 = true,
	ean13_5 = true,
	upce    = true,
	upce_2  = true,
	upce_5  = true,
	upca    = true,
	upca_2  = true,
	upca_5  = true,
}

mod.is_databar = {
	databar          = true,
	databar_expanded = true,
	databar_limited  = true,
}

mod.is_composite = {
	cca = true,
	ccb = true,
	ccc = true,
}

mod.is_dmre = {
	dmre              = true,
	dmre_fnc1_1_5     = true,
	dmre_fnc1_2_6     = true,
	dmre_eci          = true,
	dmre_eci_fnc1_1_5 = true,
	dmre_eci_fnc1_2_6 = true,
}

function mod.fullsymb_str(symb, subsymb)
	if (symb == "qr") and (subsymb == "2005_fnc1_1") then
		return _("GS1 QR")
	elseif (symb == "dmx") and (subsymb == "ecc_200_fnc1_1_5") then
		return _("GS1 DataMatrix")
	elseif (symb == "linear") or (symb == "pdf417") then
		return _(mod.symbology_desc[subsymb])
	else
		return _(mod.symbology_desc[symb]) .. " - " .. _(mod.symbology_desc[subsymb])
	end
end

mod.param_desc = require "dms_param"
setmetatable(mod.param_desc, return_key_meta)

function mod.get_ansi(grade)
	if not grade then
		return nil
	end
	if grade < 0.45 then
		return "F"
	elseif grade < 1.45 then
		return "D"
	elseif grade < 2.45 then
		return "C"
	elseif grade < 3.45 then
		return "B"
	else
		return "A"
	end
end

function mod.base_param_str(p)
	if not p then
		return nil
	end
	if p.type == "int" then
		return tostring(p.val)
	elseif p.type == "char" then
		if string.byte(p.val) < 20 or string.byte(p.val) > 126 then
			return '?'
		end

		return string.format("'%s'", p.val)
	elseif p.type == "double" then
		return string.format("%." .. p.precision .. "f",
		                     mod.round(p.val, p.precision))
	elseif p.type == "bool" then
		return _(mod.bool_str[p.booltype][p.val])
	else
		return nil
	end
end

-- If param array contains any childen with statuses, show a count of each.
-- Order declared in show table determines the order in the output string.
-- If there are no statuses, concatenate all string params.  If no strings,
-- return "N/a"
local function param_array_str(pa)
	local show = { "pass", "warn", "fail" }
	local count = {}

	if (not pa) or (type(pa) ~= "table") or (#pa == 0) or
	   (pa.type ~= "param_array" and pa.type ~= "param_loc_array") then
		return _(mod.plugin_status_str.na)
	end

	-- Initialise all counts to 0
	for _, st in ipairs(show) do
		count[st] = 0
	end

	for _, v in ipairs(pa) do
		local st = v.status

		if count[st] then
			count[st] = count[st] + 1
		end
	end

	local s = {}
	for k, st in ipairs(show) do
		if count[st] > 0 then
			table.insert(s, string.format("%d %s", count[st],
			                              _(mod.plugin_status_str[st])))
		end
	end

	if #s > 0 then
		return table.concat(s, ", ")
	end

	-- Concatenate child params, space separated
	for _, v in ipairs(pa) do
		if v.type == "string" then
			table.insert(s, mod.parse_msg(v.litstr or v.str, v.litstr and {}))
		end
	end

	if #s > 0 then
		return table.concat(s, " ")
	end

	return _(mod.plugin_status_str.na)
end

function mod.param_str(p)
	if not p.valid then
		return _("Invalid")
	end

	if p.type == "base" then
		return mod.base_param_str(p.param)
	elseif p.type == "dim" then
		return mod.dim_str(p)
	elseif p.type == "percent" then
		return mod.base_param_str(p.param) .. "%"
	elseif p.type == "check" then
		local isok = (p.actual.val == p.expected.val)
		local s = string.format("%s: %s", _(mod.bool_str.pass_fail[isok]),
		                        mod.base_param_str(p.actual))
		if not isok then
			s = string.format("%s (%s: %s)", s, _("Expected"),
			                  mod.base_param_str(p.expected))
		end
		return s
	elseif p.type == "param_array" or p.type == "param_loc_array" then
		return param_array_str(p)
	elseif p.type == "string" then
		return mod.parse_msg(p.litstr or p.str, p.litstr and {})
	else
		return _("N/a")
	end
end

-- When creating ints, truncate val as if using printf("%d")
local function mk_bp(bp_type, bp_val, bp_prec)
	if bp_type == "int" then
		bp_val = math.floor(bp_val)
	end
	return {
		type      = bp_type,
		val       = bp_val,
		precision = bp_prec,  -- Only if fp type
	}
end

-- Converts mm to inches, or um to mils
local function metric_to_imperial(param, unit)
	if (unit == "mm" or unit == "um") and
	   (param.type == "int" or param.type == "double") then
		local imperial_prec = (param.precision or 0) + 2

		return mk_bp("double", mod.round_param(param) / 25.4, imperial_prec)
	end

	return param
end

-- Convert metric dimensions to imperial if necessary.
local function dim_str_fmt(param, unit, opt)
	local can_convert = { um = true, mm = true }
	local is_imperial = can_convert[unit] and (not mod.using_metric)
	local val_str
	opt = opt or { show_units = true }

	-- Catch imperial
	if is_imperial then
		val_str = mod.base_param_str(metric_to_imperial(param, unit))
	else
		val_str = mod.base_param_str(param)
	end

	if not opt.show_units then
		return val_str
	end

	local sep_str, unit_str = "", ""

	if is_imperial then
		if unit == "um" then
			sep_str = " "
			unit_str = _("mil")
		elseif unit == "mm" then
			unit_str = _('"')
		end

	elseif mod.unit_str[unit] then
		sep_str = " "
		unit_str = _(mod.unit_str[unit])
	elseif unit == "ratio" then
		unit_str = ":1"
	elseif unit == "modules" then
		unit_str = "x"
	end

	return val_str .. sep_str .. unit_str
end

function mod.mk_dim(d_unit, bp_type, bp_val, bp_prec)
	return {
		type = "dim",
		unit = d_unit,
		param = mk_bp(bp_type, bp_val, bp_prec),
	}
end

-- Get the displayed resolution of a param.  E.g. if the precision of a double
-- type is 1, the smallest visible change in results is 0.1.
-- All non-fp types are treated as integers, so have a resolution of 1.
local function get_epsilon(p)
	if p.type ~= "double" then
		return 1
	end

	return 10 ^ -p.precision
end

local function is_range_symmetrical(min, max)
	return ((mod.round_param(min) + mod.round_param(max)) <
	        math.min(get_epsilon(min), get_epsilon(max)))
end

-- Convert a dimension parameter to a string.
-- Format:  val[_units][ (abs min|max|range)][, pc%[ (pc% min|max|range)]]
-- E.g:  "330 µm (264..660), 100% (80%..200%)"
function mod.dim_str(p)
	local s = dim_str_fmt(p.param, p.unit)
	local opt = { show_units = false }

	if p.min and p.max then
		if is_range_symmetrical(p.min, p.max) then
			s = string.format("%s (±%s)", s, dim_str_fmt(p.max, p.unit, opt))
		else
			s = string.format("%s (%s..%s)", s, dim_str_fmt(p.min, p.unit, opt),
			                  dim_str_fmt(p.max, p.unit, opt))
		end
	elseif p.min then
		s = string.format("%s (min %s)", s, dim_str_fmt(p.min, p.unit, opt))
	elseif p.max then
		s = string.format("%s (max %s)", s, dim_str_fmt(p.max, p.unit, opt))
	end

	if p.percent then
		s = string.format("%s, %s%%", s, mod.base_param_str(p.percent))
	end

	if p.min_percent and p.max_percent then
		if is_range_symmetrical(p.min_percent, p.max_percent) then
			s = string.format("%s (±%s%%)", s, mod.base_param_str(p.max_percent))
		else
			s = string.format("%s (%s%%..%s%%)", s,
			       mod.base_param_str(p.min_percent), mod.base_param_str(p.max_percent))
		end
	elseif p.min_percent then
		s = string.format("%s (min %s%%)", s, mod.base_param_str(p.min_percent))
	elseif p.max_percent then
		s = string.format("%s (max %s%%)", s, mod.base_param_str(p.max_percent))
	end

	return s
end

function mod.get_print_spec(symb, is_dpm)
	local pqs_str = _("N/a")
	if is_dpm then
		pqs_str = "ISO/IEC 29158"
	elseif symb == "qr" or symb == "dmx" or symb == "pdf417" then
		pqs_str = "ISO/IEC 15415"
	elseif symb == "linear" then
		pqs_str = "ISO/IEC 15416"
	end
	return pqs_str
end

function mod.get_verif_spec(symb)
	local vcs_str = _("N/a")
	if symb == "qr" or symb == "dmx" or symb == "pdf417" then
		vcs_str = "ISO/IEC 15426-2"
	elseif symb == "linear" then
		vcs_str = "ISO/IEC 15426-1"
	end
	return vcs_str
end

function mod.get_symb_spec(symb, subsymb, qz_size)
	local symbspec_str = _("N/a")

	if symb == "qr" then
		-- qz_size only exists for non-standard QR
		if not qz_size then
			symbspec_str = "ISO/IEC 18004"
		end

	elseif symb == "dmx" then
		if mod.is_dmre[subsymb] then
			symbspec_str = "ISO/IEC 21471"
		else
			symbspec_str = "ISO/IEC 16022"
		end

	elseif symb == "pdf417" then
		if mod.is_composite[subsymb] then
			symbspec_str = "ISO/IEC 24723"

		elseif subsymb == "micro_pdf417" then
			symbspec_str = "ISO/IEC 24728"

		else  -- if subsymb == "standard_pdf417" then
			symbspec_str = "ISO/IEC 15438"
		end

	elseif symb == "linear" then
		if mod.is_ean_upc[subsymb] or subsymb == "ean_unknown" then
			symbspec_str = "ISO/IEC 15420"

		elseif subsymb == "code39" or subsymb == "code39_tri" then
			symbspec_str = "ISO/IEC 16388"

		elseif subsymb == "itf" or subsymb == "itf14" then
			symbspec_str = "ISO/IEC 16390"

		elseif subsymb == "codabar" then
			symbspec_str = "EN 798"

		elseif subsymb == "code128" or subsymb == "gs1_128" then
			symbspec_str = "ISO/IEC 15417"

		elseif mod.is_databar[subsymb] or subsymb == "databar_unknown" then
			symbspec_str = "ISO/IEC 24724"

		elseif subsymb == "code93" then
			symbspec_str = "ANSI/AIM BC5-1995"
		end
	end

	return symbspec_str
end

function mod.get_msg_table(msg)
	local t = {}
	local got_bslash = false

	for c in msg:gmatch('.') do
		if got_bslash then
			if c:find("[1234%?\\]") then
				table.insert(t, '\\' .. c)
			else
				table.insert(t, c)
			end
			got_bslash = false
		elseif c == '\\' then
			got_bslash = true
		else
			table.insert(t, c)
		end
	end

	return t
end

-- True if no high bits are set
function mod.is_ascii(str)
	str = str or ""
	return string.match(str, "[%z\1-\127]+") == str
end

-- The message string must contain a check digit (can be a dummy).
-- Returns expected check digit (or nil) & read check digit.
function mod.mod10_check(msg)
	local sum = 0
	local triple = true
	local mst = mod.get_msg_table(msg)

	for i = #mst - 1, 1, -1 do
		local n = tonumber(mst[i])
		if not n then
			sum = nil
			break
		end

		if triple then
			n = (n * 3)
		end
		sum = sum + n
		triple = not triple
	end

	local cdx
	if sum then
		cdx = tostring(-sum % 10)
	end

	return cdx, mst[#mst]
end

-- The message string must contain a check digit (can be a dummy).
-- Returns expected check digit (or nil) & read check digit.
-- https://en.wikipedia.org/wiki/Luhn_algorithm
function mod.luhn_mod10_check(msg)
	-- double the input (0-9) and sum the digits of the result
	local twice_digit_sum = { [0] = 0, 2, 4, 6, 8, 1, 3, 5, 7, 9 }
	local alternate = true  -- Which positions are "weighted"
	local sum = 0
	local mst = mod.get_msg_table(msg)

	for i = #mst - 1, 1, -1 do
		local n = tonumber(mst[i])
		if not n then
			sum = nil
			break
		end

		if alternate then
			n = twice_digit_sum[n]
		end
		sum = sum + n
		alternate = not alternate
	end

	local cdx
	if sum then
		cdx = tostring(-sum % 10)
	end

	return cdx, mst[#mst]
end

-- The message string must contain a check digit (can be a dummy).
-- Returns expected check value ('X' for 10) or nil, & read check digit.
function mod.mod11_pzn_check(msg)
	local sum = 0
	local mst = mod.get_msg_table(msg)

	for i = 1, #mst - 1 do
		local n = tonumber(mst[i])
		if not n then
			sum = nil
			break
		end

		sum = sum + (n * i)
	end

	local cdx
	if sum then
		cdx = tostring(sum % 11)
		if #cdx ~= 1 then
			cdx = 'X'
		end
	end

	return cdx, mst[#mst]
end

-- The message string must contain a check digit (can be a dummy).
-- Returns expected check character (or nil) & read check character.
function mod.mod43_check(msg)
	local charset = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%"
	-- Lookup by char, values 0 to 42.
	local t = {}
	for i = 1, #charset do
		t[charset:sub(i, i)] = i - 1
	end

	local sum = 0
	local mst = mod.get_msg_table(msg)

	for i = #mst - 1, 1, -1 do
		local n = t[mst[i]]
		if not n then
			sum = nil
			break
		end

		sum = sum + n
	end

	local cdx
	if sum then
		local pos = (sum % 43) + 1
		cdx = charset:sub(pos, pos)
	end

	return cdx, mst[#mst]
end

local function fnc_str(c, str_enclose_func, opt)
	opt = opt or {}

	if opt.fnc then
		return str_enclose_func('FNC' .. c)
	else
		return c
	end
end

local function badchar_str(str_enclose_func)
	return str_enclose_func('BAD')
end

local function char_str(c, str_enclose_func, opt)
	local ctrl_abbrev = {}
	opt = opt or {}
	opt.esc_func = opt.esc_func or tostring

	if opt.ctrl then
		ctrl_abbrev = {
			[0] = 'NUL', 'SOH', 'STX', 'ETX', 'EOT', 'ENQ', 'ACK', 'BEL',
			      'BS',   nil,   nil,  'VT',  'FF',   nil,  'SO',  'SI',
			      'DLE', 'DC1', 'DC2', 'DC3', 'DC4', 'NAK', 'SYN', 'ETB',
			      'CAN', 'EM',  'SUB', 'ESC', 'FS',  'GS',  'RS',  'US'
		}
	end

	if opt.whitespace then
		ctrl_abbrev[ 9] = 'HT'
		ctrl_abbrev[10] = 'LF'
		ctrl_abbrev[13] = 'CR'
		ctrl_abbrev[32] = 'SP'
	end

	local b = string.byte(c)

	if ctrl_abbrev[b] then
		return str_enclose_func(ctrl_abbrev[b])
	elseif (b > 126) and opt.hibit then
		return str_enclose_func(string.format("0x%.2X", b))
	else
		return opt.esc_func(c)
	end
end

local function do_parse_msg(msg, str_enclose_func, opt)
	if not msg or #msg == 0 then
		return ''
	end

	local parsed_msg = {}
	local got_bslash = false

	for i = 1, #msg do
		local c = msg:sub(i, i)

		if got_bslash then
			local n = tonumber(c)

			if n and (n >= 1) and (n <= 4) then
				table.insert(parsed_msg, fnc_str(c, str_enclose_func, opt))
			elseif c == '?' then
				table.insert(parsed_msg, badchar_str(str_enclose_func))
			else
				table.insert(parsed_msg, char_str(c, str_enclose_func, opt))
			end
			got_bslash = false
		elseif c == '\\' then
			got_bslash = true
		else
			table.insert(parsed_msg, char_str(c, str_enclose_func, opt))
		end
	end

	return table.concat(parsed_msg, opt and opt.char_sep)
end

-- Parse a decoded message and display control and function characters in a
-- non-ambiguous way.
function mod.parse_msg(msg, opt)
	opt = opt or { fnc = true, ctrl = true, whitespace = true, hibit = true }

	return do_parse_msg(msg, function(s)
		return '«' .. s .. '»'
	end, opt)
end

local HTML_ENTITIES = {
	["&"] = "&amp;",
	["<"] = "&lt;",
	[">"] = "&gt;",
	['"'] = "&quot;",
	["'"] = "&#39;",
	["/"] = "&#47;"
}

function mod.html_escape(s)
	-- gsub also returns number of substitutions, but we only want the string
	return (s:gsub("[\">/<'&]", HTML_ENTITIES))
end

function mod.html_msg(msg, span_class, opt)
	opt = opt or { fnc = true, ctrl = true, whitespace = true, hibit = true,
	               esc_func = mod.html_escape, char_sep = "<wbr>" }

	return do_parse_msg(msg, function(s)
		local r

		if span_class then
			r = string.format('<span class="%s">', span_class)
		else
			r = "<span>"
		end

		return r .. s .. "</span><wbr>"
	end, opt)
end

-- Compute the difference in seconds between local time and UTC.
function mod.get_local_utc_offset(ts)
	local local_t = os.date("*t", ts)
	local_t.isdst = false
	local utc_t = os.date("!*t", ts)

	-- return os.difftime(os.time(utc_t), os.time(local_t))
	return os.time(local_t) - os.time(utc_t)
end

-- Return a timezone string in ISO 8601:2000 standard form (+hhmm or -hhmm)
function mod.get_tz_str()
	local tz = mod.get_local_utc_offset(os.time())
	local h, m = math.modf(tz / 3600)

	return string.format("%+.4d", (100 * h) + (60 * m))
end

-- Returns a date & time string with time zone.
-- Since we have to manually add on the time zone, ensure that the order of the
-- date, time & zone makes sense i.e. "<date> <time> (<tz>)".
function mod.get_time_str(t)
	return string.format("%s (%s)", os.date("%x %X", t), mod.get_tz_str())
end

-- Round params that may be stored as fp types to the relevant precision
function mod.round_param(param)
	if param.type == "double" then
		return mod.round(param.val, param.precision)

	elseif param.type == "int" then
		return param.val
	end

	return nil
end

local function get_mod_size_params(param_list)
	local xdim, ydim

	for _, v in ipairs(param_list) do
		if v.valid then
			if v.id == "xdim" and v.param.type then
				-- Can't remember why we added this catch in the first place.
				-- It probably shouldn't happen but keep it just in case.
				if v.unit == "mm" then  -- x1000 if stored in millimetres
					v.param.val = v.param.val * 1000
					v.param.precision = (v.param.precision or 0) + 3
				end
				xdim = v.param

			elseif v.id == "ydim" and v.param.type then
				if v.unit == "mm" then
					v.param.val = v.param.val * 1000
					v.param.precision = (v.param.precision or 0) + 3
				end
				ydim = v.param
			end
		end
	end

	return xdim, ydim
end

function mod.get_mod_sizes_um(param_list)
	local xdim, ydim = get_mod_size_params(param_list)

	if xdim then
		xdim = mod.round(xdim.val)
	end

	if ydim then
		ydim = mod.round(ydim.val)
	end

	return (xdim or -1), (ydim or -1)
end

-- Get the smallest of the x & y dims (just xdim for linears)
function mod.get_min_mod_size(param_list)
	local xdim, ydim = get_mod_size_params(param_list)

	if ydim and ((not xdim) or ydim.val < xdim.val) then
		return mod.round(ydim.val)
	elseif xdim then
		return mod.round(xdim.val)
	end

	return -1
end

-- actual_grade: Measured symbol grade
-- pass_grade:   Minimum grade to pass application spec (required)
-- target_grade: Minimum target grade, warn if symbol grade is lower (optional)
function mod.check_grade(plugin_details, actual_grade, pass_grade, target_grade)
	local actual_grade_1dp = mod.round(actual_grade, 1)
	local pass_grade_1dp = mod.round(pass_grade, 1)
	local target_grade_1dp = target_grade and mod.round(target_grade, 1)
	local pg_str = string.format("%.1f (%s)", pass_grade_1dp,
	                             mod.get_ansi(pass_grade))

	local new = {
		title = _("Grade"),
		value = string.format("%.1f (%s)", actual_grade_1dp,
		                      mod.get_ansi(actual_grade)),
		desc  = string.format(_("Grade must be at least %s"), pg_str),
	}

	if actual_grade_1dp < pass_grade_1dp then
		new.status = mod.plugin_status.fail
	--	new.errors = { string.format("%s < %s", _("Grade"), pg_str) }

	elseif target_grade_1dp and (actual_grade_1dp < target_grade_1dp) then
		local tg_str = string.format("%.1f (%s)", target_grade_1dp,
		                                          mod.get_ansi(target_grade))
		new.warnings = { string.format(_("Grade should be at least %s"), tg_str) }
		new.status = mod.plugin_status.warn
	else
		new.status = mod.plugin_status.pass
	end

	table.insert(plugin_details, new)
end

-- spec_aperture between 0 & 1 is a percentage of the module size
function mod.check_aperture(plugin_details, aperture, spec_aperture, zdim)
	local req_aperture = mod.spec_to_req_aperture(spec_aperture, zdim)
	local new = {
		title = _("Aperture"),
		value = mod.get_aperture_str(aperture),
		desc  = string.format("%s: %s", _("Required aperture"),
		                      mod.get_aperture_str(req_aperture))
	}

	if mod.is_aperture_similar(aperture, spec_aperture, zdim) then
		new.status = mod.plugin_status.pass
	else
		new.status = mod.plugin_status.fail
	end
	table.insert(plugin_details, new)
end

local module_strs = {
	xdim = {
		title      = N_("X dimension"),
		range_desc = N_("Allowed X dimension range"),
		range_fmt  = N_("%s to %s"),
		min        = N_("Minimum X dimension"),
		max        = N_("Maximum X dimension"),
	},

	ydim = {
		title      = N_("Y dimension"),
		range_desc = N_("Allowed Y dimension range"),
		range_fmt  = N_("%s to %s"),
		min        = N_("Minimum Y dimension"),
		max        = N_("Maximum Y dimension"),
	},

	ac_warning = N_("Out of range, but within ±2% acceptance criteria"),
}

-- Format a generic string about requirements in microns using the given vargs.
-- Display in imperial if required.
local function notes_fmt_um(fmt, args, units)
	if type(fmt) ~= "string" then
		return ""
	elseif (type(args) ~= "table") or (#args == 0) then
		return fmt
	end

	local dims = {}
	for k, v in ipairs(args) do
		table.insert(dims, mod.dim_str(mod.mk_dim(units, "int", v)))
	end

	return string.format(fmt, table.unpack(dims))
end

-- Check X or Y dimension, whichever is passed, using appropriate strings
local function check_xydim(plugin_details, xydim_val, req_mod_size_um, str_t)
	-- Create param wrapper so we can convert to imperial if necessary
	local xydim = mk_bp("int", xydim_val)
	local min_param, max_param, min_str, max_str
	-- Use the same base units everywhere
	local unit_id = "um"
	-- If adding ±2% to the fail boundaries, also check the warn boundaries
	local use_gs_ac = req_mod_size_um.apply_gs_ac
	local min_val_warn, max_val_warn

	if req_mod_size_um.min then
		min_param = mk_bp("int", req_mod_size_um.min)
		min_str = dim_str_fmt(min_param, unit_id)

		if use_gs_ac then
			min_val_warn = min_param.val
			min_param.val = min_param.val * 0.98
		end
	end

	if req_mod_size_um.max then
		max_param = mk_bp("int", req_mod_size_um.max)
		max_str = dim_str_fmt(max_param, unit_id)

		if use_gs_ac then
			max_val_warn = max_param.val
			max_param.val = max_param.val * 1.02
		end
	end

	local new = {
		title  = _(str_t.title),
		value  = dim_str_fmt(xydim, unit_id),
		status = mod.plugin_status.pass,
	}

	-- Show GS1 caveats (if applicable)
	if type(req_mod_size_um.note_fmt) == "string" and (#req_mod_size_um.note_fmt > 0) then
		new.notes = notes_fmt_um(_(req_mod_size_um.note_fmt),
		                         req_mod_size_um.note_args, "um")
	end

	if min_param and max_param then
		new.desc = string.format("%s: %s", _(str_t.range_desc),
		                         string.format(_(str_t.range_fmt), min_str, max_str))

		if (xydim.val < min_param.val) or (xydim.val > max_param.val) then
			new.status = mod.plugin_status.fail
		elseif use_gs_ac and ((xydim.val < min_val_warn) or
		                      (xydim.val > max_val_warn)) then
			new.warnings = { _(module_strs.ac_warning) }
			new.status = mod.plugin_status.warn
		end

	elseif min_param then
		new.desc = string.format("%s: %s", _(str_t.min), min_str)

		if xydim.val < min_param.val then
			new.status = mod.plugin_status.fail
		elseif use_gs_ac and (xydim.val < min_val_warn) then
			new.warnings = { _(module_strs.ac_warning) }
			new.status = mod.plugin_status.warn
		end

	elseif max_param then
		new.desc = string.format("%s: %s", _(str_t.max), max_str)

		if xydim.val > max_param.val then
			new.status = mod.plugin_status.fail
		elseif use_gs_ac and (xydim.val > max_val_warn) then
			new.warnings = { _(module_strs.ac_warning) }
			new.status = mod.plugin_status.warn
		end
	end

	table.insert(plugin_details, new)
end

function mod.check_mod_size(plugin_details, param_list, req_mod_size_um)
	-- Round measued dim to nearest micron, even if we have more precision
	local xdim, ydim = mod.get_mod_sizes_um(param_list)

	if xdim > 0 then
		check_xydim(plugin_details, xdim, req_mod_size_um, module_strs.xdim)
	end

	if ydim > 0 then
		check_xydim(plugin_details, ydim, req_mod_size_um, module_strs.ydim)
	end
end

function mod.check_qr_qz(plugin_details, param_list)
	local qzs_str

	for k, v in ipairs(param_list) do
		if v.valid and v.id == "qz_size" then
			-- qzs = mod.round_param(v.param)
			qzs_str = mod.dim_str(v)
		end
	end

	if not qzs_str then
		return
	end

	local new = {
		title  = _("Quiet zone"),
		value  = qzs_str,
		desc   = _("Quiet zone must be at least 4x"),
		status = mod.plugin_status.fail,
	}
--[[
	if qzs >= 4 then
		new.status = mod.plugin_status.pass
	end
--]]
	table.insert(plugin_details, new)
end

-- Get the overall status based on all the entries in plugin_details
function mod.get_status(plugin_details)
	local status = mod.plugin_status.pass

	for _, v in pairs(plugin_details) do
		if v.status == mod.plugin_status.fail then
			return v.status
		elseif v.status == mod.plugin_status.warn then
			status = v.status
		end
	end

	return status
end

-- Round away from 0, like C lround().
-- Optional precision in "prec", e.g.:
-- round(0.276, 2) gives 0.28
function mod.round(x, prec)
	prec = prec or 0

	local m = 10 ^ prec
	if x >= 0 then
		return math.floor(x * m + 0.5) / m
	else
		return math.ceil(x * m - 0.5) / m
	end
end

function mod.get_fullsymb_list(list, default_symb)
	local str_list = { "" }  -- Make sure the separator string appears in front
	for k, v in pairs(list) do
		local parent_symb = (type(v) == "string") and v or default_symb
		table.insert(str_list, mod.fullsymb_str(parent_symb, k))
	end

	return table.concat(str_list, "\n • ")
end

function mod.clamp(val, min, max)
	if max and val > max then
		return max
	elseif min and val < min then
		return min
	end

	return val
end

-- "tol" is (nil or) a table, containing optional min, max & scalar members.
-- If empty, no tolerance so val must be within the min & max limits exactly.
-- (De/in)crement min & max by their tol.scalar product (useful on its own).
-- (De/in)crement min & max by at least tol.min (useful on its own).
-- (De/in)crement min & max by no more than tol.max (only useful with scalar).
function mod.is_in_range(val, min, max, tol)
	if not (min or max) then
		return false
	end

	if type(tol) ~= "table" then
		tol = {}
	end

	tol.scalar = tol.scalar or 0

	if min and (val < mod.clamp(min * (1 - tol.scalar),
	                            tol.max and (min - tol.max),
	                            tol.min and (min - tol.min))) then
		return false
	elseif max and (val > mod.clamp(max * (1 + tol.scalar),
	                                tol.min and (max + tol.min),
	                                tol.max and (max + tol.max))) then
		return false
	end

	return true
end

function mod.spec_to_req_aperture(spec_aperture, zdim)
	local req_aperture = spec_aperture
	if req_aperture > 0 and req_aperture < 1 then
		req_aperture = req_aperture * zdim
	end

	return req_aperture
end

-- target between 0 & 1 is a percentage of the module size
function mod.is_aperture_similar(actual, target, zdim)
	if target >= 1 then
		return (actual == target)
	elseif (target <= 0) or not zdim then
		return false
	end

	target = mod.spec_to_req_aperture(target, zdim)

	-- Allow ±20 µm
	return mod.is_in_range(actual, target, target, { min = 20 })
end

-- Get string with aperture and units in metric/imperial (& aperture ref #)
function mod.get_aperture_str(ap_microns)
	local ap_dim = mod.mk_dim("um", "int", ap_microns)
	return string.format("%s (%.2d)", mod.dim_str(ap_dim),
	                     mod.round(ap_microns / 25.4))
end

function mod.filesize(file)
	local current = file:seek()
	local size = file:seek("end")
	file:seek("set", current)
	return size
end

local b64 = { [0] =
	'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P',
	'Q','R','S','T','U','V','W','X','Y','Z','a','b','c','d','e','f',
	'g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v',
	'w','x','y','z','0','1','2','3','4','5','6','7','8','9','+','/',
}

function mod.base64_enc(s)
	local byte, rep = string.byte, string.rep
	local pad = 2 - ((#s - 1) % 3)
	s = (s .. rep('\0', pad)):gsub("...", function(cs)
		local a, b, c = byte(cs, 1, 3)
		return b64[a >> 2] .. b64[(a & 3) << 4 | b >> 4] ..
		       b64[(b & 15) << 2 | c >> 6] .. b64[c & 63]
	end)
	return s:sub(1, #s - pad) .. rep('=', pad)
end

function mod.get_tmpname(basedir, prefix, extension)
	local p = prefix or "acvlua"
	local e = extension or ""
	local c = 0
	local fn = ""

	local b = ""
	if type(basedir) == "string" and #basedir > 0 then
		b = basedir .. package.config:sub(1, 1)
	end

	local f
	repeat
		c = c + 1
		fn = string.format("%s%s%.6d%s", b, p, c, e)
		f = io.open(fn, "r")
		if f then
			f:close()
		else
			break
		end
	until false

	--io.open(fn, "w"):close()
	return fn
end

function mod.is_dpm(codeinfo)
	if codeinfo then
		local vi = codeinfo.verification_info
		local di = codeinfo.decode_info

		return ((vi and vi.dpm_verify_mul and vi.dpm_verify_mul > 0) or
		        (di and di.dpm_decode_mul and di.dpm_decode_mul > 0))
	end

	return false
end

-- Add overall fail status & description to existing results table
function mod.get_results_no_msg(plugin_results)
	plugin_results.status = mod.plugin_status.fail
	plugin_results.output = _("Error: Code message empty")

	return plugin_results
end

function mod.set_results_na(plugin_results, always_show, is_dpm)
	if not always_show then
		plugin_results.status = mod.plugin_status.na
		return
	end

	local fmt
	if is_dpm then
		fmt = _("Error: DPM not permitted for %s")
	else
		fmt = _("Error: Invalid symbology for %s")
	end

	plugin_results.status = mod.plugin_status.fail
	plugin_results.output = string.format(fmt, plugin_results.desc)
end

--------------------------------------------------------------------------------
--[[ Deprecated: used by old hpdf report
function mod.isoparam_str(p)
	local t = {
		grade   = "",
		percent = ""
	}
	if p.grade then
		t.grade = mod.base_param_str(p.grade)
		t.ansi = mod.get_ansi(p.grade.val)
	end
	if p.percent then
		t.percent = mod.base_param_str(p.percent) .. "%"
	end
	if not p.grade and not p.percent then
		t = { grade = "N/a", percent = "N/a" }
	end
	return t
end

function mod.strip_nonprintable(s)
	return string.gsub(s, "[^\33-\126]", mod.get_bytecode)
end

function mod.get_bytecode(s)
	return string.format("<0x%.2x>", string.byte(s))
end

function mod.strip_unicode(s)
	local str = string.gsub(s, "µ", "\xb5")
	return string.gsub(str, "±", "\xb1")
end
--]]
--------------------------------------------------------------------------------

return mod
