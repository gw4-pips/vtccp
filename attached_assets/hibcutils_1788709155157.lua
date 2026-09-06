-- require "strict"
local u = require "utils"
local hs = require "hibc-secondary"
require "gettext-utils"

local M = {}

-- Passthru to utils.lua
M.plugin_status  = u.plugin_status -- { pass = 0, fail = 1, warn = 2, na = 3 }
M.check_grade    = u.check_grade
M.check_mod_size = u.check_mod_size
M.get_status     = u.get_status
M.set_metric     = u.set_metric

--[[
M.is_hibc_linear = {
	code39  = true,
	code128 = true,
	gs1_128 = true,  -- Common error to use FNC1, so run plugin and fail
}
--]]

local is_hibc_2d = {
	--[[ QR: must quote key if it begins with digit ]]--
	["2005"]            = true,
	["2005_eci"]        = true,
	["2005_fnc1_1"]     = true,
	["2005_eci_fnc1_1"] = true,
	["2005_fnc1_2"]     = true,
	["2005_eci_fnc1_2"] = true,
	--[[ DMX ]]--
	ecc_000_140          = true,
	ecc_200              = true,
	ecc_200_fnc1_1_5     = true,
	ecc_200_fnc1_2_6     = true,
	ecc_200_eci          = true,
	ecc_200_eci_fnc1_1_5 = true,
	ecc_200_eci_fnc1_2_6 = true,
}

--[[
local is_2d_subsymb_valid = {
	["2005"] = true,
	ecc_200  = true,
}
--]]

-- Block type
local b_t = {
	empty      = -1,
	unset      =  0,
	primary    =  1,
	secondary  =  2,
	supplement =  3,
}

-- Set the character used to separate parts of the HIBC human-readable
local ELEMENT_SEP = ' '

-- local specials = { '-', '.', ' ', '$', '/', '+', '%' }
-- Only allow uppercase letters [%u], digits [%d], and special chars [-. $/+%]
-- Lua char classes only work with ASCII
local valid_filter = {
	[b_t.unset]      = "[%u%d%-%. %$%/%+%%]",  -- A-Z, 0-9, specials: [-. $/+%]
	[b_t.primary]    = "[%u%d]",               -- A-Z, 0-9, specials: none
	[b_t.secondary]  = "[%u%d%-%.%$%+]",       -- A-Z, 0-9, specials: [-.$+]
	[b_t.supplement] = "[%u%d%-%.]",           -- A-Z, 0-9, specials: [-.]
}

-- Generic field requirements:
-- { valid_str_pattern, min_len, max_len, check_missing }
local field_reqs = {
	lic  = { "^%u[%u%d][%u%d][%u%d]$", 4,  4, true },
	pcn  = { "^[%u%d]+$",              1, 18, true },
	umid = { "^%d$",                   1,  1, true },

	lot = { "^[%u%d%-%.]+$", 0, 18, false },
	sn  = { "^[%u%d%-%.]+$", 0, 18, false },
}

local res_str = {
	encoding = {
		title = N_("Encoding"),
		desc = {
			[b_t.unset]      = N_("Invalid characters in decoded message"),
			[b_t.primary]    = N_("Invalid characters in primary data"),
			[b_t.secondary]  = N_("Invalid characters in secondary data"),
			[b_t.supplement] = N_("Invalid characters in supplemental data"),
		},
		notes = {
			[b_t.unset]      = N_("Only [A-Z], [0-9], '-', '.', ' ' (space), '$', \z
			                       '/', '+' or '%' characters may be encoded."),
			[b_t.primary]    = N_("Only [A-Z] or [0-9] characters may be encoded."),
			[b_t.secondary]  = N_("Only [A-Z], [0-9], '-', '.', '$' or '+' \z
			                       characters may be encoded."),
			[b_t.supplement] = N_("Only [A-Z], [0-9], '-' or '.' characters \z
			                       may be encoded."),
		},
	},

	generic = {
		invalid  = N_("Invalid format"),
		bad_cont = N_("Invalid content"),
		space    = N_("space"),
		short    = N_("Field too short"),
		long     = N_("Field too long"),
		missing  = N_("Field not present"),
		warn_dep = N_("Deprecated format, see notes"),
		notreal  = N_("Not a real date!"),
	},

	--[[ Data structure blocks ]]--
	[b_t.empty] = {
		title = N_("Structure"),
		desc  = N_("Structure content is missing"),
		notes = N_("Empty structures are not permitted. Leading, trailing or \z
		            sequential concatenation characters ('/') may be interpreted \z
		            as the start and/or end of a zero-length structure."),
	},

	[b_t.unset] = {
		title = N_("Unknown Data"),
		desc  = N_("Data structure does not fit primary or secondary format"),
	},

	[b_t.primary] = {
		title = N_("Primary Data"),
		desc  = N_("Overall primary data structure OK"),
		missing_1d = N_("Primary Data not present, cannot fully validate"),
		missing_2d = N_("Primary Data not present (HIBC data should not be \z
		                 split between multiple 2D symbols)"),
		notes = N_("The primary data structure contains an indication of the \z
		            labeler of the item, the item, the packaging level, and \z
		            (if at the end of the symbol) a Check Character."),
		err_note = N_("HIBC spec. 2.6 (Section 3.2):\nWhen using a 2D symbol, \z
		               a single 2D code should be used to carry all Primary, \z
		               secondary and supplemental HIBC data as required."),
	},

	[b_t.secondary] = {
		title = N_("Secondary Data"),
		desc  = N_("Overall secondary data structure OK"),
		notes = N_("Optional secondary data elements are used in conjunction \z
		            with primary data elements, for example to encode expiry \z
		            date and/or Lot/Batch/Serial Number."),
	},

	[b_t.supplement] = {
		title = N_("Supplemental Data"),
		desc  = N_("Overall secondary supplemental data OK"),
		notes = N_("Additional supplemental data can be used when a manufacturer \z
		            wishes to encode both lot number and serial number in the \z
		            same symbol, date of manufacture, expiry date in the \z
		            YYYYMMDD format, and/or quantity."),
		na_desc = N_("Content not checked (full DI validation not yet supported)"),
		warn_1d = N_("Supplemental data used with linear symbology"),
		w_note  = N_("HIBC spec. 2.6 (Section 2.3):\nIt is strongly recommended \z
		              that Additional Supplemental Data be used in the \z
		              concatenated format and with 2D symbologies to reduce the \z
		              risk of creating a linear bar code that may be too long \z
		              for practical use."),
	},

	--[[ Primary data fields ]]--
	flag = {
		title = N_("SL Flag"),
		desc  = N_("HIBC Supplier Labeling Data Identifier Flag Character ('+')"),
		notes = N_("'+' required at the start of a HIBC Supplier Data Structure."),
		err_missing  = N_("Not present at start of code message"),
		err_unwanted = N_("Only use once at start of code message"),
	},

	lic = {
		title = N_("LIC"),
		desc  = N_("Labeler Identification Code (LIC)"),
		notes = N_("Uppercase alphanumeric (A-Z, 0-9), 4 characters, first \z
		            character always alphabetic."),
	},

	pcn = {
		title = N_("PCN"),
		desc  = N_("Labelers Product or Catalog Number (PCN)"),
		notes = N_("Uppercase alphanumeric (A-Z, 0-9) data, 1 to 18 characters."),
	},

	umid = {
		title = N_("U/M"),
		desc  = N_("Unit of Measure Identifier"),
		notes =
			N_("Numeric value only (0-9). 0 is for unit-of-use items. 1 to 8 \z
			    are used to indicate different packaging levels above the unit \z
			    of use. 9 is used for variable quantity containers when manual \z
			    key entry or scan of a secondary will be used to collect \z
			    specific quantity data. The labeler should ensure consistency \z
			    in this field within their packaging process."),
	},

	--[[ Secondary data fields ]]--
	qty_fmt = {
		title = N_("Quantity (format)"),
		desc  = N_("Quantity field format indicator"),
		notes = N_("HIBC spec. 2.6 (Appendix H):\nFrom this point forward labelers \z
		            that wish to include quantity will do so in the supplemental \z
		            data field as indicated in section 2.3.2.4 of this document. \z
		            Existing labels are still valid, but should not be used for \z
		            Unique Device Identification (UDI)."),
		[hs.qf_t.OLD_QTY2] = N_("QQ - 2 digits"),
		[hs.qf_t.OLD_QTY5] = N_("QQQQQ - 5 digits"),
	},

	qty_val = {
		title = N_("Quantity"),
		desc  = N_("Number of units-of-use"),
	},

	exp_fmt = {
		title = N_("Expiry Date (format)"),
		desc  = N_("Date field format indicator"),
		notes = N_("When using MMYY format (0 or 1), this is also the first \z
		            number of the encoded month."),
		-- [0] and [1] both treated as MMYY
		[hs.df_t.MMYY]     = N_("Month and year (MMYY)"),
		[hs.df_t.MMDDYY]   = N_("Month, day and year (MMDDYY)"),
		[hs.df_t.YYMMDD]   = N_("Year, month and day (YYMMDD)"),
		[hs.df_t.YYMMDDHH] = N_("Year, month, day and hour (YYMMDDHH)"),
		[hs.df_t.YYJJJ]    = N_("Julian date (YYJJJ)"),
		[hs.df_t.YYJJJHH]  = N_("Julian date and hour (YYJJJHH)"),
		[hs.df_t.NO_DATE]  = N_("No date encoded"),
	},

	exp_val = {
		title = N_("Expiry Date"),
		desc  = N_("Decoded date"),
		err_past = N_("Using passed date"),
	},

	lot = {
		title = N_("Lot"),
		desc  = N_("Lot or batch number, uppercase alphanumeric \z
		            (A-Z, 0-9, '-', '.'), 0 to 18 characters"),
		w_desc = N_("Lot or batch number, uppercase alphanumeric \z
		             (A-Z, 0-9, '-', '.'), 0 to 13 characters"),
		w_note = N_("HIBC spec. 2.6 (Appendix F, Note 1):\nUsers who wish to \z
		            encode a five-digit Julian date followed by a lot/batch \z
		            field should use the current format of the secondary data \z
		            field \"+$$5\"."),
	},

	sn = {
		title = N_("Serial Number"),
		desc  = N_("Serial number, uppercase alphanumeric \z
		           (A-Z, 0-9, '-', '.'), 0 to 18 characters"),
		sup_desc = N_("Serial number, uppercase alphanumeric \z
		               (A-Z, 0-9, '-', '.'), 1 to 18 characters"),
	},

	unknown_sec = {
		title = N_("Unknown Field"),
		desc  = N_("Data does not fit any secondary field format"),
	},

	--[[ Supplemental fields ]]--
	prod_date = {
		title = N_("Production Date"),
		desc  = N_("Decoded date"),
	},

	unknown_di = {
		title = N_("Unknown DI"),
		desc  = N_("Data does not fit any suitable DI format"),
		notes = N_("Comprehensive DI validation is not yet supported. Some \z
		            commonly used DIs may appear as \"Unknown DI\"."),
	},

	--[[ Check fields ]]--
	check = {
		title = N_("Check Character"),
		desc  = N_("Modulo 43 check character"),
		exp   = N_("Modulo 43 check character - expected: '%s'"),
		notes = N_("Any one of [A-Z], [0-9], '-', '.', ' ' (space), '$', '/', \z
		            '+' or '%'. The human-readable interpretation shall use \z
		            an underscore ('_') to represent the space character."),
	},

	link = {
		title = N_("Link Character"),
		desc  = N_("Value should match the modulo 43 check character from the \z
		            primary data symbol"),
		-- For notes, reuse check.notes
	},
}

res_str.link.notes = res_str.check.notes
res_str.exp_fmt['0'] = res_str.exp_fmt[hs.df_t.MMYY]


-- Convenience function to see if symbol is valid 1D/2D HIBC
--[[
function M.is_hibc(subsymb)
	return M.is_hibc_linear[subsymb] or is_hibc_2d[subsymb]
end
--]]

local function get_block_details(block_t, str)
	local block_details = {
		title = _(res_str[block_t].title),
		value = str,
		desc  = _(res_str[block_t].desc),
		notes = _(res_str[block_t].notes),
	}

	if block_t <= b_t.unset then
		block_details.status = M.plugin_status.fail
	-- Until supplemental is properly parsed, default to na
	elseif block_t == b_t.supplement then
		block_details.status = M.plugin_status.na
	else
		block_details.status = M.plugin_status.pass
	end

	return block_details
end

-- TODO: Do more of the pattern-matching boilerplate inside this function once,
-- rather than at every place it is called
local function get_encoding_err(block_t, str)
	local encoding = {
		title  = _(res_str.encoding.title),
		value  = u.parse_msg(str),
		desc   = _(res_str.encoding.desc[block_t]),
		notes  = _(res_str.encoding.notes[block_t]),
		status = M.plugin_status.fail,
	}

	return encoding
end

local function get_field_details(field, str)
	local field = {
		title    = _(res_str[field].title),
		value    = str,
		desc     = _(res_str[field].desc),
		notes    = _(res_str[field].notes),
		status   = M.plugin_status.pass,  -- default
		errors   = {},
		warnings = {},
	}

	return field
end

-- Convenience function to perform simple length and pattern checks
-- For special cases, pass along modified min/max/check_empty in override
local function check_field(field, str, override)
	if not field_reqs[field] then
		return nil
	end

	local new = get_field_details(field, u.parse_msg(str))
	local ptrn, min, max, check_empty = table.unpack(field_reqs[field])
	str = str or ""

	if override then
		min         = override.min or min
		max         = override.max or max
		check_empty = override.check_empty
	end

	if check_empty and (#str == 0) then
		table.insert(new.errors, _(res_str.generic.missing))

	else
		if str:match(ptrn) ~= str then
			table.insert(new.errors, _(res_str.generic.invalid))
		end

		if #str < min then
			table.insert(new.errors, _(res_str.generic.short))

		elseif #str > max then
			table.insert(new.errors, _(res_str.generic.long))
		end
	end

	if #new.errors > 0 then
		new.status = M.plugin_status.fail
	end

	return new
end

-- Filter non-base43 chars without losing function characters
-- If no filter is passed, use the most lenient charset
local function filter_invalid_chars(msg, filter)
	filter = filter or valid_filter[b_t.unset]
	local invalid_msg = { "" }  -- "string builder"
	local got_bslash = false

	for c in msg:gmatch('.') do
		if got_bslash then
			if c:find("[1234%?\\]") then
				table.insert(invalid_msg, '\\' .. c)

			elseif not c:find(filter) then
				table.insert(invalid_msg, c)
			end

			got_bslash = false

		elseif c == '\\' then
			got_bslash = true

		elseif not c:find(filter) then
			table.insert(invalid_msg, c)
		end
	end

	return table.concat(invalid_msg)
end

-- Append varargs to a string builder table, but only if not nil or empty strings
local function add_str(str_tab, ...)
	local args = table.pack(...)

	for i = 1, args.n do
		if not args[i] then
			-- Do nothing
		elseif type(args[i]) ~= "string" then
			table.insert(str_tab, tostring(args[i]))

		elseif #args[i] > 0 then
			table.insert(str_tab, args[i])
		end
	end
end

-- Use first char (or second if first is '+') to distinguish primary/secondary
-- Set overall "has_(primary|secondary)" flag if found
local function detect_block_type(flags, pos, char)
	if (not char) or (#char == 0) then
		return b_t.empty

	-- Only check for primary if at start
	elseif (pos == 1) and (char:match("[A-Z]") == char) then
		flags.has_primary = true
		return b_t.primary

	-- Only check for secondary if it doesn't already exist
	elseif (not flags.has_secondary) and (char:match("[0-9%$]") == char) then
		flags.has_secondary = true
		return b_t.secondary

	-- Only check for DI if main secondary exists
	elseif flags.has_secondary and (char:match("[A-Z0-9]") == char) then
		-- flags.n_supplement = flags.n_supplement + 1
		return b_t.supplement
	end

	return b_t.unset
end

local function get_block_status(plugin_details, block_start)
	local status = M.plugin_status.pass
	block_start = block_start or 1  -- Start from 1 (or block_start if given)

	for k = block_start, #plugin_details do
		local st = plugin_details[k].status

		if st == M.plugin_status.fail then
			return st

		elseif st == M.plugin_status.warn then
			status = st
		end
	end

	return status
end

-- Append results to plugin_details, return (escaped) human-readable
local function check_primary(plugin_details, flags, s)
	local block_start = #plugin_details + 1
	local valid_chars = valid_filter[b_t.primary]

	if s:match(valid_chars .. '+') ~= s then
		local err = get_encoding_err(b_t.primary, s:gsub(valid_chars, ''))
		table.insert(plugin_details, err)
	end

	local lic, pcn, umid
	lic = s:sub(1, 4)

	-- 6 or more chars, both PCN and UMID can be interpreted
	if #s >= 6 then
		pcn = s:sub(5, -2)
		umid = s:sub(-1)

	-- Only 4 or 5 chars, cannot interpret UMID, possibly can read PCN
	else
		pcn = s:sub(5) or ""
		umid = ""
	end

	local new, hr

	-- Check LIC
	new = check_field("lic", lic)
	hr = { new.value }
	table.insert(plugin_details, new)

	-- Check PCN
	new = check_field("pcn", pcn)
	add_str(hr, new.value)
	table.insert(plugin_details, new)

	-- Check U/M
	new = check_field("umid", umid)
	add_str(hr, new.value)
	table.insert(plugin_details, new)

	-- Create block summary
	local hr_str = table.concat(hr, ELEMENT_SEP)
	new = get_block_details(b_t.primary, hr_str)
	new.status = get_block_status(plugin_details, block_start)

	if new.status == M.plugin_status.fail then
		new.desc = _(res_str.generic.bad_cont)
	end

	table.insert(plugin_details, block_start, new)
	return hr_str
end

-- Input size as either a number (fixed-length), or a table with .min and .max
local function get_length_notes(size)
	local min, max

	if type(size) == "number" then
		min, max = size, size
	elseif type(size) == "table" and size.min and size.max then
		min, max = size.min, size.max
	else
		return nil
	end

	if min == max then
		return string.format(_("Fixed length, numeric, %d digits."), max)
	else
		return string.format(_("Variable length, numeric, %d to %d digits."), min, max)
	end
end

-- Return new result and remaining secondary data
local function check_quantity(q_fmt, s)
	local new, qty_info

	-- Get qty_info, and rest of string (minus qty flag and value)
	qty_info, s = hs.get_qty_info(q_fmt, s)

	-- TODO: display without leading zeros?
	new = get_field_details("qty_val", u.parse_msg(qty_info.encode))
	new.notes = get_length_notes(hs.fmt_len[q_fmt])

	if not qty_info.value then
		new.status = M.plugin_status.fail
		table.insert(new.errors, _(res_str.generic.invalid))

		if qty_info.short then
			table.insert(new.errors, _(res_str.generic.short))

		elseif qty_info.long then
			table.insert(new.errors, _(res_str.generic.long))
		end
	end

	return new, s
end

-- Return new result and remaining secondary data
local function check_date(d_fmt, s, decode_time, check_exp)
	local new, date_info, esc_str, date_str

	-- Get date_info, and rest of string (minus date flag and value)
	date_info, s = hs.get_date_info(d_fmt, s, decode_time, check_exp)
	esc_str = u.parse_msg(date_info.encode)
	local len_str = get_length_notes(date_info.length)

	if check_exp then
		new = get_field_details("exp_val", esc_str)
	else
		new = get_field_details("prod_date", esc_str)
	end

	if (d_fmt == hs.df_t.YYJJJHH) or (d_fmt == hs.df_t.YYMMDDHH) then
		new.notes = string.format("%s\n%s", len_str,
			_("HIBC expiry dates with times are encoded for UTC+00:00."))
	else
		new.notes = len_str
	end

	if date_info.valid then
		date_str = date_info.hr

		if date_info.exp then
			new.status = M.plugin_status.fail
			table.insert(new.errors, _(res_str.exp_val.err_past))
		end

	elseif date_info.notreal then
		date_str = _(res_str.generic.notreal)
		new.status = M.plugin_status.fail
	else
		date_str = _(res_str.generic.invalid)
		new.status = M.plugin_status.fail

		if date_info.short then
			table.insert(new.errors, _(res_str.generic.short))

		elseif date_info.long then
			table.insert(new.errors, _(res_str.generic.long))
		end
	end

	new.desc = string.format("%s: %s", new.desc, date_str)
	return new, s
end

-- Append results to plugin_details, return (escaped) human-readable
local function check_secondary(plugin_details, flags, s)
	local block_start = #plugin_details + 1
	local valid_chars = valid_filter[b_t.secondary]

	if s:match(valid_chars .. '+') ~= s then
		local err = get_encoding_err(b_t.secondary, s:gsub(valid_chars, ''))
		table.insert(plugin_details, err)
	end

	local hr = {}  -- hr string builder table
	local ref_id       -- $, $+, $$, $$+ or nil
	local date_fmt     -- If date format flag exists, store it
	local has_yyjjj    -- Else if +YYJJJ date format is used, set to true
	local has_sn       -- Trailing field is sn if true, lot if false/nil
	local has_unknown  -- If format is broken, display an error

	local secondary_fmts = {
		julian   = "^(%d.*)$",               -- "[0-9]"
		lot      = "^(%$)([%u%d%-%.].*)$",   -- "$[A-Z0-9-.]"
		d_lot    = "^(%$%$)([0-7])(.*)$",    -- "$$[0-7]"
		qt_d_lot = "^(%$%$)([89])(.*)$",     -- "$$[89]"
		sn       = "^(%$%+)([%u%d%-%.].*)$", -- "$+[A-Z0-9-.]"
		d_sn     = "^(%$%$%+)([0-7])(.*)$",  -- "$$+[0-7]"
	}

	has_unknown = true  -- Unset when a complete match is found

	for k, v in pairs(secondary_fmts) do
		local caps = table.pack(s:match(v))

		if caps[1] then  -- match successful
			if k == "qt_d_lot" then
				local q_fmt, new
				ref_id, q_fmt, s = table.unpack(caps)

				-- Display quantity format flag
				new = get_field_details("qty_fmt", q_fmt)
				new.desc = string.format("%s: %s", new.desc, _(res_str.qty_fmt[q_fmt]))
				table.insert(new.warnings, _(res_str.generic.warn_dep))
				new.status = M.plugin_status.warn
				table.insert(plugin_details, new)

				new, s = check_quantity(q_fmt, s)
				table.insert(plugin_details, new)

				add_str(hr, ref_id, q_fmt, new.value)

				-- If more text after quantity, may contain date
				if #s > 0 then
					date_fmt = s:match("^[0-7]")

					if date_fmt then
						has_unknown = false
						s = s:sub(2)  -- Trim the date_fmt digit
					end
				else
					has_unknown = false
				end
			else
				if k == "julian" then
					-- TODO: warn deprecated if using +YYJJJ, even without lot?
					has_yyjjj = true
				elseif (k == "lot") or (k == "sn") then
					-- TODO: is deprecated? Spec is not very clear on "$" or "$+"
					ref_id, s = table.unpack(caps)
				else
					ref_id, date_fmt, s = table.unpack(caps)
				end

				if (k == "sn") or (k == "d_sn") then
					has_sn = true
				end

				has_unknown = false
				add_str(hr, ref_id)
			end

			break  -- exit for loop after first match
		end
	end

	if date_fmt or has_yyjjj then
		-- TODO: display something for missing fmt char (implicit yyjjj fmt)
		local new = get_field_details("exp_fmt", date_fmt)
		new.desc = string.format("%s: %s", new.desc,
		                         _(res_str.exp_fmt[date_fmt or hs.df_t.YYJJJ]))
		table.insert(plugin_details, new)

		if has_yyjjj or (date_fmt ~= hs.df_t.NO_DATE) then
			local check_exp = true
			new, s = check_date(date_fmt, s, flags.time, check_exp)
			table.insert(plugin_details, new)
		else
			new = nil
		end

		add_str(hr, date_fmt, (new and new.value))
	end

	local esc_str = u.parse_msg(s)  -- escape remaining lot/sn/garbage content
	add_str(hr, esc_str)

	if has_unknown then
		local new = get_field_details("unknown_sec", esc_str)
		new.status = M.plugin_status.fail
		table.insert(plugin_details, new)

	elseif #s > 0 then  -- Lot or sn at the end
		local field, new, override

		if has_sn then
			field = "sn"
		else
			field = "lot"
			override = { min = 0, max = (has_yyjjj and 13) }
		end

		new = check_field(field, s, override)

		if has_yyjjj then
			new.desc  = _(res_str[field].w_desc)
			new.notes = _(res_str[field].w_note)
			table.insert(new.warnings, _(res_str.generic.warn_dep))

			if new.status ~= M.plugin_status.fail then
				new.status = M.plugin_status.warn
			end
		end

		table.insert(plugin_details, new)
	end

	-- Create block summary
	local hr_str = table.concat(hr, ELEMENT_SEP)
	local new = get_block_details(b_t.secondary, hr_str)
	new.status = get_block_status(plugin_details, block_start)

	if new.status == M.plugin_status.fail then
		new.desc = _(res_str.generic.bad_cont)
	end

	table.insert(plugin_details, block_start, new)
	return hr_str
end

-- Append results to plugin_details, return (escaped) human-readable
local function check_supplemental(plugin_details, flags, s)
	local block_start = #plugin_details + 1
	local valid_chars = valid_filter[b_t.supplement]

	if s:match(valid_chars .. '+') ~= s then
		local err = get_encoding_err(b_t.supplement, s:gsub(valid_chars, ''))
		table.insert(plugin_details, err)
	end

	local subcat, category, content = s:match("^(%d?%d?%d?)(%u)(.*)$")
	local di = subcat .. category
	local hr = {}  -- hr string builder table
	local has_unknown  -- Keep track of unknown (but possibly valid) DIs

	if #category == 0 then  -- Cannot possibly be a valid DI, so fail
		local hr_str = u.parse_msg(s)
		table.insert(plugin_details, block_start, get_block_details(b_t.unset, hr_str))
		return hr_str

	else
		local subcat_val = tonumber(subcat) or 0
		local new

		if (category == 'D') and (subcat_val == 14) then
			local check_exp = true
			new = check_date(hs.df_t.SUP_DATE, content, flags.time, check_exp)
			new.title = string.format("DI: %s (%s)", di, new.title)

		elseif (category == 'D') and (subcat_val == 16) then
			new = check_date(hs.df_t.SUP_DATE, content, flags.time)
			new.title = string.format("DI: %s (%s)", di, new.title)

		elseif (category == 'Q') and (subcat_val == 0) then
			new = check_quantity(hs.qf_t.SUP_QTY, content)
			new.title = string.format("DI: %s (%s)", di, new.title)

		elseif (category == 'S') and (subcat_val == 0) then
			local override = { min = 1 }
			new = check_field("sn", content, override)
			new.title = string.format("DI: %s (%s)", di, new.title)
			new.desc  = _(res_str.sn.sup_desc)
		else
			new = get_field_details("unknown_di", u.parse_msg(content))
			new.title  = string.format("%s: %s", new.title, di)
			new.status = M.plugin_status.na
			has_unknown = true
		end

		add_str(hr, di, new.value)
		table.insert(plugin_details, new)
	end

	-- Create block summary
	local hr_str = table.concat(hr, ELEMENT_SEP)
	local new = get_block_details(b_t.supplement, hr_str)
	new.status = get_block_status(plugin_details, block_start)

	-- Warn if using linear symbol with supplemental data
	if not flags.is_2d then
		if new.status ~= M.plugin_status.fail then
			new.status = M.plugin_status.warn
		end

		new.notes    = _(res_str[b_t.supplement].w_note)
		new.warnings = { _(res_str[b_t.supplement].warn_1d) }
	end

	if new.status == M.plugin_status.fail then
		new.desc = _(res_str.generic.bad_cont)

	-- For now, unknown supplementals should never actually pass as they are not
	-- fully checked. Of course, they can still outright fail or produce warnings.
	elseif has_unknown and (new.status == M.plugin_status.pass) then
		new.desc   = _(res_str[b_t.supplement].na_desc)  -- Temporary disclaimer
		new.status = M.plugin_status.na
	end

	table.insert(plugin_details, block_start, new)
	return hr_str
end

-- Append results to plugin_details, return (escaped) human-readable
local function check_block_content(plugin_details, flags, pos, s)
	local hr = {}  -- hr string builder table

	-- Check for '+' flag at start of string
	local first = s:sub(1, 1)

	if first == '+' then
		local new = get_field_details("flag", first)
		flags[pos].has_di_flag = true

		-- If not the first block, '+' should never be the first character
		if pos > 1 then
			new.status = M.plugin_status.fail
			table.insert(new.errors, _(res_str.flag.err_unwanted))
		end

		table.insert(plugin_details, new)
		add_str(hr, first)

		s = s:sub(2)         -- Shift '+' out of primary/secondary block
		first = s:sub(1, 1)  -- Update first

	elseif pos == 1 then  -- '+' is missing
		local new = get_field_details("flag")
		new.status = M.plugin_status.fail
		table.insert(new.errors, _(res_str.flag.err_missing))

		table.insert(plugin_details, 1, new)  -- Always insert at top
	end

--	First non '+' char is used to determine block type.
--
--	If last block, empty and primary data exists: glue the "link" char back on
--	before detecting type, and clear the link flag.
--
--	Otherwise: check the block type, and then glue the link char back on if
--	applicable (last block, and primary data exists).
	if (#s == 0) and (pos == #flags) and flags.has_primary then
		s = flags.link_char
		flags.link_char = nil
		flags[pos].type = detect_block_type(flags, pos, s)
	else
		flags[pos].type = detect_block_type(flags, pos, first)

		if (pos == #flags) and flags.has_primary then
			s = s .. flags.link_char
			flags.link_char = nil
		end
	end

	local block_str

	if flags[pos].type == b_t.primary then
		block_str = check_primary(plugin_details, flags, s)

	elseif flags[pos].type == b_t.secondary then
		block_str = check_secondary(plugin_details, flags, s)

	elseif flags[pos].type == b_t.supplement then
		block_str = check_supplemental(plugin_details, flags, s)

	elseif flags[pos].type == b_t.empty then
		block_str = ""
		table.insert(plugin_details, get_block_details(b_t.empty))

	else  -- if flags[pos].type == b_t.unset then
		block_str = u.parse_msg(s)
		table.insert(plugin_details, get_block_details(b_t.unset, block_str))
	end

	add_str(hr, block_str)

	-- Warn/fail if primary is missing
	if pos == 1 and not flags.has_primary then
		local new = get_block_details(b_t.primary)

		if flags.is_2d then
			new.desc   = _(res_str[b_t.primary].missing_2d)
			new.notes  = _(res_str[b_t.primary].err_note)
			new.status = M.plugin_status.fail
		else
			new.desc   = _(res_str[b_t.primary].missing_1d)
			new.status = M.plugin_status.warn
		end

		table.insert(plugin_details, 1, new)  -- Insert at top
	end

	return table.concat(hr, ELEMENT_SEP)
end

-- Mostly just so that we don't need to repeat the pattern matching
local function replace_spaces(char)
	return (char:gsub("[ ]", '_'))
end

-- Convert spaces to underscores for human-readable and displayed value
local function escape_check_chars(msg)
	local cdx, cdr = u.mod43_check(msg)

	return replace_spaces(cdx), replace_spaces(cdr)
end

-- In plugin results, qualify underscores as "_ (space)" so their meaning is clear
local function label_spaces(char, fmt)
	local str = fmt or "%s"

	if char == '_' then
		return string.format("%s (%s)", str:format(char), _(res_str.generic.space))
	end

	return str:format(char)
end

-- Run the data content check, return plugin_details{} & human readable
function M.get_core_results(msg, subsymb, decode_time)
	local plugin_details = {}
	local blocks, hr = {}, {}
	local flags = {
		time          = decode_time or os.time(),
		is_2d         = is_hibc_2d[subsymb],
		has_primary   = false,
		has_secondary = false,
		-- n_supplement  = 0,  -- How many supplemental data blocks
		-- link_char: penultimate char i.e. msg:sub(-2, -2)
	}

	-- First check for empty msg, if so stop parsing
	if (not msg) or (#msg == 0) then
		table.insert(plugin_details, get_block_details(b_t.empty))
		return plugin_details

	-- Check for invalid characters, stop parsing if any are found
	elseif msg:match(valid_filter[b_t.unset] .. '+') ~= msg then
		local err = get_encoding_err(b_t.unset, filter_invalid_chars(msg))

		table.insert(plugin_details, err)
		return plugin_details, u.parse_msg(msg)

	-- Check that msg is at least as long as the smallest valid HIBC code "+$LC"
	elseif #msg < 4 then
		local esc_msg = u.parse_msg(msg)
		table.insert(plugin_details, get_block_details(b_t.unset, esc_msg))
		return plugin_details, esc_msg
	end

	-- Store expected and read check characters (also possible link char)
	local check_exp, check_read = escape_check_chars(msg)
	flags.link_char = msg:sub(-2, -2)
	-- Cut check and (possible) link characters from string
	msg = msg:sub(1, -3)

	-- Add a '/' onto msg so we can use '/' as the anchor for end of blocks
	for block in string.gmatch(msg .. '/', "([^/]*)/") do
		table.insert(blocks, block)

		local new = {
			has_di_flag = false,
			type        = b_t.unset,
		}

		table.insert(flags, new)
	end

	-- Parse each '/' delimited block, replace raw data with hr representation
	for k, v in ipairs(blocks) do
		-- Function params: PD table, flags, block index, raw block str
		blocks[k] = check_block_content(plugin_details, flags, k, v)
	end

	-- Replace slashes between blocks, with extra separators
	add_str(hr, table.concat(blocks, ELEMENT_SEP .. '/' .. ELEMENT_SEP))

	-- If link_char is still set, display it (replace ' ' with '_' if necessary)
	if flags.link_char then
		local link = replace_spaces(flags.link_char)

		local new = get_field_details("link", label_spaces(link))
		new.status = M.plugin_status.na  -- Cannot be validated, so no status
		table.insert(plugin_details, new)

		add_str(hr, link)
	end

	-- Validate check character
	local new = get_field_details("check", label_spaces(check_read))
	local overall_st = u.get_status(plugin_details)

	if overall_st == M.plugin_status.fail then
		-- Plugin has already failed, check character can't really be validated
		new.status = M.plugin_status.na

	elseif check_exp ~= check_read then
		new.desc   = label_spaces(check_exp, _(res_str.check.exp))
		new.status = M.plugin_status.fail
	end

	table.insert(plugin_details, new)
	-- Restore check character for human-readable
	add_str(hr, check_read)

	return plugin_details, table.concat(hr, ELEMENT_SEP)
end

function M.check_min_height(plugin_details, param_list, req_height)
	local aspect_val
	local req_str = req_height.min .. '%'

	do
		local height_val, width_val

		for k, v in ipairs(param_list) do
			if v.valid then
				if v.id == "height" and v.unit == "mm" then
					height_val = v.param.val
				elseif v.id == "width" and v.unit == "mm" then
					width_val = v.param.val
				end
			end
		end

		if not (height_val and width_val) then
			return
		end

		aspect_val = u.round(height_val / width_val * 100)
	end

	local new = {
		title  = _("Bar Height"),
		value  = string.format("%d%%", aspect_val),
		desc   = string.format(_("Height of bars should be at least %s of symbol width"),
		                       req_str),
		notes  = _("Measured symbol width includes all quiet zones"),
		status = M.plugin_status.pass,
	}

	-- If given error status is anything other than warn, default to fail
	-- TODO: Spec uses "should" instead of "must". Default to warn?
	if req_height.err ~= M.plugin_status.warn then
		req_height.err = M.plugin_status.fail
	end

	-- If height (as X% of symbol width) is < required bar height %
	if aspect_val < req_height.min then
		new.status = req_height.err
	end

	table.insert(plugin_details, new)
end

function M.check_wn_ratio(plugin_details, param_list, req_wnr)
	local wnr = -1

	for k, v in ipairs(param_list) do
		if v.valid and v.id == "ratio" and v.param.type then
			wnr = u.round_param(v.param)
		end
	end

	if wnr > 0 then
		local new = {
			title  = _("Wide:Narrow Ratio"),
			value  = string.format("%s:1", wnr),
			desc   = string.format(_("Wide:Narrow ratio should be %s"), req_wnr.str),
			status = M.plugin_status.pass,
		}

		if (wnr < req_wnr.min) or (wnr >= req_wnr.max) then
			new.status = M.plugin_status.fail
		elseif (wnr < req_wnr.target_min) or (wnr >= req_wnr.target_max) then
			new.status = M.plugin_status.warn
		end

		table.insert(plugin_details, new)
	end
end

--[[
function M.check_2d_symb(plugin_details, symb, subsymb)
	if not is_2d_subsymb_valid[subsymb] then
		local new = {
			title  = _("Symbology"),
			value  = u.fullsymb_str(symb, subsymb),
			status = M.plugin_status.fail,
		}

		if subsymb == "ecc_000_140" then
			new.desc = _("Invalid symbology version, only ECC 200 is permitted")
		else
			new.desc  = _("Symbol uses implied FNC1 character or ECI mode")
			new.notes = _("All 2D GS1 symbols will include at least one \z
			               implied FNC1 character.")
		end

		table.insert(plugin_details, new)
	end
end
--]]

return M
