--require "strict"
--dbg = require "dbg"
local u = require "utils"
require "gettext-utils"

local mod = {}

-- mod.status_hr = { [0] = _("Pass"), [1] = _("Fail"), [2] = _("Warning") }

-- Passthru to utils.lua
mod.plugin_status    = u.plugin_status -- { pass = 0, fail = 1, warn = 2, na = 3 }
mod.check_grade      = u.check_grade
mod.check_aperture   = u.check_aperture
mod.check_mod_size   = u.check_mod_size
mod.get_status       = u.get_status
mod.get_mod_sizes_um = u.get_mod_sizes_um
mod.get_min_mod_size = u.get_min_mod_size
mod.set_metric       = u.set_metric

-- TODO: rename.
-- Really this is a list of sub-symbologies which contain GS1 AIs.
mod.ai_checked_with_fnc1 = {
	gs1_128          = true,
	databar_expanded = true,
	ecc_200_fnc1_1_5 = true,
	["2005_fnc1_1"]  = true,
	cca              = true,
	ccb              = true,
	ccc              = true,
}

-- List of sub-symbologies which *MIGHT* have digital link data.
mod.digital_link_symbs = {
	ecc_200  = true,
	["2005"] = true,
}

-- Map AICheck status keys to generic plugin status
local ai_status = {
	ok      = mod.plugin_status.pass,
	error   = mod.plugin_status.fail,
	warning = mod.plugin_status.warn,
	na      = mod.plugin_status.na,
}

-- Run the aicheck, return the aicheck output table.
-- TODO: cache results in a weak table?
function mod.aicheck(msg, subsymb, opt)
	opt = opt or {}
	local r
	-- Strip off the leading "\1"
	if (subsymb == "gs1_128") or (subsymb == "ecc_200_fnc1_1_5") then
		msg = msg:sub(3)
	end

	if (subsymb ~= "gs1_128") and (subsymb ~= "databar_expanded") then
		opt.ignore_order = true

		if (subsymb == "ecc_200_fnc1_1_5") or (subsymb == "2005_fnc1_1") then
			opt.ignore_gs_for_f1 = true
		elseif (subsymb == "ecc_200") or (subsymb == "2005") then
			opt.forced_separator = true
		end
	end

	r = acv_aicheck(msg, opt)

	--dbg.table_print(r)
	for k, v in ipairs(r.ais) do
		if opt.forced_separator then
			v.data = (v.data:gsub("\\1", ''))
		end

		v.data = u.parse_msg(v.data)

		if opt.healthcare and (v.ai == "11" or v.ai == "17") and
		   #v.data == 6 and v.data:sub(5) == "00" then
			v.warnings = v.warnings or {}
			table.insert(v.warnings, _("Deprecated format, see notes"))
			-- TODO: error instead of warning?
			--table.insert(v.errors, _("Invalid format"))

			if v.status == "ok" then
				v.status = "warning"
			end
		end
	end

	if opt.forced_separator then
		r.human_readable = (r.human_readable:gsub("\\1", ''))
	end
	r.human_readable = u.parse_msg(r.human_readable)

	return r
end

-- Only one key is permitted in the path
local primary_ais = {
	"00", "01",
	"253", "255",
	"401", "402", "414", "415", "417",
	"8003", "8004", "8006", "8010", "8013", "8017", "8018",
}

-- TPX exceptions for 01 & 414 are handled in code, see: q_list
local qualifier_ais = {
	["01"  ] = { "22", "10", "21" },
--	["01"  ] = { "235" },
	["414" ] = { "254" },
--	["414" ] = { "7040" },
	["415" ] = { "8020" },
	["417" ] = { "7040" },
	["8004"] = { "7040" },
	["8006"] = { "22", "10", "21" },
	["8010"] = { "8011" },
	["8017"] = { "8019" },
	["8018"] = { "8019" },
}

-- How long is the bit in brackets, based on 2 leading digits
local ai_length = {
	["00"] = 2, ["01"] = 2, ["02"] = 2,
--	["03"] = 2, ["04"] = 2,
	["10"] = 2, ["11"] = 2, ["12"] = 2, ["13"] = 2,
--	["14"] = 2,
	["15"] = 2, ["16"] = 2, ["17"] = 2,
--	["18"] = 2, ["19"] = 2,
	["20"] = 2, ["21"] = 2, ["22"] = 2, ["23"] = 3, ["24"] = 3, ["25"] = 3,
	["30"] = 2, ["31"] = 4, ["32"] = 4, ["33"] = 4, ["34"] = 4, ["35"] = 4,
	["36"] = 4, ["37"] = 2, ["39"] = 4,
	["40"] = 3, ["41"] = 3, ["42"] = 3, ["43"] = 4,
	["70"] = 4, ["71"] = 3, ["72"] = 4,
	["80"] = 4, ["81"] = 4, ["82"] = 4,
	["90"] = 2, ["91"] = 2, ["92"] = 2, ["93"] = 2, ["94"] = 2,
	["95"] = 2, ["96"] = 2, ["97"] = 2, ["98"] = 2, ["99"] = 2,
}

local function is_ai_plausible(ai)
	if (type(ai) ~= "string") or (#ai < 2) then
		return false
	end

	return (#ai == ai_length[ai:sub(1, 2)])
end

function mod.split_uri(msg)
	-- No part can be empty.  No case-insensitive option, thanks Lua.
	local m1, m2, rem = msg:match("^([hH][tT][tT][pP][sS]?://)([^/?#]+)(.+)$")

	if not rem then
		return nil
	end

	-- Prefix is everything before the domain
	local r = { "prefix", "domain", prefix = m1, domain = m2,
	            ais = {}, reserved = {} }

	-- Any/all parts can be empty
	local path, query, fragment = rem:match("^/?([^?#]*)%??([^#]*)#?(.*)$")

	if #path > 0 then
		table.insert(r, "path")
		r.path = path

		local pieces = {}
		-- Add a '/' onto path so we can use '/' as the anchor for end
		for s in string.gmatch(path .. '/', "([^/]*)/") do
			table.insert(pieces, s)
		end

		local i = 1
		while i < #pieces do
			local ai   = pieces[i    ]:match("^(%d%d%d?%d?)$")
			local data = pieces[i + 1]:match("^(.+)$")

			if is_ai_plausible(ai) and data then
				-- 3rd member is to say this came from the path, not query.
				table.insert(r.ais, { ai, data, true })
				i = i + 2
			elseif #r.ais > 0 then
				break
			else
				i = i + 1
			end
		end

		-- No more AIs, so the path should have already ended.
		-- TODO: allow override, to always show at least one "unknown AI"?
		if #r.ais > 0 and i <= #pieces then
			-- We'll allow a single trailing slash (final empty piece)
			if #pieces[i] > 0 or i < #pieces then
				-- Show an error about trailing path
				-- Include the leading '/' to ensure aicheck will fail
				pieces[i] = '/' .. pieces[i]
				table.insert(r.ais, { "", table.concat(pieces, '/', i), true })
			end
		end
	end

	-- TODO: only parse query if path already has valid AIs?
	if #query > 0 then
		table.insert(r, "query")
		r.query = query

		local pieces = {}
		-- Add a '&' onto path so we can use '&' as the anchor for end
		for s in string.gmatch(query .. '&', "([^&;]*)[&;]") do
			table.insert(pieces, s)
		end

		for i, v in ipairs(pieces) do
			local ai, data = v:match("^(%d%d%d?%d?)=(.*)$")

			if is_ai_plausible(ai) then
				table.insert(r.ais, { ai, data })
			else
				ai, data = v:match("^(%d+)=(.*)$")
				if ai then
					table.insert(r.reserved, { ai, data })
				end
			end
		end
	end

	if #fragment > 0 then
		table.insert(r, "fragment")
		r.fragment = fragment
	end

	return r
end

local function decode_percents(str)
	if type(str) ~= "string" then
		return nil
	end

	local unescaped = {}
	local bad = {}

	local i = 1
	while i <= #str do
		local ch, xx = str:match("^(.)(%x?%x?)", i)

		if ch == '%' then
			if #xx == 2 then
				ch = string.char(tonumber(xx, 16))
				i = i + 2
			else
				-- Eat the percent, spit out a badchar
				ch = "\\?"
				table.insert(bad, str:sub(i, i + 2))
			end
		end

		table.insert(unescaped, ch)
		i = i + 1
	end

	return unescaped, bad
end

local function unescape_percents(str)
	return table.concat((decode_percents(str)))
end

-- Return a comma separated list of each invalid percent sequence.  Follow each
-- '%' with the next 2 characters for context.  If input is valid, return nil.
local function get_bad_percent_decode_str(str)
	local dummy, bad = decode_percents(str)

	if #bad > 0 then
		return u.parse_msg(table.concat(bad, ", "), {})
	end
end

-- Return human-readable representation of all forbidden characters (according
-- to RFC 3986) found within input string.  If input is valid, return nil.
local function get_invalid_uri_chars_str(str)
	local inval = {}
	local mst = u.get_msg_table(str)

	for i, v in ipairs(mst) do
		if #v == 2 then
			if v ~= "\\\\" then
				-- FNC or BAD char
				table.insert(inval, v)
			end
		elseif v:find("[\x00-\x20\"<>\\%^`{|}\x7F-\xFF]") then
			-- Control chars, hi-bit set and these literals: [ "<>\^`{|}]
			table.insert(inval, v)
		end
	end

	if #inval > 0 then
		return u.parse_msg(table.concat(inval))
	end
end

local function gs1_uri_breakdown(msg)
	local r = {}

	r.bad_lit = get_invalid_uri_chars_str(msg)

	if not r.bad_lit then
		r.bad_pc = get_bad_percent_decode_str(msg)

		if not r.bad_pc then
			r.parts = mod.split_uri(msg)
		end
	end

	return r
end

-- Wrapper for GS1 plugins that only use direct aicheck output (no extra reqs).
-- Just return original msg if symbol content is fixed (EAN/UPC/ITF/DataBar).
function mod.get_aicheck_results(msg, subsymb, opt)
	local plugin_details, hr

	if mod.ai_checked_with_fnc1[subsymb] then
		local aicheck_output = mod.aicheck(msg, subsymb, opt)
		plugin_details = mod.get_plugin_details(aicheck_output)
		hr = aicheck_output.human_readable

	elseif mod.digital_link_symbs[subsymb] then
		local uri = gs1_uri_breakdown(msg)

		-- Make sure first AI was in the path, not query
		if uri.parts and (#uri.parts.ais > 0) and (uri.parts.ais[1][3]) then
			local combined = {}
			for i, v in ipairs(uri.parts.ais) do
				table.insert(combined, v[1] .. unescape_percents(v[2]))
			end
			-- TODO: should we make sure every AI ends with FNC1?
			--table.insert(combined, "")
			combined = table.concat(combined, "\\1")

			local aicheck_output = mod.aicheck(combined, subsymb)
			plugin_details = mod.get_plugin_details(aicheck_output)
			hr = aicheck_output.human_readable

			local is_primary = {}
			for i, v in ipairs(primary_ais) do
				is_primary[v] = true
			end

			local pkey = uri.parts.ais[1][1]

			if not is_primary[pkey] then
				local det = plugin_details[1]
				det.status = mod.plugin_status.fail
				det.errors = det.errors or {}
				table.insert(det.errors, _("Invalid primary key"))

				pkey = nil
			end

			local q_list = qualifier_ais[pkey] or {}
			if pkey == "01" then
				for i = 2, #uri.parts.ais do
					if uri.parts.ais[i][1] == "235" then
						q_list = { "235" }
						break
					end
				end
			elseif pkey == "414" then
				for i = 2, #uri.parts.ais do
					if uri.parts.ais[i][1] == "7040" then
						q_list = { "7040" }
						break
					end
				end
			end

			local is_qualifier = {}
			local order_info = {}
			for i, v in ipairs(q_list) do
				is_qualifier[v] = true
				table.insert(order_info, { ai = v })
			end

			for i = 2, #uri.parts.ais do
				local from_path = uri.parts.ais[i][3]
				-- If qualifier is in query, or attribute is in path
				if is_qualifier[uri.parts.ais[i][1]] ~= from_path then
					local det = plugin_details[i]
					det.status = mod.plugin_status.fail
					det.errors = det.errors or {}
					table.insert(det.errors, from_path and
					             _("Invalid key qualifier") or
					             _("Invalid attribute"))
				end
			end
			mod.check_order(plugin_details, aicheck_output, order_info)

			for i, v in ipairs(uri.parts.reserved) do
				table.insert(plugin_details, {
					title  = v[1],
					value  = u.parse_msg(v[2]),
					desc   = _("Invalid extension - all numeric keys are reserved"),
					status = mod.plugin_status.fail,
				})
			end
		else
			local new = { status = mod.plugin_status.fail }

			if uri.bad_lit then
				new.value = uri.bad_lit
				new.title = _("Encoding")
				new.desc  = _("Invalid characters")
			elseif uri.bad_pc then
				new.value = uri.bad_pc
				new.title = _("Encoding")
				new.desc  = _("Invalid percent encoding")
			else
				new.value = _(u.plugin_status_str.fail)
				new.title = _("Structure")
				new.desc  = _("Invalid GS1 Digital Link URI")
			end

			plugin_details = { new }
			hr = u.parse_msg(msg)
		end
	else
		plugin_details = {}
		-- Don't escape whitespace between ean/upc & addon
		hr = u.parse_msg(msg, subsymb:find("_[25]$") and {} or nil)
	end

	return plugin_details, hr
end

local function get_notes(info, assoc)
	local notes = {}

	if type(info) == "string" and #info > 0 then
		table.insert(notes, info)
	end

	if type(assoc) == "string" and #assoc > 0 then
		table.insert(notes, assoc)
	end

	if #notes < 1 then
		return nil
	end

	return table.concat(notes, "\n")
end

-- Create a plugin details table from an aicheck output table.
function mod.get_plugin_details(aicheck_output, gs1_ovr, fields_ovr_func)
	local plugin_details = {}

	if #aicheck_output.ais < 1 then
		plugin_details[1] = {
			title  = _("Missing AI"),
			status = mod.plugin_status.fail,
			desc   = _("Error: Code message empty"),
		}

		aicheck_output.human_readable = plugin_details[1].desc
		return plugin_details
	end

	local friendly_hr = gs1_ovr and {} or nil
	gs1_ovr = gs1_ovr or {}

	for k, v in ipairs(aicheck_output.ais) do
		local ai_label, ai_desc, fields_str, f_elm_str
		local show_notes = true
		if (type(v.fields) == "table") and (#v.fields > 0) then
			fields_str = table.concat(v.fields, ", ")

			if fields_ovr_func and (not v.errors) then
				fields_str = fields_ovr_func(v.ai, v.data, fields_str)
			end
		end

		if type(gs1_ovr[v.ai]) == "table" then
			-- If GS1 data title is used, try to display AI value as human-friendly
			-- interpretation in hr.  If not, always display "as-encoded" in hr.
			if gs1_ovr[v.ai].title then
				-- E.g. "17180226" will become: "EXPIRY" & "EXPIRY: 26/02/2018"
				ai_label = gs1_ovr[v.ai].title
				f_elm_str = string.format("%s: %s",
				                          ai_label, fields_str or v.data)
			end

			if gs1_ovr[v.ai].desc then
				ai_desc = _(gs1_ovr[v.ai].desc)
			end

			if gs1_ovr[v.ai].hide_notes then
				show_notes = nil
			end
		end

		-- Set any strings that have not already been overridden
		-- E.g. "22ARBITRARYTEXT" will become: "(22)" & "(22)ARBITRARYTEXT"
		ai_label = ai_label or string.format("(%s)", v.ai)
		f_elm_str = f_elm_str or string.format("(%s)%s", v.ai, v.data)
		ai_desc = ai_desc or v.title

		plugin_details[k] = {
			title    = ai_label,
			value    = fields_str or v.data,
			status   = ai_status[v.status],
			desc     = ai_desc,
			notes    = show_notes and get_notes(v.info, v.assoc),
			warnings = v.warnings,
			errors   = v.errors,
		}

		if friendly_hr then
			table.insert(friendly_hr, f_elm_str)
		end
	end

	if friendly_hr then
		aicheck_output.human_readable = table.concat(friendly_hr, "  ")
	end

	return plugin_details
end

-- Checks for missing ais and add errors to plugin details if missing ais are
-- found.  ai_info is an array containing tables { .ai, .desc }, e.g.
-- { ai = "01", desc = "Product Code" }
function mod.check_missing(plugin_details, aicheck_output, ai_info)
	local ai_present = {}

	for _, v in ipairs(aicheck_output.ais) do
		ai_present[v.ai] = true
	end

	for k, v in ipairs(ai_info) do
		if (not v.obsolete) and (not ai_present[v.ai]) then
			local new_error = {
				title = string.format(_("Missing AI: %s"), v.ai),
				status = mod.plugin_status.fail,
				desc = _(v.desc),
			}
			table.insert(plugin_details, new_error)
		end
	end
end

function mod.check_unwanted(plugin_details, aicheck_output, ai_info)
	local wanted_ais = {}

	for _, v in ipairs(ai_info) do
		wanted_ais[v.ai] = true
	end

	for k, v in ipairs(aicheck_output.ais) do
		if not wanted_ais[v.ai] then
			local unwanted_str = string.format(_("Unwanted AI: %s"), v.ai)
			plugin_details[k].errors = plugin_details[k].errors or {}
			table.insert(plugin_details[k].errors, unwanted_str)
			plugin_details[k].status = mod.plugin_status.fail
		end
	end
end

function mod.check_obsolete(plugin_details, aicheck_output, ai_info)
	local obsolete_ais = {}

	for _, v in ipairs(ai_info) do
		obsolete_ais[v.ai] = v.obsolete
	end

	for k, v in ipairs(aicheck_output.ais) do
		local obsai = obsolete_ais[v.ai]
		if obsai then
			-- TODO: is "unwanted" the right word?
			local obsolete_str = string.format(_("Unwanted AI: %s"), v.ai)
			plugin_details[k].warnings = plugin_details[k].warnings or {}
			table.insert(plugin_details[k].warnings, obsolete_str)

			if plugin_details[k].status == mod.plugin_status.pass then
				plugin_details[k].status = mod.plugin_status.warn
			end

			if (type(obsai) == "string") and (#obsai > 0) then
				plugin_details[k].notes = _(obsai)
			end
		end
	end
end

-- Checks that the order of AIs in aicheck_output corresponds to the order
-- in ai_info.
function mod.check_order(plugin_details, aicheck_output, ai_info)
	local ai_order = {}

	for k, v in ipairs(ai_info) do
		ai_order[v.ai] = v.order or k
	end

	for k, v in ipairs(aicheck_output.ais) do
		local order1 = ai_order[v.ai] or 0
		local follow_ais = {}

		if k == #aicheck_output.ais then
			break
		end

		-- Look ahead to see if there are any ais that we need to follow
		for i = k + 1, #aicheck_output.ais do
			local v2 = aicheck_output.ais[i]
			local order2 = ai_order[v2.ai] or math.huge

			if order1 > order2 then
				table.insert(follow_ais, v2.ai)
			end
		end

		if #follow_ais > 0 then
			local follow_str

			if #follow_ais == 1 then
				follow_str = _("Must follow AI: %s")
			else  -- if #follow_ais > 1 then
				follow_str = _("Must follow AIs: %s")
			end

			local ai_list = {}
			for k, v in ipairs(follow_ais) do
				table.insert(ai_list, v)
			end
			ai_list = table.concat(ai_list, ", ")

			plugin_details[k].errors = plugin_details[k].errors or {}
			table.insert(plugin_details[k].errors, follow_str:format(ai_list))
			plugin_details[k].status = mod.plugin_status.fail
		end
	end
end

function mod.dim_in_range(dim, limits)
	-- limits table must exist and contain at least one boundary to test
	if type(limits) ~= "table" then
		return false
	end

	-- Allow ±20 µm on xdim (roughly 2% @ 990 µm module size)
	return u.is_in_range(dim, limits.min, limits.max, { min = 20 })
end

-- Return min-max (inclusive) range of allowed X-dimensions for given symbology
function mod.get_xdim_range(spec, subsymb)
	if spec[subsymb] then
		return spec[subsymb].mod_size or spec.generic.mod_size
	end

	return nil
end

-- Return min-max (inclusive) range of X-dimensions for which the plugin should
-- make recommendations.  No limits implies any size is applicable.  Since
-- dim_in_range() expects at least one limit, use min = 0 if open-ended.
function mod.get_cutoff_range(spec, subsymb)
	if spec[subsymb] then
		return spec[subsymb].cutoff or spec.generic.cutoff or { min = 0 }
	end

	return nil
end

-- Return req. aperture and min. pass grade for given symbology, or {-1, -1} if N/a
function mod.get_quality_spec(spec, subsymb, xdim)
	local aperture, min_grade = -1, -1

	if spec[subsymb] then
		-- Catch ITF-14 larger aperture/lower pass grade
		if spec[subsymb].big_spec and (xdim >= spec[subsymb].big_spec.threshold) then
			aperture = (spec[subsymb].big_spec.aperture or
			            spec[subsymb].aperture or
			            spec.generic.aperture)
			min_grade = (spec[subsymb].big_spec.pass_grade or
			             spec[subsymb].pass_grade or
			             spec.generic.pass_grade)
		else
			aperture = spec[subsymb].aperture or spec.generic.aperture
			min_grade = spec[subsymb].pass_grade or spec.generic.pass_grade
		end
	end

	return aperture, min_grade
end

-- Returns aperture in accordance with the supplied spec
function mod.get_rec_aperture(spec, subsymb, mod_width, is_dpm)
	local aperture = -1

	if not is_dpm and spec[subsymb] then
		local xdim_range = mod.get_cutoff_range(spec, subsymb)

		if mod.dim_in_range(mod_width, xdim_range) then
			aperture = mod.get_quality_spec(spec, subsymb, mod_width)
		end
	end

	return u.spec_to_req_aperture(aperture, mod_width)
end

-- Returns minimum pass grade in accordance with the supplied spec
function mod.get_pass_grade(spec, subsymb, xdim)
	local pass_grade, dummy = -1

	if spec[subsymb] then
		local xdim_range = mod.get_cutoff_range(spec, subsymb)

		if mod.dim_in_range(xdim, xdim_range) then
			dummy, pass_grade = mod.get_quality_spec(spec, subsymb, xdim)
		end
	end

	return pass_grade
end

-- Return true if the decoded msg contains at least one of the required AIs
function mod.msg_contains_req_ai(req_ais, msg, subsymb)
	-- If list of required AIs is not set, return true
	if not req_ais then
		return true
	end

	local aicheck_output = mod.aicheck(msg, subsymb)

	for k, v in ipairs(aicheck_output.ais) do
		if req_ais[v.ai] then
			return true
		end
	end

	return false
end

-- Check if decoded msg contains at least one of the required AIs.
-- If none are found, add an error to plugin_details.
function mod.check_contains_req_ai(plugin_details, req_ais, msg, subsymb)
	if mod.msg_contains_req_ai(req_ais, msg, subsymb) then
		return
	end

	local ai_list_str = {}
	for k, v in pairs(req_ais) do
		table.insert(ai_list_str, k)
	end

	local new = {
		title  = _("Missing AI"),
		status = mod.plugin_status.fail,
		desc   = string.format(_("Plugin requires at least one of the following AIs: %s"),
		                       table.concat(ai_list_str, ", ")),
	}

	table.insert(plugin_details, new)
end

function mod.get_ai(msg, subsymb, ai)
	if mod.ai_checked_with_fnc1[subsymb] then
		local aicheck_output = mod.aicheck(msg, subsymb)

		for k, v in ipairs(aicheck_output.ais) do
			if v.ai == ai then
				return v.data
			end
		end
	elseif mod.digital_link_symbs[subsymb] then
		local uri = gs1_uri_breakdown(msg)

		if uri.parts then
			for k, v in ipairs(uri.parts.ais) do
				if v[1] == ai then
					return u.parse_msg(v[2])
				end
			end
		end
	end

	return nil
end

-- Return true if full plugin results should be displayed, else return false
-- Last 2 params (req_ais and msg) only required for SSTs 5, 9 and 11
function mod.results_applicable(spec, subsymb, aperture, xdim,
                                always_show, is_dpm, req_ais, msg)
	if is_dpm or not spec[subsymb] then
		return false
	elseif always_show then
		return true
	end

	local xdim_range = mod.get_cutoff_range(spec, subsymb)

	if mod.dim_in_range(xdim, xdim_range) then
		local spec_aperture = mod.get_quality_spec(spec, subsymb, xdim)

		if u.is_aperture_similar(aperture, spec_aperture, xdim) then
			return mod.msg_contains_req_ai(req_ais, msg, subsymb)
		end
	end

	return false
end

local function get_min_height_dim(subsymb, xdim, xdim_range, nrows, override)
	local min_height_dim = u.mk_dim("mm", "double", -1, 1)
	local min_val  --, symb_note
	-- Clamp X-dim to valid range
	xdim = u.clamp(xdim, xdim_range.min, xdim_range.max)

	if u.is_ean_upc[subsymb] then
		local eanupc_rec_min, eanupc_rec_max, x_height_ratio

		if subsymb == "ean8" then
			-- X:Height ratio = Height of (nominal) 100% EAN-8 / Xdim of 100% EAN-8
			x_height_ratio = 18.23 / 330
			eanupc_rec_min = 14.58  -- EAN-8 height in mm @ 80%
			eanupc_rec_max = 36.46  -- EAN-8 height in mm @ 200%
		else
			-- X:Height ratio = Height of (nominal) 100% EAN-13 / Xdim of 100% EAN-13
			x_height_ratio = 22.85 / 330
			eanupc_rec_min = 18.28  -- EAN-13/UPC height in mm @ 80%
			eanupc_rec_max = 45.70  -- EAN-13/UPC height in mm @ 200%
		end

		min_val = xdim * x_height_ratio
		-- Clamp min height to 80%..200% range
		min_val = u.clamp(min_val, eanupc_rec_min, eanupc_rec_max)

	elseif subsymb == "gs1_128" or subsymb == "itf14" then
		if override and override.absolute then
			min_val = override.absolute
		else
			min_val = 31.75  -- Nominal ITF-14 / Code-128 height
		end

	elseif u.is_databar[subsymb] then
		local xmul  -- Don't include separator patterns

		if override and override.scalar then
			-- Only used by tables 1 & 3, which require over square aspect ratio
			xmul = override.scalar * (nrows or 1)
		elseif subsymb == "databar_limited" then
			xmul = 10
		elseif subsymb == "databar_expanded" then
			xmul = 34 * (nrows or 1)
		else  -- if subsymb == "databar" then
			-- Assume truncated or stacked. TODO: detect omni?
			-- The minimum overall height for any "databar" is 13X for both the
			-- truncated (one-row) and stacked variants.  But we don't include
			-- the separator in the overall height, so the absolute minimum (as
			-- we measure it) for "GS1 DataBar Stacked" is 12X.
			if nrows and nrows == 2 then
				xmul = 12
			else
				xmul = 13
			end
		end

		min_val = xdim * xmul / 1000
--[[
	elseif u.is_composite[subsymb] then
		if subsymb == "cca" then
			-- Minimum 3 rows, 2x xdim per row.  TODO: (nrows * 2 * xdim)
			min_val = 6 * xdim / 1000
		elseif subsymb == "ccb" then
			-- Minimum 10 rows, 2x xdim per row.  TODO: (nrows * 2 * xdim)
			min_val = 20 * xdim / 1000
		elseif subsymb == "ccc" then
			-- Minimum 3 rows, 3x xdim per row.  TODO: (nrows * 3 * xdim)
			min_val = 9 * xdim / 1000
		end
--]]
	end

	min_height_dim.param.val = min_val
--	min_height_dim.note = symb_note

	return min_height_dim
end

function mod.check_bar_height(plugin_details, spec, param_list, subsymb)
	local actual_height, actual_height_str, nrows

	for k, v in ipairs(param_list) do
		if v.valid then
			if v.id == "height" and v.unit == "mm" then
				-- Doesn't make sense to compare values if not the same precision,
				-- so always use 1 d.p. in metric (+2 = 3 d.p. in imperial).
				v.param.precision = 1
				actual_height     = u.round_param(v.param)
				actual_height_str = u.dim_str(v)
			elseif (v.id == "rows" and v.type == "base" and
			        v.param.type == "int" and v.param.val > 0) then
				nrows = v.param.val
			end
		end
	end

	if not actual_height then
		return
	end

	local min_height, min_height_str, min_height_dim
	do
		local xdim = u.get_mod_sizes_um(param_list)
		local xdim_range = mod.get_xdim_range(spec, subsymb)
		local override = spec[subsymb].height
		min_height_dim = get_min_height_dim(subsymb, xdim, xdim_range, nrows, override)
	end

	min_height     = u.round_param(min_height_dim.param)
	min_height_str = u.dim_str(min_height_dim)

	local new = {
		title  = _("Bar Height"),
		value  = actual_height_str,
		desc   = string.format(_("Height of bars should be at least %s"), min_height_str),
	--	notes  = min_height_dim.note,
		status = mod.plugin_status.na,
	}

--[[ Nobody pays attention to minimum height rules so don't check, only display
	if actual_height < min_height then
		new.status = mod.plugin_status.fail
	end
--]]

	table.insert(plugin_details, new)
end

function mod.check_itf14_wnr(plugin_details, param_list)
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
			desc   = _("Wide:Narrow ratio should be between 2.25:1 and 3.0:1"),
			status = mod.plugin_status.pass,
		}

		-- Warn if wnr is between 3.05 and 3.15
		if (wnr < 2.25) or (wnr >= 3.15) then
			new.status = mod.plugin_status.fail
		elseif wnr >= 3.05 then
			new.status = mod.plugin_status.warn
		end

		table.insert(plugin_details, new)
	end
end

-- Wrapper for GS1 plugins that always check the grade/aperture/module sizes
-- (and bar heights) against the symbol specification tables (genspec 5.12.3)
function mod.do_application_checks(plugin_details, spec, param_list,
                                   subsymb, overall_grade, aperture)
	local xdim = u.get_min_mod_size(param_list)
	local spec_ap, spec_min_grade = mod.get_quality_spec(spec, subsymb, xdim)

	u.check_grade(plugin_details, overall_grade, spec_min_grade)
	u.check_aperture(plugin_details, aperture, spec_ap, xdim)

	local xdim_range = mod.get_xdim_range(spec, subsymb)
	-- Set this for all GS1 plugins to allow ±2% acceptance criteria on xdim
	xdim_range.apply_gs_ac = true
	u.check_mod_size(plugin_details, param_list, xdim_range)

	if subsymb == "itf14" then
		mod.check_itf14_wnr(plugin_details, param_list)
	end

	if not u.is_composite[subsymb] then
		mod.check_bar_height(plugin_details, spec, param_list, subsymb)
	end

	if subsymb:sub(1, 4) == "2005" then
		u.check_qr_qz(plugin_details, param_list)
	end
end

--------------------------------------------------------------------------------
--[[ Deprecated: used by old aicheck plugin

mod.ai_check_missing_fnc1 = {
	code128         = true,
	code128_kraft   = true,
	code128_kraft_a = true,
	code128_kraft_b = true,
}

-- Returns aperture and pass grade in accordance with the GS1 Genereral
-- specifications.
function mod.gs1_recs(subsymb, mod_size)
	if u.is_ean_upc[subsymb] then
		return 150, 1.5
	elseif subsymb == "itf14" then
		if (mod_size < 635) then
			return 250, 1.5
		else
			return 500, 0.5
		end
	elseif subsymb == "gs1_128" then
		return 250, 1.5
	elseif subsymb == "databar_expanded" then
		return 250, 1.5
	elseif subsymb == "databar" then
		return 150, 1.5
	elseif subsymb == "ecc_200_fnc1_1_5" then
		return 200, 1.5
	elseif subsymb == "2005_fnc1_1" then
		return 200, 1.5
	else
		return -1, -1
	end
end

--]]
--------------------------------------------------------------------------------

return mod
