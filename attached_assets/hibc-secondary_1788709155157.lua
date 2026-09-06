-- require "strict"
local u = require "utils"

local M = {}

-- Date format types
M.df_t = {
	MMYY     = '1',
	MMDDYY   = '2',
	YYMMDD   = '3',
	YYMMDDHH = '4',
	YYJJJ    = '5',
	YYJJJHH  = '6',
	NO_DATE  = '7',
	SUP_DATE = 'D',
}

-- Quantity format types
M.qf_t = {
	OLD_QTY2 = '8',
	OLD_QTY5 = '9',
	SUP_QTY  = 'Q',
}

-- Lookup date/qty field size based on format type
M.fmt_len = {
	[M.df_t.MMYY]     = 3,  -- First digit is in format flag
	[M.df_t.MMDDYY]   = 6,
	[M.df_t.YYMMDD]   = 6,
	[M.df_t.YYMMDDHH] = 8,
	[M.df_t.YYJJJ]    = 5,
	[M.df_t.YYJJJHH]  = 7,
--	[M.df_t.NO_DATE]  = 0,
	[M.df_t.SUP_DATE] = 8,

	[M.qf_t.OLD_QTY2] = { min = 2, max = 2 },
	[M.qf_t.OLD_QTY5] = { min = 5, max = 5 },
	[M.qf_t.SUP_QTY]  = { min = 1, max = 5 },
}

local monthdays = { 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 }
local month_has_leap = { [2] = true }

local function get_leapdays(year)
	if type(year) ~= "number" then
		return nil

	elseif (year % 4) > 0 then
		return 0

	elseif (year % 100) > 0 then
		return 1

	elseif (year % 400) > 0 then
		return 0
	end

	return 1
end

-- Set date table month and day (both nil if invalid)
-- Return true if valid, false if invalid
local function julian_to_date(dt)
	if (not dt) or (not dt.julian) or (not dt.leap) then
		return false
	end

	local daysinyear = 365 + dt.leap

	if (dt.julian > daysinyear) then
		dt.month = nil
		dt.day   = nil
		return false
	end

	-- Reduce day by size of each month until a valid date appears
	local month, day = 1, dt.julian

	while day > monthdays[month] do
		day = day - monthdays[month]

		if month_has_leap[month] then
			day = day - dt.leap
		end

		month = month + 1
	end

	dt.month = month
	dt.day   = day
	return true
end

-- Given the last 2 digits of the encoded year & current year, work out the full year
local function interpret_year(read_yy, cur_yyyy)
	if (not read_yy) or (not cur_yyyy) then
		return nil

	elseif read_yy > 99 then  -- Encoded year is already complete
		return read_yy
	end

	local cur_cent = math.modf(cur_yyyy / 100)
	local cur_yy   = cur_yyyy % 100
	local diff     = cur_yy - read_yy

	local implied_cent = cur_cent

	if diff < -50 then
		implied_cent = cur_cent - 1

	elseif diff >= 50 then
		implied_cent = cur_cent + 1
	end

	return (implied_cent * 100) + read_yy
end

-- Pass in whatever components of a date are available, fill in the rest
-- Return true if date is logical, false if nonsense e.g. 32nd of 13th month
local function complete_date(dt)
	if dt.month then
		if dt.month < 1 or dt.month > 12 then
			return false
		end

		local max_days = monthdays[dt.month]
		if month_has_leap[dt.month] then
			max_days = max_days + dt.leap
		end

		if dt.day then  -- Either MMDDYY, YYMMDD or YYMMDDHH (or YYYYMMDD 2nd sup)
			if (dt.day < 1) or (dt.day > max_days) then
				return false
			end

		else  -- Has month but no day, can only be MMYY
			dt.day = max_days  -- Use last day of the month
		end

	-- No month, can only be YYJJJ or YYJJJHH
	elseif not julian_to_date(dt) then
		return false
	end

	if dt.hour and dt.hour > 23 then
		return false
	end

	return true
end

-- Return date_info table, and rest of untouched secondary data.
--
--	length:   Target length of encoded date, displayed in notes
--	encode:   Actual encoded characters, to be displayed as "value"
--	ts:       Unix time in seconds, UTC (exp only)
--	hr:       Date(time) - expiration formatted as "YYYY-MM-DD [hh:mm:ss]", local
--	                     - production formatted as "YYYY-MM-DD"
--	valid:    Is the encoded date a sensible calendar day?
--	exp:      Error flag - expired
--	short:    Error flag - encoding too short
--	long:     Error flag - encoding too long
--	notreal:  Error flag - illogical date
--
function M.get_date_info(datefmt_t, s, decode_time, check_exp)
	local date_info = {}
	local dt = {}     -- Date table: year, month, day, hour, min, sec
	local mmyy_first  -- If MMYY type, store first char from fmt

	if (not s) or (#s == 0) then
		date_info.encode = ""
		date_info.short  = true
		return date_info, ""

	-- Only date type with no explicit format digit is YYJJJ
	elseif not datefmt_t then
		datefmt_t = M.df_t.YYJJJ

	-- If using MMYY, store the first digit, and set the date format type
	elseif datefmt_t:byte() <= M.df_t.MMYY:byte() then
		mmyy_first = datefmt_t  -- Set to '0' or '1'
		datefmt_t  = M.df_t.MMYY
	end

	date_info.length = M.fmt_len[datefmt_t]
	local date_pattern = string.format("^%s$", string.rep("%d", date_info.length))
	local remain_str

	if datefmt_t == M.df_t.SUP_DATE then -- Supplemental date, do not trim
		date_info.encode = s
	else
		date_info.encode = s:sub(1, date_info.length)
		remain_str = s:sub(date_info.length + 1)
	end

	if not date_info.encode:find(date_pattern) then
		if #date_info.encode < date_info.length then
			date_info.short = true

		elseif #date_info.encode > date_info.length then
			date_info.long = true
		end

		return date_info, remain_str  -- Do nothing, return bad encoding

	elseif datefmt_t == M.df_t.MMYY then
		s = mmyy_first .. s
		dt.month, dt.year = s:match("(%d%d)(%d%d)")

	elseif datefmt_t == M.df_t.MMDDYY then
		dt.month, dt.day, dt.year = s:match("(%d%d)(%d%d)(%d%d)")

	elseif datefmt_t == M.df_t.YYMMDD then
		dt.year, dt.month, dt.day = s:match("(%d%d)(%d%d)(%d%d)")

	elseif datefmt_t == M.df_t.YYMMDDHH then
		dt.year, dt.month, dt.day, dt.hour = s:match("(%d%d)(%d%d)(%d%d)(%d%d)")

	elseif datefmt_t == M.df_t.YYJJJ then
		dt.year, dt.julian = s:match("(%d%d)(%d%d%d)")

	elseif datefmt_t == M.df_t.YYJJJHH then
		dt.year, dt.julian, dt.hour = s:match("(%d%d)(%d%d%d)(%d%d)")

	elseif datefmt_t == M.df_t.SUP_DATE then
		dt.year, dt.month, dt.day = s:match("(%d%d%d%d)(%d%d)(%d%d)")
	end

	for k, v in pairs(dt) do
		dt[k] = tonumber(v)
	end

	local is_utc = (dt.hour ~= nil)
	local decode_dt = os.date(is_utc and "!*t" or "*t", decode_time)
	dt.year = interpret_year(dt.year, decode_dt.year)
	dt.leap = get_leapdays(dt.year)

	if complete_date(dt) then
		if check_exp then
			if is_utc then
				-- Time is encoded, so show it (converted to local)
				date_info.ts = os.time(dt) + u.get_local_utc_offset(decode_time)
				date_info.hr = os.date("%Y-%m-%d %X", date_info.ts)
			else
				-- With no time set, os.time() will have inferred 12pm (midday),
				-- so add another 12 hours to get to 00:00 the next day.
				date_info.ts = os.time(dt) + 43200
			end

			date_info.exp = date_info.ts <= decode_time
		end

		if not date_info.hr then
			date_info.hr = string.format("%04d-%02d-%02d", dt.year, dt.month, dt.day)
		end

		date_info.valid = true
	else
		date_info.notreal = true  -- Formed of digits, but nonsense
	end

	return date_info, remain_str
end

-- Return qty_info table, and rest of untouched str (nil if supplemental)
--
--	encode:  Actual encoded characters, to be displayed as "value"
--	value:   Numeric value (or nil if not valid)
--	short:   Error flag - encoding too short
--	long:    Error flag - encoding too long
--
function M.get_qty_info(qtyfmt_t, s)
	local qty_info = {}
	local remain_str, qty_pattern
	local qty_len = M.fmt_len[qtyfmt_t]  -- min/max table

	if (not s) or (#s == 0) or (not qty_len) then
		qty_info.encode = ""
		return qty_info, ""

	elseif qtyfmt_t == M.qf_t.SUP_QTY then  -- Supplemental qty
		qty_info.encode = s
	else  -- Deprecated secondary 2 or 5 digit qty
		qty_info.encode = s:sub(1, qty_len.max)
		remain_str = s:sub(qty_len.max + 1)
	end

	qty_pattern = string.format("^%s%s$", string.rep("%d", qty_len.min),
		                        string.rep("%d?", qty_len.max - qty_len.min))

	-- Make sure encode is correct number of digits (2, 5 or 1-5)
	if qty_info.encode:find(qty_pattern) then
		qty_info.value = tonumber(qty_info.encode)

	elseif #qty_info.encode < qty_len.min then
		qty_info.short = true

	elseif #qty_info.encode > qty_len.max then
		qty_info.long = true
	end

	return qty_info, remain_str
end

return M
