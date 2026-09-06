-- acv_gettext() should be exported to all scripts.  Export some more idiomatic
-- functions here.  Scripts may of course just use acv_gettext() directly.

-- The no-op function.  Used usually in variable declarations in C.  Not
-- strictly necessary in lua, since function calls are allowed anywhere.
function N_(s)
	return s
end

-- What should gettext return with an input of nil or ""?  The no-op function
-- returns whatever was passed in (i.e. nil, ""), but acv_gettext() may not,
-- depending on the implementation of dgettext().  To avoid the inconsistency
-- entirely, catch UB-inducing input and return whatever N_() returns.
local function safe_gettext(s)
	if (not s) or (s == "") then
		return N_(s)
	end

	return acv_gettext(s)
end

-- The usual gettext alias, falling back to N_
if acv_gettext then
	_ = safe_gettext
else
	_ = N_
end
