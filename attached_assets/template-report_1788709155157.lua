function acv_report_info(info)
	local tl = require "template-loader"
	return tl.report_info(info)
end

function acv_report(info)
	local tl = require "template-loader"
	return tl.create_report(info)
end
