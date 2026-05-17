//Outputs various code quality values
//Please note that different metrics are available in different modes
//See documentation for further details
function onResult (decodeResults, readerProperties, output)
{
	// single code read?
	if(decodeResults[0].decoded && decodeResults.length == 1)
	{
		if(decodeResults[0].symbology.name == "Data Matrix" || decodeResults[0].symbology.name == "Direct Mark QR" || decodeResults[0].symbology.name == "QR Code")
		{
			output.content = 
				" Data: " + decodeResults[0].content + ", " +
			    " Symbology: " + decodeResults[0].symbology.name + ", " +
				" Grading Standard: " + decodeResults[0].trucheck.overall.gradingStandard + ", " +
			    " Application Standard: " + decodeResults[0].trucheck.overall.applicationStandardName + ", " +
			    " Acceptance Criteria: " + decodeResults[0].trucheck.overall.applicationStandardPass + ", " +
			    " Overall Grade: " + decodeResults[0].trucheck.overall.gradeLetter +
				" (" + decodeResults[0].trucheck.overall.gradeValue + ")" + ", " +
                " Nominal X-Dimension: " + decodeResults[0].trucheck.general.xDimension + " mils" + ", " +
				" ANU: " + decodeResults[0].trucheck.axialNonuniformity.grade +
				" (" + decodeResults[0].trucheck.axialNonuniformity.raw + "%" + ")" + ", " + 
				" GNU: " + decodeResults[0].trucheck.gridNonuniformity.grade +
				" (" + decodeResults[0].trucheck.gridNonuniformity.raw + "%" + ")" + ", " +  
				" UEC: " + decodeResults[0].trucheck.unusedErrorCorrection.grade + ", " + 
				" FPD: " + decodeResults[0].trucheck.fixedPatternDamage.grade +
				" (" + decodeResults[0].trucheck.fixedPatternDamage.raw + ")" + ", " + 
				" CU: " + decodeResults[0].trucheck.general.contrastUniformity + ", " +
				" Horizontal BWG: " + decodeResults[0].trucheck.general.horizontalBWG + "%" + ", " +
				" Vertical BWG: " + decodeResults[0].trucheck.general.verticalBWG + "%" + ", "; 
		}
			if (decodeResults[0].trucheck.overall.gradingStandard == "ISO 15415")
			{
				output.content = output.content + " SC: " + decodeResults[0].trucheck.symbolContrast.grade +
				" (" + decodeResults[0].trucheck.symbolContrast.raw + "%" + ")" + ", "; 
			}
			else	
			{
				output.content = output.content + " CC: " + decodeResults[0].trucheck.cellContrast.grade +
				" (" + decodeResults[0].trucheck.cellContrast.raw + "%" + ")" + ", "; 
			}
			if (decodeResults[0].trucheck.overall.gradingStandard == "ISO 15415")
			{
				output.content = output.content + " MOD: " + decodeResults[0].trucheck.modulation.grade + ", ";
			}
			else	
			{
				output.content = output.content + " CMOD: " + decodeResults[0].trucheck.cellModulation.grade + ", "; 
			}	
	}
}
