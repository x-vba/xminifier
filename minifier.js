"use strict";


/**
 * This function combines a series of smaller functions that each perform a
 * single step in the minification process.
 *
 * @author Anthony Mancini
 * @version 1.0.0
 * @license MIT
 * @param {string} vbaSourceCode is the source code from a VBA module
 * @returns {string} the VBA source code in minified form
 */
function vbaMinifier(vbaSourceCode) {
	vbaSourceCode = _removeExcessLines(vbaSourceCode);
	vbaSourceCode = _removeExcessWhitespace(vbaSourceCode);
	vbaSourceCode = _removeSingleLineComments(vbaSourceCode);
	vbaSourceCode = _removeEndOfLineComments(vbaSourceCode);
	vbaSourceCode = _removeMulitline(vbaSourceCode);
	//vbaSourceCode = _removeSyntaxWhitespace(vbaSourceCode);
	//vbaSourceCode = _minifyVariableNames(vbaSourceCode);
	vbaSourceCode = _replaceVariableTypeNames(vbaSourceCode);

	return vbaSourceCode
}


/**
 * This is one of the minification functions. It simply removes all excess
 * lines (lines that when whitespace is removed contain empty strings)
 *
 * @author Anthony Mancini
 * @version 1.0.0
 * @license MIT
 * @param {string} vbaSourceCode is the source code from a VBA module
 * @returns {string} the VBA code with excess lines removed
 */
function _removeExcessLines(vbaSourceCode) {
	vbaSourceCode = vbaSourceCode.split("\n");
	vbaSourceCode = vbaSourceCode.filter(codeLine => {
		if (codeLine.trim() === "") {
			return false
		}
		
		return true
	});
	
	return vbaSourceCode.join("\n")
}


/**
 * This function removes all excess whitespace (for example the left tab
 * whitespace and excess whitespace at the end of a line) on each line
 *
 * @author Anthony Mancini
 * @version 1.0.0
 * @license MIT
 * @param {string} vbaSourceCode is the source code from a VBA module
 * @returns {string} the VBA code with excess white space removed
 */
function _removeExcessWhitespace(vbaSourceCode) {
	vbaSourceCode = vbaSourceCode.split("\n");
	vbaSourceCode = vbaSourceCode.map(codeLine => codeLine.trim());
	
	return vbaSourceCode.join("\n")
}


/**
 * This function removes comments from the VBA source code. It currently only
 * supports comments that on their own seperate line and does not support
 * comments that come after a statement. For example, this comment will be
 * removed currently:
 * 
 * ' This is my comment
 * 
 * But this comment will not be removed currently:
 * 
 * Dim myVar as String ' This is my comment
 *
 * @author Anthony Mancini
 * @version 1.0.0
 * @license MIT
 * @param {string} vbaSourceCode is the source code from a VBA module
 * @returns {string} the VBA code with comments removed
 */
function _removeSingleLineComments(vbaSourceCode) {
	vbaSourceCode = vbaSourceCode.split("\n");
	vbaSourceCode = vbaSourceCode.filter(codeLine => {
		if (codeLine.trim().startsWith("'")) {
			return false
		}
		
		return true
	});
	
	return vbaSourceCode.join("\n")
}


/**
 * This function removes end of line comments by checking if the comment start
 * character (') is outside of a string and then removing all characters to the
 * right of the comment start. Technically this function should remove all
 * comments in general and in the future the single line comment function should
 * potentially be deprecated
 *
 * @author Anthony Mancini
 * @version 1.0.0
 * @license MIT
 * @param {string} vbaSourceCode is the source code from a VBA module
 * @returns {string} the VBA code with end of line comments removed
 */
function _removeEndOfLineComments(vbaSourceCode) {
	vbaSourceCode = vbaSourceCode.split("\n");
	vbaSourceCode = vbaSourceCode.map(codeLine => {
		let doubleQuoteFlag = false;
		let singleQuoteLocation;
		
		for (let i = 0; i <= codeLine.length; i++) {
			if (codeLine.charAt(i) === '"') {
				doubleQuoteFlag = !doubleQuoteFlag
			}
			
			// When we are not currently within a string literal, get the location
			// of the single quote
			if (codeLine.charAt(i) === "'" && doubleQuoteFlag === false) {
				singleQuoteLocation = i;
				break;
			} 
		}
		
		// Get all characters to the left of the singleQuoteLocation
		if (singleQuoteLocation !== undefined) {
			codeLine = codeLine.substr(0, singleQuoteLocation);
		}

		return codeLine
	});
	
	return vbaSourceCode.join("\n")
}


/**
 * This function combines multiline VBA statements into a single line. These
 * VBA statements end with an underscore _ and act as a continutation on the
 * next line.
 *
 * @author Anthony Mancini
 * @version 1.0.0
 * @license MIT
 * @param {string} vbaSourceCode is the source code from a VBA module
 * @returns {string} the VBA code with multiline statements combined
 */
function _removeMulitline(vbaSourceCode) {
	vbaSourceCode = vbaSourceCode.split("\n");
	vbaSourceCode = vbaSourceCode.map(codeLine => codeLine.trimRight());
	
	for (let lineNumber = vbaSourceCode.length - 1; lineNumber >= 0; lineNumber--) {
		if (vbaSourceCode[lineNumber].endsWith("_")) {
			vbaSourceCode[lineNumber] = vbaSourceCode[lineNumber].substr(0, vbaSourceCode[lineNumber].lastIndexOf("_"));
			vbaSourceCode[lineNumber] = vbaSourceCode[lineNumber] + vbaSourceCode[lineNumber + 1];
			vbaSourceCode[lineNumber + 1] = null;
		}
	}
	
	vbaSourceCode = vbaSourceCode.filter(codeLine => codeLine !== null);
	
	return vbaSourceCode.join("\n")
}

/**
 * This function removes excess syntax whitespace (which is often generated by
 * the VBE when creating VBA code). For example, VBA code like:
 * 
 * c=a+b
 * 
 * Will be converted by the VBE editor to:
 * 
 * c = a + b
 *
 * This function will remove this excess whitespace
 *
 * @author Anthony Mancini
 * @version 1.0.0
 * @license MIT
 * @param {string} vbaSourceCode is the source code from a VBA module
 * @returns {string} the VBA code with excess syntax whitespace removed
 */
function _removeSyntaxWhitespace(vbaSourceCode) {
	let sytnaxReplacementArray = [
		["= ", "="],
		[" =", "="],
		["+ ", "+"],
		[" +", "+"],
		["- ", "-"],
		[" -", "-"],
		["* ", "*"],
		[" *", "*"],
		["/ ", "/"],
		[" /", "/"],
		["( ", "("],
		[" (", "("],
		[") ", ")"],
		[" )", ")"],
		[", ", ","],
		[" ,", ","],
		["< ", "<"],
		[" <", "<"],
		["> ", ">"],
		[" >", ">"],
	];
	
	sytnaxReplacementArray.forEach(syntaxArray => {
		vbaSourceCode = vbaSourceCode.split(syntaxArray[0]).join(syntaxArray[1]);	
	});
	
	return vbaSourceCode
}


/**
 * This function will reduce variable name size by replacing the varible names
 * defined by the user to smaller names. For example, a variable declared as:
 *
 * Dim myVariable As String
 *
 * Will become:
 *
 * Dim b As String
 *
 * @author Anthony Mancini
 * @version 1.0.0
 * @license MIT
 * @todo Currently this function does not work since i need to find a proper
 * and variables with names like i will have difficulty replacing
 * @param {string} vbaSourceCode is the source code from a VBA module
 * @returns {string} the VBA code variable names minified
 */
function _minifyVariableNames(vbaCode) {
	// Variable names from Dim statments
	let allVariableNames = vbaCode.match(/Dim\s[a-zA-Z0-9_]+?(\s|\n)/gmi);
	allVariableNames = allVariableNames.map(variableCode => {
		variableCode = variableCode.replace("Dim ", "");
		variableCode = variableCode.trim();
		return variableCode
	});
	
	// Variable names from arguments
	let combinedAllArgumentNames = [];
	let allArgumentNames = vbaCode.match(/(Function|Sub)\s[a-zA-Z0-9_]+?\s*?[(][\S\s]*?[)]/gmi);
	allArgumentNames = allArgumentNames.map(functionDeclaration => {
		functionDeclaration = functionDeclaration.split("(")[1];
		functionDeclaration = functionDeclaration.replace(")", "");
		functionDeclaration = functionDeclaration.split(",");
		functionDeclaration = functionDeclaration.map(functionName => functionName.split(" As")[0].trim());
		functionDeclaration = functionDeclaration.map(functionName => functionName.substr(functionName.lastIndexOf(" "), functionName.length - 1).trim());
		
		functionDeclaration.forEach(variableName => combinedAllArgumentNames.push(variableName));
		
		return functionDeclaration
	});
	
	// Removing duplicates
	let allVariableNamesSet = new Set();
	allVariableNames.forEach(variableName => allVariableNamesSet.add(variableName));
	combinedAllArgumentNames.forEach(variableName => allVariableNamesSet.add(variableName));
	allVariableNames = Array.from(allVariableNamesSet);
	
	// Performing the replacement
	console.log(allVariableNames);
	
	return vbaCode
}


/**
 * This function replaces variable Type Names with their shorthand version.
 * For example, if you declare a variable as Type String:
 *
 * Dim myString As String
 *
 * A shorthand Type Name would be:
 *
 * Dim myString$
 *
 * @author Anthony Mancini
 * @version 1.0.0
 * @license MIT
 * @todo make the Variant replacement code more robust
 * @param {string} vbaSourceCode is the source code from a VBA module
 * @returns {string} the VBA code variable Type Names minified with shorthand versions
 */
function _replaceVariableTypeNames(vbaSourceCode) {
	// Changing the argument types to their shorthand versions
	let allArgumentNames = vbaSourceCode.match(/(Function|Sub)\s[a-zA-Z0-9_]+?\s*?[(][\S\s]*?[)]/gmi);
	
	if (allArgumentNames !== null) {
		allArgumentNames.forEach(argumentName => {
			let modifiedArgumentName = argumentName;
			
			modifiedArgumentName = modifiedArgumentName.replace(/\sAs\sInteger/gmi, "%");
			modifiedArgumentName = modifiedArgumentName.replace(/\sAs\sLong/gmi, "&");
			modifiedArgumentName = modifiedArgumentName.replace(/\sAs\sDecimal/gmi, "@");
			modifiedArgumentName = modifiedArgumentName.replace(/\sAs\sSingle/gmi, "!");
			modifiedArgumentName = modifiedArgumentName.replace(/\sAs\sDouble/gmi, "#");
			modifiedArgumentName = modifiedArgumentName.replace(/\sAs\sString/gmi, "$");
			
			vbaSourceCode = vbaSourceCode.split(argumentName).join(modifiedArgumentName);
		});
	}

	// Changing the Dim declarations with their shorthand
	let allVariableNames = vbaSourceCode.match(/Dim\s[a-zA-Z0-9_]+?\sAs\s[a-zA-Z0-9_]+?\s*?\n/gmi);
	
	if (allVariableNames !== null) {
		allVariableNames.forEach(variableName => {
			let modifiedVariableName = variableName;
			
			modifiedVariableName = modifiedVariableName.replace(/\sAs\sInteger/gmi, "%");
			modifiedVariableName = modifiedVariableName.replace(/\sAs\sLong/gmi, "&");
			modifiedVariableName = modifiedVariableName.replace(/\sAs\sDecimal/gmi, "@");
			modifiedVariableName = modifiedVariableName.replace(/\sAs\sSingle/gmi, "!");
			modifiedVariableName = modifiedVariableName.replace(/\sAs\sDouble/gmi, "#");
			modifiedVariableName = modifiedVariableName.replace(/\sAs\sString/gmi, "$");
			
			vbaSourceCode = vbaSourceCode.split(variableName.trim()).join(modifiedVariableName.trim());
		});
	}
	
	// Removing all Variant type declarations, as Variant is the default. One
	// thing to note here is that if a string contains the text " As Variant",
	// this step will erroneously remove those as well.
	vbaSourceCode = vbaSourceCode.replace(/\sAs\sVariant/gmi, "");

	return vbaSourceCode
}


module.exports = {
	vbaMinifier: vbaMinifier,	
};
