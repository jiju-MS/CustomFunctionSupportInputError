/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = "./src/functions/functions.ts");
/******/ })
/************************************************************************/
/******/ ({

/***/ "./src/functions/functions.ts":
/*!************************************!*\
  !*** ./src/functions/functions.ts ***!
  \************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */

/* global clearInterval, console, setInterval */

Object.defineProperty(exports, "__esModule", {
  value: true
});

function add(first, second) {
  return first + second + 20000;
}

exports.add = add;

function add400(first, second) {
  return first + second + 400;
}

exports.add400 = add400;
/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */

function clock(invocation) {
  var timer = setInterval(function () {
    var time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = function () {
    clearInterval(timer);
  };
}

exports.clock = clock;
/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */

function currentTime() {
  return new Date().toLocaleTimeString();
}

exports.currentTime = currentTime;
/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */

function increment(incrementBy, invocation) {
  var result = 0;
  var timer = setInterval(function () {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = function () {
    clearInterval(timer);
  };
}

exports.increment = increment;

function customErrorOut(errorCase) {
	switch (errorCase) {
		case 1: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.notAvailable,	// #N/A
				"An error *case 1* was detected in customErrorOut"
			);
			throw error;
		}
		break;
		case 2: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.invalidValue, // #VALUE!
				"An error *case 2* was detected in customErrorOut"
			);
			throw error;
		}
		break;
		case 3: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.divisionByZero, // #DIV/0!
				"This message should not appear in UI"
			);
			throw error;
		}
		break;
		case 4: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.invalidNumber, // #NUM!
				"This message should not appear in UI"
			);
			throw error;
		}
		break;
		case 5: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.nullReference, // #NULL!
				"This message should not appear in UI"
			);
			throw error;
		}
		break;
		case 6: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.notAvailable // #N/A!
			);
			throw error;
		}
		break;
		case 7: {
			var error = new CustomFunctions.Error(); // #VALUE!
			throw error;
		}
		break;
		case 8: {
			var error = new CustomFunctions.Error(
				undefined, // #VALUE!
				"An error *case 8* was detected in customErrorOut"
			);
			throw error;
		}
		break;
		case 9: {
			var error = new CustomFunctions.Error(new Error()); // #VALUE!
			throw error;
		}
		break;
		case 10: {
			throw new Error("This message should not appear in UI"); // #VALUE!
		}
		break;
		// case 11: {
		// 	return new Promise(function(resolve) {
		// 		var error = new CustomFunctions.Error(
		// 		  CustomFunctions.ErrorCode.nullReference // #N/A!
		// 		);
		// 		throw error;
		// 		setTimeout(function () {
		// 		  resolve(error.code);
		// 		}, 1000);
		// 	  });
		// }
		// break;
		default: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.notAvailable,
				"An unknown error case was detected in customErrorOut"
			);
			throw error;
		}
		break;
	}
}

function customErrorOut(errorCase) {
	switch (errorCase) {
		case 1: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.notAvailable,	// #N/A
				"An error *case 1* was detected in customErrorOut"
			);
			throw error;
		}
		break;
		case 2: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.invalidValue, // #VALUE!
				"An error *case 2* was detected in customErrorOut"
			);
			throw error;
		}
		break;
		case 3: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.divisionByZero, // #DIV/0!
				"This message should not appear in UI"
			);
			throw error;
		}
		break;
		case 4: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.invalidNumber, // #NUM!
				"This message should not appear in UI"
			);
			throw error;
		}
		break;
		case 5: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.nullReference, // #NULL!
				"This message should not appear in UI"
			);
			throw error;
		}
		break;
		case 6: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.notAvailable // #N/A!
			);
			throw error;
		}
		break;
		case 7: {
			var error = new CustomFunctions.Error(); // #VALUE!
			throw error;
		}
		break;
		case 8: {
			var error = new CustomFunctions.Error(
				undefined, // #VALUE!
				"An error *case 8* was detected in customErrorOut"
			);
			throw error;
		}
		break;
		case 9: {
			var error = new CustomFunctions.Error(new Error()); // #VALUE!
			throw error;
		}
		break;
		case 10: {
			throw new Error("This message should not appear in UI"); // #VALUE!
		}
		break;
		case 11: {
			return new Promise(function(resolve) {
				var error = new CustomFunctions.Error(
				  CustomFunctions.ErrorCode.nullReference // #N/A!
				);
				throw error;
				setTimeout(function () {
				  resolve(error.code);
				}, 1000);
			  });
		}
		break;
		default: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.notAvailable,
				"An unknown error case was detected in customErrorOut"
			);
			throw error;
		}
		break;
	}
}

exports.customErrorOut = customErrorOut;

function customErrorReturn(errorCase) {
	switch (errorCase) {
		case 1: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.notAvailable,	// #N/A
				"An error *case 1* was detected in customErrorReturn"
			);
			return error;
		}
		case 2: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.invalidValue, // #VALUE!
				"An error *case 2* was detected in customErrorReturn"
			);
			return error;
		}
		case 3: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.divisionByZero, // #DIV/0!
				"This message should not appear in UI"
			);
			return error;
		}
		case 4: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.invalidNumber, // #NUM!
				"This message should not appear in UI"
			);
			return error;
		}
		case 5: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.nullReference, // #NULL!
				"This message should not appear in UI"
			);
			return error;
		}
		case 6: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.notAvailable // #N/A!
			);
			return error;
		}
		case 7: {
			var error = new CustomFunctions.Error(); // #VALUE!
			return error;
		}
		case 8: {
			var error = new CustomFunctions.Error(
				undefined, // #VALUE!
				"An error *case 8* was detected in customErrorReturn"
			);
			return error;
		}
		case 9: {
			var error = new CustomFunctions.Error(
				"Customized", // #VALUE!
				"This message should not appear in UI"
			);
			return error;
		}
		case 10:{
			var error = new CustomFunctions.Error(new Error()); // #VALUE!
			return error;
		}
		case 11: {
			return new Error("This message should not appear in UI"); // #VALUE!
		}
		case 12: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.invalidName // #NAME?
			);
			return error;
		}
		case 13: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.invalidReference // #REF!
			);
			return error;
		}
		default: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.notAvailable,
				"An unknown error case was detected in customErrorReturn"
			);
			return error;
		}
	}
}

exports.customErrorReturn = customErrorReturn;

function customErrorReturnArray(errorCase) {
	switch (errorCase) {
		case 1: {
			var error = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.notAvailable,	// #N/A
				"An error *case 1* was detected in customErrorReturnArray "
			);
			return [['Hello'],[error]];
		}
		case 2: {
			var error1 = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.notAvailable,	// #N/A
				"An error *case 2* was detected in customErrorReturnArray "
			);
			var error2 = new CustomFunctions.Error(
				CustomFunctions.ErrorCode.invalidValue, // #VALUE!
				"An error *case 2* was detected in customErrorReturnArray "
			);
			return [[error1],[error2]];
		}
	}
}

exports.customErrorReturnArray = customErrorReturnArray;

function customErrorInput(inputAllowError, inputAllowErrorOptional, inputAllowErrorRepeating) {
	if (inputAllowError instanceof CustomFunctions.Error) {
		return inputAllowError.code + " detected in param1";
	}
	else if (inputAllowErrorOptional instanceof CustomFunctions.Error) {
		return inputAllowErrorOptional.code + " detected in param2";
	}
	else if (Array.isArray(inputAllowErrorRepeating)){
		for (let i = 0; i < inputAllowErrorRepeating.length; i++) {
			if (inputAllowErrorRepeating[i] instanceof CustomFunctions.Error) {
				return inputAllowErrorRepeating[i].code + " detected in param repeating";
			}
		}

		return "no error detected";
	}
	else {
		return "no error detected";
	}
}

exports.customErrorInput = customErrorInput;

function customErrorInputInvalid(numberAllowError, stringAllowError, boolAllowError) {
	var ret = [];
	var oneRow = [];

	if (numberAllowError instanceof CustomFunctions.Error) {
		oneRow.push(numberAllowError.code + " detected");
	}
	else {
		oneRow.push(numberAllowError);
	}

	if (stringAllowError instanceof CustomFunctions.Error) {
		oneRow.push(stringAllowError.code + " detected");
	}
	else {
		oneRow.push(stringAllowError);
	}

	if (boolAllowError instanceof CustomFunctions.Error) {
		oneRow.push(boolAllowError.code + " detected");
	}
	else {
		oneRow.push(boolAllowError);
	}

	ret.push(oneRow);
	return ret;
}

exports.customErrorInputInvalid = customErrorInputInvalid;

function customErrorTest(singleAny, singleString, singleDefault, multipleAny, multipleDefault, multipleString) {
	return "succuess"
}

exports.customErrorTest = customErrorTest;

function customErrorTest2(singleAny, singleString, singleDefault, multipleAny, multipleDefault, multipleString) {
	var ret = [];
	var oneRow = [];

	if (singleAny instanceof CustomFunctions.Error) {
		oneRow.push(singleAny.code + " detected");
	}
	else {
		oneRow.push(singleAny);
	}

	if (singleString instanceof CustomFunctions.Error) {
		oneRow.push(singleString.code + " detected");
	}
	else {
		oneRow.push(singleString);
	}

	if (singleDefault instanceof CustomFunctions.Error) {
		oneRow.push(singleDefault.code + " detected");
	}
	else {
		oneRow.push(singleDefault);
	}

	for (var i = 0; i < multipleAny.length; ++i) {
		for (var j = 0; j < multipleAny[i].length; ++j) {
			if (multipleAny[i][j] instanceof CustomFunctions.Error) {
				oneRow.push(multipleAny[i][j].code + " detected");
			}
			else {
				oneRow.push(multipleAny[i][j]);
			}
		}
	}

	for (var i = 0; i < multipleDefault.length; ++i) {
		for (var j = 0; j < multipleDefault[i].length; ++j) {
			if (multipleDefault[i][j] instanceof CustomFunctions.Error) {
				oneRow.push(multipleDefault[i][j].code + " detected");
			}
			else {
				oneRow.push(multipleDefault[i][j]);
			}
		}
	}
	
	for (var i = 0; i < multipleString.length; ++i) {
		for (var j = 0; j < multipleString[i].length; ++j) {
			if (multipleString[i][j] instanceof CustomFunctions.Error) {
				oneRow.push(multipleString[i][j].code + " detected");
			}
			else {
				oneRow.push(multipleString[i][j]);
			}
		}
	}

	ret.push(oneRow);
	return ret;
}

exports.customErrorTest2 = customErrorTest2;

function customErrorInputArray(inputAllowError) {
	var ret = [];
	for (var i = 0; i < inputAllowError.length; ++i) {
		var oneRow = [];
		for (var j = 0; j < inputAllowError[i].length; ++j) {
			if (inputAllowError[i][j] instanceof CustomFunctions.Error) {
				oneRow.push(inputAllowError[i][j].code + " detected");
			}
			else {
				oneRow.push(inputAllowError[i][j]);
			}
		}
		ret.push(oneRow);
	}
	return ret;
}

exports.customErrorInputArray = customErrorInputArray;

function customErrorInputInvalidArray(inputAllowError) {
	var ret = [];
	for (var i = 0; i < inputAllowError.length; ++i) {
		var oneRow = [];
		for (var j = 0; j < inputAllowError[i].length; ++j) {
			if (inputAllowError[i][j] instanceof CustomFunctions.Error) {
				oneRow.push(inputAllowError[i][j].code + " detected");
			}
			else {
				oneRow.push(inputAllowError[i][j]);
			}
		}
		ret.push(oneRow);
	}
	return ret;
}

exports.customErrorInputInvalidArray = customErrorInputInvalidArray;

function logMessage(message) {
  console.log(message);
  return message;
}

function GetParameterAddresses(firstParameter, secondParameter, thirdParameter, invocationContext) {
    var items = [
        [invocationContext.parameterAddresses[0]],
        [invocationContext.parameterAddresses[1]],
        [invocationContext.parameterAddresses[2]]
    ];
    return items;
}

exports.GetParameterAddresses = GetParameterAddresses;
 

function GetParameterAddressesRepeating(firstParameter, secondParameter, invocationContext) {
    var resultArray = [];
    for (let i = 0; i < invocationContext.parameterAddresses.length; i++)
    {
        var parameterAddresses = [invocationContext.parameterAddresses[i]];
        resultArray.push(parameterAddresses);
    }
    return resultArray;
}

exports.GetParameterAddressesRepeating = GetParameterAddressesRepeating;

function GetParameterAddressesOptional(firstParameter, secondParameter, invocationContext) {
	var resultArray = [];
    for (let i = 0; i < invocationContext.parameterAddresses.length; i++)
    {
		var parameterAddresses = [invocationContext.parameterAddresses[i]];
        resultArray.push(parameterAddresses);
    }
    return resultArray;
}

exports.GetParameterAddressesOptional = GetParameterAddressesOptional;


function GetParameterAddressesRange(firstParameter, secondParameter, invocationContext) {
	var items = [
		[invocationContext.parameterAddresses[0]],
        [invocationContext.parameterAddresses[1]]
    ];
    return items;
}

exports.GetParameterAddressesRange = GetParameterAddressesRange;


function GetParameterAddressesFalse(firstParameter, secondParameter, invocationContext) {
	return invocationContext.parameterAddresses;
}

exports.GetParameterAddressesFalse = GetParameterAddressesFalse;


function GetParameterAddressesOff(firstParameter, secondParameter, invocationContext) {
	return invocationContext.parameterAddresses;
}

exports.GetParameterAddressesOff = GetParameterAddressesOff;

function customErrorInput2(inputAllowError, inputAllowErrorOptional, inputAllowErrorRepeating) {
	var ret = [];
	var oneRow = [];
	if (inputAllowError instanceof CustomFunctions.Error) {
		oneRow.push(inputAllowError.code + " detected");
	}
	else
	{
		oneRow.push(inputAllowError);
	}

	if (inputAllowErrorOptional instanceof CustomFunctions.Error) {
		oneRow.push(inputAllowErrorOptional.code + " detected");
	}
	else
	{
		oneRow.push(inputAllowErrorOptional);
	}

	if (inputAllowErrorRepeating instanceof CustomFunctions.Error) {
		oneRow.push(inputAllowErrorRepeating.code + " detected");
	}
	else
	{
		oneRow.push(inputAllowErrorRepeating);
	}

	ret.push(oneRow);
	return ret;
}

exports.customErrorInput2 = customErrorInput2;

function SetRichError(errorCase) {
	switch (errorCase) {
		case 1: {
			var error = new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable);
			return error;
		}
		case 2: {
			var error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue);
			return error;
		}
		case 3: {
			var error = new CustomFunctions.Error(CustomFunctions.ErrorCode.divisionByZero);
			return error;
		}
		case 4: {
			var error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
			return error;
		}
		case 5: {
			var error = new CustomFunctions.Error(CustomFunctions.ErrorCode.nullReference);
			return error;
		}
		case 6: {
			return {
				type: "Error",
				basicValue: "#N/A",
				basicType: "Error",
				errorSubType: "HlookupValueNotFound",
				errorType: "NotAvailable"
			};
		}
		default: {
			return "Invalid test case";
		}
	}
}

exports.SetRichError = SetRichError;

function SetFormattedNumber(formattedNumberCase) {
	switch (formattedNumberCase) {
		case 1: {
			var formattedNumber = new CustomFunctions.FormattedNumber(
				1.234,
				"0.00"
			);
			return formattedNumber;
		}
		case 2: {
			var formattedNumber = new CustomFunctions.FormattedNumber(
				1.234,
				"$#,##0.00"
			);
			return formattedNumber;
		}
		case 3: {
			var formattedNumber = new CustomFunctions.FormattedNumber(
				1.23,
				"_($* #,##0.00_);_($* (#,##0.00);_($*??_);_(@_)"
			);
			return formattedNumber;
		}
		case 4: {
			var formattedNumber = new CustomFunctions.FormattedNumber(
				43789,
				"m/d/yyyy"
			);
			return formattedNumber;
		}
		case 5: {
			var formattedNumber = new CustomFunctions.FormattedNumber(
				45678,
				"[$-x-systime]h:mm:ss AM/PM"
			);
			return formattedNumber;
		}
		case 6: {
			var formattedNumber = new CustomFunctions.FormattedNumber(
				12345678,
				"0.00E+00"
			);
			return formattedNumber;
		}
		case 7: {
			var formattedNumber = new CustomFunctions.FormattedNumber(
				1.23,
				"0.00%"
			);
			return formattedNumber;
		}
		default: {
			return "Invalid test case";
		}
	}
}

exports.SetFormattedNumber = SetFormattedNumber;

function SetEntity(entityCase) {
	switch (entityCase) {
		case 1: {
			var properties = {
				"TestString": {
					type: "String",
					basicValue: "Test"
				},
				"TestDouble":{
					type: "Double",
					basicValue: 1
				}
			};
			var Entity = new CustomFunctions.Entity("Basic Entity", properties);
			return Entity;
		}

		case 2: {
			var properties = {
				"TestString": {
					type: "String",
					basicValue: "Test"
				},
				"TestDouble":{
					type: "Double",
					basicValue: 1
				},
				"TestError": new CustomFunctions.Error(CustomFunctions.ErrorCode.divisionByZero)
			};
			var Entity = new CustomFunctions.Entity("Entity With simple Error", properties);
			return Entity;
		}

		case 3: {
			var properties = {
				"TestString": {
					type: "String",
					basicValue: "Test"
				},
				"TestDouble":{
					type: "Double",
					basicValue: 1
				},
				"TestFormattedNumber": new CustomFunctions.FormattedNumber(1.234, "0.00")
			};
			var Entity = new CustomFunctions.Entity("Entity With Formatted Number", properties);
			return Entity;
		}
		
		case 4: {
			var nestedProperties = {
				"TestString2": {
					type: "String",
					basicValue: "Test2"
				},
				"TestDouble2":{
					type: "Double",
					basicValue: 2
				},
			};
			
			var parentProperties = {
				"TestString": {
					type: "String",
					basicValue: "Test"
				},
				"TestDouble":{
					type: "Double",
					basicValue: 1
				},
				"TestNestedEntity": new CustomFunctions.Entity("Nested Entity", nestedProperties)
			};
			
			var parentEntity = new CustomFunctions.Entity("Entity With Nested Object", parentProperties);
			return parentEntity;
		}

		case 5: {
			var EmptyEntity = new CustomFunctions.Entity("Entity Without Properties");
			return EmptyEntity;
		}

		case 6: {
			var EmptyEntity = new CustomFunctions.Entity("Entity Issue", {
				"TestString":"Test",
				"TestDouble": 1
			});
			return EmptyEntity;
		}

		case 7: {
			var EntityWithPorpertyMetadata = new CustomFunctions.Entity("Entity With PropertyMetadata", {
				TestString: {
					type: "String",
					basicValue: "Test",
					propertyMetadata: {
						excludeFrom:{
							"autoComplete":false,
							"calcCompare":false,
							"dotNotation":false,
							"cardView":false
						},
						sublabel:"string in entity"
					}
				},
				TestDouble:{
					type: "Double",
					basicValue: 1
				}
			})

			return EntityWithPorpertyMetadata;
		}
		case 8: {
			var EntityWithPorpertyMetadata = new CustomFunctions.Entity("Entity With PropertyMetadata", {
				TestString: {
					type: "String",
					basicValue: "Test",
					propertyMetadata: {
						excludeFrom:{
							"autoComplete":false,
							"calcCompare":false,
							"dotNotation":false,
							"cardView":true
						},
						sublabel:"string in entity"
					}
				},
				TestDouble:{
					type: "Double",
					basicValue: 1
				}
			})

			return EntityWithPorpertyMetadata;
		}
		case 9:{
			return {"type":"Entity","basicType":"Error","basicValue":"#VALUE!","text":"Entity With API","properties":{"TestDouble":{"type":"Double","basicType":"Double","basicValue":1,"propertyMetadata":{"excludeFrom":{"cardView":true}}},"TestString":{"type":"String","basicType":"String","basicValue":"Test"}}}
		}
		case 10:{
			return {"type":"Entity","basicType":"Error","basicValue":"#VALUE!","text":"Supermetrics Entity",
			"properties":{
				"Total spent":{"type":"Double","basicType":"Double","basicValue":1,},
				"Impressions":{"type":"Double","basicType":"Double","basicValue":304626},
				"Total spent":{"type":"Double","basicType":"Double","basicValue":4378},
				"Clicks":{"type":"Double","basicType":"Double","basicValue":880}
			}}
	
		}
		default: {
			return "Invalid test case";
		}
	}
}

exports.SetEntity = SetEntity;

function SetWebImage(imageCase) {
	switch (imageCase) {
		case 1: {
			var Image = new CustomFunctions.WebImage("https://upload.wikimedia.org/wikipedia/commons/3/3a/Cat03.jpg");
			return Image;
		}
		case 2: {
			var Entity = new CustomFunctions.Entity("Animal", {
				Cat: new CustomFunctions.WebImage("https://upload.wikimedia.org/wikipedia/commons/3/3a/Cat03.jpg")
			});
			return Entity;
		}
		case 3: {
			var attribution = new CustomFunctions.Attribution("https://bing.com", "license", "https://bing.com", "source");
			var image = new CustomFunctions.WebImage("https://upload.wikimedia.org/wikipedia/commons/3/3a/Cat03.jpg", undefined, undefined, [attribution]);
			return image;
		}
		case 4: {
			var attribution1 = new CustomFunctions.Attribution("https://bing.com", "licenseA", "https://bing.com", "sourceA");
			var attribution2 = new CustomFunctions.Attribution("https://bing.com", "licenseB", "https://bing.com", "sourceB");
			var provider = new CustomFunctions.Provider("Powered by Power BI", "https://app.powerbi.com/PowerBILogo.png", "https://msit.powerbi.com/home");
			var image = new CustomFunctions.WebImage("https://upload.wikimedia.org/wikipedia/commons/3/3a/Cat03.jpg", "Cat Picture", "https://upload.wikimedia.org/wikipedia/commons/thumb/0/0b/Cat_poster_1.jpg/1920px-Cat_poster_1.jpg", [attribution1, attribution2], provider);
			return image;
		}
		case 5: {
			var provider = new CustomFunctions.Provider("Powered by Selfhost");
			var image = new CustomFunctions.WebImage("https://upload.wikimedia.org/wikipedia/commons/3/3a/Cat03.jpg", undefined, undefined, undefined, provider);
			return image;
		}
		case 6: {
			var attribution = new CustomFunctions.Attribution(undefined, "license");
			var image = new CustomFunctions.WebImage("https://upload.wikimedia.org/wikipedia/commons/3/3a/Cat03.jpg", undefined, undefined, [attribution]);
			return image;
		}
	}
}

exports.SetWebImage = SetWebImage;

function SetArray(entityCase) {
	switch (entityCase) {
		case 1: {
			var entity = new CustomFunctions.Entity("First Entity", {
				"TestString": {
					type: "String",
					basicValue: "Test"
				},
				"TestDouble":{
					type: "Double",
					basicValue: 1
				}
			});
			var formattedNumber = new CustomFunctions.FormattedNumber(
				1.234,
				"0.00"
			);
			var error = new CustomFunctions.Error(CustomFunctions.ErrorCode.divisionByZero);
			return [[entity], [formattedNumber], [error]];
		}

		case 2: {
			var Entity1 = new CustomFunctions.Entity("First Entity", {
				"TestString": {
					type: "String",
					basicValue: "Test"
				},
				"TestDouble":{
					type: "Double",
					basicValue: 1
				}
			});
			var Entity2 = new CustomFunctions.Entity("Second Entity", {
				"TestString": {
					type: "String",
					basicValue: "Test2"
				},
				"TestDouble":{
					type: "Double",
					basicValue: 2
				}
			});
			return [
				["First", Entity1],
				["Second", Entity2]
			];
		}

		default: {
			return "Invalid entityCase"
		}
	}
}

exports.SetArray = SetArray

function getRichData(value, attribute) {
	// return JSON.stringify(value);
	if (value.type == CustomFunctions.Entity.valueType) {
		if (attribute == "text")
			return value.text;
		else
		{
			return value.properties[attribute].basicValue;
		}
	}
	else if(value.type == CustomFunctions.FormattedNumber.valueType) {
		return value[attribute];
	}
	else if(value.type == CustomFunctions.Error.valueType) {
		return value[attribute];
	}
	else if (value.type == CustomFunctions.WebImage.valueType) {
		if (attribute == "attribution" || attribute == "provider")
			return JSON.stringify(value[attribute]);
		else
			return value[attribute];
	}
	else {
		return "no richData detected";
	}
}

exports.getRichData = getRichData

function getRichDataArray(value) {
	var ret = [];
	var oneRow = [];
	if (value instanceof Array)
	{
		for (var i = 0; i < value.length; ++i) {
			for (var j = 0; j < value[i].length; ++j) {
				var item = value[i][j];
				if (item.type == CustomFunctions.Entity.valueType) {
					oneRow.push(item.text);
				}
				else if(item.type == CustomFunctions.FormattedNumber.valueType) {
					oneRow.push(item.basicValue);
				}
				else if(item.type == CustomFunctions.Error.valueType) {
					oneRow.push(item.basicValue);
				}
				else if (item.type == CustomFunctions.WebImage.valueType) {
					oneRow.push(item.address);
				}
				else {
					oneRow.push(item);
				}
			}
		}
	}

	ret.push(oneRow);
	return ret;
}

exports.getRichDataArray = getRichDataArray;

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
exports.logMessage = logMessage;

CustomFunctions.associate("ADD", add);
CustomFunctions.associate("ADD400", add400);
CustomFunctions.associate("CLOCK", clock);
CustomFunctions.associate("INCREMENT", increment);
CustomFunctions.associate("LOG", logMessage);
CustomFunctions.associate("customErrorOut", customErrorOut);
CustomFunctions.associate("customErrorReturn", customErrorReturn);
CustomFunctions.associate("customErrorReturnArray", customErrorReturnArray);
CustomFunctions.associate("customErrorInput", customErrorInput);
CustomFunctions.associate("customErrorInputInvalid", customErrorInputInvalid);
CustomFunctions.associate("customErrorInputArray", customErrorInputArray);
CustomFunctions.associate("GetParameterAddresses", GetParameterAddresses);
CustomFunctions.associate("GetParameterAddressesRepeating", GetParameterAddressesRepeating);
CustomFunctions.associate("GetParameterAddressesOptional", GetParameterAddressesOptional);
CustomFunctions.associate("GetParameterAddressesRange", GetParameterAddressesRange);
CustomFunctions.associate("GetParameterAddressesFalse", GetParameterAddressesFalse);
CustomFunctions.associate("GetParameterAddressesOff", GetParameterAddressesOff);
CustomFunctions.associate("customErrorInputInvalidArray", customErrorInputInvalidArray);
CustomFunctions.associate("customErrorInput2", customErrorInput2);
CustomFunctions.associate("SetRichError", SetRichError);
CustomFunctions.associate("SetFormattedNumber", SetFormattedNumber);
CustomFunctions.associate("SetEntity", SetEntity);
CustomFunctions.associate("SetWebImage", SetWebImage);
CustomFunctions.associate("SetArray", SetArray);
CustomFunctions.associate("getRichData", getRichData);
CustomFunctions.associate("getRichDataArray", getRichDataArray);
/***/ })

/******/ });
//# sourceMappingURL=functions.js.map
