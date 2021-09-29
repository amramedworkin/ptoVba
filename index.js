const fs = require('fs')
const path = require('path')
const ptoUtils = require('./ptoUtils')
const ptoFindServers = require('./FindServersNSLookup')
const ptoFindComponents = require('./FindComponentinPML')
const ptoMatch = require('./ptoMatchFns')
const ptoServerSw = require('./ServerSWUpdates')

String.prototype.left = function(n) {
    return this.substring(0, n);
}
String.prototype.removeNum = function() {
	return this.replace(/\d+/g, '')
}
String.prototype.trimRight = function(n) {
	return this.replace(new RegExp("["+n+"]+$"), '')
}
String.prototype.trimLeft = function(n) {
	return this.replace(new RegExp("^["+n+"]+"), '')
}

console.log('====================================================')
console.log('OCTO SP Site Clean - Initiated')
console.log('----------------------------------------------------')

// List of inputs
// - file = excel file (will have to change for SP list read)
// - sheet = sheet in excel workbook containing data
// - column = column on worksheet containing computer or server value
// - attribute = what group of data is this (used to reference list)
// - component:name - property containing Component name (optional)
// - component:id - property containing Component id (optional)
const inputs = [
	{file: './data/test/input/NewEcmoReport.xlsx', sheet: 'ECMO', outSheet: 'Cyber', column: 'Computer Name',attribute: 'ecmo',component: { name: '', id: ''}, newfields: ["Label" ,"SDAP EXACT MATCH" ,"GEARS EXACT MATCH" ,"Component Manual" ,"Component from PML" ,"Component from GEARS" ,"LOGIC Match" ,"Diamond lookup" ,"Final Combined" ,"Software Name" ,"Software ID" ,"Database" ,"Server" ,"NSLookup"]}
	,{file: './data/test/input/NewGears.xlsx', sheet: 'GEARS', column: 'Server',attribute: 'gears',component: { name: '', id: 'ComponentLookup:ComponentID'} }
	,{file: './data/sdap.xlsx', sheet: 'SDAP', column: 'Hostname',attribute: 'sdap',component: { name: 'ADDED: Component Name', id: 'ADDED: Component ID'}}
	,{file: './data/manual.xlsx', sheet: 'Manual', column: 'Server Name',attribute: 'manual',component: { name: 'Component', id: ''} }
	,{file: './data/test/input/NewDiamond.xlsx', sheet: 'Diamond', column: 'Hostname',attribute: 'diamond',component: { name: 'GEARS Component', id: 'Component ID'}, specialConditioning: "key.replace(String.fromCharCode(160),' ')"}
	,{file: './data/pml.xlsx', sheet: 'PML', column: 'ComponentAcronym',attribute: 'pml',component: { name: 'ComponentAcronym', id: 'ID'}}
	,{file: './data/test/input/NewServers.xlsx', sheet: 'Server', column: 'ServerName',attribute: 'server'}
	,{attribute: 'comparens', sheet: 'CompareNSLookupServers'}
	,{attribute: 'cyberlist', sheet: 'CyberListfromCNSLookupS'}
	,{attribute: 'serversw',sheet: 'ServerSWUpdates'}
	,{attribute: 'dev',sheet: 'Dev'}
	,{attribute: 'exactmatch',sheet: 'ExactMatch'}
]

console.time('MAIN PROGRAM')

// Read all excel tables
var tables = ptoUtils.getAllTables(inputs);

// ptoCheckAndMatch.:ExactCyberToGEARSServerMatch()
try
{
	ptoMatch.ClearUsptoDotGov(tables['ecmo'])
	ptoMatch.ClearUsptoDotGov(tables['gears'])
	// ptoMatch.ClearUsptoDotGov(tables['sdap'])
	ptoMatch.ExactCyberToGEARSServerMatch(tables)
	// writeNewWorkbook('NODE-exactcybertogearsservermatch',inputs,tables,['ecmo'],true)	
	ptoMatch.IdentifyType(tables)
	// writeNewWorkbook('NODE-identifytype',inputs,tables,['ecmo','gears'],true)	
	ptoMatch.MoveCyberNonProd(tables)
	// writeNewWorkbook('NODE-movecybernonprod',inputs,tables,['ecmo','dev','exactmatch'],true)	
	ptoMatch.MatchOnSDAP(tables)
	// writeNewWorkbook('NODE-matchonsdap',inputs,tables,['ecmo','sdap'],true)	
	ptoMatch.CheckManual(tables)
	// writeNewWorkbook('NODE-checkmanual',inputs,tables,['ecmo','manual'],true)	
	ptoMatch.DiamondCheck(tables)
	// writeNewWorkbook('NODE-DiamondCheck',inputs,tables,['ecmo','diamond'],true)	
	ptoFindComponents.FindComponentinPML(tables)
	ptoServerSw.ServerSWUpdates(tables)
	writeNewWorkbook('NODE-serverswupdates',inputs,tables,['ecmo','server','serversw'])
	process.exit()
	ptoFindServers.FindServersNSLookup(tables)
	writeNewWorkbook('comparens',inputs,tables,'*')
} catch(err) {
	console.error(2,err)
	console.error(2,err.stack)
} finally {
	console.timeEnd('MAIN PROGRAM')
}

function writeNewWorkbook(prefix,inputs,tables,includeList,force=false) {
	var writeAr = []
	var folderPath = './data/test/output/'
	// Delete all prior versions of this prefixed sheet
	if (force) {
		var filesToDelete = fs.readdirSync(folderPath).filter(fn => fn.startsWith(prefix));
		filesToDelete.forEach(fileToDelete => {
			var filePath = path.join(folderPath,fileToDelete)
			console.log("Deleting " + filePath)
			fs.unlinkSync(filePath)
		})
	}
	inputs.forEach(input => {
		if((typeof includeList === 'string' 
					&& (includeList == '*' || includeList == input.attribute))
				|| (Array.isArray(includeList) 
					&& includeList.includes(input.attribute))) {
			var sheetName = input.outSheet || input.sheet
			writeAr.push({ name: sheetName, data: tables[input.attribute]})
		}
	})
	ptoUtils.writeWorkbook(prefix,writeAr)
}