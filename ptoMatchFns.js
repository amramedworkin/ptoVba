function ClearUsptoDotGov(table) {
	// So we can test as we go
	if (typeof table === 'undefined') { return }
	const searchRegExp = /\.uspto.gov/gi;
	const replaceWith = '';

	table.forEach(row => {
		for (var prop in row) {
			if (Object.prototype.hasOwnProperty.call(row, prop)) {
				if (typeof prop === 'string' && prop.toLowerCase().match(/\.uspto.gov/i) )
				{
					prop = prop.replace(searchRegExp,replaceWith)
				}
			}
		}
	})
}
function ExactCyberToGEARSServerMatch(tables) {
	var ecmoTable = tables['ecmo']
	var gearsTable = tables['gears']

	console.time('ExactCyberToGEARSServerMatch')
	ecmoTable.forEach(ecmoRow => {
		gearsTable.forEach(gearsRow => {
			if(ecmoRow._key == gearsRow._key) {
				ecmoRow['GEARS EXACT MATCH'] = gearsRow._key
			}
		})
	})
	console.log('ExactCyberToGEARSServerMath=%d matched',ecmoTable.filter(row=>{return row['GEARS EXACT MATCH'] != null}).length)
	console.timeEnd('ExactCyberToGEARSServerMatch')
}

function IdentifyType(tables) {
	var ecmoTable = tables['ecmo']
	console.time('IdentifyType')
	// var today = new Date().setHours(0,0,0,0)
	// Get Max Date
	var maxLastResponse = -Infinity
	ecmoTable.map(
			row => new Date(row['Last Report Time'])
				.setHours(0,0,0,0))
				.forEach(lastResponse => { 
		if(lastResponse > maxLastResponse) { maxLastResponse = lastResponse }
	})
	// var maxLastResponse = Math.max.apply(maxLastResponses)
	// eslint-disable-next-line no-unused-vars
	var today = new Date().setHours(0,0,0,0)
	ecmoTable.forEach(ecmoRow => {
		if( ecmoRow.Relay.match(/dev-/i)) {
			ecmoRow.Label = 'dev-dev' 
		}
		else if( ecmoRow.OS.match(/win10/i)) { 
			ecmoRow.Label = 'dev-os'
		}
		else if( new Date(ecmoRow['Last Report Time']).setHours(0,0,0,0) != maxLastResponse) { 
			ecmoRow.Label = 'dev-date' 
		}
		else if(ecmoRow['GEARS EXACT MATCH']) { 
			ecmoRow.Label = 'Exact Match'
		}
	})
	console.log('IdentifyType-dev-dev=%d matched',ecmoTable.filter(row=>{return row.Label == 'dev-dev'}).length)
	console.log('IdentifyType-dev-os=%d matched',ecmoTable.filter(row=>{return row.Label == 'dev-os'}).length)
	console.log('IdentifyType-dev-date=%d matched',ecmoTable.filter(row=>{return row.Label == 'dev-date'}).length)
	console.log('IdentifyType-exactmatch=%d matched',ecmoTable.filter(row=>{return row.Label == 'Exact Match'}).length)
	console.log('IdentifyType-unmatched',ecmoTable.filter(row=>{return row.Label == null}).length)
	// Convert back from counting temp values to expected values
	// tables['ecmo'].forEach(row => { 
	// 	if (['dev-dev','dev-os','dev-date'].includes(row.Label)) { row.Label = 'dev' } 
	// })
	console.timeEnd('IdentifyType')
}

function MoveCyberNonProd(tables) {
	console.time('MoveCyberNonProd')
	var ecmoTable = tables['ecmo']
	console.log('MoveCyberNonProd=%d ON ENTRY',ecmoTable.length)
	tables['dev'] = ecmoTable.filter(ecmoRow => { return ["dev-dev","dev-os","dev-date"].includes(ecmoRow.Label) })
	tables['exactmatch'] = ecmoTable.filter(ecmoRow => { return ecmoRow.Label === 'Exact Match' })
	tables['ecmo'] = ecmoTable.filter(ecmoRow => { return ecmoRow.Label === null })
	console.log('MoveCyberNonProd-DEV=%d matched',tables['dev'].length)
	console.log('MoveCyberNonProd-EXACTMATCH=%d matched',tables['exactmatch'].length)
	console.log('MoveCyberNonProd-ECMO=%d AFTER REMOVE',ecmoTable.length)
	console.timeEnd('MoveCyberNonProd')
}

function MatchOnSDAP(tables) {
	var ecmoTable = tables['ecmo']
	var sdapTable = tables['sdap']
	console.time('MatchOnSDAP')
	ecmoTable.forEach(ecmoRow => {
		sdapTable.forEach(sdapRow => {
			if(ecmoRow._key == sdapRow._key) {
				if (sdapRow['ADDED: Component Name']) {
					ecmoRow['SDAP EXACT MATCH'] = 
						(sdapRow['ADDED: Component Name'] || '')
							.toUpperCase()
							.trim()
				}
				
			}
		})
	})
	console.log('MatchOnSDAP=%d matched',ecmoTable.filter(row=>{return row['SDAP EXACT MATCH'] != null}).length)
	console.timeEnd('MatchOnSDAP')
}

function CheckManual(tables) {
	var ecmoTable = tables['ecmo']
	var manualTable = tables['manual']
	console.time('CheckManual')
	ecmoTable.forEach(ecmoRow => {
		manualTable.forEach(manualRow => {
			if(ecmoRow._key == manualRow._key) {
				if ( manualRow['Component'] ) {
					ecmoRow['Component Manual'] = 
						(manualRow['Component'] || '')
							.toUpperCase()
							.trim()
				}
				
			}
		})
	})
	console.log('CheckManual=%d matched',ecmoTable.filter(row=>{return row['Component Manual'] != null}).length)
	console.timeEnd('CheckManual')
}

function DiamondCheck(tables) {
	var ecmoTable = tables['ecmo']
	var diamondTable = tables['diamond']
	console.time('DiamondCheck')
	ecmoTable.forEach(ecmoRow => {
		diamondTable.forEach(diamondRow => {
			if(ecmoRow._key == diamondRow._key) {
				if ( diamondRow['GEARS Component'] ) {
					ecmoRow['Diamond Lookup'] =
						(diamondRow['GEARS Component'] || '')
							.toUpperCase()
							.trim()
				}
				
			}
		})
	})
	console.log('DiamondCheck=%d matched',ecmoTable.filter(row=>{return row['Diamond Lookup'] != null}).length)
	console.timeEnd('DiamondCheck')
}

module.exports.ExactCyberToGEARSServerMatch = ExactCyberToGEARSServerMatch
module.exports.IdentifyType = IdentifyType
module.exports.MoveCyberNonProd = MoveCyberNonProd
module.exports.MatchOnSDAP = MatchOnSDAP
module.exports.CheckManual = CheckManual
module.exports.DiamondCheck = DiamondCheck
module.exports.ClearUsptoDotGov = ClearUsptoDotGov