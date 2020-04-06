#!/usr/bin/env node
let fs = require('fs')	
let path = require('path')
let xlsx = require('xlsx')
let dirpath = process.argv[2]
if (dirpath == undefined) {
  dirpath = "."
}
function getLastChar(str) {
	let len = str.length
	if(!len || len < 1) {
		return -1
	}
	for(let i = len - 1; i >= 0; --i) {
		let ch = str[i]
		if((ch >= 'a' && ch <= 'z') || 
			(ch >= 'A' && ch <= 'Z'))
			{
				return i
			}
	}
	return -1
}

function findExcelFile(path) {
  const files = fs.readdirSync(path)
  const excelfile = []
  files.forEach(function(file){
    if (file.indexOf('.xlsx') >= 0 && file.indexOf('~$') == -1) {
      excelfile.push(file)
    }
	})
  return excelfile
}

function clearDir(dir) {
  if( fs.existsSync(dir)) {
    fs.readdirSync(dir).forEach(function(file) {
      var curPath = path.join(dir, file)
      if(fs.statSync(curPath).isDirectory()) {
        clearDir(curPath)
      } else {
        fs.unlinkSync(curPath)
      }
  })
  fs.rmdirSync(dir)
  }
}

function createOutputDir(dir) {
	dir = path.join(dir, 'output')
	if (fs.existsSync(dir)) {
		clearDir(dir)
	}
	fs.mkdirSync(dir)
}

function createClass(workbook, sheetname) {
  var worksheet = workbook.Sheets[sheetname]
  let c = []
  for(z in worksheet) {
    if(z[0] === '!'){
      delete worksheet[z]
      continue
    }
    var lastChar = getLastChar(z)
		if(lastChar == -1) {
			console.log('error data', sheetname)
		}
    var row = parseInt(z.substring(lastChar + 1))
    var value = worksheet[z].v;
    if(row == 1) {
      c.push(value)
      delete worksheet[z]
		} else {
			break
		}
  }
  return c
}

function createObj(C, workbook, sheetname) {
	var worksheet = workbook.Sheets[sheetname]
	var objlist = []
	var currentrow = 2
	let values = []
	for(z in worksheet) {
		if(z[0] === '!') {
			// 处理最后一行
			if (values.length > 0) {
				if (C.length != values.length) throw new Error(sheetname + " at Line :" + row)
				let obj = {}
				for(let i = 0; i < values.length; i++) {
					obj[C[i]] = values[i]
				}
				values = []
				objlist.push(obj)
			}
			continue
		}  
		var lastChar = getLastChar(z)
		if(lastChar == -1) {
			console.log('error data', sheetname)
		}
		var row = parseInt(z.substring(lastChar + 1))
		if (row != currentrow) {
			currentrow = row
      let obj = {}
      for(let i = 0; i < values.length; i++) {
        obj[C[i]] = values[i]
      }
			values = []
      objlist.push(obj)
		}
		var value = worksheet[z].v
		values.push(value)	
	}
	return objlist
}

function genJSConfig(jsobj, filename, configData) {
	let s = ''
	let left, right, gap
	if (jsobj.length == 1) {
		left = ''
		right = ''
		gap = '\n    '
	} else {
		left = '\n    {'
		right = '\n    },'
		gap = '\n      '
	}

	for(let i = 0; i < jsobj.length; i++) {
		s += left
		for (var j in jsobj[i]) {
			s += gap
			let value = jsobj[i][j]
			if (typeof value == 'string' && value != 'false' && value != 'true') {
				value = '\'' + jsobj[i][j] + '\''
			}
			s += j + ': ' + value +','
		}
		s += right
	}
	
	if (jsobj.length == 1) {
		return `${configData}  ${filename}: {${s}\n  },\n`
	} else {
		return `${configData}  ${filename}: [${s}\n  ],\n`
	}
}

function genConfig(dirpath, files) {
  var newfilepath = path.join(dirpath, 'output', 'Config.js')
  let writerStream =fs.createWriteStream(newfilepath)
  let configData = 'const Conf = {\n'
  let i
  for(i = 0; i < files.length; i++) {
    let workbook = xlsx.readFile(path.join(dirpath, files[i]), {cellStyles:true, bookFiles:true})
    let sheetNames = workbook.SheetNames;
    let C = createClass(workbook, sheetNames[0])
    let objs = createObj(C, workbook, sheetNames[0])
    if (objs.length > 0) {
      configData = genJSConfig(objs, sheetNames[0], configData)
    }
  }
  configData+='}\nexport default Conf'
  writerStream.write(configData, 'utf-8')
  writerStream.end()
  console.log('Config.js is finished')
}

let files = findExcelFile(dirpath)
createOutputDir(dirpath)
genConfig(dirpath, files)




