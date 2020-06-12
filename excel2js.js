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

function newObj(c, values) {
  const obj = {}
  for(let i = 0; i < c.length; i++) {
    let key = c[i] 
    let value = values[i]
    if (!value) {
      value = ''
    }
    let k
    let v
    if (key[key.length - 1] == ']') {
      let i = key.indexOf('[')
      if (i == -1) throw new Error('can not found [')
      k = key.substring(0, i)
      for(let j = i + 1; j < key.length -1; j++) {
        let t = key[j]
        switch (t) {
          case 's':
            v = String(value)
            break
          case 'n':
            if (value == '') {
              value = 0
            }
            v = Number(value)
            break
          case 'a':
            if (value == '') {
              v = '[]'
            } else {
              v = '[' + value + ']'
            }
            break
        }
      }
    } else {
      k = key
      v = value
    }
    obj[k] = v
  }
  return obj
}

function createObj(c, workbook, sheetname) {
	var worksheet = workbook.Sheets[sheetname]
	var objlist = []
	var currentrow = 2
	let values = []
	for(z in worksheet) {
		if(z[0] === '!') {
			// 处理最后一行
			if (values.length > 0) {
				let obj = newObj(c, values)
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
      let obj = newObj(c, values)
			values = []
      objlist.push(obj)
		}
    const value = worksheet[z].v
		values.push(value)	
	}
	return objlist
}



function genJSConfig(jsobjs, filename, configData) {
  let s = ''
  let left, right, gap
  if (jsobjs.length == 1) {
    left = ''
    right = ''
    gap = '\n    '
  } else {
    left = '\n    {'
    right = '\n    },'
    gap = '\n      '
  }
	for(let i = 0; i < jsobjs.length; i++) {
    const jsobj = jsobjs[i]
    const keys = Object.keys(jsobj)
    if (keys.length == 1 && keys[0] == 'array') {
      const value = `\n    '${jsobj.array}',`
      s += value
    } else {
      s += left
      for (let j = 0; j < keys.length; j++) {
        s += gap
        const key = keys[j]
        let value = jsobj[key]
        if (typeof value == 'string' && value != 'false' && value != 'true' && value[0] != '[') {
          value = '\'' + jsobj[key] + '\''
        }
        s += `${key}: ${value},`
      }
      s += right
    }
	}
	
	if (jsobjs.length == 1) {
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
    let workbook = xlsx.readFile(path.join(dirpath, files[i]), {cellStyles:true, bookFiles:true, sheetStubs: true})
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




