// 公共库
const xlsx = require('node-xlsx')
const fs = require('fs')
// excel文件夹路径（把要合并的文件放在excel文件夹内）
const _file = `${__dirname}/excel/`
const _output = `${__dirname}/result/`

// 排除 sheet
let excludeSheet = ['PMO汇总', '角色（仅筛选使用）']
// 原 excel 表中的数据集
init('项目&人员统计表.xlsx', [mergeData, countFunctionUsers])
// init('项目&人员统计表.xlsx', [mergeData])

async function init (excelName, callbacks) {
	if (!excelName) {
		console.log('需读取的 Excel 名不能为空！')
		for (let callback of callbacks) {
			callback([])
		}
		return false
	}

	fs.readdir(_file, function(err, files) {
		if (err) {
			throw err
		}

		// files是一个数组
		// 每个元素是此目录下的文件或文件夹的名称
		// console.log(`${files}`);
		files.forEach((item, index) => {
			try {
				// console.log(`${_file}${item}`)
				console.log(`开始读取：${item}`)
				let excelData = xlsx.parse(`${_file}${item}`)

				if (files.length===1 || excelName === item) {
					if (excelData && excelData[0] && excelData[0]['data'].length > 0) {
						// console.log(excelData)
						for (let callback of callbacks) {
							callback(excelData)
						}
						return true
					}
				}
			} catch (e) {
				console.log(e)
				console.log('excel表格读取出错，请检查后再试。')
			}
		})
	})
}

// 合并各 sheet 数据
function mergeData(excelData) {
	if (!excelData || excelData.length === 0) {
		console.log('读取不到对应 Excel 表数据！')
		return false
	}

	console.log(`开始处理合并数据`)
	let data = getRoleUserFunctions(excelData)

	writeFile('总数据', data)
}

// 统计 项目-功能-对应人数 并写入 Excel 中
function countFunctionUsers(excelData) {
	if (!excelData || excelData.length === 0) {
		console.log('读取不到对应 Excel 表数据！')
		return false
	}

	console.log(`开始统计项目功能人数`)
	let functionUsers = getFunctionsUsers(excelData)

	writeFile('项目功能人数统计', functionUsers)
}

// 统计 项目-功能-对应人数 并写入 Excel 中
function writeFile(name, data) {
	if (!name || !data || data.length === 0) {
		console.log('需要写的文件名和内容不能为空！')
		return false
	}

	let writeData = [
		{
			name: 'sheet',
			data: data
		}
	]
	// 文件名
	let writeFileName = `${name}.${new Date().getTime()}.xlsx`
	// xlsx 合并参数
	let merges = [] //[{s:{c:0,r:0}, e:{c:0,r:3}}] // c:列，r:行
	// 写入 xlsx
	let buffer = xlsx.build(writeData, {'!merges': merges})
	fs.writeFile(_output+writeFileName, buffer, function (err) {
		if (err) {
			throw err
		}
		console.log('\x1B[33m%s\x1b[0m', `写入完成：${_output}${writeFileName}`)
	})
}

// 获取 项目-功能-对应人数
function getFunctionsUsers(excelData) {
	let functionUsers = [] // [[f1,f2,f3,f4,userNum]]
	for (let sheet of excelData) {
		// 排除 sheet
		if (excludeSheet.indexOf(sheet["name"]) >= 0)
			continue

		let indexRoles = {}	// { index, roleName }
		let roles = []
		let users = []
		let f1,f2,f3,f4
		for (let i=4; i<sheet['data'].length; i++) {
			let row = sheet['data'][i]
			let mf1 = row[0] && row[0].trim()
			let mf2 = row[1] && row[1].trim()
			let mf3 = row[2] && row[2].trim()
			let mf4 = row[3] && row[3].trim()
			if (!mf1 && !mf2 && !mf3 && !mf4) // 最后一列
				continue

			if (mf1) {	// 新的项目，所有模块功能重新赋值
				f1 = mf1
				f2 = mf2
				f3 = mf3
			} else {	// 还是旧的项目
				if (mf2) {	// 新的模块，模块下面的功能分类重新赋值
					f2 = mf2
					f3 = mf3
				} else {	// 旧的模块
					if (mf3)
						f3 = mf3
				}
			}
			f4 = mf4

			if ((i === 4 && '模块' === f2)	// 模块/功能 标题那一列
				|| ('……' === f2 && '……' === f3)) // 最后一列
				continue
			
			let userCount = 0
			for (let ii=0; ii<row.length; ii++) {
				if (row[ii] && row[ii].trim().indexOf('√')>=0) {
					userCount++
				}
			}
			functionUsers.push([f1,f2,f3,f4,userCount])
		}
	}
	return functionUsers
}

// 获取 角色 用户 对应项目模块功能 [{role,user,[f1,f2,f3,f4]}]
function getRoleUserFunctions(excelData) {
	let newData = [
			['','','',''],	// 角色行
			['','','','']	// 用户行
		]
	for (let sheet of excelData) {
		// 排除 sheet
		if (excludeSheet.indexOf(sheet["name"]) >= 0)
			continue

		// 角色所在列范围
		let rolesIndex = {}
		let roles = []
		let users = []
		let f1,f2,f3,f4
		for (let i=0; i<sheet['data'].length; i++) {
			let row = sheet['data'][i]
			if (row.indexOf('角色') === 0) { // 角色
				// roles 最后一个角色后面的空列拿不到，需要用户那边去匹配
				roles = row
			} else if (row.indexOf('人员') === 0) { // 人员
				users = row
				// 每个 sheet 处理一次就行了
				// 因为 roles 最后一个角色后面的空列拿不到，需要用户那边去匹配
				let lastRole = ''
				for (let ri=1; ri<users.length; ri++) {
					if (!users[ri]) continue
					// 处理 roles
					if (roles[ri]) {
						lastRole = roles[ri]
					} else {
						roles[ri] = lastRole
					}
					// 把当前 sheet 角色、用户合并进来
					let endRoleIndex = newData[0].lastIndexOf(lastRole)
					if (endRoleIndex === -1) {	// 原来没有角色
						newData[0].push(lastRole)
						newData[1].push(users[ri])
						// 原来的项目模块功能行对应加一列
						for (let rowIndex=2; rowIndex<newData.length; rowIndex++) {
							newData[rowIndex].push('')
						}
					} else {
						// 需要看下原来这个角色下有没有用户
						let existRoleUser = false
						let beginRoleIndex = newData[0].indexOf(lastRole)
						for (let ui=beginRoleIndex; ui<=endRoleIndex; ui++) {
							if (newData[1][ui] === users[ri]) {
								existRoleUser = true
								break
							}
						}
						if (!existRoleUser) {
							newData[0].splice(endRoleIndex+1, 0, lastRole)
							newData[1].splice(endRoleIndex+1, 0, users[ri])
							// 原来的项目模块功能行对应加一列
							for (let rowIndex=2; rowIndex<newData.length; rowIndex++) {
								newData[rowIndex].splice(endRoleIndex+1, 0, '')
							}
						}
					}
				}
			} else if (i>=4) {
				let mf1 = row[0] && row[0].trim()
				let mf2 = row[1] && row[1].trim()
				let mf3 = row[2] && row[2].trim()
				let mf4 = row[3] && row[3].trim()
				if (!mf1 && !mf2 && !mf3 && !mf4) // 后面的空行
					continue

				if (mf1) {	// 新的项目，所有模块功能重新赋值
					f1 = mf1
					f2 = mf2
					f3 = mf3
				} else {	// 还是旧的项目
					if (mf2) {	// 新的模块，模块下面的功能分类重新赋值
						f2 = mf2
						f3 = mf3
					} else {	// 旧的模块
						if (mf3)
							f3 = mf3
					}
				}
				f4 = mf4


				if ((i === 4 && '模块' === f2)	// 模块/功能 标题那一行
					|| ('……' === f2 && '……' === f3)) // 最后一行
					continue

				// 把当前行 项目-模块-功能 对应的 角色-用户 坐标 筛选出来
				let checkedIndexs = []
				for (let ii=0; ii<row.length; ii++) {
					if (row[ii] && row[ii].indexOf('√')>=0) {
						let role = roles[ii]
						let user = users[ii]
						if (!role || !user){
							console.log(`角色[${ii}]【${role}】人员【${user}】功能【${f1}-${f2}-${f3}-${f4}】`)
						}

						// 先去处对应角色开始和结束index，缩小遍历范围
						let beginRoleIndex = newData[0].indexOf(role)
						let endRoleIndex = newData[0].lastIndexOf(role)
						for (let ui=beginRoleIndex; ui<=endRoleIndex; ui++) {
							if (newData[1][ui] === user) {
								checkedIndexs.push(ui)
							}
						}
					}
				}

				let functionsRow = [f1,f2,f3,f4]
				for (let columnIndex=4; columnIndex<newData[0].length; columnIndex++) {
					if (checkedIndexs.indexOf(columnIndex) === -1) {
						functionsRow.push('')
					} else {
						functionsRow.push('√')
					}
				}
				
				newData.push(functionsRow)
			}
		}
	}

	return newData
}
