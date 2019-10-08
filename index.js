/**
 * @description excel批量合成工具
 * @author cmyh100
 */


// 公共库
const xlsx = require('node-xlsx')
const fs = require('fs')
// excel文件夹路径（把要合并的文件放在excel文件夹内）
const _file = `${__dirname}/excel/`
const _output = `${__dirname}/result/`
// 合并数据的结果集
let dataList = [
	{
		name: '提交成绩',
		data: []
	}]

init()
function init () {
	fs.readdir(_file, function(err, files) {
		console.log(files)
		if (err) {
			throw err
		}
		// files是一个数组
		// 每个元素是此目录下的文件或文件夹的名称
		// console.log(`${files}`);
		files.forEach((item, index) => {
			try {
				// console.log(`${_file}${item}`)
				console.log(`开始合并：${item}`)
				let excelData = xlsx.parse(`${_file}${item}`)
				console.log(excelData[0].data)

/*				var title=excelData[0].data.splice(0, 1)
			var myData = excelData[0].data
				console.log(myData[0])
				myData.forEach(function (value,index) {
					var count=0;
					myData.forEach(function (val,ind) {
						if(value[0]===val[0]){
							count++;
						}
                    })
					if(count>=2){
						myData[index].push(count)
					}
                })
				console.log(myData)*/
			if (excelData) {
					if (dataList[0].data.length > 0) {
						excelData[0].data.splice(0, 1)
                        //excelData[1].data.splice(0, 1)
                        //console.log(excelData[1].data)
					}
                	//dataList[0].data = myData;
					dataList[0].data = dataList[0].data.concat(excelData[0].data)
					//dataList[1].data = dataList[1].data.concat(excelData[1].data)

				}
			} catch (e) {
				console.log('excel表格内部字段不一致，请检查后再合并。')
			}
		})
		// 写xlsx
		var buffer = xlsx.build(dataList)
		fs.writeFile(`${_output}resut.${new Date().getTime()}.xlsx`, buffer, function (err) {
			if (err) {
				throw err
			}
			console.log('\x1B[33m%s\x1b[0m', `完成合并：${_output}resut.${new Date().getTime()}.xlsx`)
		})
	})
}