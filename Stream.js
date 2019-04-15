const xlsx = require('node-xlsx');
const fs = require('fs');

var now = new Date();
var year = now.getFullYear();
var month = now.getMonth()+1;
var week = "周"+now.getDay();

const timeDetail = formatTime(now);
const timeDate = formatDate(now);

const fileName = `${year}年下班时间统计`;
const sheetName = year+'年'+month+'月';

var data;
var preData=null;
let index=-1;
try{
	preData = xlsx.parse(`./${fileName}.xlsx`);
	index = mapParamHasVal(preData,'name',sheetName);
}catch(e){
	console.log(e)
}
if(preData){
	data = preData;
	if(index===-1){ // 新的月份
		let obj = {
			name : sheetName,
			data : [
				[
					'日期',
					'周',
					'下班时间',
					'下班时间转秒数',
				],
				[
					timeDate,
					week,
					timeDetail.time,
					timeDetail.seconds
				],
			]
		}
		data.push(obj)
	}else{
		let arr = data[index].data;
		for(let i=0;i<=arr.length;i++){
			let item = arr[i];
			if(item){
				if(item[0]===timeDate || item[0]==='平均下班时间：' || item.length===0){
					arr.splice(i,1);
					i--;
				}
			}
			
		}
		data[index].data.push([timeDate,week,timeDetail.time,timeDetail.seconds])
	}
}else{
	data = [
	    {
	        name : year+'年'+month+'月',
	        data : [
	            [
					'日期',
	                '周',
	                '下班时间',
					'下班时间转秒数',
	            ],
	            [
	                timeDate,
	                week,
	                timeDetail.time,
					timeDetail.seconds
	            ],
	        ]
	    },
	]
}
let sheetIndex = mapParamHasVal(data,'name',sheetName);
let s=0;
let average,averageTime;
data[sheetIndex].data.forEach((item,index)=>{
	if(index==0 || item === undefined) return;
	s+=item[3];
})
let length = data[sheetIndex].data.length-1;
average = s/length;
averageTime = seconds2Time(average);
data[sheetIndex].data.push([]);
data[sheetIndex].data.push(['平均下班时间：',averageTime]);

// 写xlsx
var buffer = xlsx.build(data);
fs.writeFile(`./${fileName}.xlsx`, buffer, function (err)
{
    if (err){
		// 创建一个可以写入的流，写入到文件 output.txt 中
		var writerStream = fs.createWriteStream('./output.txt');
		// 使用 utf8 编码写入数据
		writerStream.write(err,'UTF8');
		writerStream.end();
		throw err;
	}
	
}
);

function seconds2Time(seconds){
	let s = seconds%60;
	let m = (seconds-s)/60%60;
	let h = Math.floor(seconds/60/60);
	
	if(h < 10){
		h = '0' + h;
	}
	if(m < 10){
		m = '0' + m;
	}
	if(s < 10){
		s = '0' + s;
	}
	
	return (h + ':' + m+ ':' + s)
}

function formatTime(date){//格式化日期 hh:mm:ss
	let hour = date.getHours();
	let minute = date.getMinutes();
	let second = date.getSeconds();
	let seconds = (hour*60 + minute)*60 + second;
	
	if(hour < 10){
		hour = '0' + hour;
	}
	if(minute < 10){
		minute = '0' + minute;
	}
	if(second < 10){
		second = '0' + second;
	}
	
	
	return {
		time : (hour + ':' + minute+ ':' + second),
		seconds : seconds
	};
}
function formatDate(date){//格式化日期 YYYY-MM-DD
	let year = date.getFullYear();
	let month = date.getMonth()+1;
	let weekday = date.getDate();

	if(month < 10){
		month = "0" + month;
	}
	if(weekday < 10){
		weekday = "0" + weekday;
	}
	return (year+"/"+month + "/" + weekday);
}
function mapParamHasVal(map, param, val) { // 判断数组对象中的某个参数书否有这个值 没有返回-1，有返回索引
	if (!Array.isArray(map)) return -1;
	let index = map.findIndex(item => {
		if (item && item[param] !== undefined) {
			return item[param] == val;
		} else {
			return false;
		}
	})
	return index;
}