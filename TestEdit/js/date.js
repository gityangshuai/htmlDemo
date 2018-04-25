$(function() {
	'use strict';
	// 对Date的扩展，将 Date 转化为指定格式的String
	// 月(M)、日(d)、小时(h)、分(m)、秒(s)、季度(q) 可以用 1-2 个占位符， 
	// 年(y)可以用 1-4 个占位符，毫秒(S)只能用 1 个占位符(是 1-3 位的数字) 
	// 例子： 
	// new Date(time).format("yyyy-MM-dd hh:mm:ss");==>2017-04-26 10:20:25
	Date.prototype.format = function(fmt) { 
	     var o = { 
	        "M+" : this.getMonth()+1,                 //月份 
	        "d+" : this.getDate(),                    //日 
	        "h+" : this.getHours(),                   //小时 
	        "m+" : this.getMinutes(),                 //分 
	        "s+" : this.getSeconds(),                 //秒 
	        "q+" : Math.floor((this.getMonth()+3)/3), //季度 
	        "S"  : this.getMilliseconds()             //毫秒 
	    }; 
	    if(/(y+)/.test(fmt)) {
	            fmt=fmt.replace(RegExp.$1, (this.getFullYear()+"").substr(4 - RegExp.$1.length)); 
	    }
	     for(var k in o) {
	        if(new RegExp("("+ k +")").test(fmt)){
	             fmt = fmt.replace(RegExp.$1, (RegExp.$1.length==1) ? (o[k]) : (("00"+ o[k]).substr((""+ o[k]).length)));
	         }
	     }
	    return fmt; 
	};
	/**
	 * 传入yyyyMMddHHmmss,返回类型     返回yyyy或mm或dd或hh或mi或ss或 yyyy-mm-dd hh:mi:ss
	 */
	var format2dd = function(dd,ret){
		var ss="";
		if(dd!=null&&dd!=""){
			var year = dd.substring(0,4);
			var moth = dd.substring(4,6);
			var day = dd.substring(6,8);
			var shi = dd.substring(8,10);
			var fen = dd.substring(10,12);
			var miao = dd.substring(12,14);
			if(ret=="yyyy"){
				ss = year;
			}else if(ret=="mm"){
				ss = moth;
			}else if(ret=="dd"){
				ss = day;
			}else if(ret=="hh"){
				ss = shi;
			}else if(ret=="mi"){
				ss = fen;
			}else if(ret=="ss"){
				ss = miao;
			}else{
				ss = year+"-"+moth+"-"+day+" "+shi+":"+fen+":"+miao;
			}
		}
		return ss;
	};
	var dd2format = function(ss) {
		var ret = "";
		if(ss!=""&&ss!=null){
			ret = ss.replace("-","").replace(" ","").replace(":","");
		}
		return ret;
	}
});