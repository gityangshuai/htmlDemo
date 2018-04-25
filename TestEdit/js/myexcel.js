/*jshint browser:true */  
/*global XLSX */  
var X = XLSX;  
var XW = {  
  /* worker message */  
  msg: 'xlsx',  
  /* worker scripts */  
  rABS: '../js/xlsxworker2.js',  
  norABS: '../js/xlsxworker1.js',  
  noxfer: '../js/xlsxworker.js'  
};  
  
function ab2str(data) {  
  var o = "", l = 0, w = 10240;  
  for(; l<data.byteLength/w; ++l) o+=String.fromCharCode.apply(null,new Uint16Array(data.slice(l*w,l*w+w)));  
  o+=String.fromCharCode.apply(null, new Uint16Array(data.slice(l*w)));  
  return o;  
}  
  
function s2ab(s) {  
  var b = new ArrayBuffer(s.length*2), v = new Uint16Array(b);  
  for (var i=0; i != s.length; ++i) v[i] = s.charCodeAt(i);  
  return [v, b];  
}  
  
function xw_xfer(data, cb) {  
  var worker = new Worker(rABS ? XW.rABS : XW.norABS);  
  worker.onmessage = function(e) {  
    switch(e.data.t) {  
      case 'ready': break;  
      case 'e': console.error(e.data.d); break;  
      default: xx=ab2str(e.data).replace(/\n/g,"\\n").replace(/\r/g,"\\r");cb(JSON.parse(xx)); break;  
    }  
  };  
  if(rABS) {  
    var val = s2ab(data);  
    worker.postMessage(val[1], [val[1]]);  
  } else {  
    worker.postMessage(data, [data]);  
  }  
}  
  
function xw(data, cb) {  
  transferable = document.getElementsByName("xferable")[0].checked;  
  if(transferable) xw_xfer(data, cb);  
  else xw_noxfer(data, cb);  
}  
  
function get_radio_value( radioName ) {  
  var radios = document.getElementsByName( radioName );  
  for( var i = 0; i < radios.length; i++ ) {  
    if( radios[i].checked || radios.length === 1 ) {  
      return radios[i].value;  
    }  
  }  
}  
  
function to_json(workbook) {  
  var result = {};  
  workbook.SheetNames.forEach(function(sheetName) {  
    var roa = X.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);  
    if(roa.length > 0){  
      result[sheetName] = roa;  
    }  
  });  
  return result;  
}  
  
  
  
  
var tarea = document.getElementById('b64data');  
function b64it() {  
  if(typeof console !== 'undefined') console.log("onload", new Date());  
  var wb = X.read(tarea.value, {type: 'base64',WTF:wtf_mode});  
  process_wb(wb);  
}  
 var token="sdfsdf";//Cookies.get("token");  
  console.log(token);  
  if(token==null){  
    // alert("你是怎么进来的？请先登录");  
    // window.location.href="../login.html";  
  }  
var global_wb;  
function process_wb(wb) {  
  global_wb = wb;  
  var output = "";
  //这里是成绩批量录入代码，可以忽略不计  
  $(".submit_all").on("click",function(){  
    var array=Object.values(to_json(wb))[0];  
    var len=array.length;  
    var array1="",array2="",array3="",array4="",array5="";  
    for(var i=0;i<len;i++){  
      array1+=array[i].学号+",";  
      array2+=array[i].成绩+",";  
      array3+=array[i].总分+",";  
      array4+=array[i].时间+",";  
      array5+=array[i].评语+",";  
    }  
    array1=array1.substring(0,array1.length-1);  
    array2=array2.substring(0,array2.length-1);  
    array3=array3.substring(0,array3.length-1);  
    array4=array4.substring(0,array4.length-1);  
    array5=array5.substring(0,array5.length-1);  
    console.log(array1);  
    var url='grade/addall';  
    var params={  
      "token":token,  
      "gradeUserId":array1,  
      "gradeCourse":array2,  
      "gradeRank":array3,  
      "gradeRemark":array5,  
      "gradeTime":array4  
    };  
    pullData(url,params,function(res){  
      if(res.state==0){  
        alert("成绩录入成功！");  
        window.location.href="Techrecord.html";  
      }  
      else if(res.state==1){  
        alert("您输入的学号不存在");  
      }  
    });  
  })  
    
  
  output = JSON.stringify(to_json(wb), 2, 2);  
  if(out.innerText === undefined) out.textContent = output;  
  else out.innerText = output;  
  if(typeof console !== 'undefined') console.log(output);  
}  
function setfmt() {if(global_wb) process_wb(global_wb);}  
  
var xlf = document.getElementById('xlf');  
function handleFile(e) {  
  rABS = document.getElementsByName("userabs")[0].checked;  
  use_worker = document.getElementsByName("useworker")[0].checked;  
  var files = e.target.files;  
  var f = files[0];  
  {  
    var reader = new FileReader();  
    var name = f.name;  
    reader.onload = function(e) {  
      if(typeof console !== 'undefined')  
      var data = e.target.result;  
      if(use_worker) {  
        xw(data, process_wb);  
      } else {  
        var wb;  
        if(rABS) {  
          wb = X.read(data, {type: 'binary'});  
        } else {  
          var arr = fixdata(data);  
          wb = X.read(btoa(arr), {type: 'base64'});  
        }  
  
        process_wb(wb);  
      }  
    };  
    if(rABS) reader.readAsBinaryString(f);  
    else reader.readAsArrayBuffer(f);  
  }  
}  
  
if(xlf.addEventListener) xlf.addEventListener('change', handleFile, false);  
