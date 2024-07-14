/*
    作者: imoki
    仓库: https://github.com/imoki/
    公众号：默库
    更新时间：20240714
    脚本：PY.js 主程序，动态修改定时任务时间
    说明：在运行此PY脚本前，请先运行PY_INIT脚本，并配置好PY表格的内容。
          将PY.js加入定时任务即可自动python脚本。
*/

// 修改名称为“wps”表内的值，需要填“wps_sid”和“文档名”
// wps_sid抓包获得，文档名就是你这个文档的名称

// 不要修改代码，修改wps表表格内的值即可
var filename = "" // 文件名
var cookie = ""
var file_id = file_id // 文件id
var cronArray = []  // 存放定时任务
let sheetNameSubConfig = "wps"; // 分配置表名称
let sheetName = "PY"
var pushHM = [] // 记录PUSH任务的推送时间
var hourMin = 0
var hourMax = 23
var line = 100

function sleep(d) {
  for (var t = Date.now(); Date.now() - t <= d; );
}

// 激活工作表函数
function ActivateSheet(sheetName) {
    let flag = 0;
    try {
      // 激活工作表
      let sheet = Application.Sheets.Item(sheetName);
      sheet.Activate();
      console.log("🥚 激活工作表：" + sheet.Name);
      flag = 1;
    } catch {
      flag = 0;
      console.log("🍳 无法激活工作表，工作表可能不存在");
    }
    return flag;
}

// 获取wps_sid、cookie
function getWpsSid(){
  flagConfig = ActivateSheet(sheetNameSubConfig); // 激活wps配置表
  // 主配置工作表存在
  if (flagConfig == 1) {
    console.log("🍳 开始读取wps配置表");
    for (let i = 2; i <= 100; i++) {
      wps_sid = Application.Range("A" + i).Text; // 以第一个wps为准
      // name = Application.Range("H" + i).Text;
      break
    }
  }
  cookie = "wps_sid=" + wps_sid
  // filename = name
}

// 是否排除文件
function juiceExclude(script_name){
  let flagExclude = 0
  let i = 2
  let key = Application.Range("I" + i).Text;
  let keyarry= key.split("&") // 使用|作为分隔符
  for(let j = 0; j < keyarry.length; j ++){
    if(script_name == keyarry[j]){ // 默认排除定时任务为CRON 和PUSH的脚本
      flagExclude = 1
      console.log( "🍳 排除任务：" , keyarry[j])
      break
    }
  }
  return flagExclude
}

// 数组字符串转整形
function arraystrToint(array){
  let result = []
  for(let i=0; i<array.length; i++){
    result.push(parseInt(array[i]))
  }
  return result
}

// 数组升序排序
function arraySortUp(value){
  value.sort(function(a, b) {
    return a - b; // 升序排序
  });
  return value
}

// 获取脚本内容
function getPyScript(url, headers){

  let resp = HTTP.get(
    url,
    { headers: headers }
  );
  resp = resp.json()
  script = resp["script"]

  return script
}

// 执行脚本
function runScript(url, headers, script){
  let data = {"sheet_name":"PY","script":script}

  let resp = HTTP.post(
      url,
      data,
      { headers: headers },
  );
  resp = resp.json()
  let result = resp["result"]
  return result
}

function main(){
  
  getWpsSid() // 获取cookie
  headers= {
    "Cookie": cookie,
    "Content-Type" : "application/json",
    "Origin":"https://www.kdocs.cn",
    "Priority":"u=1, i",
  //   "Content-Type":"application/x-www-form-urlencoded",
  }
  // console.log(headers)

  
  // 设置定时任务
  ActivateSheet(sheetName);

  let file_name = ""
  let file_id = ""
  let script_name = ""
  let script_id = ""
  for (let i = 2; i <= line; i++) {
      file_name = Application.Range("A" + i).Text;
      if (file_name == "") {
          // 如果为空行，则提前结束读取
          break;
      }
      // ABCDE
      exec = Application.Range("E" + i).Text;  // 是否执行

      if (exec == "是") {  // 是代表进行调整，则进行修改
        file_id = Application.Range("B" + i).Value;
        // console.log(file_id)
        script_name = Application.Range("C" + i).Text;
        script_id = Application.Range("D" + i).Text;
        console.log("🧑 开始执行任务：" , file_name, "-", script_name )

        // 读取脚本内容
        url = "https://www.kdocs.cn/api/v3/ide/file/" + file_id + "/script/" + script_id
        let script = getPyScript(url, headers)
        // console.log(script)

        // 执行脚本
        url = "https://www.kdocs.cn/api/aigc/pyairscript/v2/" + file_id + "/script/" + script_id + "/exec"
        let result = runScript(url, headers, script)
        if(result == "ok"){
          console.log("✨ " + script_name + " 已执行")
        }else{
          console.log("📢 " + script_name + "执行失败")
        }

        sleep(3000)
      }
  } 

}

main()