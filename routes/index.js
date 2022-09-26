var express = require('express');
var axios = require('axios');
var fs = require('fs');//获取文件系统工具，负责读写文件
var path = require('path')//工具模块，处理文件路径的小工具
var jwt = require('jsonwebtoken')
var formidable = require('formidable')
var router = express.Router()
var dayjs = require('dayjs')
var db = require("../db/db")
// 导入DocxTemplater
var PizZip = require('pizzip')
var Docxtemplater = require('docxtemplater');
// 使用JSZIP打包文件夹
var JSZIP = require("jszip");
var zip = new JSZIP();


const { finished } = require('stream');
const compression = require('compression');


var root = path.resolve(__dirname,'../')
var clone =(e)=> {
  return JSON.parse(JSON.stringify(e))
}

const SECRET_KEY = 'ANSAIR-SYSTEM'

var callSQLProc = (sql, params, res) => {
  return new Promise (resolve => {
    db.procedureSQL(sql,JSON.stringify(params),(err,ret)=>{
      if (err) {
        res.status(500).json({ code: -1, msg: '提交请求失败，请联系管理员！', data: null})
      }else{
        resolve(ret)
      }
    })
  })
}

var callP = async (sql, params, res) => {
  return  await callSQLProc(sql, params, res)
}

//解密token
var decodeUser = (req)=>{
  let token = req.headers.authorization
  return  JSON.parse(token?.split(' ')[1])
}


router.post('/login',async (req, res, next) =>{
  let params = req.body
  let sql = `CALL PROC_LOGIN(?)`
  let r = await callP(sql, params, res)

  if (r.length > 0) {
    let ret = clone(r[0])
    let token = jwt.sign(ret, SECRET_KEY)
    res.status(200).json({code: 200, data: ret, token: token, msg: '登录成功'})
  } else {
    res.status(200).json({code: 301, data: null, msg: '用户名或密码错误'})
  }
})
//组件加载时自动运行
router.post('/qryCls', async (req, res, next) =>{
  //从token中解密出uid
  let uid = decodeUser(req).uid
  //console.log(uid);
  let params = {uid:uid}
  // 获取课程名称（name）和课程代码（code，数量等于课程数量）
  let sql= `CALL PROC_QRY_CLS(?)`
  let r = await callP(sql, params, res)
  res.status(200).json({ code: 200, data: r })
});
// 组件加载时自动运行，获取历史课程名称和课程代码
router.post('/qryClsOld',async(req,res,next)=>{
  let uid = decodeUser(req).uid
  let params = {uid:uid}
  let sql = `CALL PROC_QRY_CLS_OLD(?)`
  let r = await callP(sql,params,res)
  res.status(200).json({code:200,data:r})
});
router.post('/qryClsMain', async (req, res, next) =>{
  let uid = decodeUser(req).uid
  let params = {uid:uid, code: req.body.code}

  // console.log(params)
  // 读取了tab_tech_main表的所有信息 
  let sql1= `CALL PROC_QRY_CLS_MAIN(?)`
  // 读取了tab_tech_tp的所有信息
  let sql2= `CALL PROC_QRY_TECH(?)`
  // 读取了tab_tech_ep的所有信息
  let sql3= `CALL PROC_QRY_EXP(?)`
  let r = await callP(sql1, params, res)
  let s = await callP(sql2, params, res)
  let t = await callP(sql3, params, res)
  res.status(200).json({ code: 200, data: r, tecList:s, expList:t })
});


router.post('/savCls', async (req, res, next) =>{
  let uid = decodeUser(req).uid
  req.body.uid = uid
  let params = req.body
  console.log(params)
  let sql1= `CALL PROC_SAV_CLS(?)`
  let sql2= `CALL PROC_SAV_TECH(?)`
  let sql3= `CALL PROC_SAV_EXP(?)`
  let r = await callP(sql1, params, res)
  let s = await callP(sql2, params, res)
  let t = await callP(sql3, params, res)
  res.status(200).json({ code: 200, data: r, tecList:s, expList:t })
});

//导出历史课程
router.post('/impHC', async (req,res,next)=>{
  let uid = decodeUser(req).uid
  let params = {uid:uid, code: req.body.code}
  //console.log(params)
  let sql1 = `CALL PROC_GET_CLS_MAIN_OLD(?)`
  let r = await callP(sql1,params,res)
  res.status(200).json({code:200,data:r}) 
})
//导出同类课程
router.post('/impSCLS',async (req,res,next)=>{
  let uid = decodeUser(req).uid
  let params = {uid:uid,code:req.body.code}
  let sql = `CALL PROC_QRY_SCLS(?)`
  let r = await callP(sql,params,res)
  res.status(200).json({code:200,data:r});
})

// 读取目录及文件
function  readDir(obj,Path){
   let files = fs.readdirSync(Path);//读取目录中的所有文件及文件夹（同步操作）
   files.forEach(fileName =>{
    let  fillPath = Path + "/" +fileName;
    let  file = fs.statSync(fillPath);// 获取文件信息状态
    if(file.isDirectory()){//如果是目录的话继续查询
      let  dirZip = zip.folder(fileName);//压缩对象中生成该目录
      readDir(dirZip,fillPath);//重新检索目录文件
    }else{
      obj.file(fileName,fs.readFileSync(fillPath));//压缩目录添加文件
    }
   })
}
// 生成一个压缩包
function startZip(){
  const sourceDir = path.join(__dirname,"../export"); // path.join就是把每个路径进行拼接
  readDir(zip,sourceDir);
  zip.generateAsync({//设置压缩格式，开始打包
      type:"nodebuffer",//nodejs用
      compression:"DEFLATE",//压缩算法
      compressionOptions:{//压缩级别
          level:9
      }
  }).then((content)=>{
    const dest = path.join(__dirname,"../public");
    fs.mkdirSync(dest,{
      recursive:true
    })
    fs.writeFileSync(path.resolve(dest,'hznu.zip'),content);
  });

}
// 导出为docx
router.get('/export',async(req,res)=>{

  let sql1 = `CALL PROC_EXPORT_UID()`
  let sql2 = `CALL PROC_EXPORT_EXP()`
  let sql3 = `CALL PROC_EXPORT_TECH()`
  let sql4 = `CALL PROC_EXPORT_CLS()`
  let r = await callP(sql1,null,res)
  let s = await callP(sql2,null,res)
  let t = await callP(sql3,null,res)
  let p = await callP(sql4,null,res)
  // 将docx文件作为二进制内容加载
  for(let i = 0;i<r.length;i++){

    let content = fs.readFileSync(path.resolve(__dirname,'../public/hznu.docx'),'binary');
    let zip = new PizZip(content);
    let doc = new Docxtemplater(zip);
    //try { doc = new Docxtemplater(zip) } catch(error){errorHandler(error);}
    let data1 = r[i];
    // 实验进度
    let data2 = s.filter((item,index,arr)=>{
      return item.uid == data1.uid && (item.code.includes(`${data1.code}`));
    });
    // 教学进度
    let data3 = t.filter((item,index,arr)=>{
      return item.uid == data1.uid && (item.code.includes(`${data1.code}`));
    });
    // 基本信息
    let data4 = p.filter((item,index,arr)=>{
      return item.uid == data1.uid && item.code == data1.code;
    })
    let w_hour = parseInt(data4[0].t_hour) + parseInt(data4[0].e_hour);
    let a_hour = w_hour*16;
    doc.setData();
    try{ doc.render({
      clsterm:`${data4[0].term}`,
      name:`${data4[0].name}`,
      clspos:`${data4[0].pos}`,
      clscol:`${data4[0].col}`,
      clsmark:`${data4[0].mark}`,
      clsweek:`${data4[0].week}`,
      clse_hour:`${data4[0].e_hour}`,
      clst_hour:`${data4[0].t_hour}`,
      clsm_tech:`${data4[0].m_tech}`,
      clss_tach:`${data4[0].s_tach}`,
      clsq_time:`${data4[0].q_time}`,
      clsq_addr:`${data4[0].q_addr}`,
      w_hour:`${w_hour}`,
      a_hour:`${a_hour}`,
      cls:data4,
      tecList:data3,
      expList:data2,
      desc:`${data4[0].desc}`,
      mate:`${data4[0].mate}`,
      exam:`${data4[0].exam}`,
      method:`${data4[0].method}`,
    }
    )} catch (error) { errorHandler(error);}
    var buf = doc.getZip().generate({
      type: 'nodebuffer',
      compression: "DEFLATE",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    fs.writeFileSync(path.resolve(__dirname,`../export/${data4[0].uname}${data4[0].name}.docx`),buf);
    
  }
  startZip();
  console.log(`finished....`);
  res.status(200).json({code:200,data:'http://124.220.20.66:8000/hznu.zip',msg:"finished"});
  
});

// 上传文件
router.post('/upload', function (req, res) {
  const form = formidable({uploadDir: `${__dirname}/../img`});

  form.on('fileBegin', function (name, file) {
    file.filepath = `img/sys_${dayjs().format('YYYYMMDDhhmmss')}.jpeg`
  })
 
  form.parse(req, (err, fields, files) => {
    if (err) {
      next(err);
      return;
    }
    res.status(200).json({
      code: 200,
      msg: '上传照片成功',
      data: {path: files.file.filepath}
    })
    
  })
});

module.exports = router