var express = require('express');
var axios = require('axios');
var fs = require('fs');
var path = require('path')
var jwt = require('jsonwebtoken')
var formidable = require('formidable')
var router = express.Router()
var dayjs = require('dayjs')
var db = require("../db/db")

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
  });
})

module.exports = router