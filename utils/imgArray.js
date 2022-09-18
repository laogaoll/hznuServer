var _ = require('lodash');
var fs = require('fs');

module.exports.getImgFromArray = function (array) {  
  return _.map(array, 'path');
}

module.exports.getImgAfterDeleting = function (img, deleteImgs) {  
  
  // 过滤数组，去除需删除图片的路径，同时删除本地图片
  img = img.filter( function (el) {
    if (deleteImgs.includes(el) && el) fs.unlinkSync(el);
    return !deleteImgs.includes(el) 
  });

  // 返回数组
  return img;
}

module.exports.getImgAfterAdding = function (img, array) {  

  // 获取新的图片数组
  var newImgArr = _.map(array, 'path');

  // 合并后，返回数组
  return img.concat(newImgArr);
}

// module.exports = { getImgFromArray, getImgAfterDeleting, getImgAfterAdding };