
var jwt = require('jsonwebtoken');
var User = require('../models/user');

module.exports = (req, res, next) => {
  const header = req.headers.authorization;
  let token;

  if (header) token = header.split(' ')[1];
  if (token) {
    jwt.verify(token, process.env.JWT_SECRET, (err, decoded) => {
      if (err) {
        res.status(401).json({ errors: { global: "Invalid token"} });
      } else {
        User.findOne({ account: decoded.account }).then(user => {
          req.currentUser = user;
          next();
        });
      }
    });
  } else {
    res.status(401).json({ errors: { global: "No token"} });
  }
}