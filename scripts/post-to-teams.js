// Description:
//   Post anything the bot hears to a Teams channel
//
// Notes:
//   Hardcoded stuff at the moment

var util = require('util')
require('dotenv').config();
var passport = require('passport');
var express = require('express');
var path = require('path');
var createError = require('http-errors');
var cookieParser = require('cookie-parser');
var OIDCStrategy = require('passport-azure-ad').OIDCStrategy;
var logger = require('morgan');
var request = require('request');

function getBearerToken(robot) {
  var access_token

  data = `client_id=${process.env.client_id}&client_secret=${process.env.secret}&scope=${process.env.scope}&grant_type=client_credentials`

  return new Promise(function (resolve, reject) {
    robot.http(`https://login.microsoftonline.com/${process.env.tenant}/oauth2/v2.0/token`)
      .header('Content-Type', 'application/x-www-form-urlencoded')
      .post(data)(function (err, resp, body) {
        if (err) {
          throw (err)
        }
        if (!body) {
          throw new Error("No token obtained")
        }
        parsedBody = JSON.parse(body)
        access_token = parsedBody.access_token
        if (!access_token) {
          throw new Error("Cannot obtain token: " + body)
        }
        resolve(access_token)
      });
  });
}

module.exports = async function (robot) {
  access_token = await getBearerToken(robot)

  var app = robot.router;
  var authRouter = require('../routes/auth');
  app.use('/hubot/auth', authRouter);
  var teamsDataRouter = require('../routes/teamsData')(robot, access_token);
  app.use('/hubot/teamsData', teamsDataRouter);


  app.use(logger('dev'));
  app.use(express.json());
  app.use(express.urlencoded({ extended: false }));
  app.use(cookieParser());
  app.use(express.static(path.join(__dirname, 'public')));

  app.use(passport.initialize());
  app.use(passport.session());

  app.use(function (req, res, next) {
    next(createError(404));
  });

  // error handler
  app.use(function (err, req, res, next) {
    // set locals, only providing error in development
    res.locals.message = err.message;
    res.locals.error = req.app.get('env') === 'development' ? err : {};

    // render the error page
    res.status(err.status || 500);
    res.json({
      message: err.message,
      error: err
    });
  });

  robot.hear(/.*/i, function (msg) {
    var today = new Date();

    msg.send("Sent that to Teams at " + today);
  });
}