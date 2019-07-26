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


require('http')

function getBearerToken(robot) {
  var access_token

  data = `client_id=${process.env.client_id}&client_secret=${process.env.secret}&scope=${process.env.scope}&grant_type=client_credentials`

  return new Promise(function (resolve, reject) {
    robot.http(`https://login.microsoftonline.com/${process.env.tenant}/oauth2/v2.0/token`)
      .header('Content-Type', 'application/x-www-form-urlencoded')
      .post(data)(function (err, resp, body) {
        if(err) {
          throw(err)
        }
        if(!body) {
          throw new Error("No token obtained")
        }
        parsedBody = JSON.parse(body)
        access_token = parsedBody.access_token
        if(!access_token) {
          throw new Error("Cannot obtain token: " + body)
        }
        resolve(access_token)
      });
  });
}

function getTeamData(robot, access_token) {
  return new Promise(function (resolve, reject) {
  robot.http(`https://graph.microsoft.com/beta/groups?$select=id,displayname,resourceProvisioningOptions`)
    .header('Authorization', `Bearer ${access_token}`)
    .get()(function (err, resp, body) {
      if(err) {
        throw(err)
      }
      if(!body) {
        throw new Error("No teams found")
      }
      parsedBody = JSON.parse(body)
      resolve(parsedBody)
    });
  });
}

function getAllGroupsWithTeams(robot, access_token) {
  return new Promise(function (resolve, reject) {
  robot.http(`https://graph.microsoft.com/beta/groups?$select=id,displayname,resourceProvisioningOptions`)
    .header('Authorization', `Bearer ${access_token}`)
    .get()(function (err, resp, body) {
      if(err) {
        throw(err)
      }
      if(!body) {
        throw new Error("No groups found")
      }
      groups = JSON.parse(body).value
      if(!groups) {
        throw new Error("No groups found")
      }
      groupsWithTeams = []
      groups.forEach((group) => {if(group.resourceProvisioningOptions == "Team") groupsWithTeams.push(group)})
      resolve(groupsWithTeams)
    });
  });
}

function getTeamData(robot, access_token, group_id) {
  return new Promise(function (resolve, reject) {
    robot.http(`https://graph.microsoft.com/beta/teams/${group_id}`)
      .header('Authorization', `Bearer ${access_token}`)
      .get()(function (err, resp, body) {
        if(err) {
          throw(err)
        }
        if(!body) {
          throw new Error("Can't get data for group id " + group_id)
        }
        groups = JSON.parse(body)
        resolve(groups)
      });
    });
}

function getChannelsInTeam(robot, access_token, team_id) {
  
  return new Promise(function (resolve, reject) {
    robot.http(`https://graph.microsoft.com/beta/teams/${team_id}/channels`)
      .header('Authorization', `Bearer ${access_token}`)
      .get()(function (err, resp, body) {
        if(err) {
          throw(err)
        }
        if(!body) {
          throw new Error("No channels found in team " + team_id)
        }
        channels = JSON.parse(body)
        resolve(channels)
      });
    });
}

function getMessagesFromChannel(robot, access_token, team, channel) {
  return new Promise(function (resolve, reject) {
    robot.http(`https://graph.microsoft.com/beta/teams/{id}/channels/{id}/messages`)
      .header('Authorization', `Bearer ${access_token}`)
      .get()(function (err, resp, body) {
        if(err) {
          throw(err)
        }
        if(!body) {
          throw new Error("No messages found in channel " + channel.name + " in team " + team.name)
        }
        console.log("messages raw: " + body)
        console.log("messages: " + util.inspect(body))
        console.log("resp.statusCode: " + resp.statusCode)
        messages = JSON.parse(body)
        resolve(messages)
      });
    });
}

async function doGraphStuff(robot) {
  access_token = await getBearerToken(robot)

  groupsWithTeams = await getAllGroupsWithTeams(robot, access_token)

  teams = []
  for(const group of groupsWithTeams) {
    teams_data = await getTeamData(robot, access_token, group.id)
    channels = await getChannelsInTeam(robot, access_token, teams_data.id)
    for(const channel of channels.value) {
      console.log("Group: " + group.displayName + ", channel id = " + channel.id + " name = " + channel.displayName)
    }
  }

  team = {id: "fb5cb1df-ad0a-4ae4-b979-21db4a48f68c", name: "fb5cb1df-ad0a-4ae4-b979-21db4a48f68c"}
  channel = {id: "19:6e6280be5ce240b28eca8ee00de866c3@thread.skype", name: "slacktest channel 2"}

  messages = await getMessagesFromChannel(robot, access_token, team, channel)
}

module.exports = function (robot) {

  doGraphStuff(robot)





  var app = robot.router;

  var authRouter = require('../routes/auth');
  app.use('/hubot/auth', authRouter);

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
    console.log("Got an error!!!!")
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
    console.log("sent it!")
  });
}