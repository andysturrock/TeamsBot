var express = require('express');
var router = express.Router();
var util = require('util')
var robot;
var access_token

router.get('/teams', async function (req, res) {
  groupsWithTeams = await getAllGroupsWithTeams(robot, access_token)

  for (const group of groupsWithTeams) {
    teams_data = await getTeamData(robot, access_token, group.id)
    res.write(`{id:${teams_data.id}, displayName:"${teams_data.displayName}"}`)
  }
  res.end();
}
);

router.get('/channels', async function (req, res) {
  team_id = req.query.team_id

  channels = await getChannelsInTeam(robot, access_token, team_id)

  for (const channel of channels.value) {
    res.write(`{id:${channel.id}, displayName:"${channel.displayName}"}`)
  }
  res.end();
}
);

router.get('/messages', async function (req, res) {
  team_id = req.query.team_id
  channel_id = req.query.channel_id

  messages = await getMessagesFromChannel(robot, access_token, team_id, channel_id)

  console.log("messages = " + util.inspect(messages))

  for (const message of messages.value) {
    res.write(`{id:${message.id}, displayName:"${message.displayName}"}`)
  }
  res.end();
}
);

module.exports = function (theRobot, theAccessToken) {
  robot = theRobot
  access_token = theAccessToken
  return router;
}

function getAllGroupsWithTeams(robot, access_token) {
  return new Promise(function (resolve, reject) {
    robot.http(`https://graph.microsoft.com/beta/groups?$select=id,displayname,resourceProvisioningOptions`)
      .header('Authorization', `Bearer ${access_token}`)
      .get()(function (err, resp, body) {
        if (err) {
          throw (err)
        }
        if (!body) {
          throw new Error("No groups found")
        }
        groups = JSON.parse(body).value
        if (!groups) {
          throw new Error("No groups found")
        }
        groupsWithTeams = []
        groups.forEach((group) => { if (group.resourceProvisioningOptions == "Team") groupsWithTeams.push(group) })
        resolve(groupsWithTeams)
      });
  });
}

function getTeamData(robot, access_token, group_id) {
  return new Promise(function (resolve, reject) {
    robot.http(`https://graph.microsoft.com/beta/teams/${group_id}`)
      .header('Authorization', `Bearer ${access_token}`)
      .get()(function (err, resp, body) {
        if (err) {
          throw (err)
        }
        if (!body) {
          throw new Error("Can't get data for group id " + group_id)
        }
        teams = JSON.parse(body)
        resolve(teams)
      });
  });
}

function getChannelsInTeam(robot, access_token, team_id) {

  return new Promise(function (resolve, reject) {
    robot.http(`https://graph.microsoft.com/beta/teams/${team_id}/channels`)
      .header('Authorization', `Bearer ${access_token}`)
      .get()(function (err, resp, body) {
        if (err) {
          throw (err)
        }
        if (!body) {
          throw new Error("No channels found in team " + team_id)
        }
        channels = JSON.parse(body)
        resolve(channels)
      });
  });
}

function getMessagesFromChannel(robot, access_token, team_id, channel_id) {
  return new Promise(function (resolve, reject) {
    robot.http(`https://graph.microsoft.com/beta/teams/${team_id}/channels/${channel_id}/messages`)
      .header('Authorization', `Bearer ${access_token}`)
      .get()(function (err, resp, body) {
        if (err) {
          throw (err)
        }
        if (!body) {
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