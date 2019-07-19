# Description:
#   Post anything the bot hears to a Teams channel
#
# Notes:
#   Hardcoded stuff at the moment

module.exports = (robot) ->

    robot.hear (res) ->
      res.send "Sent that to Teams"
  