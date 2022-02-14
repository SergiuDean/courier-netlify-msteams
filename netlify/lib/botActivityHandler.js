const { TurnContext, TeamsActivityHandler } = require("botbuilder");
const { CourierClient } = require("@trycourier/courier");
const courier = CourierClient();

class BotActivityHandler extends TeamsActivityHandler {
  constructor() {
    super();

    // Registers an activity event handler for the message event, emitted for every incoming message activity.
    this.onMessage(async (context, next) => {
      TurnContext.removeRecipientMention(context.activity);
      const text = context.activity.text.trim().toLocaleLowerCase();
      if (text.includes("show-user") || text.includes("show-channel")) {
        await this.updateCourierProfile(context);
      } else if (text === "test") {
        await context.sendActivity(`Gravity bot has been successfully added.`);
      } else if (text === "info") {
        await context.sendActivity("service url: "+service_url+"\ntenant id: "+tenant_id+"\nuser id: "+context.activity.from.id);
      }

      await next();
    });
  }

  async updateCourierProfile(context) {
    const {
      serviceUrl: service_url,
      channelData: {
        tenant: { id: tenant_id }
      }
    } = context.activity;
    let profile = {
      ms_teams: {
        tenant_id,
        service_url
      }
    };
    const text = context.activity.text.trim();
    const [cmd, recipientId] = text.split(" ");

    if (cmd.toLowerCase() === "show-channel") {
      if (!context.activity.channelData.channel) {
        await context.sendActivity(
          `show-channel must be called from a channel.`
        );
        return;
      }
      await context.sendActivity("channel id: "+context.activity.channelData.channel.id);
    } else if (cmd.toLowerCase() === "show-user") {
      await context.sendActivity("user id: "+context.activity.from.id);
    } else {
      await context.sendActivity(
        `Error. Unsupported action.`
      );
      return;
    }

//     try {
//       await courier.mergeProfile({
//         recipientId,
//         profile
//       });

//       await context.sendActivity(`Your profile has been updated.`);
//     } catch (err) {
//       console.log(err);
//       await context.sendActivity(`An error occurred updating your profile.`);
//     }
  }
}

module.exports.BotActivityHandler = BotActivityHandler;
