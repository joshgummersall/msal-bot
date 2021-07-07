import { ActivityHandler, MessageFactory } from "botbuilder";

export class EchoBot extends ActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context, next) => {
      const replyText = `Echo: ${context.activity.text}`;
      await context.sendActivity(MessageFactory.text(replyText, replyText));

      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded ?? [];
      const welcomeText = "Hello and welcome!";

      await Promise.all(
        membersAdded
          .filter((member) => member.id !== context.activity.recipient.id)
          .map(() =>
            context.sendActivity(MessageFactory.text(welcomeText, welcomeText))
          )
      );

      await next();
    });
  }
}
