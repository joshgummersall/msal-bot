require("dotenv").config();

import restify from "restify";
import { CallerIdConstants, CloudAdapter } from "botbuilder";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { EchoBot } from "./bot";
import { MsalServiceClientCredentialsFactory } from "./msalServiceClientCredentialsFactory";
import { ProxyHttpClient } from "./networkClient";

import {
  AuthenticationConfiguration,
  AuthenticationConstants,
  BotFrameworkAuthenticationFactory,
} from "botframework-connector";

const server = restify.createServer();
server.use(restify.plugins.acceptParser(server.acceptable));
server.use(restify.plugins.queryParser());
server.use(restify.plugins.bodyParser());

server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`${server.name} listening to ${server.url}`);
});

const clientApplication = new ConfidentialClientApplication({
  auth: {
    clientId: process.env.MicrosoftAppId ?? "",
    clientSecret: process.env.MicrosoftAppPassword ?? "",
  },
  system: {
    networkClient: new ProxyHttpClient({
      host: "localhost",
      port: 8080,
    }),
  },
});

const botFrameworkAuthentication = BotFrameworkAuthenticationFactory.create(
  "",
  true,
  AuthenticationConstants.ToChannelFromBotLoginUrl,
  AuthenticationConstants.ToChannelFromBotOAuthScope,
  AuthenticationConstants.ToBotFromChannelTokenIssuer,
  AuthenticationConstants.OAuthUrl,
  AuthenticationConstants.ToBotFromChannelOpenIdMetadataUrl,
  AuthenticationConstants.ToBotFromEmulatorOpenIdMetadataUrl,
  CallerIdConstants.PublicAzureChannel,
  new MsalServiceClientCredentialsFactory(
    process.env.MicrosoftAppId ?? "",
    clientApplication
  ),
  new AuthenticationConfiguration(),
  undefined,
  undefined
);

const adapter = new CloudAdapter(botFrameworkAuthentication);

const onTurnErrorHandler: typeof adapter.onTurnError = async (
  context,
  error
) => {
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  await context.sendTraceActivity(
    "OnTurnError Trace",
    `${error}`,
    "https://www.botframework.com/schemas/error",
    "TurnError"
  );

  await context.sendActivity("The bot encountered an error or bug.");
  await context.sendActivity(
    "To continue to run this bot, please fix the bot source code."
  );
};

adapter.onTurnError = onTurnErrorHandler;

const bot = new EchoBot();

server.post("/api/messages", (req, res) => {
  adapter.process(req, res, async (context) => {
    await bot.run(context);
  });
});

server.on("upgrade", (req, socket, head) => {
  const streamingAdapter = new CloudAdapter(botFrameworkAuthentication);
  streamingAdapter.onTurnError = onTurnErrorHandler;

  streamingAdapter.process(req, socket, head, async (context) => {
    await bot.run(context);
  });
});
