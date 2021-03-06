require("dotenv").config();

import { JwtTokenExtractor } from 'botframework-connector/lib/auth/jwtTokenExtractor';
import { ProxyOpenIdMetadata } from './proxyOpenIdMetadata';
import { memoize } from 'lodash';

(JwtTokenExtractor as any).getOrAddOpenIdMetadata = memoize((url: string) =>
  new ProxyOpenIdMetadata(url, {
    host: 'localhost',
    port: 8080
  }));

import axios from "axios";
import restify from "restify";
import { CallerIdConstants, CloudAdapter, Response } from "botbuilder";
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
  async (input, init) => {
    const response = await axios.post(input, JSON.parse(init.body), {
      headers: init.headers,
      proxy: {
        host: "localhost",
        port: 8080,
      },
      validateStatus: () => true,
    });

    return ({
      status: response.status,
      json: async () => response.data,
    } as any) as Response;
  },
  {
    proxySettings: {
      host: "localhost",
      port: 8080,
    },
  }
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
