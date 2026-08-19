import "dotenv/config";
import open from "open";
import { createAuthService } from "../src/auth.js";
import { loadConfig } from "../src/config.js";

const config = loadConfig();
const auth = await createAuthService(config);
console.log("Opening the Microsoft sign-in page in your default browser...");
const account = await auth.loginInteractively(async (url) => {
  await open(url);
});
console.log(`Signed in as ${account.username}. The encrypted local token cache is ready.`);
