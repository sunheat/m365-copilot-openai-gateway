import "dotenv/config";
import { createAuthService } from "../src/auth.js";
import { loadConfig } from "../src/config.js";

const config = loadConfig();
const auth = await createAuthService(config);
const account = await auth.loginWithDeviceCode((message) => console.log(message));
console.log(`Signed in as ${account.username}. The encrypted local token cache is ready.`);
