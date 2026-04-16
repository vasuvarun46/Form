import { defineConfig } from "@playwright/test";

export default defineConfig({
  testDir: "./e2e",
  timeout: 60000,
  use: {
    baseURL: "https://suparaco.sharepoint.com/sites/EForm",
    storageState: "e2e/.auth/user.json",
  },
});
