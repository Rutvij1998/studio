import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { studioPlugin } from "./server/plugin";

export default defineConfig({
  plugins: [react(), studioPlugin()],
  server: { host: "127.0.0.1", port: 4310, strictPort: true },
  preview: { host: "127.0.0.1", port: 4310, strictPort: true },
});
