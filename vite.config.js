import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// If you deploy to GitHub Pages at https://<user>.github.io/<repo>/,
// uncomment and set base to '/<repo>/'. For Vercel/Netlify, leave it as '/'.
export default defineConfig({
  plugins: [react()],
  // base: '/attendance-ledger/',
});
