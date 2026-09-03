import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// IMPORTANT: change "data-workbench" below to your GitHub repo name
// before deploying to GitHub Pages (e.g. base: "/my-repo-name/").
// If you're not deploying to GitHub Pages (e.g. Netlify/Vercel/custom domain),
// change this back to base: "/"
export default defineConfig({
  plugins: [react()],
  base: "/data-workbench/",
});
