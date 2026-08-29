import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// IMPORTANT: base must match your GitHub repo name exactly (case-sensitive)
// Your site is at https://dhirajbaikerikar.github.io/Readnexa/
// so base MUST be "/Readnexa/"
export default defineConfig({
  plugins: [react()],
  base: "/Readnexa/",
});
