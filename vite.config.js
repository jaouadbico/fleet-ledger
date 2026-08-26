import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { VitePWA } from "vite-plugin-pwa";

export default defineConfig({
  // Relative base so the built site works at any GitHub Pages path,
  // e.g. https://username.github.io/repo-name/
  base: "./",
  plugins: [
    react(),
    VitePWA({
      registerType: "autoUpdate",
      includeAssets: ["icon-192.png", "icon-512.png"],
      manifest: {
        name: "Fleet Ledger",
        short_name: "Fleet Ledger",
        description: "Track trucks and their booked contracts.",
        start_url: "./",
        scope: "./",
        display: "standalone",
        background_color: "#12161B",
        theme_color: "#12161B",
        icons: [
          { src: "icon-192.png", sizes: "192x192", type: "image/png" },
          { src: "icon-512.png", sizes: "512x512", type: "image/png" },
          { src: "icon-512.png", sizes: "512x512", type: "image/png", purpose: "maskable" },
        ],
      },
      workbox: {
        globPatterns: ["**/*.{js,css,html,png,svg,ico,webmanifest}"],
        // Take control immediately on activate instead of waiting for every
        // open tab to close - Safari in particular is prone to getting
        // stuck serving an old cached version without this.
        skipWaiting: true,
        clientsClaim: true,
        cleanupOutdatedCaches: true,
      },
    }),
  ],
});
