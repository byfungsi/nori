import DefaultTheme from "vitepress/theme";
import "./style.css";
import LivePlayground from "./LivePlayground.vue";
import type { Theme } from "vitepress";
export default {
  extends: DefaultTheme,
  enhanceApp({ app }) {
    app.component("LivePlayground", LivePlayground);
  },
} satisfies Theme;
