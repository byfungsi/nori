<script setup lang="ts">
import { withBase } from "vitepress";
</script>

<template>
  <div class="nori-live-example">
    <p>
      <a :href="withBase('/demo/')" target="_blank" rel="noopener"
        >Open playground in a new tab ↗</a
      >
    </p>
    <iframe
      :src="withBase('/demo/?embed=1')"
      title="Nori interactive spreadsheet playground"
      loading="lazy"
    />
  </div>
</template>

<style scoped>
.nori-live-example {
  width: 100%;
  min-width: 0;
}
iframe {
  display: block;
  scroll-margin-top: 80px;
  width: 100%;
  height: 900px;
  border: 1px solid var(--vp-c-divider);
  border-radius: 12px;
  background: #f7f9f2;
}
@media (max-width: 640px) {
  iframe {
    height: 780px;
  }
}
</style>
