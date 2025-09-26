<script setup lang="ts">
import ExcelReader from './components/ExcelReader.vue'
</script>

<template>
  <div id="app">
    <ExcelReader />
  </div>
</template>

<style>
* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}

body {
  font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
  background-color: #f5f5f5;
  line-height: 1.6;
}

#app {
  min-height: 100vh;
}
</style>
