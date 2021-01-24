<template>
  <div id="app">
    <AppHeader />
    <router-view />
  </div>
</template>

<script>
import AppHeader from "@/components/Header.vue";

export default {
  name: "app",
  components: {
    AppHeader
  }
};
</script>

<style lang="scss">
@import "~bulma/css/bulma.css";
@import '~bulma-calendar/dist/css/bulma-calendar.min.css';

html { height:100%; }
body{
  height:100%;
  margin: 0px;
}

#app {
  display: flex;
  flex-direction: column;
  font-family: "Avenir", Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  height: 100%;
}

div.v-dialog-toast__container .v-dialog-toast__content {
    padding: 12px 10px 0 70px;
    font-size: 14px;
    overflow: hidden;
    min-height: 68px;
}
</style>
