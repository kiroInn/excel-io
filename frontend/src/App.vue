<template>
  <div id="app">
    <div id="nav">
      <router-link to="/transform">转换</router-link> |
      <router-link to="/check">检查</router-link>
    </div>
    <router-view />
  </div>
</template>

<style lang="less">
@import "./reset.less";

#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
}

#nav {
  padding: 30px;

  a {
    font-weight: bold;
    color: #2c3e50;
    font-size: 20px;
    padding: 4px 8px;
    border-radius: 4px;

    &.router-link-exact-active {
      color: #42b983;
    }
  }
}
</style>
