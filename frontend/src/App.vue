<template>
  <div id="app">
    <div id="nav">
      <div class="logo">
        <img src="./assets/logo.png" alt="logo" /><span>excellent</span>
      </div>
      <router-link to="/transform">fill</router-link> |
      <router-link to="/check">check</router-link>
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

.logo {
  display: flex;
  align-items: center;
  position: absolute;
  right: 150px;
  img {
    width: 20px;
  }
  span {
    margin-left: 6px;
    color: #ffd732;
    font-weight: 500;
  }
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
