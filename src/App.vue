<template>
  <div id="app">
    <div id="nav">
      <div class="logo">
        <img src="./assets/logo.png" alt="logo" /><span>excelio</span>
      </div>
      <router-link to="/split">split</router-link> |
      <router-link to="/evaluation">evaluation</router-link>
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
  color: #595c60;
}

.logo {
  display: flex;
  align-items: flex-end;
  position: absolute;
  left: 150px;
  top: 20px;
  img {
    width: 35px;
  }
  span {
    font-size: 20px;
    margin-left: 6px;
    color: #ffd732;
    font-weight: 500;
  }
}
#nav {
  padding: 35px;
  margin-bottom: 10px;
  text-align: center;
  a {
    font-weight: bold;
    color: #2c3e50;
    font-size: 20px;
    padding: 27px 8px;

    &.router-link-exact-active {
      color: #ffd732;
      border-top: 3px solid #ffd732;
    }
  }
}
</style>
