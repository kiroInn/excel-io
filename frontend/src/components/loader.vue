<template>
  <div class="container">
    <div class="circleloader"></div>
  </div>
</template>

<script lang="ts">
import { Component, Vue } from "vue-property-decorator";

@Component
export default class Loader extends Vue {}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped lang="less">
@import "../css/color.less";
.container {
  position: relative;
}
.circleloader {
  position: absolute;
  box-sizing: border-box;
  top: 50%;
  transform: translateX(-50%);
  border-radius: 16px;
  width: 80px;
  height: 20px;
  padding: 4px;
  background: #389c6e;

  &:before {
    content: "";
    position: absolute;
    border-radius: 16px;
    width: 20px;
    height: 12px;
    left: 0;
    background: @primary-color;
    animation: push 1s infinite linear;
  }
}

@keyframes push {
  50% {
    left: 56px;
  }
}
</style>
