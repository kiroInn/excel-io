<template>
  <div class="loading-mask">
    <img class="loading" src="../assets/loading__flatline.png" />
  </div>
</template>

<script lang="ts">
import { Component, Vue } from "vue-property-decorator";

@Component
export default class Loader extends Vue {}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped lang="less">
.loading-mask {
  position: absolute;
  top: 0;
  right: 0;
  left: 0;
  bottom: 0;
  background-color: rgba(238, 238, 238, 0.8);
  display: flex;
  justify-content: center;
  align-items: center;
}
.loading {
  width: 300px;
}
</style>
