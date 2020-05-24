<template>
  <div class="home">
    <input type="file" id="file" ref="file" v-on:change="handleFileUpload()" class="inputfile" />
    <label for="file">Choose a file</label>
  </div>
</template>

<script>
export default {
  name: "Transform",
  data() {
    return {
      file: ""
    };
  },
  methods: {
    handleFileUpload() {
      this.file = this.$refs.file.files[0];
      const formData = new FormData();
      formData.append("file", this.file);
      this.$http
        .post("upload", formData, {
          headers: {
            "Content-Type": "multipart/form-data"
          }
        })
        .then(function(data) {
          console.log("SUCCESS!!", data);
        })
        .catch(function() {
          console.log("FAILURE!!");
        });
    }
  }
};
</script>
<style scoped lang="less">
.home {
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: center;
  height: calc(100vh - 250px);
}
.inputfile {
	width: 0.1px;
	height: 0.1px;
	opacity: 0;
	overflow: hidden;
	position: absolute;
	z-index: -1;
}
.inputfile + label {
    width: 200px;
    font-size: 1.25em;
    font-weight: 700;
    color: white;
    background-color: rgb(50, 49, 49);
    display: inline-block;
    padding: 0.4em;
}

.inputfile:focus + label,
.inputfile + label:hover {
    background-color: rgb(36, 36, 36);
}
.inputfile + label {
	cursor: pointer;
}
</style>
