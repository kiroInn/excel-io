<template>
  <div class="home">
    <h1 class="message">{{ message }}</h1>
    <input
      type="file"
      id="file"
      ref="file"
      v-on:change="handleFileUpload()"
      class="inputfile"
    />
    <label for="file">Choose a file</label>
    <h1 class="slogan">Smart <strong>Excel</strong> Transform</h1>
    <ul class="results">
      <li v-for="file in files" :key="file.key">
        <a href="javascript:void(0);" v-on:click="download(file.key)">{{file.key}}</a>
      </li>
    </ul>
  </div>
</template>

<script>
import * as Excel from "exceljs";
import { saveAs } from "file-saver";
import { write, validateTB, loadTemplate, SHEET_NAME, parse, fillData } from "./service";

export default {
  name: "Transform",
  data() {
    return {
      message: "",
      files: [],
    };
  },
  methods: {
    async download(key){
      console.log('files', this.files);
        const workbook = this.files.filter(f => f.key === key)[0].workbook;
        const buffer = await workbook.xlsx.writeBuffer();
        const fileType =
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        const fileExtension = ".xlsx";

        const blob = new Blob([buffer], { type: fileType });

        saveAs(blob, "Sample" + fileExtension);
    },
    handleFileUpload() {
      this.message = "";
      const file = this.$refs.file.files[0];
      const reader = new FileReader();
      reader.onload = async () => {
        const workbook = new Excel.Workbook();
        await workbook.xlsx.load(reader.result);

        console.log(workbook);

        this.message = validateTB(workbook);
        if (!this.message) {
          // const data = parse(workbook, SHEET_NAME[196000])
          // const tempWrokSheet = loadTemplate(196000)
          // const workbook = fillData(tempWrokSheet, data);
          this.files = [{
            key: SHEET_NAME[196000],
            workbook,
          }];
          console.log('files', this.files);
        }
      };
      reader.readAsArrayBuffer(file);
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
  background-color: mintcream;
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
.slogan {
  margin-top: 40px;
  font-size: 44px;
  strong {
    font-weight: 600;
  }
}
.results{
  margin-top: 20px;
}
</style>
