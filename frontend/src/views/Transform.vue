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
        <a href="javascript:void(0);" v-on:click="download(file.key)">{{
          file.key
        }}</a>
      </li>
    </ul>
  </div>
</template>

<script>
import * as Excel from "exceljs";
import { saveAs } from "file-saver";
import { validateTB, loadTemplate, SHEET_NAME, fillData } from "./service";
import {isXlsx} from "@/util/file";

export default {
  name: "Transform",
  data() {
    return {
      message: "",
      isLoading: false,
      files: []
    };
  },
  methods: {
    async download(key) {
      const file = this.files.filter(f => f.key === key)[0];
      const workbook = file.workbook;
      const buffer = await workbook.xlsx.writeBuffer();
      const fileType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
      const fileExtension = ".xlsx";
      const blob = new Blob([buffer], { type: fileType });
      saveAs(blob, `${file.name + fileExtension}`);
    },
    handleFileUpload() {
      this.message = "";
      this.isLoading = true;
      const file = this.$refs.file.files[0];
      const reader = new FileReader();
      reader.onload = async () => {
        const fromWorkBook = new Excel.Workbook();
        await fromWorkBook.xlsx.load(reader.result);
        this.message = validateTB(fromWorkBook);
        if (!this.message) {
          const toWorkbook = await loadTemplate(SHEET_NAME._196000);
          const resultWrokbook = fillData(
            fromWorkBook,
            toWorkbook,
            SHEET_NAME._196000
          );
          this.files = [
            {
              key: SHEET_NAME._196000,
              name: `202004-CNCDU-${SHEET_NAME._196000}`,
              workbook: resultWrokbook
            }
          ];
          console.log("parsing success", this.files);
        }
        this.isLoading = false;
      };
      if(isXlsx(file.name)){
        reader.readAsArrayBuffer(file);
      } else {
        this.message = "File type is not supported"
        this.isLoading = false;
      }
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
.results {
  margin-top: 20px;
}
.message {
  margin-bottom: 25px;
}
</style>
