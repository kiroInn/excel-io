<template>
  <div class="home">
    <h1 v-if="!!message" class="message">{{ message }}</h1>
    <input
      type="file"
      id="file"
      ref="file"
      v-on:change="handleFileUpload()"
      class="inputfile"
    />
    <label for="file">Choose a file</label>
    <button v-if="!!files.length" class="downloadAll">download all</button>
    <ul class="results">
      <li v-for="file in files" :key="file.key">
        <span>{{ file.name }}</span>
        <span class="status-success">success</span>
        <a
          href="javascript:void(0);"
          class="download"
          v-on:click="download(file.key)"
          >download</a
        >
      </li>
    </ul>
    <h1 class="slogan">Smart <strong>Excel</strong> Transform</h1>
  </div>
</template>

<script>
import * as Excel from "exceljs";
import { saveAs } from "file-saver";
import { validateTB, loadTemplate, SHEET_NAME, fillData } from "./service";
import { isXlsx } from "@/util/file";

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
      const blob = new Blob([buffer], { type: fileType });
      saveAs(blob, file.name);
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
              name: `202004-CNCDU-${SHEET_NAME._196000}.xlsx`,
              workbook: resultWrokbook,
              buffer: resultWrokbook.xlsx.writeBuffer()
            }
          ];
          console.log("parsing success", this.files);
        }
        this.isLoading = false;
      };
      if (isXlsx(file.name)) {
        reader.readAsArrayBuffer(file);
      } else {
        this.message = "File type is not supported";
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
  border-radius: 3px;
}

.inputfile:focus + label,
.inputfile + label:hover {
  background-color: rgb(36, 36, 36);
}
.inputfile + label {
  cursor: pointer;
}
.slogan {
  font-size: 44px;
  margin-top: 24px;
  strong {
    font-weight: 600;
  }
}
.downloadAll {
  border-radius: 3px;
  margin-top: 12px;
  background-color: #02c487;
  padding: 5px 8px;
  font-size: 16px;
  color: white;
  font-weight: 600;
}
.results {
  margin-top: 20px;
  li {
    padding: 8px;
    background-color: white;
    border: 1px solid gray;
    span {
      color: #2d3e50;
      text-decoration: none;
      margin-right: 16px;
    }
    .status-success {
      background-color: #49cb90;
      color: white;
      border-radius: 3px;
      padding: 1px 5px;
    }
    .download {
      color: #42b983;
    }
  }
}
.message {
  margin-bottom: 25px;
  padding: 8px;
  border: 1px solid #ffccc7;
  background-color: #fff2f0;
  color: black;
  border-radius: 3px;
}
</style>
