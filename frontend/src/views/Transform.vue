<template>
  <div class="home">
    <h1 v-if="!!message" class="message">{{ message }}</h1>
    <div>
    <input
      type="file"
      id="file"
      ref="file"
      v-on:change="handleFileUpload()"
      class="inputfile"
    />
    <label class="select-file" for="file">Choose a file</label>
    <label class="configMapping" v-on:click="isEditMapping = true">Config</label>
    </div>
    
    <loader v-show="isLoading" class="loader"></loader>
    <a
      v-if="!!files.length"
      class="downloadAll"
      v-on:click="downloadAll"
      href="javascript:void(0)"
      >Download all</a
    >
    <ul v-if="!!files.length" class="results">
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
    <modal
      v-if="isEditMapping"
      v-on:submit="onEditConfig"
      v-on:close="isEditMapping = false"
    >
      <template v-slot:header>
        <h1>Config Mapping</h1>
      </template>
      <template v-slot:body>
        <div v-for="(mapping, index) in mappings" :key="index">
          templateName: <input type="text" v-model="mapping.templateName" />
          <ul>
            <li v-for="(mp, index) in mapping.values" :key="index">
              from: <input type="text" v-model="mp.from" /> to:
              <input type="text" v-model="mp.to" /> type:
              <select v-model="mp.type">
                <option :selected="mp.type === 'string'" value="string"
                  >String</option
                >
                <option :selected="mp.type === 'image'" value="image"
                  >Image</option
                >
                <option :selected="mp.type === 'date'" value="date"
                  >Date</option
                >
              </select>
              <div v-if="mp.type === 'image'">
                <label>tl:col</label
                ><input type="text" v-model="mp.range.tl.col" />
                <label>tl:row</label
                ><input type="text" v-model="mp.range.tl.row" />
                <label>br:col</label
                ><input type="text" v-model="mp.range.br.col" />
                <label>br:row</label
                ><input type="text" v-model="mp.range.br.row" />
              </div>
            </li>
          </ul>
        </div>
        <div></div>
      </template>
    </modal>
  </div>
</template>

<script>
import Loader from "@/components/loader";
import Modal from "@/components/modal";
import * as Excel from "exceljs";
import { saveAs } from "file-saver";
import { validateFrom, loadTemplate, fillData } from "@/service/transform";
import { isXlsx } from "@/util/file";
import _ from "lodash";
import { DEFAULT_MAPPING } from "@/service/mapping";

export default {
  name: "Transform",
  components: {
    Loader,
    Modal
  },
  created() {
    this.mappings = _.clone(DEFAULT_MAPPING);
  },
  data() {
    return {
      message: "",
      isLoading: false,
      isEditMapping: false,
      files: [],
      mappings: []
    };
  },
  methods: {
    onEditConfig() {
      console.log(this.mappings);
      this.isEditMapping = false;
    },
    downloadAll() {
      if (this.files.length > 0) {
        this.files.forEach(file => this.download(file.key));
      }
    },
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
      this.files = [];
      const file = this.$refs.file.files[0];
      const reader = new FileReader();
      reader.onload = async () => {
        const fromWorkBook = new Excel.Workbook();
        await fromWorkBook.xlsx.load(reader.result);
        this.message = validateFrom(fromWorkBook, this.mappings);
        if (!this.message) {
          this.mappings.forEach(async mapping => {
            const toWorkbook = await loadTemplate(
              _.get(mapping, "templateName")
            );
            const resultWrokbook = fillData(fromWorkBook, toWorkbook, mapping);
            this.files.push({
              key: _.get(mapping, "fileName"),
              name: _.get(mapping, "fileName"),
              workbook: resultWrokbook,
              buffer: resultWrokbook.xlsx.writeBuffer()
            });
            if (this.files.length === this.mappings.length)
              this.isLoading = false;
          });
          console.log("parsing success", this.files);
        } else {
          this.isLoading = false;
        }
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
@import "../css/color.less";
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
  margin-top: 18px;
  background-color: #02c487;
  padding: 8px 16px;
  font-size: 16px;
  color: white;
  font-weight: 600;
  text-decoration: none;
}
.results {
  margin-top: 24px;
  border: 1px solid gray;
  padding: 8px;
  border-radius: 2px;
  li {
    padding: 8px;
    background-color: #ececec;
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
      color: @primary-color;
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
.configMapping {
  margin-left: 8px;
  font-size: 14px;
  padding: 4px;
  border-radius: 2px;
  text-decoration: underline;
  cursor: pointer;
}
.loader {
  margin-top: 24px;
}
</style>
