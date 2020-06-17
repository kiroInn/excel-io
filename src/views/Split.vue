<template>
  <div class="container">
    <div class="description"><b>"split"</b> excel into multiple excel 🧬</div>
    <h1 v-if="!!message" class="message">{{ message }}</h1>
    <div v-if="files.length <= 0" class="select-file-container">
      <input
        type="file"
        id="file"
        ref="file"
        v-on:change="handleFileUpload()"
        class="inputfile"
      />
      <label class="select-file" for="file">Choose a file</label>
      <label class="configBtn" v-on:click="isEditMapping = true">Config</label>
    </div>
    <loader v-show="isLoading" class="loader"></loader>
    <modal
      v-if="isEditMapping"
      v-on:submit="onEditConfig"
      v-on:close="isEditMapping = false"
    >
      <template v-slot:header>
        <h1>Config Mapping</h1>
        <input type="checkbox" id="isJSON" v-model="isJSON" />
        <label for="isJSON" class="isJSON configBtn">switch JSON</label>
      </template>
      <template v-slot:body>
        <div class="configMapping">
          <div v-if="isJSON">
            <textarea
              name=""
              id=""
              cols="30"
              rows="10"
              v-model="mappingsJSON"
            ></textarea>
          </div>
          <table v-if="!isJSON">
            <thead>
              <tr>
                <th rowspan="2">source</th>
                <th colspan="2">to</th>
                <th rowspan="2">type</th>
                <th rowspan="2">operate</th>
              </tr>
              <tr>
                <!-- <th>fromFile</th> -->
                <td>destination</td>
                <td>fileName</td>
              </tr>
            </thead>
            <tbody>
              <tr v-for="(mapping, index) in mappings" :key="index">
                <!-- <td>
                   <select v-model="mapping.fromFile">
                     <option value="">选择文件</option>
                  </select>
                </td> -->
                <td>
                  <input type="text" v-model="mapping.from" />
                </td>
                <td :class="{ imageType: mapping.type === 'image' }">
                  <div class="to">
                    <input type="text" v-model="mapping.to" />
                  </div>
                  <div v-if="mapping.type === 'image'">
                    <div>
                      <label>tl:col</label
                      ><input type="text" v-model="mapping.range.tl.col" />
                    </div>
                    <div>
                      <label>tl:row</label
                      ><input type="text" v-model="mapping.range.tl.row" />
                    </div>
                    <div>
                      <label>br:col</label
                      ><input type="text" v-model="mapping.range.br.col" />
                    </div>
                    <div>
                      <label>br:row</label
                      ><input type="text" v-model="mapping.range.br.row" />
                    </div>
                  </div>
                </td>
                <td>
                  <input type="text" v-model="mapping.toFile" />
                </td>
                <td>
                  <select v-model="mapping.type">
                    <option
                      :selected="mapping.type === 'sheet-capture'"
                      value="sheet-capture"
                      >Sheet-Capture</option
                    >
                    <option :selected="mapping.type === 'sheet'" value="sheet"
                      >Sheet</option
                    >
                    <option :selected="mapping.type === 'string'" value="string"
                      >String</option
                    >
                    <option :selected="mapping.type === 'date'" value="date"
                      >Date</option
                    >
                    <option :selected="mapping.type === 'image'" value="image"
                      >Image</option
                    >
                  </select>
                </td>
                <td>
                  <a
                    v-on:click="removeMapping(index)"
                    href="javascript:void(0);"
                    >delete</a
                  >
                </td>
              </tr>
              <tr>
                <td colspan="5">
                  <a
                    class="addMapping"
                    v-on:click="addMapping"
                    href="javascript:void(0);"
                    >Add</a
                  >
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </template>
    </modal>
    <div v-if="!!files.length" class="result">
      <div class="downloadOperate">
        <a
          class="downloadAll"
          v-on:click="downloadAll"
          href="javascript:void(0)"
          >Download all</a
        >
        <input type="checkbox" id="prefixing" v-model="isPrefixing" />
        <label for="prefixing" class="prefixing">enable prefixing</label>
      </div>
      <div v-if="isPrefixing" class="filePrefix">
        <input
          v-model="prefixFileName"
          type="text"
          placeholder="file name prefix"
        />
      </div>
      <ul v-if="!!files.length" class="resultList">
        <li v-for="file in files" :key="file.key">
          <span>{{ prefixFileName }}{{ file.name }}</span>
          <span class="status-success">success</span>
          <a
            href="javascript:void(0);"
            class="download"
            v-on:click="download(file.key)"
            >download</a
          >
        </li>
      </ul>
      <a class="button" v-on:click="restart">ReStart</a>
    </div>
  </div>
</template>

<script>
import Loader from "@/components/loader";
import Modal from "@/components/modal";
import * as Excel from "exceljs";
import * as JSZip from "jszip";
import { saveAs } from "file-saver";
import { fillData } from "@/service/transform";
import { isXlsx } from "@/util/file";
import { DEFAULT_MAPPING } from "@/const/split";
import _ from "lodash";
import { transformMappings, reverseTransformMappings } from "@/service/mapping";

export default {
  name: "Split",
  components: {
    Loader,
    Modal
  },
  computed: {
    mappingsJSON: {
      get: function() {
        return JSON.stringify(this.mappings);
      },
      set: function(newValue) {
        try {
          this.mappings = JSON.parse(newValue);
        } catch (e) {
          console.error(e);
        }
      }
    }
  },
  created() {
    this.mappings = transformMappings(DEFAULT_MAPPING);
  },
  data() {
    return {
      prefixFileName: "",
      message: "",
      isLoading: false,
      isEditMapping: false,
      isPrefixing: false,
      isJSON: false,
      files: [],
      mappings: []
    };
  },
  methods: {
    restart() {
      this.files = [];
    },
    addMapping() {
      this.mappings.push({
        type: "string",
        range: {
          tl: { col: 0, row: 0 },
          br: { col: 1, row: 1 }
        }
      });
    },
    removeMapping(deleteIndex) {
      this.mappings = _.remove(
        this.mappings,
        (n, index) => deleteIndex !== index
      );
    },
    onEditConfig() {
      this.isEditMapping = false;
    },
    downloadAll() {
      const filesLength = this.files.length;
      if (filesLength > 0) {
        this.isLoading = true;
        const zip = new JSZip();
        new Promise(resovle => {
          this.files.forEach(async ({ key }, index) => {
            const { fileName, blob } = await this.getFileEntity(key);
            zip.file(fileName, blob);
            if (index === filesLength - 1) resovle();
          });
        })
          .then(() => {
            return zip.generateAsync({ type: "blob" });
          })
          .then(content => {
            const date = new Date();
            saveAs(
              content,
              `excel-io-${date.getFullYear()}${date.getMonth() +
                1}${date.getDate()}${date.getHours()}${date.getMinutes()}`
            );
            this.isLoading = false;
          })
          .catch(err => {
            console.error(err);
            this.message = "download failed";
            this.isLoading = false;
          });
      } else {
        this.message = "no files";
      }
    },
    async getFileEntity(key) {
      const file = this.files.filter(f => f.key === key)[0];
      const workbook = file.workbook;
      const buffer = await workbook.xlsx.writeBuffer();
      const fileType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
      const blob = new Blob([buffer], { type: fileType });
      const fileName = `${this.prefixFileName}${file.name}`;
      return { fileName, blob };
    },
    async download(key) {
      const { fileName, blob } = await this.getFileEntity(key);
      saveAs(blob, fileName);
    },
    handleFileUpload() {
      this.message = "";
      this.isLoading = true;
      this.files = [];
      const file = this.$refs.file.files[0];
      const reader = new FileReader();
      const mappings = reverseTransformMappings(this.mappings);
      reader.onload = async () => {
        const fromWorkBook = new Excel.Workbook();
        await fromWorkBook.xlsx.load(reader.result);
        mappings.forEach(mapping => {
          const toWorkbook = new Excel.Workbook();
          const resultWrokbook = fillData(fromWorkBook, toWorkbook, mapping);
          if (_.get(resultWrokbook, "worksheets.length", 0) > 0) {
            this.files.push({
              key: _.get(mapping, "templateName"),
              name: _.get(mapping, "templateName"),
              workbook: resultWrokbook,
              buffer: resultWrokbook.xlsx.writeBuffer()
            });
          }
        });
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
@import "../css/color.less";
@import "../css/table.less";
@import "../css/common.less";
.container {
  padding-top: 40px;
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: center;
  background: no-repeat center/750px
    url("../assets/data_arranging_flatline.png");
}
.slogan {
  font-size: 44px;
  margin-top: 24px;
  strong {
    font-weight: 600;
  }
}
.downloadOperate {
  display: flex;
  align-items: center;
  margin-top: 20px;
}
.filePrefix {
  input {
    box-sizing: border-box;
    margin: 0;
    padding: 0;
    font-variant: tabular-nums;
    list-style: none;
    font-feature-settings: "tnum";
    position: relative;
    display: inline-block;
    min-width: 0;
    padding: 4px 11px;
    color: rgba(0, 0, 0, 0.65);
    font-size: 14px;
    line-height: 1.5715;
    background-color: #fff;
    background-image: none;
    border: 1px solid #d9d9d9;
    border-radius: 2px;
    transition: all 0.3s;
  }
  margin-top: 10px;
}
.downloadAll {
  border-radius: 3px;
  background-color: #02c487;
  padding: 8px 16px;
  font-size: 16px;
  color: white;
  font-weight: 600;
  text-decoration: none;
}
.result {
  padding: 50px;
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: center;
  background-color: rgba(255, 255, 255, 0.7);
}
.resultList {
  margin: 16px;
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
      text-decoration: underline;
      cursor: pointer;
    }
  }
}
.configMapping {
  height: calc(100vh - 500px);
  overflow-y: scroll;
}
#prefixing {
  margin-left: 6px;
}
</style>
