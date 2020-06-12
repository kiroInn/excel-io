<template>
  <div class="container">
    <div class="description">
      <b>"evaluation"</b> different excel cell address value 🔎
    </div>
    <h1 v-if="!!message" class="message">{{ message }}</h1>
    <div v-if="dataSource.length <= 0" class="select-file-container">
      <input
        type="file"
        id="file"
        ref="file"
        v-on:change="handleFileUpload()"
        class="inputfile"
        multiple
      />
      <label class="select-file" for="file">Choose multiple files</label>
      <label class="configBtn" v-on:click="showModal = true">Config</label>
      <modal
        v-if="showModal"
        v-on:submit="onEditConfig"
        v-on:close="showModal = false"
      >
        <template v-slot:header>
          <h1>Config Variable</h1>
        </template>
        <template v-slot:body>
          <div class="mappings">
            <table>
              <tr>
                <th>Current</th>
                <th>Totals</th>
                <th>Balance per the GL</th>
                <th>Variance</th>
                <th>￥</th>
                <th>Actual</th>
              </tr>
              <tr v-for="(positions, index) in expressMatrix" :key="index">
                <td v-for="(position, index) in positions" :key="index">
                  <input type="text" :value="position" />
                </td>
              </tr>
            </table>
          </div>
          <!-- <input type="text" v-model="columName" /><button
            v-on:click="addVariable"
          >
            Add variable
          </button> -->
        </template>
      </modal>
    </div>
    <evaluation-result
      v-if="dataSource.length > 0"
      :data-source="dataSource"
    ></evaluation-result>
    <a class="button" v-if="dataSource.length > 0" v-on:click="restart"
      >ReStart</a
    >
    <loader v-show="isLoading" class="loader"></loader>
  </div>
</template>
<script>
import _ from "lodash";
import { isXlsx } from "@/util/file";
import { parse, DEFAULT_EXPRESS } from "@/service/evaluation";
import Modal from "@/components/modal";
import Loader from "@/components/loader";
import EvaluationResult from "@/views/EvaluationResult";
import * as Excel from "exceljs";

export default {
  name: "Evaluation",
  components: {
    Modal,
    Loader,
    EvaluationResult
  },
  data() {
    return {
      message: "",
      isLoading: false,
      showModal: false,
      colums: [],
      columName: "",
      expressMatrix: [],
      dataSource: []
    };
  },
  created() {
    this.expressMatrix = DEFAULT_EXPRESS;
  },
  methods: {
    restart() {
      this.dataSource = [];
    },
    handleFileUpload() {
      this.message = "";
      this.dataSource = [];
      this.isLoading = true;
      const files = this.$refs.file.files;
      new Promise((resolve, reject) => {
        const result = [];
        _.each(files, file => {
          const reader = new FileReader();
          reader.onload = async () => {
            const workbook = new Excel.Workbook();
            await workbook.xlsx.load(reader.result);
            result.push({
              fileName: file.name,
              workbook
            });
            if (result.length === files.length) resolve(result);
          };
          if (isXlsx(file.name)) {
            reader.readAsArrayBuffer(file);
          } else {
            this.message = `${file.name} file type is not supported`;
            reject();
          }
        });
      })
        .then(files => {
          this.dataSource = parse(files, this.expressMatrix);
          this.isLoading = false;
        })
        .catch(err => {
          this.isLoading = false;
          console.log("err: ", err);
        });
    },
    onEditConfig() {
      this.isEditMapping = false;
    },
    addVariable() {
      if (this.columName) {
        this.colums.push(this.columName);
        if (!this.expressMatrix)
          this.expressMatrix.push(new Array(this.colums.length));
      }
    }
  }
};
</script>
<style lang="less" scoped>
@import "../css/common.less";
@import "../css/table.less";
.container {
  padding: 40px 0px;
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: center;
  background: no-repeat center/ 600px
    url("../assets/report_analysis__flatline.png");
}
.mappings {
  height: calc(100vh - 500px);
  overflow-y: scroll;
}
</style>
