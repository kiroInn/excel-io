<template>
  <div class="result-container">
    <div class="result-count">
      <div>
        imported count: <b>{{ importCount - 1 }}</b>
      </div>
      <div>
        success count:<b>{{ successCount }}</b>
      </div>
    </div>
    <div class="config">
      <input type="checkbox" id="isShowAll" v-model="isShowAll" />
      <label for="isShowAll" class="isShowAll configBtn">enable show all</label>
      <input type="checkbox" id="isShowDollar" v-model="isShowDollar" />
      <label for="isShowDollar" class="configBtn">enable show dollar</label>
      status margin: <input type="number" v-model="margin" />
    </div>
    <table>
      <thead>
        <tr>
          <th>Code</th>
          <th>Current</th>
          <th>Totals</th>
          <th>Balance Per</th>
          <th>Variance</th>
          <th v-if="isShowDollar">￥</th>
          <th v-if="isShowDollar">Actual</th>
          <th>Current-Balance Per</th>
          <th>Totals-Balance Per</th>
          <th>Status</th>
        </tr>
      </thead>
      <tr v-for="(values, index) in dataArray" :key="index">
        <td class="codeName">{{ values[1] }}</td>
        <td>{{ financial(values[0]) }}</td>
        <td>{{ financial(values[2]) }}</td>
        <td>{{ financial(values[3]) }}</td>
        <td>{{ financial(values[4]) }}</td>
        <td v-if="isShowDollar">{{ financial(values[5]) }}</td>
        <td v-if="isShowDollar">{{ financial(values[6]) }}</td>
        <td>{{ financial(financial(values[0]) - financial(values[3])) }}</td>
        <td>{{ financial(financial(values[2]) - financial(values[3])) }}</td>
        <td
          class="status"
          :class="{
            success:
              isValidate(values[0], values[3]) &&
              isValidate(values[2], values[3])
          }"
        >
          {{
            isValidate(values[0], values[3]) && isValidate(values[2], values[3])
              ? "OK"
              : "ERR"
          }}
        </td>
      </tr>
    </table>
  </div>
</template>
<script>
import { EXPRESS_VALUE_NOT_FOUND } from "@/service/evaluation";
import _ from "lodash";

export default {
  name: "EvaluationResult",
  props: ["dataSource", "importCount"],
  computed: {
    dataArray() {
      if (this.isShowAll) {
        return this.dataSource;
      } else {
        return _.filter(
          this.dataSource,
          data =>
            _.filter(data, value => value !== EXPRESS_VALUE_NOT_FOUND).length >
            2
        );
      }
    },
    successCount() {
      return _.get(
        _.union(_.map(this.dataArray, items => _.get(items, 1))),
        "length"
      );
    }
  },
  data() {
    return {
      isShowAll: false,
      isShowDollar: false,
      margin: 0.5
    };
  },
  methods: {
    financial(x) {
      return _.isNumber(x) ? Number.parseFloat(x).toFixed(2) : x;
    },
    isValidate(value1, value2) {
      return (
        Math.abs(this.financial(value1) - this.financial(value2)) <= this.margin
      );
    }
  }
};
</script>
<style lang="less" scoped>
@import "../css/table.less";
@import "../css/common.less";
.result-container {
  padding: 50px 50px;
  background-color: rgba(255, 255, 255, 0.8);
  .result-count {
    display: flex;
    justify-content: center;
    div {
      margin: 6px;
      b {
        font-weight: 700;
      }
    }
  }
  .config {
    display: flex;
    justify-content: center;
    align-items: center;
    padding: 8px;
  }
  td.status {
    background-color: #f50000;
    color: white;
  }
  td.success {
    background-color: #4caf50;
    color: white;
  }
  thead th {
    position: sticky;
    top: 0;
    background-color: white;
  }
  td.codeName {
    max-width: 240px;
    display: inline-flex;
    text-overflow: ellipsis;
    white-space: nowrap;
    overflow: hidden;
  }
}
</style>
