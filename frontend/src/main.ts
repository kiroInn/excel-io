import Vue from "vue";
import App from "./App.vue";
import "./registerServiceWorker";
import router from "./router";
import VueResource from "vue-resource";

Vue.use(VueResource);
Vue.config.productionTip = false;

let baseUrl = ''
if (process.env.NODE_ENV === 'production') {
   baseUrl = '/api/'
}
else {
   baseUrl = 'http://localhost:3000/api/'
}

Vue.http.options.root = baseUrl;
new Vue({
  router,
  render: h => h(App),
}).$mount("#app");
